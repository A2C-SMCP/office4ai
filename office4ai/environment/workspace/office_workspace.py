"""
Office Workspace Implementation

实现 Office 特定的 Workspace 环境，集成 Socket.IO 服务器。
"""

import asyncio
import logging
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import socketio  # type: ignore[import-untyped]
from aiohttp import web

from .base import BaseWorkspace, DocumentStatus, OfficeAction, OfficeObs
from .socketio.config import SocketIOConfig, default_config
from .socketio.factory import build_sio_server
from .socketio.services.connection_manager import connection_manager, normalize_document_uri

logger = logging.getLogger(__name__)

# Tool names whose results populate the caches
_VISIBLE_CONTENT_TOOLS = {"word_get_visible_content"}
_STRUCTURE_TOOLS = {"word_get_document_structure"}


@dataclass
class LastActivity:
    """Records the most recent tool activity on a document."""

    document_uri: str
    tool_name: str
    timestamp: float = field(default_factory=time.time)


class OfficeWorkspace(BaseWorkspace):
    """
    Office Workspace 实现

    负责：
    1. 启动/停止 Socket.IO 服务器
    2. 管理 Office Add-In 连接
    3. 执行 Office 动作 (通过 Socket.IO)
    4. 维护文档状态
    """

    def __init__(
        self,
        host: str = "127.0.0.1",
        port: int = 3000,
        config: SocketIOConfig = default_config,
        use_https: bool = True,
        cert_dir: Path | None = None,
    ):
        """
        初始化 Office Workspace

        Args:
            host: 绑定地址 (默认 localhost)
            port: 绑定端口 (默认 3000)
            config: Socket.IO 配置
            use_https: 是否启用 HTTPS (默认 True)
            cert_dir: 证书目录 (默认使用 certs.get_cert_dir())
        """
        self.host = host
        self.port = port
        self.config = config
        self.use_https = use_https
        self.cert_dir = cert_dir

        # Socket.IO 服务器和应用
        self.sio_server: socketio.AsyncServer | None = None
        self.app: web.Application | None = None
        self.runner: web.AppRunner | None = None
        self._site: web.TCPSite | None = None
        self._https_site: web.TCPSite | None = None

        # 运行状态
        self._running = False

        # 活动追踪 | Activity tracking
        self._last_activity: LastActivity | None = None
        self._content_cache: dict[str, str] = {}  # document_uri → visible content
        self._structure_cache: dict[str, str] = {}  # document_uri → document structure

        # Server-side OOXML mutations (e.g. chart engine) need to tell MCP
        # subscribers that the underlying file changed. The owning MCP server
        # wires this callback at startup; until then it stays None.
        self._resource_update_callback: Callable[[list[str]], None] | None = None

    # ── 活动追踪 API | Activity tracking API ──

    def update_last_activity(self, document_uri: str, tool_name: str, result_data: dict[str, Any]) -> None:
        """
        Record the latest tool activity and optionally cache content/structure.

        Called by BaseTool after a successful workspace.execute().
        """
        document_uri = normalize_document_uri(document_uri)
        self._last_activity = LastActivity(document_uri=document_uri, tool_name=tool_name)

        # Extract text content from result_data for caching
        content_text = result_data.get("content") if isinstance(result_data, dict) else None

        if tool_name in _VISIBLE_CONTENT_TOOLS and isinstance(content_text, str):
            self._content_cache[document_uri] = content_text

        if tool_name in _STRUCTURE_TOOLS and isinstance(content_text, str):
            self._structure_cache[document_uri] = content_text

    def get_last_activity(self) -> LastActivity | None:
        """Return the most recent activity, or None."""
        return self._last_activity

    def get_cached_content(self, document_uri: str) -> str | None:
        """Return cached visible content for a document, or None."""
        return self._content_cache.get(normalize_document_uri(document_uri))

    def get_cached_structure(self, document_uri: str) -> str | None:
        """Return cached document structure for a document, or None."""
        return self._structure_cache.get(normalize_document_uri(document_uri))

    def _clear_document_cache(self, document_uri: str) -> None:
        """Remove all cached data for a disconnected document."""
        self._content_cache.pop(document_uri, None)
        self._structure_cache.pop(document_uri, None)
        if self._last_activity and self._last_activity.document_uri == document_uri:
            self._last_activity = None

    # ── Resource-update notifications | Resource update notification API ──

    def set_resource_update_callback(self, callback: Callable[[list[str]], None] | None) -> None:
        """Wire a callback the workspace fires after Server-side OOXML mutations.

        OASP /ppt chart events bypass Office.js and modify .pptx OOXML directly;
        we use this hook to push MCP ``resource_updated`` notifications so subscribers
        (window resources for /ppt etc.) know to refetch.
        """
        self._resource_update_callback = callback

    def notify_resource_updated(self, uris: list[str]) -> None:
        """Fire the resource-update callback if one is registered."""
        if not uris or self._resource_update_callback is None:
            return
        try:
            self._resource_update_callback(uris)
        except Exception:
            logger.exception("Resource update callback raised")

    async def start(self) -> None:
        """
        启动 Workspace Socket.IO 服务器

        创建并启动 Socket.IO 服务器，开始监听 Add-In 连接
        """
        if self._running:
            logger.warning("Workspace is already running")
            return

        try:
            # 注册断连回调清理缓存 | Register disconnect callback to clear caches
            connection_manager.register_disconnect_callback(self._clear_document_cache)

            # 创建 Socket.IO 服务器 + 注册命名空间（单一事实源：socketio.factory）
            self.sio_server = build_sio_server(self.config)

            # 创建 aiohttp 应用
            self.app = web.Application()
            self.sio_server.attach(self.app)

            # 添加健康检查路由
            async def health_check(request: web.Request) -> web.Response:
                return web.json_response(
                    {
                        "status": "ok",
                        "service": "office4ai-workspace",
                        "connections": connection_manager.get_connection_count(),
                        "documents": connection_manager.get_document_count(),
                    }
                )

            self.app.router.add_get("/health", health_check)

            # 创建并启动 runner
            self.runner = web.AppRunner(self.app)
            await self.runner.setup()

            # 启动 HTTP 站点
            self._site = web.TCPSite(self.runner, self.host, self.port)
            await self._site.start()

            self._running = True

            logger.info("=" * 60)
            logger.info("✅ Office Workspace started")
            logger.info(f"HTTP:  http://{self.host}:{self.port}")

            # 启动 HTTPS 站点（如果启用）
            if self.use_https:
                from office4ai.certs.paths import create_ssl_context

                try:
                    ssl_context = create_ssl_context(self.cert_dir)
                    https_port = self.port + 1443  # 3000 + 1443 = 4443
                    self._https_site = web.TCPSite(self.runner, self.host, https_port, ssl_context=ssl_context)
                    await self._https_site.start()
                    logger.info(f"HTTPS: https://{self.host}:{https_port}")
                except FileNotFoundError as e:
                    logger.warning(f"SSL certificates not found, skipping HTTPS: {e}")

            logger.info("=" * 60)
            logger.info(f"Health check: http://{self.host}:{self.port}/health")
            logger.info(f"Bind address: {self.host} (localhost only)")
            logger.info("=" * 60)

        except Exception as e:
            logger.error(f"❌ Failed to start Workspace: {e}")
            await self.stop()
            raise

    async def stop(self) -> None:
        """
        停止 Workspace Socket.IO 服务器

        关闭服务器，清理所有连接。
        先显式关闭 Socket.IO 以避免 runner cleanup 等待 ping_timeout。
        """
        if not self._running:
            return

        try:
            # 先显式关闭 Socket.IO，断开所有客户端连接
            # 避免 runner.cleanup() 触发的 on_shutdown 等待 ping_timeout (60s)
            if self.sio_server:
                try:
                    await asyncio.wait_for(self.sio_server.shutdown(), timeout=3.0)
                except asyncio.TimeoutError:
                    logger.warning("Socket.IO shutdown timed out (3s), forcing cleanup")

            # 停止 HTTPS 站点
            if self._https_site:
                await self._https_site.stop()
                self._https_site = None

            # 停止 HTTP 站点
            if self._site:
                await self._site.stop()
                self._site = None

            # 清理 runner (Socket.IO 已关闭，不会再阻塞)
            if self.runner:
                await self.runner.cleanup()
                self.runner = None

            self.sio_server = None
            self._running = False
            self.app = None

            logger.info("✅ Office Workspace stopped")

        except Exception as e:
            logger.error(f"❌ Error stopping Workspace: {e}")
            raise

    @property
    def is_running(self) -> bool:
        """Workspace 是否正在运行"""
        return self._running

    async def execute(self, action: OfficeAction) -> OfficeObs:
        """
        执行统一动作接口

        Args:
            action: Office 动作对象

        Returns:
            OfficeObs: 执行结果
        """
        # 提取 document_uri
        document_uri = action.params.get("document_uri")
        if not document_uri:
            return OfficeObs(
                success=False,
                data={},
                error="Missing document_uri in params",
            )

        # 检查文档状态
        status = self.get_document_status(document_uri)
        if status != DocumentStatus.CONNECTED:
            return OfficeObs(
                success=False,
                data={},
                error=f"Document not connected: {document_uri}",
            )

        # 构造事件名称
        event = f"{action.category}:{action.action_name}"

        # 发送 Socket.IO 事件
        try:
            response = await self.emit_to_document(document_uri, event, action.params)

            # 从响应中提取业务数据（response["data"]）
            # 响应格式: {requestId, success, data, timestamp, duration}
            if response.get("success"):
                # 协议层成功，检查业务层结果
                if "data" in response:
                    business_data = response["data"]
                    # 检查业务层的 success 字段（如果存在）
                    business_success = business_data.get("success", True)
                    return OfficeObs(success=business_success, data=business_data)
                else:
                    # 协议层成功但没有业务数据（可能是不需要返回数据的操作）
                    return OfficeObs(success=True, data={})
            else:
                # 失败：返回错误信息
                error_msg = response.get("error", "Unknown error")
                # 将 dict 类型的 error 转换为字符串（前端返回 {code, message} 格式）
                if isinstance(error_msg, dict):
                    error_msg = f"{error_msg.get('code', 'Unknown')}: {error_msg.get('message', 'Unknown error')}"
                logger.error(f"Action failed: {error_msg}")
                return OfficeObs(success=False, data={}, error=error_msg)
        except Exception as e:
            logger.error(f"Error executing action: {e}")
            return OfficeObs(success=False, data={}, error=str(e))

    def get_document_status(self, document_uri: str) -> DocumentStatus:
        """
        获取文档状态

        Args:
            document_uri: 文档 URI

        Returns:
            DocumentStatus: 文档连接状态
        """
        if connection_manager.is_document_active(document_uri):
            return DocumentStatus.CONNECTED
        return DocumentStatus.DISCONNECTED

    async def emit_to_document(self, document_uri: str, event: str, data: dict[str, Any]) -> dict[str, Any]:
        """
        向指定文档发送 Socket.IO 事件

        Args:
            document_uri: 目标文档 URI
            event: 事件名称
            data: 事件数据

        Returns:
            dict: Add-In 返回的响应数据

        Raises:
            ValueError: 如果文档未连接
            TimeoutError: 如果请求超时
        """
        # 查找 socket_id
        socket_id = connection_manager.get_socket_by_document(document_uri)
        if not socket_id:
            raise ValueError(f"No socket found for document: {document_uri}")

        # 获取客户端信息（包含命名空间）
        client_info = connection_manager.get_client_info(socket_id)
        if not client_info:
            raise ValueError(f"Client not found: {socket_id}")

        if not self.sio_server:
            raise RuntimeError("Socket.IO server is not running")

        logger.info(f"Calling {event} on {socket_id} for document {document_uri} (namespace={client_info.namespace})")

        # Auto-wrap business parameters into BaseRequest format
        from office4ai.environment.workspace.socketio.request_wrapper import (
            RequestWrapperError,
            wrap_request,
        )

        try:
            wrapped_data = wrap_request(
                event=event,
                business_params=data,
                document_uri=document_uri,
            )
            logger.debug(
                f"Wrapped request for {event}: "
                f"requestId={wrapped_data.get('requestId')}, "
                f"documentUri={wrapped_data.get('documentUri')}"
            )
        except RequestWrapperError as e:
            logger.error(f"Failed to wrap request for {event}: {e}")
            raise ValueError(f"Request wrapping failed: {e}") from e

        # 使用 Socket.IO 的 .call() 方法（自动处理 callback）
        try:
            response: dict[str, Any] = await self.sio_server.call(
                event,
                wrapped_data,
                to=socket_id,
                namespace=client_info.namespace,
                timeout=self.config.request_timeout // 1000,
            )
            logger.info(f"Received response from {socket_id}")
            return response
        except TimeoutError:
            logger.error(f"Timeout waiting for response from {socket_id}")
            raise
        except Exception as e:
            logger.error(f"Error emitting event: {e}")
            raise

    async def wait_for_addin_connection(self, timeout: float = 30.0) -> bool:
        """
        等待 Add-In 连接

        Args:
            timeout: 超时时间（秒）

        Returns:
            bool: True 如果有 Add-In 连接，False 如果超时
        """
        logger.info(f"Waiting for Add-In connection (timeout: {timeout}s)...")

        start_time = asyncio.get_event_loop().time()
        while (asyncio.get_event_loop().time() - start_time) < timeout:
            if connection_manager.get_connection_count() > 0:
                logger.info("✅ Add-In connected!")
                return True
            await asyncio.sleep(0.5)

        logger.warning("⏱️ Timeout waiting for Add-In connection")
        return False

    def get_connected_documents(self) -> list[str]:
        """
        获取所有已连接的文档 URI

        Returns:
            list[str]: 文档 URI 列表
        """
        clients = connection_manager.get_all_clients()
        return list({client.document_uri for client in clients})
