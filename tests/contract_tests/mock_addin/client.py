"""
Mock Add-In Client

模拟 Office Add-In 客户端行为的测试替身。
用于契约测试，验证服务器与 Add-In 之间的协议。
"""

from __future__ import annotations

import asyncio
import logging
from collections.abc import Awaitable, Callable

from socketio import AsyncClient  # type: ignore[import-untyped]

from office4ai.environment.workspace.socketio.versioning import SERVER_VERSION

logger = logging.getLogger(__name__)


class MockAddInClient:
    """
    模拟 Office Add-In 客户端。

    这个类模拟真实 Add-In 的行为：
    1. 连接到 Workspace Socket.IO 服务器
    2. 完成握手（发送 clientId、documentUri）
    3. 接收服务器发送的事件
    4. 返回符合协议的响应

    Examples:
        ```python
        # 创建客户端
        client = MockAddInClient(
            server_url="http://127.0.0.1:3002",
            namespace="/word",
            client_id="test_client_001",
            document_uri="file:///tmp/test.docx",
        )

        # 连接到服务器
        await client.connect()

        # 注册响应
        client.register_static_response(
            "word:get:selectedContent",
            {
                "success": True,
                "data": {"text": "Hello World", "elements": [], "metadata": {}}
            }
        )

        # 断开连接
        await client.disconnect()
        ```
    """

    def __init__(
        self,
        server_url: str,
        namespace: str,
        client_id: str,
        document_uri: str,
        oasp_version: str | None = None,
    ) -> None:
        """
        初始化 Mock Add-In 客户端。

        Args:
            server_url: Socket.IO 服务器 URL
            namespace: 命名空间 (/word, /ppt, /excel)
            client_id: 客户端 ID（用于握手）
            document_uri: 文档 URI（用于握手）
            oasp_version: 握手声明的 OASP 协议版本（OASP 0.3.0 起强制）。
                默认与 Server 同版本（``SERVER_VERSION``，单一事实源，避免漂移）；
                传入不兼容版本可用于测试 PROTOCOL_VERSION_MISMATCH 拒绝路径。
        """
        self.server_url = server_url
        self.namespace = namespace
        self.client_id = client_id
        self.document_uri = document_uri
        self.oasp_version = oasp_version if oasp_version is not None else str(SERVER_VERSION)

        # Socket.IO 客户端
        self._client = AsyncClient()

        # 响应注册表 {event_name: response_factory}
        self._response_registry: dict[str, Callable[[dict], dict | Awaitable[dict]]] = {}

        # 记录接收到的所有事件 [(event_name, data), ...]
        self._received_events: list[tuple[str, dict]] = []

        # 连接状态
        self._connected = False

    async def connect(self) -> None:
        """
        连接到服务器并完成握手。

        Raises:
            ConnectionError: 如果连接失败
        """
        try:
            # 连接到服务器
            await self._client.connect(
                self.server_url,
                transports=["websocket"],
                namespaces=[self.namespace],
                auth={
                    "clientId": self.client_id,
                    "documentUri": self.document_uri,
                    "oaspVersion": self.oasp_version,
                },
            )

            self._connected = True
            logger.info(f"MockAddInClient connected: {self.client_id} to {self.server_url}{self.namespace}")

            # 等待连接建立
            await asyncio.sleep(0.1)

        except Exception as e:
            logger.error(f"MockAddInClient connection failed: {e}")
            raise ConnectionError(f"Failed to connect to {self.server_url}: {e}") from e

    async def disconnect(self) -> None:
        """断开连接。"""
        if self._connected:
            await self._client.disconnect()
            self._connected = False
            logger.info(f"MockAddInClient disconnected: {self.client_id}")

    def register_response(
        self,
        event: str,
        response_factory: Callable[[dict], dict | Awaitable[dict]],
    ) -> None:
        """
        注册事件响应工厂函数。

        当服务器发送此事件时，会调用工厂函数生成响应。

        注意：必须在 connect() 之前注册所有响应。

        Args:
            event: 事件名称（如 "word:get:selectedContent"）
            response_factory: 响应工厂函数，接收请求数据，返回响应数据

        Examples:
            ```python
            def dynamic_response(request: dict) -> dict:
                return {
                    "requestId": request["requestId"],
                    "success": True,
                    "data": {"text": "Dynamic content"}
                }

            mock_client.register_response("word:get:selectedContent", dynamic_response)
            ```
        """
        self._response_registry[event] = response_factory

        # 为此事件注册处理器
        async def handler(data: dict) -> dict:
            """处理服务器发起的 RPC 调用"""
            # 记录接收到的事件
            self._received_events.append((event, data))

            # 调用响应工厂
            response = response_factory(data)

            # 如果是协程，等待它
            if asyncio.iscoroutine(response):
                response = await response

            logger.debug(f"MockAddInClient responding to {event}: {response}")
            return response

        # 注册到 Socket.IO 客户端
        self._client.on(event, handler, namespace=self.namespace)

        logger.debug(f"Registered response factory for event: {event}")

    def register_static_response(self, event: str, response: dict) -> None:
        """
        注册静态响应。

        Args:
            event: 事件名称
            response: 静态响应数据

        Examples:
            ```python
            mock_client.register_static_response(
                "word:get:selectedContent",
                {"success": True, "data": {"text": "Hello"}}
            )
            ```
        """

        def factory(_request: dict) -> dict:
            return response

        self.register_response(event, factory)
        logger.debug(f"Registered static response for event: {event}")

    async def send_event(self, event: str, data: dict) -> None:
        """
        发送事件到服务器（模拟 Add-In 向服务器发送事件）。

        Args:
            event: 事件名称
            data: 事件数据
        """
        if not self._connected:
            raise ConnectionError("Client not connected")

        await self._client.emit(event, data, namespace=self.namespace)
        logger.debug(f"Sent event: {event}")

    @property
    def received_events(self) -> list[tuple[str, dict]]:
        """
        获取收到的所有事件（用于断言）。

        Returns:
            事件列表，每个元素是 (event_name, data) 元组
        """
        return list(self._received_events)

    def clear_events(self) -> None:
        """清空事件记录。"""
        self._received_events.clear()
