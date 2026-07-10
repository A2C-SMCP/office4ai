# filename: server.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod
from collections.abc import Iterable
from typing import Any
from urllib.parse import urlparse

from loguru import logger
from mcp.server import NotificationOptions, Server
from mcp.server.lowlevel.helper_types import ReadResourceContents
from mcp.server.sse import SseServerTransport
from mcp.server.stdio import stdio_server
from mcp.server.streamable_http_manager import StreamableHTTPSessionManager
from mcp.types import Resource, ResourcesCapability, ServerCapabilities, SubscribeRequest, Tool, UnsubscribeRequest
from pydantic import AnyUrl
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import Response
from starlette.routing import Mount, Route

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.resources.base import BaseResource
from office4ai.a2c_smcp.subscriptions import SubscriptionManager
from office4ai.a2c_smcp.tools.base import BaseTool


class BaseMCPServer(ABC):
    def __init__(self, config: MCPServerConfig, server_name: str) -> None:
        self.config = config
        self.server = Server(server_name)

        self.tools: dict[str, BaseTool] = {}
        self.resources: dict[str, BaseResource] = {}
        self.subscription_manager = SubscriptionManager()

        self._register_tools()
        self._register_resources()
        self._setup_handlers()
        self._patch_subscribe_capability()

        logger.info(
            f"MCP Server 初始化完成 | MCP Server initialized: server={server_name}, transport={config.transport}",
        )

    @abstractmethod
    def _register_tools(self) -> None:
        raise NotImplementedError

    @abstractmethod
    def _register_resources(self) -> None:
        raise NotImplementedError

    def _patch_subscribe_capability(self) -> None:
        """Patch MCP Server capabilities to declare resources.subscribe=True and listChanged.

        MCP SDK hardcodes subscribe=False in get_capabilities() even when
        subscribe/unsubscribe handlers are registered. We override
        create_initialization_options to fix this.

        W4a/W4b-1 (#63/#64): also declare ``tools.listChanged`` + ``resources.listChanged``
        by defaulting ``NotificationOptions`` when the caller passes none, so the server can
        broadcast ``tools/list_changed`` / ``resources/list_changed`` on connection changes.
        """
        original_create = self.server.create_initialization_options

        def patched_create_initialization_options(
            notification_options: NotificationOptions | None = None,
            experimental_capabilities: dict[str, Any] | None = None,
        ) -> Any:
            if notification_options is None:
                notification_options = NotificationOptions(tools_changed=True, resources_changed=True)
            options = original_create(notification_options, experimental_capabilities)
            caps = options.capabilities
            if caps.resources and (
                SubscribeRequest in self.server.request_handlers or UnsubscribeRequest in self.server.request_handlers
            ):
                options.capabilities = ServerCapabilities(
                    prompts=caps.prompts,
                    resources=ResourcesCapability(
                        subscribe=True,
                        listChanged=caps.resources.listChanged,
                    ),
                    tools=caps.tools,
                    logging=caps.logging,
                    experimental=caps.experimental,
                    completions=caps.completions,
                )
                logger.info("resources.subscribe capability enabled")
            return options

        self.server.create_initialization_options = patched_create_initialization_options  # type: ignore[method-assign]

    def _affected_resource_uris(self, tool: BaseTool, arguments: dict[str, Any]) -> list[str]:
        """Resource URIs a successful tool call should notify (``resource_updated``).

        Base default: none. Subclasses that project windows (e.g. ``OfficeMCPServer`` with
        per-file windows, W4b-1) override to target the specific document's window.
        """
        return []

    # ── 动态工具收敛（W4a / #63）| Dynamic tool convergence ──

    def _is_tool_available(self, tool: BaseTool) -> bool:
        """Whether *tool* should be exposed given the current state.

        Base: always available (no filtering). Subclasses override to converge the tool
        set by Add-In connection state (W4a).
        """
        return True

    def _visible_tools(self) -> list[BaseTool]:
        """Currently exposed tools (after connection-based convergence filtering)."""
        return [tool for tool in self.tools.values() if self._is_tool_available(tool)]

    def _current_session(self) -> Any:
        """The MCP ServerSession handling the in-flight request, or ``None`` outside one."""
        try:
            from mcp.server.lowlevel.server import request_ctx

            return request_ctx.get().session
        except Exception:
            return None

    def _resolve_resource(self, uri_str: str) -> BaseResource | None:
        """Resolve a read URI to its owning resource.

        Exact ``base_uri`` match handles window resources and skill:// roots. A
        skill:// **sub-path** (``skill://<host>/<leaf>/<rel>``) has no exact key, so
        fall back to a prefix match against registered skill roots — the owning
        ``SkillResource`` then serves the specific relative path (skill.md §3
        ``resources`` mode). The trailing ``/`` guards against sibling-prefix
        collisions (``.../leaf`` must not capture ``.../leaf-extra/...``).
        """
        parsed = urlparse(uri_str)
        base_uri = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
        resource = self.resources.get(base_uri)
        if resource is not None:
            return resource
        if parsed.scheme == "skill":
            for candidate in self.resources.values():
                if uri_str.startswith(candidate.base_uri + "/"):
                    return candidate
        return None

    async def _async_startup(self) -> None:  # noqa: B027
        """async 启动钩子, 子类可 override | Async startup hook, subclass can override"""
        pass

    async def _async_shutdown(self) -> None:  # noqa: B027
        """async 关闭钩子, 子类可 override | Async shutdown hook, subclass can override"""
        pass

    def _setup_handlers(self) -> None:
        @self.server.list_tools()  # type: ignore[no-untyped-call]
        async def list_tools() -> list[Tool]:
            session = self._current_session()
            if session is not None:
                self.subscription_manager.track_session(session)
            # W4a: converge the exposed tool set by Add-In connection state.
            return [
                Tool(
                    name=tool.name,
                    description=tool.description,
                    inputSchema=tool.input_schema,
                )
                for tool in self._visible_tools()
            ]

        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[dict[str, Any]]:
            tool = self.tools.get(name)
            if not tool:
                return [{"type": "text", "text": f"未找到工具 | Tool not found: {name}"}]

            try:
                result = await tool.execute(arguments)
                # Notify subscribers of affected resources (per-file window under W4b-1)
                affected = self._affected_resource_uris(tool, arguments)
                if affected:
                    await self.subscription_manager.notify_many(affected)
                # office4ai #42: 委托工具决定 MCP 内容类型 (截图等发 image, 避免 base64 内联 text 撑爆上下文)
                return tool.to_mcp_content(result)
            except Exception as e:
                logger.exception(f"工具执行失败 | Tool execution failed: {e}")
                return [{"type": "text", "text": f"工具执行失败 | Tool execution failed: {e}"}]

        @self.server.list_resources()  # type: ignore[no-untyped-call]
        async def list_resources() -> list[Resource]:
            session = self._current_session()
            if session is not None:
                self.subscription_manager.track_session(session)
            # Each resource contributes 1+ entries: window resources yield a single
            # Resource; a skill:// root in `resources` mode expands into a root
            # (carrying _meta.source) plus one sub-resource per packaged file.
            entries: list[Resource] = []
            for resource in self.resources.values():
                entries.extend(resource.list_entries())
            return entries

        @self.server.subscribe_resource()  # type: ignore[no-untyped-call]
        async def subscribe_resource(uri: AnyUrl) -> None:
            from mcp.server.lowlevel.server import request_ctx

            session = request_ctx.get().session
            uri_str = str(uri)
            self.subscription_manager.subscribe(uri_str, session)
            logger.debug(f"Client subscribed to resource: {uri}")

        @self.server.unsubscribe_resource()  # type: ignore[no-untyped-call]
        async def unsubscribe_resource(uri: AnyUrl) -> None:
            from mcp.server.lowlevel.server import request_ctx

            session = request_ctx.get().session
            uri_str = str(uri)
            self.subscription_manager.unsubscribe(uri_str, session)
            logger.debug(f"Client unsubscribed from resource: {uri}")

        @self.server.read_resource()  # type: ignore[no-untyped-call]
        async def read_resource(uri: Any) -> Iterable[ReadResourceContents]:
            # Convert AnyUrl to string if needed
            uri_str = str(uri)

            resource = self._resolve_resource(uri_str)
            if not resource:
                raise ValueError(f"未找到资源 | Resource not found: {uri_str}")

            # Window resources parse query params (priority/fullscreen) from the exact
            # URI; skill:// sub-paths have no params (update_from_uri is a no-op there).
            if uri_str != resource.uri:
                resource.update_from_uri(uri_str)

            return await resource.read_content(uri_str)

    async def run(self) -> None:
        await self._async_startup()
        try:
            transport = self.config.transport

            if transport == "stdio":
                await self._run_stdio()
                return

            if transport == "sse":
                await self._run_sse()
                return

            if transport == "streamable-http":
                await self._run_streamable_http()
                return

            raise ValueError(f"不支持的传输模式 | Unsupported transport mode: {transport}")
        finally:
            await self._async_shutdown()

    async def _run_stdio(self) -> None:
        async with stdio_server() as (read_stream, write_stream):
            await self.server.run(
                read_stream,
                write_stream,
                self.server.create_initialization_options(),
            )

    async def _run_sse(self) -> None:
        import uvicorn

        sse = SseServerTransport("/messages/")

        async def handle_sse(request: Request) -> Response:
            async with sse.connect_sse(request.scope, request.receive, request._send) as streams:
                await self.server.run(
                    streams[0],
                    streams[1],
                    self.server.create_initialization_options(),
                )
            return Response()

        routes = [
            Route("/sse", endpoint=handle_sse, methods=["GET"]),
            Mount("/messages/", app=sse.handle_post_message),
        ]

        app = Starlette(routes=routes)
        config = uvicorn.Config(app, host=self.config.host, port=self.config.port, log_level="info")
        server = uvicorn.Server(config)
        await server.serve()

    async def _run_streamable_http(self) -> None:
        import uvicorn

        session_manager = StreamableHTTPSessionManager(self.server)

        async def handle_streamable_http(request: Request) -> Response:
            # StreamableHTTPSessionManager.handle_request handles the request asynchronously
            await session_manager.handle_request(request.scope, request.receive, request._send)
            # Return a placeholder response (the actual response is sent by the session manager)
            return Response(content=b"", status_code=200, headers={})

        routes = [
            Route("/mcp", endpoint=handle_streamable_http, methods=["POST"]),
        ]

        app = Starlette(routes=routes)
        config = uvicorn.Config(app, host=self.config.host, port=self.config.port, log_level="info")
        server = uvicorn.Server(config)
        await server.serve()
