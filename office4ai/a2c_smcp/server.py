# filename: server.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any

from loguru import logger
from mcp.server import NotificationOptions, Server
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
from office4ai.a2c_smcp.tools.base import BaseTool, inline_schema_refs


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
        """Patch MCP Server capabilities to declare resources.subscribe=True.

        MCP SDK hardcodes subscribe=False in get_capabilities() even when
        subscribe/unsubscribe handlers are registered. We override
        create_initialization_options to fix this.
        """
        original_create = self.server.create_initialization_options

        def patched_create_initialization_options(
            notification_options: NotificationOptions | None = None,
            experimental_capabilities: dict[str, Any] | None = None,
        ) -> Any:
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

    # Category → resource URI mapping for subscription notifications.
    # Only notifies the platform-specific resource, not the root resource.
    _CATEGORY_URI_MAP: dict[str, list[str]] = {
        "word": ["window://office4ai/word"],
        "ppt": ["window://office4ai/ppt"],
        "excel": ["window://office4ai/excel"],
    }

    def _category_to_resource_uris(self, category: str) -> list[str]:
        """Map a tool category to the resource URIs it affects."""
        return self._CATEGORY_URI_MAP.get(category, [])

    async def _async_startup(self) -> None:  # noqa: B027
        """async 启动钩子, 子类可 override | Async startup hook, subclass can override"""
        pass

    async def _async_shutdown(self) -> None:  # noqa: B027
        """async 关闭钩子, 子类可 override | Async shutdown hook, subclass can override"""
        pass

    def _setup_handlers(self) -> None:
        @self.server.list_tools()  # type: ignore[no-untyped-call]
        async def list_tools() -> list[Tool]:
            return [
                Tool(
                    name=tool.name,
                    description=tool.description,
                    # office4ai #37: 内联全部 $ref (去 $defs), 否则模型把必填嵌套对象误填成 JSON 字符串
                    inputSchema=inline_schema_refs(tool.input_schema),
                )
                for tool in self.tools.values()
            ]

        @self.server.call_tool()
        async def call_tool(name: str, arguments: dict[str, Any]) -> list[dict[str, Any]]:
            tool = self.tools.get(name)
            if not tool:
                return [{"type": "text", "text": f"未找到工具 | Tool not found: {name}"}]

            try:
                result = await tool.execute(arguments)
                # Notify subscribers of affected resources
                affected = self._category_to_resource_uris(tool.category)
                if affected:
                    await self.subscription_manager.notify_many(affected)
                return [{"type": "text", "text": str(result)}]
            except Exception as e:
                logger.exception(f"工具执行失败 | Tool execution failed: {e}")
                return [{"type": "text", "text": f"工具执行失败 | Tool execution failed: {e}"}]

        @self.server.list_resources()  # type: ignore[no-untyped-call]
        async def list_resources() -> list[Resource]:
            return [
                Resource(
                    uri=AnyUrl(resource.uri),
                    name=resource.name,
                    description=resource.description,
                    mimeType=resource.mime_type,
                )
                for resource in self.resources.values()
            ]

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
        async def read_resource(uri: Any) -> str:
            from urllib.parse import urlparse

            # Convert AnyUrl to string if needed
            uri_str = str(uri)

            parsed = urlparse(uri_str)
            base_uri = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
            resource = self.resources.get(base_uri)
            if not resource:
                raise ValueError(f"未找到资源 | Resource not found: {base_uri}")

            if uri_str != resource.uri:
                resource.update_from_uri(uri_str)

            return await resource.read()

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
