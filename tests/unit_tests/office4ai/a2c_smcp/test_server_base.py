# filename: test_server_base.py
# @Time    : 2025/12/18 19:17
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
BaseMCPServer 单元测试 | BaseMCPServer unit tests
"""

import os
from unittest.mock import AsyncMock, MagicMock, patch

import pytest
from confz import DataSource
from mcp.server import NotificationOptions
from mcp.types import (
    CallToolRequest,
    CallToolRequestParams,
    ImageContent,
    SubscribeRequest,
    TextContent,
    UnsubscribeRequest,
)

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.a2c_smcp.server import BaseMCPServer
from office4ai.a2c_smcp.tools.ppt import PptGetSlideInfoTool, PptGetSlideScreenshotTool
from office4ai.environment.workspace.base import OfficeObs


class ConcreteMCPServer(BaseMCPServer):
    """用于测试的具体 MCP Server 实现 | Concrete MCP Server implementation for testing"""

    def _register_tools(self):
        pass

    def _register_resources(self):
        pass


class TestBaseMCPServer:
    """BaseMCPServer 测试类 | BaseMCPServer test class"""

    def test_initialization(self):
        """测试初始化 | Test initialization"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            assert server.config == config
            assert server.server.name == "test-server"
            assert isinstance(server.tools, dict)
            assert isinstance(server.resources, dict)

    def test_register_tools_method(self):
        """测试注册工具方法 | Test register tools method"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            # 应该调用 _register_tools | Should call _register_tools
            assert hasattr(server, "_register_tools")

    def test_register_resources_method(self):
        """测试注册资源方法 | Test register resources method"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            # 应该调用 _register_resources | Should call _register_resources
            assert hasattr(server, "_register_resources")

    @pytest.mark.asyncio
    async def test_run_stdio_transport(self):
        """测试 stdio 传输模式运行 | Test running with stdio transport"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "stdio"})):
            config = MCPServerConfig()
            server = ConcreteMCPServer(config, "test-server")

            # stdio 需要实际的输入输出流，在单元测试中应该跳过或模拟 | stdio needs actual streams, should skip or mock in unit tests
            with patch("office4ai.a2c_smcp.server.stdio_server") as mock_stdio:
                mock_stdio.side_effect = AttributeError("stdio needs actual streams")
                with pytest.raises(AttributeError):
                    await server.run()

    @pytest.mark.asyncio
    async def test_run_invalid_transport(self):
        """测试无效传输模式 | Test invalid transport mode"""
        # 创建一个模拟的无效配置 | Create a mock invalid config
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": "invalid"})):
            with pytest.raises(ValueError):  # confz will raise validation error
                MCPServerConfig()


class TestSubscribeCapability:
    """resources.subscribe capability 回归测试 | resources.subscribe capability regression tests.

    守护 issue #3 修复 (commit 09b0f55): MCP SDK 把 ``subscribe=False`` 硬编码进
    ``Server.get_capabilities()``, 即使已注册 subscribe/unsubscribe handler 也不上报。
    A2C-SMCP 客户端 ``list_windows()`` 在 ``capabilities.resources.subscribe`` 为假时返回空列表,
    导致 Desktop 拿不到窗口。``BaseMCPServer._patch_subscribe_capability`` 包装
    ``create_initialization_options`` 将其纠正为 ``subscribe=True``。

    Guards the issue #3 fix: the SDK hardcodes ``subscribe=False`` in ``get_capabilities()``
    even when subscribe/unsubscribe handlers are registered, which made the A2C-SMCP client
    return an empty window list. A future SDK bump or handler refactor could silently revert
    this; these tests pin the contract.
    """

    def _make_server(self) -> ConcreteMCPServer:
        with patch.dict(os.environ, {}, clear=True):
            return ConcreteMCPServer(MCPServerConfig(), "test-server")

    def test_create_initialization_options_declares_subscribe_true(self):
        """核心回归: 客户端实际收到的 capability 必须声明 subscribe=True。

        Core regression: the capability the client actually receives must declare subscribe=True.
        """
        server = self._make_server()

        options = server.server.create_initialization_options()
        resources_cap = options.capabilities.resources

        assert resources_cap is not None
        assert resources_cap.subscribe is True

    def test_subscribe_handlers_registered(self):
        """前置条件锚点: patch 的 guard 条件依赖 subscribe/unsubscribe handler 已注册。

        Precondition anchor for ``_patch_subscribe_capability``'s guard
        (``SubscribeRequest in request_handlers or UnsubscribeRequest in ...``):
        if a future refactor drops the handler registration, this fails first and
        localizes the cause to "handlers gone" rather than a vague "subscribe=False".
        """
        server = self._make_server()

        assert SubscribeRequest in server.server.request_handlers
        assert UnsubscribeRequest in server.server.request_handlers

    def test_streamable_http_path_declares_subscribe_true(self):
        """全 transport 覆盖: streamable-http 经 StreamableHTTPSessionManager 也须声明 subscribe=True。

        stdio/sse runners call ``self.server.create_initialization_options()`` directly, but
        streamable-http goes through ``StreamableHTTPSessionManager.app.create_initialization_options()``.
        All three share the same ``Server`` instance today, so the instance-level patch covers them;
        this pins that contract so a future wrapper/copy of the server in the streamable-http path
        can't silently regress to subscribe=False.
        """
        from mcp.server.streamable_http_manager import StreamableHTTPSessionManager

        server = self._make_server()
        manager = StreamableHTTPSessionManager(server.server)

        options = manager.app.create_initialization_options()

        assert options.capabilities.resources is not None
        assert options.capabilities.resources.subscribe is True

    def test_raw_sdk_hardcodes_subscribe_false(self):
        """上游行为文档化: 裸 SDK get_capabilities() 仍把 subscribe 钉死为 False。

        Documents the upstream behavior the patch defends against; if the SDK ever stops
        hardcoding False this test flips and signals the workaround may be removable.

        Pinned against mcp==1.20.0 (get_capabilities hardcodes subscribe=False at
        mcp/server/lowlevel/server.py). A flip here on a future mcp bump is the signal,
        not a regression — re-check whether ``_patch_subscribe_capability`` is still needed.
        """
        server = self._make_server()

        raw_cap = server.server.get_capabilities(NotificationOptions(), {})

        assert raw_cap.resources is not None
        assert raw_cap.resources.subscribe is False


class TestCallToolContentDispatch:
    """``call_tool`` 内容类型分发回归测试 (office4ai #42)。

    守护根因修复: ``call_tool`` 必须委托 ``tool.to_mcp_content(result)`` 决定 MCP 内容类型,
    不得再无条件 ``str(result)`` 包成 text。截图工具据此发 ``image`` 块, 否则十余万字符
    base64 内联进 LLM 文本上下文 → 撑爆 token → 周期性压缩 → 死循环。

    经 ``server.request_handlers[CallToolRequest]`` 触发真实 SDK 分发路径, 端到端覆盖
    "工具 to_mcp_content → SDK CallToolResult" 这一出口契约。
    """

    def _make_server(self) -> ConcreteMCPServer:
        with patch.dict(os.environ, {}, clear=True):
            return ConcreteMCPServer(MCPServerConfig(), "test-server")

    def _mock_workspace(self, obs: OfficeObs) -> MagicMock:
        ws = MagicMock()
        ws.execute = AsyncMock(return_value=obs)
        return ws

    async def _call(self, server: ConcreteMCPServer, name: str, arguments: dict):
        handler = server.server.request_handlers[CallToolRequest]
        req = CallToolRequest(method="tools/call", params=CallToolRequestParams(name=name, arguments=arguments))
        result = await handler(req)
        return result.root

    @pytest.mark.asyncio
    async def test_screenshot_dispatched_as_image_content(self):
        """核心回归: 截图经 call_tool → MCP image 内容块, base64 落在 data, 不内联进 text。"""
        server = self._make_server()
        ws = self._mock_workspace(OfficeObs(success=True, data={"base64": "ABC123", "format": "png"}))
        tool = PptGetSlideScreenshotTool(ws)
        server.tools[tool.name] = tool

        call_result = await self._call(server, tool.name, {"document_uri": "file:///t.pptx", "slideIndex": 0})

        assert call_result.isError is False
        assert len(call_result.content) == 1
        block = call_result.content[0]
        assert isinstance(block, ImageContent)
        assert block.data == "ABC123"
        assert block.mimeType == "image/png"

    @pytest.mark.asyncio
    async def test_regular_tool_dispatched_as_text_content(self):
        """回归保护: 非截图工具仍走默认 text 分支, call_tool 行为不变。"""
        server = self._make_server()
        ws = self._mock_workspace(OfficeObs(success=True, data={"slideIndex": 0, "title": "Hello"}))
        tool = PptGetSlideInfoTool(ws)
        server.tools[tool.name] = tool

        call_result = await self._call(server, tool.name, {"document_uri": "file:///t.pptx", "slideIndex": 0})

        assert call_result.isError is False
        assert len(call_result.content) == 1
        assert isinstance(call_result.content[0], TextContent)
