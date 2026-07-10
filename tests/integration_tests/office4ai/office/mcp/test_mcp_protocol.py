# filename: test_mcp_protocol.py
# @Time    : 2025/12/18 19:18
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
MCP 协议集成测试 | MCP protocol integration tests
"""

from pathlib import Path

import pytest
from mcp.client.session import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client
from pydantic import AnyUrl

# 项目根目录路径 | Project root path
project_root = Path(__file__).parent.parent.parent.parent.parent


@pytest.mark.integration
class TestMCPProtocol:
    """MCP 协议测试类 | MCP protocol test class"""

    async def test_mcp_handshake(self):
        """测试 MCP 握手 | Test MCP handshake"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                # 初始化会话 | Initialize session
                result = await session.initialize()

                # 验证服务器信息 | Verify server info
                assert result is not None
                assert result.serverInfo.name == "office4ai"

    async def test_list_tools(self):
        """测试 list_tools 返回已注册工具 | Test list_tools returns registered tools"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 获取工具列表 | Get tools list
                tools_result = await session.list_tools()
                # W4a（#63）：list_tools 按 Add-In 连接动态收敛。子进程无任何 Add-In 连接
                # → 工具集收敛到仅常驻 office_run_script（requires_connection=False）。
                # 「87 个已注册」不变量见 test_server.py::test_tools_registered（server.tools）；
                # 连接触发的收敛见 test_w4_desktop_convergence（端到端）。
                tool_names = {t.name for t in tools_result.tools}
                assert tool_names == {"office_run_script"}

    async def test_list_resources(self):
        """测试 list_resources 返回已注册资源 | Test list_resources returns registered resources"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 获取资源列表 | Get resources list
                resources_result = await session.list_resources()

                # 验证资源 URI | Verify resource URIs
                # 注：SKILL 资源在协议层被展平为「根 + 各子文件」多条目，总数随 SKILL 内容/后续
                # S5/S6 变化 → 用 presence 断言而非硬编码总数（server.resources 注册数由单测钉住）。
                resource_uris = {str(r.uri) for r in resources_result.resources}
                # W4b-1（#64）：无连接时只有根索引 + SKILL；per-type 聚合窗口已移除，
                # per-file 窗口随 Add-In 连接动态出现（子进程无连接故不含）。
                assert any(uri.startswith("window://office4ai?") for uri in resource_uris)
                assert not any("window://office4ai/word" in uri for uri in resource_uris)
                assert not any("window://office4ai/ppt" in uri for uri in resource_uris)
                assert "skill://com.a2c-smcp.office4ai/create-office-file" in resource_uris
                assert "skill://com.a2c-smcp.office4ai/edit-office-file" in resource_uris
                # 旧资源已删除
                assert not any("office://workspace/documents" in uri for uri in resource_uris)

    async def test_call_tool_not_found(self):
        """测试调用不存在的工具 | Test calling non-existent tool"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 尝试调用不存在的工具 | Try to call non-existent tool
                result = await session.call_tool("non_existent_tool", {})
                assert len(result.content) == 1
                assert "未找到工具 | Tool not found" in result.content[0].text

    async def test_read_resource_not_found(self):
        """测试读取不存在的资源 | Test reading non-existent resource"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()

                # 尝试读取不存在的资源 | Try to read non-existent resource
                from mcp.shared.exceptions import McpError

                with pytest.raises(McpError, match="未找到资源 | Resource not found"):
                    await session.read_resource(AnyUrl("office://non/existent/resource"))


@pytest.mark.integration
class TestMCPResourcesPhase1:
    """Phase 1: Window 资源 MCP 协议层集成测试"""

    async def test_list_resources_includes_window_resources(self):
        """list_resources 无连接时返回根索引 + SKILL；per-type 聚合窗口已移除（W4b-1）"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                resources = await session.list_resources()
                uris = [str(r.uri) for r in resources.resources]

                # SKILL 资源协议层展平为多条目 → presence 断言（见 test_list_resources 注释）
                assert any(u.startswith("window://office4ai?") for u in uris)
                assert not any("window://office4ai/word" in u for u in uris)
                assert not any("window://office4ai/ppt" in u for u in uris)
                assert "skill://com.a2c-smcp.office4ai/create-office-file" in uris
                assert "skill://com.a2c-smcp.office4ai/edit-office-file" in uris
                assert not any("office://workspace/documents" in u for u in uris)

    async def test_read_window_root_resource(self):
        """读取根索引：无连接时返回 '暂无文档连接'"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.read_resource(AnyUrl("window://office4ai"))
                content = result.contents[0].text
                assert "Office 工作区" in content
                assert "暂无文档连接" in content

    async def test_aggregation_windows_removed(self):
        """W4b-1：per-type 聚合窗口（word/ppt）已移除 → 读取报未找到资源。"""
        server_params = StdioServerParameters(
            command="uv",
            args=["run", "python", "-m", "office4ai.office.mcp.server"],
            cwd=str(project_root),
        )

        from mcp.shared.exceptions import McpError

        async with stdio_client(server_params) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                for uri in ("window://office4ai/word", "window://office4ai/ppt"):
                    with pytest.raises(McpError, match="未找到资源 | Resource not found"):
                        await session.read_resource(AnyUrl(uri))
