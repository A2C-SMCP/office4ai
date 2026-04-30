# filename: test_server.py
# @Time    : 2025/12/18 19:15
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

"""
OfficeMCPServer 单元测试 | OfficeMCPServer unit tests
"""

import os
from unittest.mock import patch

import pytest
from confz import DataSource

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.office.mcp.server import OfficeMCPServer


class TestOfficeMCPServer:
    """OfficeMCPServer 测试类 | OfficeMCPServer test class"""

    def test_server_initialization(self):
        """测试服务器初始化 | Test server initialization"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            assert server.config == config
            assert server.server.name == "office4ai"
            assert server.workspace is not None
            assert not server.workspace.is_running

    def test_tools_registered(self):
        """测试工具已注册 | Test tools are registered"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            # 25 Word (21 + 4 OASP v0.2.0 table tools) + 24 PPT (21 + 3 OASP v0.2.0 chart tools) = 49
            assert len(server.tools) == 49

            expected_tools = [
                # Word Get tools
                "word_get_selected_content",
                "word_get_visible_content",
                "word_get_selection",
                "word_get_document_structure",
                "word_get_document_stats",
                "word_get_styles",
                # Word Text operation tools
                "word_insert_text",
                "word_append_text",
                "word_replace_text",
                "word_replace_selection",
                "word_select_text",
                # Word Multimedia tools
                "word_insert_image",
                "word_insert_table",
                "word_insert_equation",
                "word_insert_toc",
                # Word Table operation tools (OASP /word Draft, v0.2.0)
                "word_merge_cells",
                "word_update_table_cell",
                "word_update_table_row_column",
                "word_update_table_format",
                # Word Export tool
                "word_export_content",
                # Word Comment tools
                "word_get_comments",
                "word_insert_comment",
                "word_delete_comment",
                "word_reply_comment",
                "word_resolve_comment",
                # PPT Content retrieval tools
                "ppt_get_current_slide_elements",
                "ppt_get_slide_elements",
                "ppt_get_slide_screenshot",
                "ppt_get_slide_info",
                "ppt_get_slide_layouts",
                # PPT Content insertion tools
                "ppt_insert_text",
                "ppt_insert_image",
                "ppt_insert_table",
                "ppt_insert_shape",
                # PPT Update tools
                "ppt_update_text_box",
                "ppt_update_image",
                "ppt_update_table_cell",
                "ppt_update_table_row_column",
                "ppt_update_table_format",
                "ppt_update_element",
                # PPT Chart tools (OASP /ppt Draft, v0.2.0 — Server OOXML)
                "ppt_insert_chart",
                "ppt_get_chart",
                "ppt_update_chart",
                # PPT Delete & layout tools
                "ppt_delete_element",
                "ppt_reorder_element",
                # PPT Slide management tools
                "ppt_add_slide",
                "ppt_delete_slide",
                "ppt_move_slide",
                "ppt_goto_slide",
            ]
            for tool_name in expected_tools:
                assert tool_name in server.tools, f"Tool {tool_name} not registered"

    def test_resources_registered(self):
        """测试资源已注册 | Test resources are registered"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            assert len(server.resources) == 3
            assert "window://office4ai" in server.resources
            assert "window://office4ai/word" in server.resources
            assert "window://office4ai/ppt" in server.resources

    @pytest.mark.parametrize("transport", ["stdio", "sse", "streamable-http"])
    def test_server_supports_all_transports(self, transport):
        """测试服务器支持所有传输模式 | Test server supports all transport modes"""
        with MCPServerConfig.change_config_sources(DataSource(data={"transport": transport})):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            assert server.config.transport == transport

    def test_workspace_port_from_config(self):
        """测试 workspace 使用 config 中的 socketio_port | Test workspace uses socketio_port from config"""
        with MCPServerConfig.change_config_sources(DataSource(data={"socketio_port": 4000})):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            assert server.workspace.port == 4000

    @pytest.mark.asyncio
    async def test_async_lifecycle(self):
        """测试 async 生命周期钩子 | Test async lifecycle hooks"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            # 启动 workspace
            await server._async_startup()
            assert server.workspace.is_running

            # 停止 workspace
            await server._async_shutdown()
            assert not server.workspace.is_running
