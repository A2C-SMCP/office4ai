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

            # 27 Word (21 + 4 OASP v0.2.0 table tools + 2 OASP v0.3.0 OOXML round-trip #46)
            # + 24 PPT (21 + 3 OASP v0.2.0 chart tools)
            # + 37 Excel (OASP 0.3.0 Draft: #18 read slice 3 + #19 Range CRUD/公式 7
            #   + #20 Format/条件格式/合并 6 + #21 Worksheet 管理 5 + #22 Table 操作 6
            #   + #23 Chart 操作 4 + #24 PivotTable 操作 3 + #25 Find&Filter 操作 3) = 88
            assert len(server.tools) == 88

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
                # Word OOXML fragment round-trip tools (OASP /word Draft, v0.3.0, #46)
                "word_get_ooxml",
                "word_insert_ooxml",
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
                # Excel state-awareness read tools (OASP /excel Draft 0.3.0, #18 Foundation)
                "excel_get_workbook_info",
                "excel_get_worksheet_info",
                "excel_get_selected_range",
                # Excel Range CRUD + 公式 tools (OASP /excel Draft 0.3.0, #19)
                "excel_get_range",
                "excel_set_range",
                "excel_clear_range",
                "excel_copy_range",
                "excel_delete_range",
                "excel_insert_range",
                "excel_set_formula",
                # Excel Format / 条件格式 / 合并单元格 tools (OASP /excel Draft 0.3.0, #20)
                "excel_get_range_format",
                "excel_set_range_format",
                "excel_add_conditional_format",
                "excel_clear_conditional_format",
                "excel_merge_cells",
                "excel_unmerge_cells",
                # Excel Worksheet 管理 tools (OASP /excel Draft 0.3.0, #21)
                "excel_get_worksheets",
                "excel_add_worksheet",
                "excel_delete_worksheet",
                "excel_rename_worksheet",
                "excel_activate_worksheet",
                # Excel Table 操作 tools (OASP /excel Draft 0.3.0, #22)
                "excel_insert_table",
                "excel_get_table",
                "excel_get_tables",
                "excel_add_table_row",
                "excel_delete_table_row",
                "excel_sort_table",
                # Excel Chart 操作 tools (OASP /excel Draft 0.3.0, #23)
                "excel_insert_chart",
                "excel_get_charts",
                "excel_update_chart",
                "excel_delete_chart",
                # Excel PivotTable 操作 tools (OASP /excel Draft 0.3.0, #24)
                "excel_insert_pivot_table",
                "excel_get_pivot_tables",
                "excel_delete_pivot_table",
                # Excel Find & Filter 操作 tools (OASP /excel Draft 0.3.0, #25)
                "excel_find_values",
                "excel_set_auto_filter",
                "excel_clear_auto_filter",
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
