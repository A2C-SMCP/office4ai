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

            # 26 Word (21 + 4 OASP v0.2.0 table tools + 1 run:script) + 25 PPT (21 + 3 chart + 1 run:script)
            # + 38 Excel (37 OASP 0.3.0 Draft #18–25 + 1 run:script) = 89
            # + 1 authoring standalone (milestone #4 · S1: office_run_script) = 90
            # +3 online run:script escape hatches (OASP {ns}:run:script, issue #87)
            assert len(server.tools) == 90

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
                # Word online script escape hatch (OASP word:run:script, issue #87)
                "word_run_script",
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
                # PPT online script escape hatch (OASP ppt:run:script, issue #87)
                "ppt_run_script",
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
                # Excel online script escape hatch (OASP excel:run:script, issue #87)
                "excel_run_script",
                # authoring standalone tool (milestone #4 · S1)
                "office_run_script",
            ]
            for tool_name in expected_tools:
                assert tool_name in server.tools, f"Tool {tool_name} not registered"

    def test_resources_registered(self):
        """测试资源已注册 | Test resources are registered"""
        with patch.dict(os.environ, {}, clear=True):
            config = MCPServerConfig()
            server = OfficeMCPServer(config)

            # W4b-1（#64）：per-type 聚合窗口（word/ppt）移除，per-file 窗口随连接动态注册。
            # 静态注册 = 根索引 1 + 3 生产 SKILL（S4 create / S5 edit / S6 extract）= 4。
            assert len(server.resources) == 4
            assert "window://office4ai" in server.resources
            assert "window://office4ai/word" not in server.resources
            assert "window://office4ai/ppt" not in server.resources
            assert "skill://com.a2c-smcp.office4ai/create-office-file" in server.resources
            assert "skill://com.a2c-smcp.office4ai/edit-office-file" in server.resources
            assert "skill://com.a2c-smcp.office4ai/extract-template" in server.resources

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


@pytest.fixture
def server() -> OfficeMCPServer:
    with patch.dict(os.environ, {}, clear=True):
        return OfficeMCPServer(MCPServerConfig())


@pytest.fixture
def clean_cm():
    """提供干净的全局 connection_manager（前后清空，避免跨测试污染）。"""
    from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager

    def _clear() -> None:
        for c in list(connection_manager.get_all_clients()):
            connection_manager.unregister_client(c.socket_id)

    _clear()
    yield connection_manager
    _clear()


class TestW4aToolConvergence:
    """W4a（#63）：按 Add-In 连接动态收敛工具集。"""

    def test_no_connection_only_standalone(self, server: OfficeMCPServer, clean_cm) -> None:
        assert {t.name for t in server._visible_tools()} == {"office_run_script"}

    def test_word_connection_exposes_word_plus_standalone(self, server: OfficeMCPServer, clean_cm) -> None:
        clean_cm.register_client("s1", "c1", "file:///a/x.docx", "/word")
        names = {t.name for t in server._visible_tools()}
        assert "office_run_script" in names
        assert any(n.startswith("word_") for n in names)
        assert not any(n.startswith("ppt_") for n in names)
        assert not any(n.startswith("excel_") for n in names)

    def test_chart_tool_requires_connection(self, server: OfficeMCPServer, clean_cm) -> None:
        # chart 工具（category='ppt'）无 /ppt 连接时被过滤（重分类 requires_connection=True）
        assert not any(t.name == "ppt_insert_chart" for t in server._visible_tools())
        clean_cm.register_client("s2", "c2", "file:///a/y.pptx", "/ppt")
        assert any(t.name == "ppt_insert_chart" for t in server._visible_tools())

    def test_is_tool_available_matrix(self, server: OfficeMCPServer, clean_cm) -> None:
        tools = {t.name: t for t in server.tools.values()}
        assert server._is_tool_available(tools["office_run_script"]) is True  # 常驻
        assert server._is_tool_available(tools["word_insert_text"]) is False  # 无连接
        clean_cm.register_client("s3", "c3", "file:///a/x.docx", "/word")
        assert server._is_tool_available(tools["word_insert_text"]) is True


class TestListToolsAnnotations:
    """守护 list_tools() 把 tool.annotations 接线进 Tool()（issue #87）。

    单测已覆盖工具本身的 annotations 属性；本类走**真实 list_tools handler**，防止
    server.py 里 `annotations=tool.annotations` 那行被误删而 CI 静默全绿
    （destructiveHint/openWorldHint 是「向上层披露风险面」的语义载体）。
    """

    @staticmethod
    async def _list_tools(server: OfficeMCPServer):
        import mcp.types as types

        handler = server.server.request_handlers[types.ListToolsRequest]
        result = await handler(types.ListToolsRequest(method="tools/list"))
        return {t.name: t for t in result.root.tools}

    @pytest.mark.asyncio
    async def test_run_script_tool_carries_annotations_on_wire(self, server: OfficeMCPServer, clean_cm) -> None:
        clean_cm.register_client("sa1", "ca1", "file:///a/x.docx", "/word")
        by_name = await self._list_tools(server)

        wr = by_name["word_run_script"]
        assert wr.annotations is not None
        assert wr.annotations.destructiveHint is True
        assert wr.annotations.openWorldHint is True
        # 对照：typed 工具无注解，证明接线是逐工具透传而非全局硬编码
        assert by_name["word_insert_text"].annotations is None


class TestW4bDynamicWindows:
    """W4b-1（#64）：per-file 窗口随连接动态注册/注销。"""

    def test_connect_registers_per_file_window(self, server: OfficeMCPServer, clean_cm) -> None:
        from office4ai.a2c_smcp.resources.per_file_window import WordFileWindowResource, per_file_window_base_uri

        doc = "file:///a/report.docx"
        buri = per_file_window_base_uri("/word", doc)
        assert buri not in server.resources
        server._on_doc_connect(doc, "/word")
        assert isinstance(server.resources.get(buri), WordFileWindowResource)

    def test_disconnect_removes_per_file_window(self, server: OfficeMCPServer, clean_cm) -> None:
        from office4ai.a2c_smcp.resources.per_file_window import per_file_window_base_uri

        doc = "file:///a/deck.pptx"
        server._on_doc_connect(doc, "/ppt")
        assert per_file_window_base_uri("/ppt", doc) in server.resources
        server._on_doc_disconnect(doc, "/ppt")
        assert per_file_window_base_uri("/ppt", doc) not in server.resources

    def test_excel_connect_projects_window(self, server: OfficeMCPServer, clean_cm) -> None:
        from office4ai.a2c_smcp.resources.per_file_window import ExcelFileWindowResource, per_file_window_base_uri

        doc = "file:///a/book.xlsx"
        buri = per_file_window_base_uri("/excel", doc)
        server._on_doc_connect(doc, "/excel")  # W4b-3 / #66：excel 现投射 per-file 窗口
        assert isinstance(server.resources.get(buri), ExcelFileWindowResource)

    def test_affected_resource_uris_targets_per_file_window(self, server: OfficeMCPServer, clean_cm) -> None:
        tools = {t.name: t for t in server.tools.values()}
        uris = server._affected_resource_uris(tools["word_insert_text"], {"document_uri": "file:///a/x.docx"})
        assert len(uris) == 1 and uris[0].startswith("window://office4ai/word/x.docx-")
        # 无 document_uri → 空；authoring 工具 → 空
        assert server._affected_resource_uris(tools["word_insert_text"], {}) == []
        assert server._affected_resource_uris(tools["office_run_script"], {"document_uri": "file:///a/x.docx"}) == []


class TestW4bFullscreenOwnership:
    """W4b-2（#65）：单 fullscreen 归属 = AI 最后操作文件；断连顺延次新；关闭 #4。

    直接驱动 server 的活动/连接回调（生产由 workspace.update_last_activity + connection_manager
    触发）；断言 per-file 窗口 fullscreen 状态与「至多一个 fullscreen」不变量。同步上下文下通知
    fire-and-forget 因无运行 loop 优雅 no-op，不影响状态翻转。
    """

    @staticmethod
    def _win(server: OfficeMCPServer, namespace: str, doc: str):
        from office4ai.a2c_smcp.resources.per_file_window import per_file_window_base_uri

        return server.resources[per_file_window_base_uri(namespace, doc)]

    @staticmethod
    def _fullscreen_count(server: OfficeMCPServer) -> int:
        from office4ai.a2c_smcp.resources.per_file_window import PerFileWindowResource

        return sum(1 for r in server.resources.values() if isinstance(r, PerFileWindowResource) and r.fullscreen)

    def test_activity_flips_fullscreen_and_clears_others(self, server: OfficeMCPServer, clean_cm) -> None:
        a, b = "file:///a/a.docx", "file:///a/b.pptx"
        server._on_doc_connect(a, "/word")
        server._on_doc_connect(b, "/ppt")
        assert self._fullscreen_count(server) == 0  # 连接不自动 fullscreen（须 AI 操作）

        server._on_doc_activity(a)
        assert self._win(server, "/word", a).fullscreen is True
        assert self._win(server, "/ppt", b).fullscreen is False

        # 操作 b → fullscreen 转移，a 清零（先清后置）
        server._on_doc_activity(b)
        assert self._win(server, "/word", a).fullscreen is False
        assert self._win(server, "/ppt", b).fullscreen is True

    def test_single_fullscreen_invariant_hash4_regression(self, server: OfficeMCPServer, clean_cm) -> None:
        """#4 回归：反复切换活动文件，任一时刻至多一个 fullscreen。"""
        docs = [("file:///a/a.docx", "/word"), ("file:///a/b.pptx", "/ppt"), ("file:///a/c.xlsx", "/excel")]
        for d, ns in docs:
            server._on_doc_connect(d, ns)
        for d, _ in docs * 2:  # 反复切换
            server._on_doc_activity(d)
            assert self._fullscreen_count(server) <= 1
        assert self._win(server, "/excel", "file:///a/c.xlsx").fullscreen is True  # 末次操作者独占

    def test_activity_on_unwindowed_file_is_noop(self, server: OfficeMCPServer, clean_cm) -> None:
        """无 window 的文件（office_run_script 产物 / 未连接）不参与竞争、不扰动现状。"""
        a = "file:///a/a.docx"
        server._on_doc_connect(a, "/word")
        server._on_doc_activity(a)
        assert self._win(server, "/word", a).fullscreen is True

        server._on_doc_activity("file:///a/generated.docx")  # 无对应 window
        assert self._win(server, "/word", a).fullscreen is True  # a 仍 fullscreen
        assert self._fullscreen_count(server) == 1

    def test_disconnect_fullscreen_holder_defers_to_next(self, server: OfficeMCPServer, clean_cm) -> None:
        """断连收场：fullscreen 持有者断连 → 顺延剩余最近活跃者。"""
        a, b = "file:///a/a.docx", "file:///a/b.pptx"
        server._on_doc_connect(a, "/word")
        server._on_doc_connect(b, "/ppt")
        server._on_doc_activity(a)  # a 活跃（seq1）
        server._on_doc_activity(b)  # b 活跃（seq2），fullscreen
        assert self._win(server, "/ppt", b).fullscreen is True

        server._on_doc_disconnect(b, "/ppt")  # fullscreen 持有者断连
        assert self._win(server, "/word", a).fullscreen is True  # 顺延到 a
        assert self._fullscreen_count(server) == 1

    def test_disconnect_defers_to_second_newest_not_any(self, server: OfficeMCPServer, clean_cm) -> None:
        """断连收场须让给「次新」而非任一剩余：a(seq1)/b(seq2)/c(seq3,fullscreen)，c 断 → b。"""
        a, b, c = "file:///a/a.docx", "file:///a/b.pptx", "file:///a/c.xlsx"
        server._on_doc_connect(a, "/word")
        server._on_doc_connect(b, "/ppt")
        server._on_doc_connect(c, "/excel")
        server._on_doc_activity(a)  # seq1
        server._on_doc_activity(b)  # seq2
        server._on_doc_activity(c)  # seq3，fullscreen
        assert self._win(server, "/excel", c).fullscreen is True

        server._on_doc_disconnect(c, "/excel")  # 持有者断连
        assert self._win(server, "/ppt", b).fullscreen is True  # 顺延到 b（次新）
        assert self._win(server, "/word", a).fullscreen is False  # 不是 a
        assert self._fullscreen_count(server) == 1

    def test_disconnect_fullscreen_holder_no_successor(self, server: OfficeMCPServer, clean_cm) -> None:
        """断连收场：无剩余活跃者 → 无 fullscreen（根索引主视）。"""
        a = "file:///a/a.docx"
        server._on_doc_connect(a, "/word")
        server._on_doc_activity(a)
        server._on_doc_disconnect(a, "/word")
        assert self._fullscreen_count(server) == 0

    def test_disconnect_non_fullscreen_leaves_owner(self, server: OfficeMCPServer, clean_cm) -> None:
        """断连非 fullscreen 文件 → fullscreen 持有者不变。"""
        a, b = "file:///a/a.docx", "file:///a/b.pptx"
        server._on_doc_connect(a, "/word")
        server._on_doc_connect(b, "/ppt")
        server._on_doc_activity(a)  # a fullscreen
        server._on_doc_disconnect(b, "/ppt")  # 断非 fullscreen 的 b
        assert self._win(server, "/word", a).fullscreen is True
        assert self._fullscreen_count(server) == 1

    def test_reactivating_same_file_is_idempotent(self, server: OfficeMCPServer, clean_cm) -> None:
        """重复操作同一 fullscreen 文件 → 幂等（仍单 fullscreen，无翻转）。"""
        a = "file:///a/a.docx"
        server._on_doc_connect(a, "/word")
        server._on_doc_activity(a)
        server._on_doc_activity(a)
        assert self._win(server, "/word", a).fullscreen is True
        assert self._fullscreen_count(server) == 1
