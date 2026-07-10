"""W4a + W4b-1 端到端集成回归（#63 / #64 执行手段：office4ai ↔ Computer）。

用 MCP 内存 client↔server 会话，模拟 Add-In 连接变化，断言：
- 工具集按连接即时收敛（无连接=仅常驻；连 Word=Word 工具集+常驻）；
- Computer（client）收到 tools/list_changed + resources/list_changed 通知；
- per-file window 随连接出现在 list_resources。

不起真实 Socket.IO：直接经全局 connection_manager 注册/注销客户端触发连接回调
（等价于 _async_startup 的回调接线），聚焦 MCP 协议层收敛 + 通知投递。
"""

from __future__ import annotations

import asyncio
import os
from unittest.mock import patch

import pytest
from mcp import types
from mcp.shared.memory import create_connected_server_and_client_session
from pydantic import AnyUrl

from office4ai.a2c_smcp.config import MCPServerConfig
from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager
from office4ai.office.mcp.server import OfficeMCPServer


async def _wait_for(predicate, timeout: float = 2.0) -> None:
    """轮询等待 predicate() 为真（通知经内存流异步投递）。"""
    deadline = asyncio.get_event_loop().time() + timeout
    while asyncio.get_event_loop().time() < deadline:
        if predicate():
            return
        await asyncio.sleep(0.02)


@pytest.mark.integration
class TestW4DesktopConvergenceE2E:
    async def test_connection_change_converges_tools_and_notifies_computer(self):
        # 干净的全局 connection_manager + 快照回调列表以便还原
        for c in list(connection_manager.get_all_clients()):
            connection_manager.unregister_client(c.socket_id)
        saved_connect = list(connection_manager._on_document_connect)
        saved_disc = list(connection_manager._on_document_disconnect_ns)

        with patch.dict(os.environ, {}, clear=True):
            office = OfficeMCPServer(MCPServerConfig())
        # 接线连接回调（生产在 _async_startup 中完成）
        connection_manager.register_connect_callback(office._on_doc_connect)
        connection_manager.register_disconnect_callback_ns(office._on_doc_disconnect)

        notif_methods: list[str] = []

        async def message_handler(message) -> None:
            if isinstance(message, types.ServerNotification):
                notif_methods.append(message.root.method)

        try:
            async with create_connected_server_and_client_session(
                office.server, message_handler=message_handler
            ) as client:
                # 基线：无连接 → 仅常驻工具
                tools0 = await client.list_tools()
                assert {t.name for t in tools0.tools} == {"office_run_script"}

                # 订阅根索引 → 客户端会话被跟踪，后续 list_changed 才会投递到它
                await client.subscribe_resource(AnyUrl("window://office4ai"))

                # 模拟一个 Word Add-In 连接
                connection_manager.register_client("s1", "c1", "file:///tmp/report.docx", "/word")

                # Computer 收到两类 list_changed 通知
                await _wait_for(
                    lambda: "notifications/tools/list_changed" in notif_methods
                    and "notifications/resources/list_changed" in notif_methods
                )
                assert "notifications/tools/list_changed" in notif_methods
                assert "notifications/resources/list_changed" in notif_methods

                # 工具集收敛：Word 工具出现，PPT/Excel 仍不暴露
                names1 = {t.name for t in (await client.list_tools()).tools}
                assert "office_run_script" in names1
                assert any(n.startswith("word_") for n in names1)
                assert not any(n.startswith("ppt_") for n in names1)
                assert not any(n.startswith("excel_") for n in names1)

                # per-file window 出现在资源列表
                uris = {str(r.uri) for r in (await client.list_resources()).resources}
                assert any(u.startswith("window://office4ai/word/") for u in uris)

                # 断连 → 工具集回落到仅常驻，per-file window 消失
                notif_methods.clear()
                connection_manager.unregister_client("s1")
                await _wait_for(lambda: "notifications/tools/list_changed" in notif_methods)

                names2 = {t.name for t in (await client.list_tools()).tools}
                assert names2 == {"office_run_script"}
                uris2 = {str(r.uri) for r in (await client.list_resources()).resources}
                assert not any(u.startswith("window://office4ai/word/") for u in uris2)
        finally:
            for c in list(connection_manager.get_all_clients()):
                connection_manager.unregister_client(c.socket_id)
            connection_manager._on_document_connect[:] = saved_connect
            connection_manager._on_document_disconnect_ns[:] = saved_disc

    async def test_fullscreen_ownership_and_excel_window_e2e(self):
        """W4b-2（#65）+ W4b-3（#66）端到端：AI 最后操作文件独占 fullscreen（#4 组装回归），
        断连顺延次新；excel per-file 窗口出现在资源列表。

        经内存 client↔server + workspace.update_last_activity（生产由 BaseTool.execute 触发）
        驱动 fullscreen 归属，从 list_resources 的 URI 查询串观测 fullscreen 状态。
        """
        for c in list(connection_manager.get_all_clients()):
            connection_manager.unregister_client(c.socket_id)
        saved_connect = list(connection_manager._on_document_connect)
        saved_disc = list(connection_manager._on_document_disconnect_ns)

        with patch.dict(os.environ, {}, clear=True):
            office = OfficeMCPServer(MCPServerConfig())
        connection_manager.register_connect_callback(office._on_doc_connect)
        connection_manager.register_disconnect_callback_ns(office._on_doc_disconnect)
        # 活动回调（生产在 _async_startup 接线）—— fullscreen 归属的触发源
        office.workspace.set_activity_callback(office._on_doc_activity)

        per_file_prefixes = (
            "window://office4ai/word/",
            "window://office4ai/ppt/",
            "window://office4ai/excel/",
        )

        def fullscreen_windows(resources) -> set[str]:
            """列出 fullscreen=true 的 per-file 窗口 URI（根 window 恒 false，不计入）。"""
            return {
                str(r.uri)
                for r in resources
                if str(r.uri).startswith(per_file_prefixes) and "fullscreen=true" in str(r.uri)
            }

        word_doc, ppt_doc, xlsx_doc = "file:///tmp/report.docx", "file:///tmp/deck.pptx", "file:///tmp/data.xlsx"

        try:
            async with create_connected_server_and_client_session(office.server) as client:
                connection_manager.register_client("s1", "c1", word_doc, "/word")
                connection_manager.register_client("s2", "c2", ppt_doc, "/ppt")
                connection_manager.register_client("s3", "c3", xlsx_doc, "/excel")

                # excel per-file 窗口出现（W4b-3 / #66）
                uris0 = {str(r.uri) for r in (await client.list_resources()).resources}
                assert any(u.startswith("window://office4ai/excel/") for u in uris0)
                # 未操作任何文件 → 无 fullscreen（#4：至多一个候选，此刻为零）
                assert fullscreen_windows((await client.list_resources()).resources) == set()

                # AI 操作 word 文件 → 仅 word 窗口 fullscreen
                office.workspace.update_last_activity(word_doc, "word_insert_text", {})
                fs1 = fullscreen_windows((await client.list_resources()).resources)
                assert len(fs1) == 1 and any("/word/" in u for u in fs1)

                # AI 改操作 ppt 文件 → fullscreen 转移到 ppt，word 清零（#4 单 fullscreen 不变量）
                office.workspace.update_last_activity(ppt_doc, "ppt_insert_text", {})
                fs2 = fullscreen_windows((await client.list_resources()).resources)
                assert len(fs2) == 1 and any("/ppt/" in u for u in fs2)

                # fullscreen 持有者（ppt）断连 → 顺延次新活跃者（word）。断连回调同步完成状态翻转。
                connection_manager.unregister_client("s2")
                resources3 = (await client.list_resources()).resources
                uris3 = {str(r.uri) for r in resources3}
                assert not any(u.startswith("window://office4ai/ppt/") for u in uris3)  # ppt 窗口已注销
                fs3 = fullscreen_windows(resources3)
                assert len(fs3) == 1 and any("/word/" in u for u in fs3)
        finally:
            for c in list(connection_manager.get_all_clients()):
                connection_manager.unregister_client(c.socket_id)
            connection_manager._on_document_connect[:] = saved_connect
            connection_manager._on_document_disconnect_ns[:] = saved_disc

    async def test_server_declares_list_changed_capability(self):
        """服务器初始化能力声明 tools.listChanged + resources.listChanged（+ subscribe）。"""
        with patch.dict(os.environ, {}, clear=True):
            office = OfficeMCPServer(MCPServerConfig())
        opts = office.server.create_initialization_options()
        caps = opts.capabilities
        assert caps.tools is not None and caps.tools.listChanged is True
        assert caps.resources is not None and caps.resources.listChanged is True
        assert caps.resources.subscribe is True
