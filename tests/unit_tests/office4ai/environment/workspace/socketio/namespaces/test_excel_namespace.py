"""
Test ExcelNamespace functionality

测试 ExcelNamespace 的握手能力。

Architecture Note:
    ExcelNamespace 当前仅实装握手（继承 BaseNamespace.on_connect / on_disconnect），
    不包含任何 Client → Server 事件处理器——对应事件 handler 将在后续迭代中逐步补充。

    Server → Client 的命令（excel:get:*, excel:set:*, excel:insert:* 等）通过 MCP
    BaseTool + OfficeWorkspace.emit_to_document() 直接发送，不经过 namespace handler。
"""

from typing import Any

import pytest
from socketio.exceptions import ConnectionRefusedError  # type: ignore[import-untyped]

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.namespaces.excel import ExcelNamespace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
)


class TestExcelNamespace:
    """Test ExcelNamespace class"""

    @pytest.fixture
    def excel_namespace(self) -> ExcelNamespace:
        """Create ExcelNamespace instance"""
        return ExcelNamespace()

    @pytest.mark.asyncio
    async def test_on_connect_incompatible_version_rejected(
        self, excel_namespace: ExcelNamespace, handshake_data_incompatible_version: dict[str, Any]
    ) -> None:
        """The inherited version gate fires on /excel too: incompatible oaspVersion → 2006.

        Stronger than the `hasattr(on_connect)` structural check — proves the gate actually
        runs per-namespace, so a future ExcelNamespace.on_connect override that drops the
        version check would fail here.
        """
        sid = "excel_reject_sid"
        with pytest.raises(ConnectionRefusedError) as exc_info:
            await excel_namespace.on_connect(sid, {}, handshake_data_incompatible_version)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.PROTOCOL_VERSION_MISMATCH
        assert connection_manager.get_client_info(sid) is None

    def test_namespace_init(self, excel_namespace: ExcelNamespace) -> None:
        """Test ExcelNamespace initializes with correct namespace"""
        assert excel_namespace.namespace_name == "/excel"

    def test_namespace_inherits_handshake_from_base(self, excel_namespace: ExcelNamespace) -> None:
        """Test ExcelNamespace exposes handshake lifecycle from BaseNamespace"""
        # 握手与生命周期由 BaseNamespace 提供
        assert hasattr(excel_namespace, "on_connect")
        assert hasattr(excel_namespace, "on_disconnect")

    def test_namespace_has_no_event_handlers_yet(self, excel_namespace: ExcelNamespace) -> None:
        """Excel 事件处理器尚未实装——该断言在后续补充 event handler 时需同步更新"""
        # Client → Server 事件 handler 尚未实装
        assert not hasattr(excel_namespace, "on_excel_event_selectionChanged")
        assert not hasattr(excel_namespace, "on_excel_event_worksheetActivated")

        # 命令处理器本就不应存在（走 MCP BaseTool + sio.call() RPC 路径）
        assert not hasattr(excel_namespace, "on_excel_get_selectedRange")
        assert not hasattr(excel_namespace, "on_excel_set_cellValue")
