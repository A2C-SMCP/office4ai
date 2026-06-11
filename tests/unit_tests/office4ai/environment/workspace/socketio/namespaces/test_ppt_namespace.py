"""
Test PptNamespace functionality

测试 PptNamespace 的事件处理功能。

Architecture Note:
    PptNamespace 只处理 Client → Server 的事件通知（fire-and-forget）:
    - ppt:event:slideChanged

    Server → Client 的命令（ppt:get:*, ppt:insert:* 等）通过 MCP BaseTool
    + OfficeWorkspace.emit_to_document() 直接发送，不经过 namespace handler。
"""

import logging
from typing import Any

import pytest
from socketio.exceptions import ConnectionRefusedError  # type: ignore[import-untyped]

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.namespaces.ppt import PptNamespace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
)


class TestPptNamespace:
    """Test PptNamespace class"""

    @pytest.fixture
    def ppt_namespace(self) -> PptNamespace:
        """Create PptNamespace instance"""
        return PptNamespace()

    @pytest.fixture
    def connected_session(self, ppt_namespace: PptNamespace) -> Any:
        """Create a connected session for testing"""
        sid = "test_ppt_socket_123"
        connection_manager.register_client(sid, "ppt-client1", "file:///test.pptx", "/ppt")
        yield sid
        # Cleanup
        connection_manager.unregister_client(sid)

    # ========================================================================
    # 初始化测试
    # ========================================================================

    def test_namespace_init(self, ppt_namespace: PptNamespace) -> None:
        """Test PptNamespace initializes with correct namespace"""
        assert ppt_namespace.namespace_name == "/ppt"

    @pytest.mark.asyncio
    async def test_on_connect_incompatible_version_rejected(
        self, ppt_namespace: PptNamespace, handshake_data_incompatible_version: dict[str, Any]
    ) -> None:
        """The inherited version gate fires on /ppt too: incompatible oaspVersion → 2006.

        Proves the gate runs per-namespace, so a future PptNamespace.on_connect override
        that drops the version check would fail here (the /word contract test alone wouldn't).
        """
        sid = "ppt_reject_sid"
        with pytest.raises(ConnectionRefusedError) as exc_info:
            await ppt_namespace.on_connect(sid, {}, handshake_data_incompatible_version)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.PROTOCOL_VERSION_MISMATCH
        assert connection_manager.get_client_info(sid) is None

    def test_namespace_has_only_event_handlers(self, ppt_namespace: PptNamespace) -> None:
        """Test PptNamespace only has event handlers, not command handlers"""
        # 事件处理器应存在
        assert hasattr(ppt_namespace, "on_ppt_event_slideChanged")

        # 命令处理器不应存在（已移至 MCP BaseTool）
        assert not hasattr(ppt_namespace, "on_ppt_get_currentSlideElements")
        assert not hasattr(ppt_namespace, "on_ppt_insert_text")

    # ========================================================================
    # slideChanged 事件测试
    # ========================================================================

    @pytest.mark.asyncio
    async def test_slide_changed_with_connected_client(
        self, ppt_namespace: PptNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test slideChanged event with a registered client logs correctly"""
        data = {
            "eventType": "slideChanged",
            "clientId": "ppt-client1",
            "documentUri": "file:///test.pptx",
            "data": {"fromIndex": 0, "toIndex": 2},
            "timestamp": 1234567890,
        }

        with caplog.at_level(logging.INFO):
            await ppt_namespace.on_ppt_event_slideChanged(connected_session, data)

        assert any("PPT slide changed" in record.message for record in caplog.records)
        assert any("from 0 to 2" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_slide_changed_with_unknown_client(
        self, ppt_namespace: PptNamespace, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test slideChanged event with unregistered sid does not crash"""
        data = {
            "eventType": "slideChanged",
            "clientId": "unknown",
            "documentUri": "file:///test.pptx",
            "data": {"fromIndex": 0, "toIndex": 1},
            "timestamp": 1234567890,
        }

        with caplog.at_level(logging.INFO):
            await ppt_namespace.on_ppt_event_slideChanged("unknown_sid", data)

        # 未注册的 sid 不应产生日志（get_client_info 返回 None）
        assert not any("PPT slide changed" in record.message for record in caplog.records)

    @pytest.mark.asyncio
    async def test_slide_changed_with_empty_data(
        self, ppt_namespace: PptNamespace, connected_session: Any, caplog: pytest.LogCaptureFixture
    ) -> None:
        """Test slideChanged event with missing data fields does not crash"""
        data: dict[str, Any] = {}

        with caplog.at_level(logging.INFO):
            await ppt_namespace.on_ppt_event_slideChanged(connected_session, data)

        assert any("from ? to ?" in record.message for record in caplog.records)
