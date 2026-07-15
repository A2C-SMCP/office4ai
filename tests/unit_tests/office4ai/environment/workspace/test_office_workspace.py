"""
Test OfficeWorkspace functionality

测试 OfficeWorkspace 的核心功能。
"""

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.environment.workspace.base import DocumentStatus, OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
)


class TestOfficeWorkspace:
    """Test OfficeWorkspace class"""

    @pytest.fixture
    def office_workspace(self) -> OfficeWorkspace:
        """Create OfficeWorkspace instance"""
        return OfficeWorkspace(host="127.0.0.1", port=3000, use_https=False)

    @pytest.fixture
    def connected_session(self, office_workspace: OfficeWorkspace) -> None:
        """Mock a connected Add-In session"""
        # Mock client connection
        connection_manager.register_client(
            socket_id="test_socket_123",
            client_id="client1",
            document_uri="file:///test.docx",
            namespace="/word",
        )
        yield
        # Cleanup
        connection_manager.unregister_client("test_socket_123")

    # ========================================================================
    # Test execute() method
    # ========================================================================

    @pytest.mark.asyncio
    async def test_execute_missing_document_uri(self, office_workspace: OfficeWorkspace) -> None:
        """Test execute() with missing document_uri"""
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={},  # Missing document_uri
        )

        result = await office_workspace.execute(action)

        assert result.success is False
        assert "Missing document_uri" in result.error
        assert result.data == {}

    @pytest.mark.asyncio
    async def test_execute_document_not_connected(self, office_workspace: OfficeWorkspace) -> None:
        """Test execute() with disconnected document"""
        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={"document_uri": "file:///nonexistent.docx"},
        )

        result = await office_workspace.execute(action)

        assert result.success is False
        assert "Document not connected" in result.error
        assert result.data == {}

    @pytest.mark.asyncio
    async def test_execute_success(self, office_workspace: OfficeWorkspace, connected_session: None) -> None:
        """Test successful execute() call"""
        # Mock sio_server.call() to return proper response format
        office_workspace.sio_server = MagicMock()
        office_workspace.sio_server.call = AsyncMock(
            return_value={
                "requestId": "test_req_001",
                "success": True,
                "data": {"text": "Selected content"},  # Business data wrapped in "data" field
                "timestamp": 1234567890000,
            }
        )

        action = OfficeAction(
            category="word",
            action_name="get:selectedContent",
            params={
                "document_uri": "file:///test.docx",
                "options": {"includeText": True},
            },
        )

        result = await office_workspace.execute(action)

        assert result.success is True
        assert result.data == {"text": "Selected content"}  # Extracted business data
        assert result.error is None

        # Verify sio_server.call was invoked with wrapped request
        office_workspace.sio_server.call.assert_called_once()
        call_args = office_workspace.sio_server.call.call_args
        assert call_args[0][0] == "word:get:selectedContent"  # event name
        assert "requestId" in call_args[0][1]  # wrapped data should have requestId
        assert call_args[0][1]["documentUri"] == "file:///test.docx"

    @pytest.mark.asyncio
    async def test_execute_error_includes_details(
        self, office_workspace: OfficeWorkspace, connected_session: None
    ) -> None:
        """execute() 摊平 wire error 时 details 必须存活（issue #82 / oasp#17）。

        协议规定 details 供双端对齐断言（如 3010 ELEMENT_NOT_FOUND 用
        details.kind 区分 worksheet/table/chart/pivotTable）。若 execute()
        只保留 "code: message"，e2e 无法断言 kind，MCP 工具层的 AI 消费者
        也拿不到可行动的定位信息。
        """
        office_workspace.sio_server = MagicMock()
        office_workspace.sio_server.call = AsyncMock(
            return_value={
                "requestId": "test_req_002",
                "success": False,
                "error": {
                    "code": "3010",
                    "message": "Worksheet not found",
                    "details": {"kind": "worksheet", "name": "GhostSheet"},
                },
                "timestamp": 1234567890000,
            }
        )

        action = OfficeAction(
            category="excel",
            action_name="get:worksheetInfo",
            params={"document_uri": "file:///test.docx", "worksheet_name": "GhostSheet"},
        )

        result = await office_workspace.execute(action)

        assert result.success is False
        assert "3010" in result.error
        assert "Worksheet not found" in result.error
        # details 必须存活于错误字符串（kind 与回带的标识都可见）
        assert "worksheet" in result.error
        assert "GhostSheet" in result.error

    @pytest.mark.asyncio
    async def test_execute_error_without_details(
        self, office_workspace: OfficeWorkspace, connected_session: None
    ) -> None:
        """无 details 的 wire error 保持 "code: message" 形态，不追加噪音。"""
        office_workspace.sio_server = MagicMock()
        office_workspace.sio_server.call = AsyncMock(
            return_value={
                "requestId": "test_req_003",
                "success": False,
                "error": {"code": "3009", "message": "Invalid range address"},
                "timestamp": 1234567890000,
            }
        )

        action = OfficeAction(
            category="excel",
            action_name="get:range",
            params={"document_uri": "file:///test.docx", "address": "ZZZZ99999999"},
        )

        result = await office_workspace.execute(action)

        assert result.success is False
        assert result.error == "3009: Invalid range address"

    # ========================================================================
    # Test get_document_status() method
    # ========================================================================

    def test_get_document_status_connected(self, office_workspace: OfficeWorkspace, connected_session: None) -> None:
        """Test get_document_status() returns CONNECTED for active document"""
        status = office_workspace.get_document_status("file:///test.docx")
        assert status == DocumentStatus.CONNECTED

    def test_get_document_status_disconnected(self, office_workspace: OfficeWorkspace) -> None:
        """Test get_document_status() returns DISCONNECTED for non-existent document"""
        status = office_workspace.get_document_status("file:///nonexistent.docx")
        assert status == DocumentStatus.DISCONNECTED

    # ========================================================================
    # Test emit_to_document() method
    # ========================================================================

    @pytest.mark.asyncio
    async def test_emit_to_document_no_socket(self, office_workspace: OfficeWorkspace) -> None:
        """Test emit_to_document() raises ValueError when socket not found"""
        office_workspace.sio_server = MagicMock()

        with pytest.raises(ValueError, match="No socket found for document"):
            await office_workspace.emit_to_document(
                "file:///nonexistent.docx",
                "word:test:event",
                {},
            )

    @pytest.mark.asyncio
    async def test_emit_to_document_success(self, office_workspace: OfficeWorkspace, connected_session: None) -> None:
        """Test successful emit_to_document() call"""
        # Mock sio_server
        office_workspace.sio_server = MagicMock()
        office_workspace.sio_server.call = AsyncMock(return_value={"result": "success"})

        # Use a registered event (word:get:selectedContent is registered in DTOs)
        result = await office_workspace.emit_to_document(
            "file:///test.docx",
            "word:get:selectedContent",
            {"options": {"includeText": True}},
        )

        assert result == {"result": "success"}

        # Verify call was made correctly
        office_workspace.sio_server.call.assert_called_once()
        call_args = office_workspace.sio_server.call.call_args
        assert call_args[0][0] == "word:get:selectedContent"
        assert "requestId" in call_args[0][1]  # Request should be wrapped
        assert call_args[1]["to"] == "test_socket_123"
        assert call_args[1]["namespace"] == "/word"

    # ========================================================================
    # Test request wrapping
    # ========================================================================

    @pytest.mark.asyncio
    async def test_request_wrapping_in_execute(
        self, office_workspace: OfficeWorkspace, connected_session: None
    ) -> None:
        """Test that business params are wrapped into BaseRequest format"""
        office_workspace.sio_server = MagicMock()
        office_workspace.sio_server.call = AsyncMock(return_value={})

        action = OfficeAction(
            category="word",
            action_name="insert:text",
            params={
                "document_uri": "file:///test.docx",
                "text": "Hello World",
                "location": "Cursor",
            },
        )

        await office_workspace.execute(action)

        # Verify wrapped request
        call_args = office_workspace.sio_server.call.call_args
        wrapped_data = call_args[0][1]

        # Should have BaseRequest fields
        assert "requestId" in wrapped_data
        assert "documentUri" in wrapped_data
        assert wrapped_data["documentUri"] == "file:///test.docx"
        assert "timestamp" in wrapped_data

        # Should have business params
        assert wrapped_data.get("text") == "Hello World"
        assert wrapped_data.get("location") == "Cursor"

    # ========================================================================
    # Test server lifecycle
    # ========================================================================

    def test_initial_state(self, office_workspace: OfficeWorkspace) -> None:
        """Test initial workspace state"""
        assert office_workspace.is_running is False
        assert office_workspace.sio_server is None
        assert office_workspace.app is None

    @pytest.mark.asyncio
    async def test_start_stop(self, office_workspace: OfficeWorkspace) -> None:
        """Test starting and stopping the workspace"""
        await office_workspace.start()
        assert office_workspace.is_running is True
        assert office_workspace.sio_server is not None

        await office_workspace.stop()
        assert office_workspace.is_running is False

    # ========================================================================
    # Test utility methods
    # ========================================================================

    def test_get_connected_documents(self, office_workspace: OfficeWorkspace, connected_session: None) -> None:
        """Test get_connected_documents() returns list of connected documents"""
        documents = office_workspace.get_connected_documents()
        assert len(documents) == 1
        assert "file:///test.docx" in documents

    @pytest.mark.asyncio
    async def test_wait_for_addin_connection_timeout(self, office_workspace: OfficeWorkspace) -> None:
        """Test wait_for_addin_connection() returns False on timeout"""
        result = await office_workspace.wait_for_addin_connection(timeout=0.5)
        assert result is False

    @pytest.mark.asyncio
    async def test_wait_for_addin_connection_success(
        self, office_workspace: OfficeWorkspace, connected_session: None
    ) -> None:
        """Test wait_for_addin_connection() returns True when already connected"""
        result = await office_workspace.wait_for_addin_connection(timeout=0.5)
        assert result is True
