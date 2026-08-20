"""
Test BaseNamespace functionality

测试 BaseNamespace 的所有核心功能。
"""

import json
import logging
from typing import Any
from unittest.mock import AsyncMock

import pytest
from socketio.exceptions import ConnectionRefusedError  # type: ignore[import-untyped]

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.namespaces.base import BaseNamespace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
)
from office4ai.environment.workspace.socketio.versioning import OASP_PROTOCOL_VERSION


def rejection_log_fields(caplog: pytest.LogCaptureFixture) -> dict[str, str]:
    """Extract the structured field payload from the single rejection log."""
    records = [record for record in caplog.records if "event=oasp_handshake_rejected" in record.getMessage()]
    assert len(records) == 1
    return json.loads(records[0].getMessage().split(" fields=", 1)[1])


class TestBaseNamespace:
    """Test BaseNamespace class"""

    @pytest.fixture
    def base_namespace(self) -> BaseNamespace:
        """Create BaseNamespace instance"""
        return BaseNamespace("/test")

    @pytest.mark.asyncio
    async def test_on_connect_success(
        self, base_namespace: BaseNamespace, valid_handshake_data: dict[str, Any]
    ) -> None:
        """Test successful connection"""
        sid = "test_socket_123"

        # Mock emit and disconnect
        base_namespace.emit = AsyncMock()  # type: ignore[method-assign]
        base_namespace.disconnect = AsyncMock()  # type: ignore[method-assign]

        # on_connect signature: (sid, environ, auth)
        await base_namespace.on_connect(sid, {}, valid_handshake_data)

        # Verify client registered
        assert connection_manager.get_client_info(sid) is not None

        # Verify confirmation sent with spec-compliant {socketId, serverVersion, timestamp}
        base_namespace.emit.assert_called_once()
        call_args = base_namespace.emit.call_args
        assert call_args[0][0] == "connection:established"
        assert call_args[1]["to"] == sid
        payload = call_args[0][1]
        assert payload["socketId"] == sid
        assert payload["serverVersion"] == str(OASP_PROTOCOL_VERSION)
        assert "timestamp" in payload

        # Cleanup
        connection_manager.unregister_client(sid)

    @pytest.mark.asyncio
    async def test_on_connect_missing_client_id(
        self,
        base_namespace: BaseNamespace,
        invalid_handshake_data_missing_client_id: dict[str, Any],
    ) -> None:
        """Test connection refused without clientId (version OK → fails business param check)"""
        sid = "test_socket_123"

        # on_connect signature: (sid, environ, auth)
        with pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, invalid_handshake_data_missing_client_id)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED
        # Not registered
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_connect_missing_document_uri(
        self,
        base_namespace: BaseNamespace,
        invalid_handshake_data_missing_document_uri: dict[str, Any],
    ) -> None:
        """Test connection refused without documentUri (version OK → fails business param check)"""
        sid = "test_socket_123"

        with pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, invalid_handshake_data_missing_document_uri)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_connect_missing_version(
        self,
        base_namespace: BaseNamespace,
        handshake_data_missing_version: dict[str, Any],
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        """Version-first: missing oaspVersion → HANDSHAKE_FAILED before business check"""
        sid = "test_socket_123"

        with caplog.at_level(logging.WARNING), pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, handshake_data_missing_version)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED
        assert connection_manager.get_client_info(sid) is None
        assert rejection_log_fields(caplog) == {
            "code": ErrorCode.HANDSHAKE_FAILED,
            "namespace": "/test",
            "reason": "missing_oasp_version",
            "sid": sid,
        }
        assert handshake_data_missing_version["documentUri"] not in caplog.text

    @pytest.mark.asyncio
    async def test_on_connect_invalid_version(
        self,
        base_namespace: BaseNamespace,
        handshake_data_invalid_version: dict[str, Any],
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        """Version-first: malformed oaspVersion → HANDSHAKE_FAILED"""
        sid = "test_socket_123"

        with caplog.at_level(logging.WARNING), pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, handshake_data_invalid_version)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED
        assert connection_manager.get_client_info(sid) is None
        assert rejection_log_fields(caplog)["reason"] == "invalid_oasp_version"
        assert handshake_data_invalid_version["oaspVersion"] not in caplog.text
        assert handshake_data_invalid_version["documentUri"] not in caplog.text

    @pytest.mark.asyncio
    async def test_on_connect_incompatible_version(
        self,
        base_namespace: BaseNamespace,
        handshake_data_incompatible_version: dict[str, Any],
        caplog: pytest.LogCaptureFixture,
    ) -> None:
        """Version-first: incompatible oaspVersion → PROTOCOL_VERSION_MISMATCH (2006)"""
        sid = "test_socket_123"

        with caplog.at_level(logging.WARNING), pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, handshake_data_incompatible_version)
        data = exc_info.value.error_args["data"]
        assert data["code"] == ErrorCode.PROTOCOL_VERSION_MISMATCH
        assert data["serverVersion"] == str(OASP_PROTOCOL_VERSION)
        assert data["clientVersion"] == "0.2.0"
        assert rejection_log_fields(caplog) == {
            "client_version": "0.2.0",
            "code": ErrorCode.PROTOCOL_VERSION_MISMATCH,
            "max_supported": "0.4.999",
            "min_supported": "0.4.0",
            "namespace": "/test",
            "reason": "protocol_version_mismatch",
            "server_version": "0.4.0",
            "sid": sid,
        }
        assert handshake_data_incompatible_version["documentUri"] not in caplog.text
        # Not registered
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_connect_version_checked_before_business_params(self, base_namespace: BaseNamespace) -> None:
        """Version-first ordering: when the version is incompatible AND business params are
        BOTH missing, the version failure (2006) wins — proving version is validated first.

        Existing tests can't prove ordering: the bad-version cases all carry valid business
        params, and the missing-param case carries a valid version. This sends both bad.
        """
        sid = "test_socket_order"

        with pytest.raises(ConnectionRefusedError) as exc_info:
            # Incompatible oaspVersion AND no clientId / documentUri.
            await base_namespace.on_connect(sid, {}, {"oaspVersion": "0.2.0"})
        # If business were checked first this would be HANDSHAKE_FAILED (2003).
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.PROTOCOL_VERSION_MISMATCH
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_disconnect(self, base_namespace: BaseNamespace, valid_handshake_data: dict[str, Any]) -> None:
        """Test client disconnection"""
        sid = "test_socket_123"

        # First connect
        base_namespace.emit = AsyncMock()  # type: ignore[method-assign]
        # on_connect signature: (sid, environ, auth)
        await base_namespace.on_connect(sid, {}, valid_handshake_data)

        # Verify connected
        assert connection_manager.get_client_info(sid) is not None

        # Now disconnect
        await base_namespace.on_disconnect(sid)

        # Verify cleaned up
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_connection_status(self, base_namespace: BaseNamespace) -> None:
        """Test connection status updates (fire-and-forget)"""
        sid = "test_socket_123"
        status_data: dict[str, Any] = {"status": "ready", "timestamp": 1234567890}

        # Should not raise any errors
        await base_namespace.on_connection_status(sid, status_data)

    def test_get_client_info(self, base_namespace: BaseNamespace, valid_handshake_data: dict[str, Any]) -> None:
        """Test getting client info from namespace"""
        sid = "test_socket_123"

        # Manually register for test
        connection_manager.register_client(sid, "client1", "file:///test.docx", "/test")

        client = base_namespace.get_client_info(sid)
        assert client is not None
        assert client.socket_id == sid

        # Cleanup
        connection_manager.unregister_client(sid)
