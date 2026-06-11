"""
Test BaseNamespace functionality

测试 BaseNamespace 的所有核心功能。
"""

from typing import Any
from unittest.mock import AsyncMock

import pytest
from socketio.exceptions import ConnectionRefusedError  # type: ignore[import-untyped]

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.namespaces.base import BaseNamespace
from office4ai.environment.workspace.socketio.services.connection_manager import (
    connection_manager,
)
from office4ai.environment.workspace.socketio.versioning import SERVER_VERSION


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
        assert payload["serverVersion"] == str(SERVER_VERSION)
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
    ) -> None:
        """Version-first: missing oaspVersion → HANDSHAKE_FAILED before business check"""
        sid = "test_socket_123"

        with pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, handshake_data_missing_version)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_connect_invalid_version(
        self,
        base_namespace: BaseNamespace,
        handshake_data_invalid_version: dict[str, Any],
    ) -> None:
        """Version-first: malformed oaspVersion → HANDSHAKE_FAILED"""
        sid = "test_socket_123"

        with pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, handshake_data_invalid_version)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED
        assert connection_manager.get_client_info(sid) is None

    @pytest.mark.asyncio
    async def test_on_connect_incompatible_version(
        self,
        base_namespace: BaseNamespace,
        handshake_data_incompatible_version: dict[str, Any],
    ) -> None:
        """Version-first: incompatible oaspVersion → PROTOCOL_VERSION_MISMATCH (2006)"""
        sid = "test_socket_123"

        with pytest.raises(ConnectionRefusedError) as exc_info:
            await base_namespace.on_connect(sid, {}, handshake_data_incompatible_version)
        data = exc_info.value.error_args["data"]
        assert data["code"] == ErrorCode.PROTOCOL_VERSION_MISMATCH
        assert data["serverVersion"] == str(SERVER_VERSION)
        assert data["clientVersion"] == "0.2.0"
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
