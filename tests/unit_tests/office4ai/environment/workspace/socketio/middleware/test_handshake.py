"""
Test handshake middleware

测试握手中间件验证功能。
"""

from typing import Any

import pytest
from socketio.exceptions import ConnectionRefusedError  # type: ignore[import-untyped]

from office4ai import __version__
from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.middleware.handshake import (
    build_connection_established,
    handshake_middleware,
    log_handshake,
    validate_handshake_data,
    validate_oasp_version,
)
from office4ai.environment.workspace.socketio.versioning import SERVER_VERSION, OaspVersion


class TestHandshakeMiddleware:
    """Test handshake validation functions"""

    def test_validate_valid_handshake(self) -> None:
        """Test validation with valid data"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="file:///tmp/test.docx")

        assert is_valid is True
        assert error == ""

    def test_validate_missing_client_id(self) -> None:
        """Test validation fails without clientId"""
        is_valid, error = validate_handshake_data(client_id="", document_uri="file:///tmp/test.docx")

        assert is_valid is False
        assert "Missing clientId" in error

    def test_validate_missing_document_uri(self) -> None:
        """Test validation fails without documentUri"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="")

        assert is_valid is False
        assert "Missing documentUri" in error

    def test_validate_invalid_uri_format(self) -> None:
        """Test validation fails with invalid URI format"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="invalid-uri")

        assert is_valid is False
        assert "Invalid documentUri format" in error

    def test_validate_file_uri(self) -> None:
        """Test validation accepts file:// URI"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="file:///C:/Users/test.docx")

        assert is_valid is True

    def test_validate_http_uri(self) -> None:
        """Test validation accepts http/https URIs"""
        is_valid, error = validate_handshake_data(client_id="test_client", document_uri="https://example.com/test.docx")

        assert is_valid is True

    def test_log_handshake(self, caplog: pytest.LogCaptureFixture) -> None:
        """Test handshake logging"""
        import logging

        caplog.set_level(logging.INFO)

        log_handshake("client1", "file:///test.docx", "/word")

        # Verify log message was created
        assert len(caplog.records) > 0
        assert "client1" in caplog.text
        assert "file:///test.docx" in caplog.text

    @pytest.mark.asyncio
    async def test_handshake_middleware(self) -> None:
        """Test handshake middleware function"""
        from socketio import AsyncServer  # type: ignore[import-untyped]

        # Create a mock server
        server = AsyncServer(async_mode="aiohttp", logger=False, engineio_logger=False)

        # Get middleware for /word namespace
        middleware = await handshake_middleware(server, "/word")

        # Test middleware accepts connection
        environ: dict[str, Any] = {
            "asgi_scope": {
                "query_string": b"",
            }
        }

        result = await middleware("test_sid", environ)

        # Middleware should accept all connections (localhost security)
        assert result is True


class TestValidateOaspVersion:
    """validate_oasp_version —— OASP 0.3.0 协议版本握手四分支"""

    def test_compatible_version_returns_parsed(self) -> None:
        """兼容版本：返回解析后的 OaspVersion，不抛异常"""
        result = validate_oasp_version({"oaspVersion": str(SERVER_VERSION)})
        assert result == SERVER_VERSION

    def test_compatible_patch_differs(self) -> None:
        """v0.x：同 MAJOR.MINOR 不同 PATCH 仍兼容"""
        client = OaspVersion(SERVER_VERSION.major, SERVER_VERSION.minor, 99)
        result = validate_oasp_version({"oaspVersion": str(client)})
        assert result == client

    def test_missing_version_rejected_handshake_failed(self) -> None:
        """缺失 oaspVersion → HANDSHAKE_FAILED (2003)"""
        with pytest.raises(ConnectionRefusedError) as exc_info:
            validate_oasp_version({"clientId": "c", "documentUri": "file:///x.docx"})
        # 两参 ConnectionRefusedError：结构化 HandshakeRejection 在 error_args["data"]
        data = exc_info.value.error_args["data"]
        assert data["code"] == ErrorCode.HANDSHAKE_FAILED

    def test_none_auth_rejected_handshake_failed(self) -> None:
        """auth 为 None → HANDSHAKE_FAILED (2003)"""
        with pytest.raises(ConnectionRefusedError) as exc_info:
            validate_oasp_version(None)
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED

    @pytest.mark.parametrize("bad", ["0.3", "0.3.0.1", "abc", "0.x.0"])
    def test_invalid_format_rejected_handshake_failed(self, bad: str) -> None:
        """格式非法 oaspVersion → HANDSHAKE_FAILED (2003)"""
        with pytest.raises(ConnectionRefusedError) as exc_info:
            validate_oasp_version({"oaspVersion": bad})
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED

    def test_non_string_version_rejected_handshake_failed(self) -> None:
        """非字符串 oaspVersion → HANDSHAKE_FAILED (2003)，不应 AttributeError"""
        with pytest.raises(ConnectionRefusedError) as exc_info:
            validate_oasp_version({"oaspVersion": 30})
        assert exc_info.value.error_args["data"]["code"] == ErrorCode.HANDSHAKE_FAILED

    def test_incompatible_version_rejected_2006_with_fields(self) -> None:
        """不兼容版本 → PROTOCOL_VERSION_MISMATCH (2006) + 扁平诊断字段"""
        # v0.x 严格 MINOR：构造一个不同 MINOR 的客户端版本
        incompatible = OaspVersion(SERVER_VERSION.major, SERVER_VERSION.minor + 1, 0)
        with pytest.raises(ConnectionRefusedError) as exc_info:
            validate_oasp_version({"oaspVersion": str(incompatible)})
        data = exc_info.value.error_args["data"]
        assert data["code"] == ErrorCode.PROTOCOL_VERSION_MISMATCH
        assert data["serverVersion"] == str(SERVER_VERSION)
        assert data["clientVersion"] == str(incompatible)
        assert "minSupported" in data
        assert "maxSupported" in data
        # error message（顶层）亦应有意义，便于无 data 解析的客户端
        assert exc_info.value.error_args["message"] == "Protocol version mismatch"


class TestBuildConnectionEstablished:
    """build_connection_established —— spec 形状 {socketId, serverVersion, timestamp}"""

    def test_shape_and_server_version(self) -> None:
        payload = build_connection_established("sid_abc")
        assert payload["socketId"] == "sid_abc"
        assert payload["serverVersion"] == str(SERVER_VERSION) == __version__
        assert isinstance(payload["timestamp"], int)
        assert set(payload.keys()) == {"socketId", "serverVersion", "timestamp"}
