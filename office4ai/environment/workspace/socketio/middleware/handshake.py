"""
Handshake Middleware

Simple handshake validation for client connections.
Validates clientId and documentUri are provided.
"""

import logging
import time
from collections.abc import Awaitable, Callable
from typing import Any

from socketio import AsyncServer  # type: ignore[import-untyped]
from socketio.exceptions import ConnectionRefusedError  # type: ignore[import-untyped]

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.versioning import (
    SERVER_MAX_SUPPORTED,
    SERVER_MIN_SUPPORTED,
    SERVER_VERSION,
    OaspVersion,
    is_compatible,
)

logger = logging.getLogger(__name__)


def handshake_rejection(message: str, code: str, **extra: str) -> ConnectionRefusedError:
    """
    构造一个握手拒绝异常，使其结构化数据正确送达客户端 ``connect_error`` 的 ``error.data``。

    python-socketio 的 ``ConnectionRefusedError`` 约定：第一个位置参数作为 error 的 ``message``，
    其余位置参数作为 ``data`` 随错误包回送客户端。故此处用两参形式
    ``ConnectionRefusedError(message, data)``，让 JS 客户端拿到 ``err.message`` 与扁平
    ``err.data``（HandshakeRejection）。单一 dict 参数会被库 ``str()`` 化而丢失结构，**勿用**。

    Args:
        message: 人类可读消息（同时作为 error message 与 data.message）
        code: 错误码（ErrorCode.HANDSHAKE_FAILED / PROTOCOL_VERSION_MISMATCH）
        **extra: 仅 PROTOCOL_VERSION_MISMATCH 附带的诊断字段（serverVersion 等）

    Returns:
        待 ``raise`` 的 ``ConnectionRefusedError``
    """
    data: dict[str, Any] = {"code": code, "message": message, **extra}
    return ConnectionRefusedError(message, data)


def validate_oasp_version(auth: dict[str, Any] | None) -> OaspVersion:
    """
    校验握手 ``auth.oaspVersion`` 的协议版本兼容性（**先于业务参数**校验）。

    规范：connection.md#protocol-version-handshake、conventions.md#compatibility-rule。

    校验顺序（版本属协议层，先于 clientId/documentUri 业务层）：

    1. 缺少 ``oaspVersion`` / 格式非法 → 拒绝，``HANDSHAKE_FAILED`` (2003)
    2. 版本不兼容（``is_compatible`` 假）→ 拒绝，``PROTOCOL_VERSION_MISMATCH`` (2006)
       并携带扁平 ``HandshakeRejection`` 结构（serverVersion/clientVersion/min/max）

    拒绝数据经 ``ConnectionRefusedError`` 原样送达客户端 ``connect_error`` 的 ``error.data``。

    Args:
        auth: 客户端握手 auth 字典

    Returns:
        解析且兼容的客户端 ``OaspVersion``

    Raises:
        ConnectionRefusedError: 缺失/非法（HANDSHAKE_FAILED）或不兼容（PROTOCOL_VERSION_MISMATCH）
    """
    raw = (auth or {}).get("oaspVersion")
    if not raw:
        raise handshake_rejection("Missing oaspVersion", ErrorCode.HANDSHAKE_FAILED)
    if not isinstance(raw, str):
        raise handshake_rejection(f"Invalid oaspVersion: {raw!r}", ErrorCode.HANDSHAKE_FAILED)
    try:
        client_version = OaspVersion.parse(raw)
    except ValueError:
        raise handshake_rejection(f"Invalid oaspVersion: {raw}", ErrorCode.HANDSHAKE_FAILED) from None

    if not is_compatible(client_version, SERVER_VERSION):
        # 扁平拒绝结构（非标准 ErrorResponse——握手期无 requestId）
        raise handshake_rejection(
            "Protocol version mismatch",
            ErrorCode.PROTOCOL_VERSION_MISMATCH,
            serverVersion=str(SERVER_VERSION),
            clientVersion=str(client_version),
            minSupported=str(SERVER_MIN_SUPPORTED),
            maxSupported=str(SERVER_MAX_SUPPORTED),
        )

    return client_version


def build_connection_established(socket_id: str) -> dict[str, Any]:
    """
    构造 ``connection:established`` 回执负载（含 ``serverVersion`` 供诊断）。

    规范：connection.md#connection-established —— ``{socketId, serverVersion, timestamp}``。
    握手通过即代表版本已兼容，``serverVersion`` 仅供 AddIn 日志/UI 展示/故障定位。
    """
    return {
        "socketId": socket_id,
        "serverVersion": str(SERVER_VERSION),
        "timestamp": int(time.time() * 1000),
    }


async def handshake_middleware(
    socketio_server: AsyncServer, namespace: str = "/"
) -> Callable[[str, Any], Awaitable[bool]]:
    """
    Create handshake middleware for a namespace.

    This is a simple validation that checks:
    1. clientId is provided
    2. documentUri is provided

    No complex authentication - just basic validation for local connections.

    Args:
        socketio_server: Socket.IO server instance
        namespace: Namespace to apply middleware to

    Returns:
        Middleware function
    """

    async def middleware(sid: str, environ: Any) -> bool:
        """
        Handshake middleware handler.

        Args:
            sid: Session ID
            environ: WSGI environ dict

        Raises:
            ValueError: If handshake data is invalid
        """
        # Get handshake auth data
        # Note: python-socketio passes auth in a different way
        # We'll get it from the socket handshake in the connection handler
        _auth = environ.get("asgi_scope", {}).get("query_string", b"").decode()

        # Parse auth data from handshake
        # Note: python-socketio passes auth in a different way
        # We'll get it from the socket handshake in the connection handler

        logger.debug(f"Handshake attempt for session {sid} on namespace {namespace}")

        # Accept all connections (localhost-only security)
        # Actual validation happens in connection handler
        return True

    return middleware


def validate_handshake_data(client_id: str, document_uri: str) -> tuple[bool, str]:
    """
    Validate handshake data from client.

    Args:
        client_id: Client-provided ID
        document_uri: Document URI

    Returns:
        Tuple of (is_valid, error_message)
    """
    if not client_id:
        return False, "Missing clientId"

    if not document_uri:
        return False, "Missing documentUri"

    # Basic format validation for document URI
    if not document_uri.startswith(("file:///", "http://", "https://")):
        return False, f"Invalid documentUri format: {document_uri}"

    return True, ""


def log_handshake(client_id: str, document_uri: str, namespace: str) -> None:
    """
    Log successful handshake.

    Args:
        client_id: Client ID
        document_uri: Document URI
        namespace: Namespace
    """
    logger.info(f"Handshake successful: client={client_id}, document={document_uri}, namespace={namespace}")
