"""
Base Namespace

Base functionality for all Socket.IO namespaces.
"""

import logging
from typing import Any

from socketio import AsyncNamespace  # type: ignore[import-untyped]
from socketio.exceptions import ConnectionRefusedError  # type: ignore[import-untyped]

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.middleware.handshake import (
    build_connection_established,
    handshake_rejection,
    log_handshake_rejection,
    validate_oasp_version,
)
from office4ai.environment.workspace.socketio.services.connection_manager import ClientInfo, connection_manager

logger = logging.getLogger(__name__)


class BaseNamespace(AsyncNamespace):
    """
    Base namespace class with common functionality.

    All namespaces (/word, /ppt, /excel) inherit from this.
    """

    def __init__(self, namespace: str) -> None:
        """
        Initialize base namespace.

        Args:
            namespace: Namespace name (/word, /ppt, /excel)
            ns_name: Optional namespace name (defaults to namespace)
        """
        super().__init__(namespace)
        self.namespace_name: str = namespace
        logger.info(f"BaseNamespace initialized: {self.namespace_name}")

    async def on_connect(self, sid: str, environ: dict, auth: dict | None = None) -> None:
        """
        Handle client connection.

        校验顺序遵循 OASP 0.3.0 协议版本握手（connection.md#protocol-version-handshake）：
        协议版本（``oaspVersion``）**先于**业务参数（``clientId`` / ``documentUri``）校验。
        任何拒绝都抛 ``ConnectionRefusedError``，其结构化数据原样送达客户端 ``connect_error``。

        Args:
            sid: Session ID
            environ: Connection environment dict (contains HTTP headers, query params, etc.)
            auth: Auth data from client (contains oaspVersion, clientId, documentUri)

        Raises:
            ConnectionRefusedError: 版本缺失/非法/不兼容，或业务握手参数缺失
        """
        data = auth if auth else {}

        try:
            # ① 协议版本先行校验：缺失/非法 → HANDSHAKE_FAILED(2003)；不兼容 → 2006
            client_version = validate_oasp_version(data)

            # ② 业务参数校验（版本兼容后再查）
            client_id = data.get("clientId")
            document_uri = data.get("documentUri")
            if not client_id or not document_uri:
                raise handshake_rejection("Missing required auth parameters", ErrorCode.HANDSHAKE_FAILED)

            # ③ 注册连接
            try:
                connection_manager.register_client(
                    socket_id=sid,
                    client_id=client_id,
                    document_uri=document_uri,
                    namespace=self.namespace_name,
                )
            except Exception as e:
                logger.error("Error registering client", exc_info=True)
                raise handshake_rejection("Internal error during connection registration", ErrorCode.UNKNOWN) from e
        except ConnectionRefusedError as rejection:
            # 所有握手拒绝在 namespace 边界统一诊断，wire 异常原样重抛。
            log_handshake_rejection(self.namespace_name, sid, rejection)
            raise

        # ④ 发送确认（含 serverVersion，仅供诊断）
        await self.emit("connection:established", build_connection_established(sid), to=sid)

        logger.info(
            f"Client connected: {client_id} ({sid}) for {document_uri} "
            f"on {self.namespace_name} (oaspVersion={client_version})"
        )

    async def on_disconnect(self, sid: str) -> None:
        """
        Handle client disconnection.

        Args:
            sid: Session ID
        """
        client_info = connection_manager.unregister_client(sid)

        if client_info:
            logger.info(f"Client disconnected: {client_info.client_id} ({sid}) from {client_info.document_uri}")
        else:
            logger.warning(f"Unknown client disconnected: {sid}")

    async def on_connection_status(self, sid: str, data: Any) -> None:
        """
        Handle connection status updates from client.

        Args:
            sid: Session ID
            data: Status data
        """
        logger.debug(f"Connection status from {sid}: {data}")
        # Can be used for health checks
        # No response needed (fire-and-forget)

    def get_client_info(self, sid: str) -> ClientInfo | None:
        """
        Get client information.

        Args:
            sid: Session ID

        Returns:
            ClientInfo if found, None otherwise
        """
        return connection_manager.get_client_info(sid)
