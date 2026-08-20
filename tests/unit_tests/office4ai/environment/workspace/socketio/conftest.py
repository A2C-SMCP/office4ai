"""
Socket.IO test fixtures

提供 Socket.IO 单元测试所需的 fixtures。
"""

from __future__ import annotations

import time
from typing import Any

import pytest

from office4ai.environment.workspace.socketio.config import SocketIOConfig
from office4ai.environment.workspace.socketio.services.connection_manager import (
    ClientInfo,
    ConnectionManager,
)
from office4ai.environment.workspace.socketio.versioning import OASP_PROTOCOL_VERSION

# 与 Server 的 OASP 协议版本一致；软件包版本 bump 不影响握手 fixture。
_COMPATIBLE_OASP_VERSION = str(OASP_PROTOCOL_VERSION)


@pytest.fixture
def socketio_config() -> SocketIOConfig:
    """创建测试用 Socket.IO 配置"""
    return SocketIOConfig(
        host="127.0.0.1",
        port=3000,
        cors_allowed_origins=["http://localhost:*"],
        ping_timeout=5000,
        ping_interval=1000,
        logger=False,
        engineio_logger=False,
    )


@pytest.fixture
def connection_manager() -> ConnectionManager:
    """创建独立的 ConnectionManager 实例

    每个测试获得一个干净的 manager 实例。
    测试后自动清理所有状态。
    """
    manager = ConnectionManager()
    yield manager
    # 清理
    manager._clients.clear()
    manager._document_to_sockets.clear()
    manager._client_id_to_socket.clear()


@pytest.fixture
def mock_client_info() -> ClientInfo:
    """创建模拟客户端信息"""
    return ClientInfo(
        socket_id="test_socket_123",
        client_id="test_client_abc",
        document_uri="file:///tmp/test.docx",
        namespace="/word",
        connected_at=time.time(),
    )


@pytest.fixture
def valid_handshake_data() -> dict[str, Any]:
    """有效的握手数据（含 OASP 0.3.0 强制的 oaspVersion）"""
    return {
        "clientId": "test_client_123",
        "documentUri": "file:///tmp/test.docx",
        "oaspVersion": _COMPATIBLE_OASP_VERSION,
    }


@pytest.fixture
def invalid_handshake_data_missing_client_id() -> dict[str, Any]:
    """缺少 clientId 的无效握手数据（oaspVersion 合法，确保停在业务参数校验）"""
    return {
        "documentUri": "file:///tmp/test.docx",
        "oaspVersion": _COMPATIBLE_OASP_VERSION,
    }


@pytest.fixture
def invalid_handshake_data_missing_document_uri() -> dict[str, Any]:
    """缺少 documentUri 的无效握手数据（oaspVersion 合法，确保停在业务参数校验）"""
    return {
        "clientId": "test_client_123",
        "oaspVersion": _COMPATIBLE_OASP_VERSION,
    }


@pytest.fixture
def invalid_handshake_data_invalid_uri() -> dict[str, Any]:
    """无效 URI 格式的握手数据"""
    return {
        "clientId": "test_client_123",
        "documentUri": "invalid-uri-format",
        "oaspVersion": _COMPATIBLE_OASP_VERSION,
    }


@pytest.fixture
def handshake_data_missing_version() -> dict[str, Any]:
    """缺少 oaspVersion 的握手数据（应 HANDSHAKE_FAILED）"""
    return {
        "clientId": "test_client_123",
        "documentUri": "file:///tmp/test.docx",
    }


@pytest.fixture
def handshake_data_invalid_version() -> dict[str, Any]:
    """oaspVersion 格式非法的握手数据（应 HANDSHAKE_FAILED）"""
    return {
        "clientId": "test_client_123",
        "documentUri": "file:///tmp/test.docx",
        "oaspVersion": "0.3",
    }


@pytest.fixture
def handshake_data_incompatible_version() -> dict[str, Any]:
    """oaspVersion 不兼容的握手数据（v0.x 严格 MINOR；应 PROTOCOL_VERSION_MISMATCH 2006）"""
    return {
        "clientId": "test_client_123",
        "documentUri": "file:///tmp/test.docx",
        "oaspVersion": "0.2.0",
    }


@pytest.fixture
def mock_session_id() -> str:
    """模拟会话 ID"""
    return "test_session_abc123"
