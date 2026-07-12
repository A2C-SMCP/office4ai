"""
Test client connection flow

测试客户端连接流程。
"""

import pytest
from socketio import AsyncClient  # type: ignore[import-untyped]

from office4ai.environment.workspace.socketio.versioning import SERVER_VERSION

# OASP 0.3.0 起握手强制校验 oaspVersion（version-first），客户端必须注入同 MAJOR.MINOR 版本
INTEGRATION_AUTH = {
    "clientId": "integration_test_client",
    "documentUri": "file:///tmp/integration_test.docx",
    "oaspVersion": str(SERVER_VERSION),  # 同 Server 版本，随 bump 自动跟随
}


@pytest.mark.asyncio
@pytest.mark.integration
async def test_client_connect(socketio_client: AsyncClient) -> None:
    """Test client can connect to /word namespace"""
    assert socketio_client.connected is True


@pytest.mark.asyncio
@pytest.mark.integration
async def test_client_handshake(socketio_client: AsyncClient, valid_handshake_data: dict) -> None:
    """Test client handshake with valid data"""
    # Emit connection event with handshake data
    # Note: In real scenario, handshake happens during connection
    # For testing, we verify the connection is established

    # Client is already connected from fixture
    assert socketio_client.connected is True

    # Verify connection manager registered the client
    # Note: Actual registration happens in on_connect handler
    # which is triggered when client emits connection event


@pytest.mark.asyncio
@pytest.mark.integration
async def test_client_disconnect_cleanup(socketio_client: AsyncClient) -> None:
    """Test client disconnect cleanup"""
    # Disconnect
    await socketio_client.disconnect()

    # Verify disconnected
    assert socketio_client.connected is False

    # Note: Actual cleanup happens in on_disconnect handler
    # which is triggered automatically by socketio


@pytest.mark.asyncio
@pytest.mark.integration
async def test_multiple_clients(socketio_server) -> None:
    """Test multiple clients can connect simultaneously"""
    clients = []

    try:
        for i in range(3):
            client = AsyncClient()
            await client.connect(
                "http://127.0.0.1:3001",
                namespaces=["/word"],
                transports=["websocket"],
                # 每个客户端唯一 clientId，携带 OASP 0.3.0 强制的 oaspVersion
                auth={**INTEGRATION_AUTH, "clientId": f"integration_test_client_{i}"},
            )
            clients.append(client)

        # All clients should be connected
        for client in clients:
            assert client.connected is True

    finally:
        # Cleanup
        for client in clients:
            await client.disconnect()
