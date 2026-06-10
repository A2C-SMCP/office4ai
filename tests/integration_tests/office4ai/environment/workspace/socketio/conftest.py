"""
Socket.IO integration test fixtures

提供 Socket.IO 集成测试所需的 fixtures，包括真实服务器和客户端。
"""

from __future__ import annotations

from typing import Any

import pytest
import pytest_asyncio
from aiohttp import web
from socketio import AsyncClient, AsyncServer  # type: ignore[import-untyped]

from office4ai.environment.workspace.socketio.config import SocketIOConfig
from office4ai.environment.workspace.socketio.server import create_socketio_server


@pytest_asyncio.fixture
async def socketio_server() -> AsyncServer:
    """启动测试用 Socket.IO 服务器（端口 3001）"""
    config = SocketIOConfig(
        host="127.0.0.1",
        port=3001,  # 使用不同端口避免冲突
        logger=False,
        engineio_logger=False,
    )

    # Create Socket.IO server
    server = create_socketio_server(config)

    # Create aiohttp app
    app = web.Application()
    server.attach(app)

    # Create runner
    runner = web.AppRunner(app)
    await runner.setup()

    # Start server
    site = web.TCPSite(runner, "127.0.0.1", 3001)
    await site.start()

    yield server

    # Cleanup
    await runner.cleanup()


# OASP 0.3.0 起 Server 在握手阶段强制校验 oaspVersion（version-first），
# 集成测试客户端必须在 auth 中注入与 Server 同 MAJOR.MINOR 的版本，否则连接被拒。
INTEGRATION_AUTH: dict[str, Any] = {
    "clientId": "integration_test_client",
    "documentUri": "file:///tmp/integration_test.docx",
    "oaspVersion": "0.3.0",
}


@pytest_asyncio.fixture
async def socketio_client(socketio_server: AsyncServer) -> AsyncClient:
    """创建连接到测试服务器的客户端（携带 OASP 0.3.0 强制的 oaspVersion）"""
    client = AsyncClient()

    await client.connect(
        "http://127.0.0.1:3001",
        transports=["websocket"],
        namespaces=["/word"],
        auth=INTEGRATION_AUTH,
    )

    yield client

    await client.disconnect()


@pytest.fixture
def valid_handshake_data() -> dict[str, Any]:
    """有效的握手数据（含 OASP 0.3.0 强制的 oaspVersion）"""
    return dict(INTEGRATION_AUTH)
