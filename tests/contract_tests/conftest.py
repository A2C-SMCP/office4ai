"""
Contract Test Fixtures

契约测试的公共 fixtures。
"""

from __future__ import annotations

import logging

import pytest
import pytest_asyncio
from aiohttp import web
from socketio import AsyncServer  # type: ignore[import-untyped]

from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.config import SocketIOConfig
from office4ai.environment.workspace.socketio.server import create_socketio_server
from tests.contract_tests.factories.excel_factories import ExcelDataFactory
from tests.contract_tests.factories.ppt_factories import PptDataFactory
from tests.contract_tests.factories.word_factories import WordDataFactory
from tests.contract_tests.mock_addin.client import MockAddInClient

logger = logging.getLogger(__name__)


@pytest_asyncio.fixture
async def contract_test_server() -> AsyncServer:
    """
    启动契约测试用 Socket.IO 服务器（端口 3003）。

    使用与生产环境相同的服务器配置，但使用不同端口避免冲突。
    """
    config = SocketIOConfig(
        host="127.0.0.1",
        port=3003,
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
    site = web.TCPSite(runner, "127.0.0.1", 3003)
    await site.start()

    logger.info("Contract test server started on port 3003")

    yield server

    # Cleanup
    await runner.cleanup()
    logger.info("Contract test server stopped")


@pytest_asyncio.fixture
async def mock_word_client_factory() -> type[MockAddInClient]:
    """返回 MockAddInClient 类，用于在测试中创建客户端"""
    return MockAddInClient


@pytest.fixture
def word_factory() -> WordDataFactory:
    """Word 数据工厂。"""
    return WordDataFactory()


@pytest.fixture
def ppt_factory() -> PptDataFactory:
    """PPT 数据工厂。"""
    return PptDataFactory()


@pytest.fixture
def excel_factory() -> ExcelDataFactory:
    """Excel 数据工厂。"""
    return ExcelDataFactory()


@pytest_asyncio.fixture
async def workspace(contract_test_server: AsyncServer) -> OfficeWorkspace:
    """
    创建测试用 OfficeWorkspace（复用契约测试服务器）。

    注意：不启动新的服务器，而是复用 contract_test_server。
    """
    # 创建一个复用服务器的 workspace 实例
    workspace = OfficeWorkspace.__new__(OfficeWorkspace)
    workspace.host = "127.0.0.1"
    workspace.port = 3003
    workspace.use_https = False
    workspace.config = SocketIOConfig(
        host="127.0.0.1",
        port=3003,
        logger=False,
        engineio_logger=False,
    )

    # 复用 contract_test_server
    workspace.sio_server = contract_test_server
    workspace._running = True

    # Activity tracking state (normally set in __init__)
    workspace._last_activity = None
    workspace._content_cache = {}
    workspace._structure_cache = {}
    # W4b-2 (#65): update_last_activity() reads this on every successful tool.execute();
    # tool-level contract tests (e.g. run:script) go through BaseTool.execute → 需此属性存在。
    workspace._activity_callback = None

    yield workspace

    # 不需要停止服务器，contract_test_server fixture 会处理
    workspace._running = False
