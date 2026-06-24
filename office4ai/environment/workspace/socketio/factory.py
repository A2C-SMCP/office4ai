"""
Socket.IO Server Factory

单一事实源：构造 AsyncServer 并注册 /word /ppt /excel 三个 namespace。

Both the production path (OfficeWorkspace.start) and the standalone debug
path (socketio/server.py) consume this factory, so namespace registration
and AsyncServer kwargs live in exactly one place.
"""

import logging

import socketio  # type: ignore[import-untyped]

from .config import SocketIOConfig, default_config
from .namespaces.excel import ExcelNamespace
from .namespaces.ppt import PptNamespace
from .namespaces.word import WordNamespace

logger = logging.getLogger(__name__)


def build_sio_server(config: SocketIOConfig = default_config) -> socketio.AsyncServer:
    """
    Build a configured Socket.IO AsyncServer with all namespaces registered.

    python-socketio 的 ping_timeout / ping_interval 期望单位是秒；
    SocketIOConfig 里保存为毫秒（与前端 JS 习惯一致），此处统一做 //1000 转换。

    Args:
        config: Socket.IO server configuration.

    Returns:
        AsyncServer with /word, /ppt, /excel namespaces registered.
    """
    sio = socketio.AsyncServer(
        async_mode="aiohttp",
        cors_allowed_origins=config.cors_allowed_origins,
        ping_timeout=config.ping_timeout // 1000,
        ping_interval=config.ping_interval // 1000,
        max_http_buffer_size=config.max_http_buffer_size,
        logger=config.logger,
        engineio_logger=config.engineio_logger,
    )

    sio.register_namespace(WordNamespace())
    sio.register_namespace(PptNamespace())
    sio.register_namespace(ExcelNamespace())

    logger.info(f"Socket.IO server built with namespaces: {', '.join(config.namespaces)}")
    return sio
