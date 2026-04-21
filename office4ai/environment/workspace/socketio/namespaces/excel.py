"""
Excel Namespace

Handles Excel-specific Socket.IO events.

Architecture Note:
    Server → Client commands use OfficeWorkspace.emit_to_document() with sio.call(),
    which returns results directly via ack mechanism. No server-side handlers needed
    for command responses.

    This namespace only handles Client → Server event reports (fire-and-forget).

Status:
    握手阶段已实现（继承 BaseNamespace.on_connect / on_disconnect）。
    Excel Add-In 可正常完成 Socket.IO 握手，获得 connection:established 回执。
    Client → Server 事件（excel:event:*）和 Server → Client 命令（excel:get:* /
    excel:set:* 等）尚未实装，将在后续迭代中逐步补充。
"""

import logging

from .base import BaseNamespace

logger = logging.getLogger(__name__)


class ExcelNamespace(BaseNamespace):
    """
    Excel namespace (/excel) for Excel Add-In communication.

    当前仅实现握手（on_connect / on_disconnect 由 BaseNamespace 提供），
    允许 Excel Add-In 建立连接并注册到 connection_manager。

    Client → Server 事件（excel:event:*）的 handler 将在后续迭代中补充，
    例如：
    - excel:event:selectionChanged
    - excel:event:worksheetActivated
    - excel:event:workbookModified

    Server → Client 命令（excel:get:*, excel:set:*, excel:insert:* 等）通过
    OfficeWorkspace.emit_to_document() 使用 sio.call() 直接 RPC 调用，
    返回值通过 ack 机制获取，不需要在此处注册 handler。
    """

    def __init__(self) -> None:
        super().__init__("/excel")
        logger.info("ExcelNamespace initialized (handshake only, events pending)")
