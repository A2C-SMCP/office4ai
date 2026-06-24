"""
Excel Namespace

Handles Excel-specific Socket.IO events.

Architecture Note:
    Server → Client commands use OfficeWorkspace.emit_to_document() with sio.call(),
    which returns results directly via ack mechanism. No server-side handlers needed
    for command responses.

Status:
    握手由 BaseNamespace.on_connect / on_disconnect 提供，Excel Add-In 可正常完成
    Socket.IO 握手并获得 connection:established 回执。

    **无 Client → Server 事件上报 handler**：OASP 0.3.0 ``events-excel.md`` 定义的 37 个
    Excel 事件全部是 Server → AddIn（请求-响应），不存在 ``excel:event:*`` 事件报告类
    （与 /word 的 ``word:event:*``、/ppt 的 ``ppt:event:*`` 不同——那两者在协议中有定义）。
    故本命名空间不实现任何事件上报 handler；若未来确需 Excel 事件上报，须先在
    oasp-protocol 定义对应事件并评审发布，代码再跟进（协议先行）。

    Server → Client 命令（excel:get:* 等）通过 OfficeWorkspace.emit_to_document() 使用
    sio.call() 直接 RPC 调用，返回值经 ack 机制获取，不需要在此处注册 handler。
"""

import logging

from .base import BaseNamespace

logger = logging.getLogger(__name__)


class ExcelNamespace(BaseNamespace):
    """
    Excel namespace (/excel) for Excel Add-In communication.

    当前仅实现握手（on_connect / on_disconnect 由 BaseNamespace 提供），
    允许 Excel Add-In 建立连接并注册到 connection_manager。

    OASP 0.3.0 未定义任何 ``excel:event:*`` Client → Server 上报事件，故不实现事件
    handler。Server → Client 命令（excel:get:*, excel:set:*, excel:insert:* 等）通过
    OfficeWorkspace.emit_to_document() 使用 sio.call() 直接 RPC 调用，返回值通过 ack
    机制获取，不需要在此处注册 handler。
    """

    def __init__(self) -> None:
        super().__init__("/excel")
        logger.info("ExcelNamespace initialized (handshake only)")
