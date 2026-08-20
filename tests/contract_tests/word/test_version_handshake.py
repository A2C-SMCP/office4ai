"""
Contract Tests for OASP 0.3.0 protocol version handshake

验证协议版本握手在**真实 Socket.IO 线缆**上的端到端行为（非 mock）：
- 兼容版本 → 放行，connection:established 携带 serverVersion；
- 不兼容版本 → connect_error，error.data 为扁平 PROTOCOL_VERSION_MISMATCH(2006)；
- 缺失版本 → connect_error，error.data 为 HANDSHAKE_FAILED(2003)。

这是唯一能证明「库真的把结构化 data 发到线上、客户端真能从 error.data 读到」的层级 ——
正是本次用两参 ConnectionRefusedError(message, data) 修正 spec 单 dict 片段的关键所在。
规范：connection.md#protocol-version-handshake。
"""

from __future__ import annotations

import asyncio
from typing import Any

import pytest
from socketio import AsyncClient, AsyncServer  # type: ignore[import-untyped]
from socketio.exceptions import ConnectionError as SioConnectionError  # type: ignore[import-untyped]

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.socketio.versioning import OASP_PROTOCOL_VERSION

SERVER_URL = "http://127.0.0.1:3003"
NAMESPACE = "/word"


@pytest.mark.asyncio
@pytest.mark.contract
async def test_handshake_compatible_version_established_carries_server_version(
    contract_test_server: AsyncServer,
) -> None:
    """兼容版本：连接成功，connection:established 携带 spec 形状 + serverVersion"""
    client = AsyncClient()
    established: list[dict[str, Any]] = []

    @client.on("connection:established", namespace=NAMESPACE)  # type: ignore[misc]
    def _on_established(data: dict[str, Any]) -> None:
        established.append(data)

    try:
        await client.connect(
            SERVER_URL,
            transports=["websocket"],
            namespaces=[NAMESPACE],
            auth={
                "clientId": "contract_handshake_ok",
                "documentUri": "file:///tmp/handshake_ok.docx",
                "oaspVersion": str(OASP_PROTOCOL_VERSION),
            },
        )
        assert client.connected is True
        await asyncio.sleep(0.1)  # 等 established 回执到达

        assert len(established) == 1
        payload = established[0]
        assert payload["serverVersion"] == str(OASP_PROTOCOL_VERSION)
        assert payload["socketId"]
        assert "timestamp" in payload
    finally:
        if client.connected:
            await client.disconnect()


@pytest.mark.asyncio
@pytest.mark.contract
async def test_handshake_incompatible_version_rejected_2006_on_wire(
    contract_test_server: AsyncServer,
) -> None:
    """不兼容版本：connect_error 的 error.data 为扁平 PROTOCOL_VERSION_MISMATCH(2006)"""
    client = AsyncClient()
    errors: list[dict[str, Any]] = []

    @client.on("connect_error", namespace=NAMESPACE)  # type: ignore[misc]
    def _on_connect_error(data: dict[str, Any]) -> None:
        errors.append(data)

    # v0.x 严格 MINOR：0.2.0 与 0.3.x 互拒
    with pytest.raises(SioConnectionError):
        await client.connect(
            SERVER_URL,
            transports=["websocket"],
            namespaces=[NAMESPACE],
            auth={
                "clientId": "contract_handshake_bad",
                "documentUri": "file:///tmp/handshake_bad.docx",
                "oaspVersion": "0.2.0",
            },
        )

    assert client.connected is False
    assert len(errors) == 1
    # 服务器经两参 ConnectionRefusedError 下发 {message, data}；扁平拒绝在 data 内层
    wire = errors[0]
    assert wire["message"] == "Protocol version mismatch"
    rejection = wire["data"]
    assert rejection["code"] == ErrorCode.PROTOCOL_VERSION_MISMATCH
    assert rejection["serverVersion"] == str(OASP_PROTOCOL_VERSION)
    assert rejection["clientVersion"] == "0.2.0"
    assert "minSupported" in rejection
    assert "maxSupported" in rejection


@pytest.mark.asyncio
@pytest.mark.contract
async def test_handshake_missing_version_rejected_handshake_failed_on_wire(
    contract_test_server: AsyncServer,
) -> None:
    """缺失版本：connect_error 的 error.data 为 HANDSHAKE_FAILED(2003)"""
    client = AsyncClient()
    errors: list[dict[str, Any]] = []

    @client.on("connect_error", namespace=NAMESPACE)  # type: ignore[misc]
    def _on_connect_error(data: dict[str, Any]) -> None:
        errors.append(data)

    with pytest.raises(SioConnectionError):
        await client.connect(
            SERVER_URL,
            transports=["websocket"],
            namespaces=[NAMESPACE],
            auth={
                "clientId": "contract_handshake_nover",
                "documentUri": "file:///tmp/handshake_nover.docx",
            },
        )

    assert client.connected is False
    assert len(errors) == 1
    assert errors[0]["data"]["code"] == ErrorCode.HANDSHAKE_FAILED
