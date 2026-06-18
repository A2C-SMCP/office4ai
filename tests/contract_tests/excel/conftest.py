"""
Excel 契约测试公共 fixture

提供 ``excel_roundtrip`` —— 把「连 /excel → 注册响应 → 握手 → workspace.execute()
→ 断开」这套样板收敛成一个可复用的异步调用，让每个契约测试只聚焦于：
  1. 在 ``response_factory`` 内断言 **wire 载荷**（camelCase，验证 snake→camel 转换）；
  2. 对返回的 ``OfficeObs`` 断言 **解析结果**（result.data / result.success / result.error）。

真实的端到端路径不变：OfficeWorkspace.execute() → wrap_request（DTO 校验 + by_alias）
→ AsyncServer.call() → MockAddInClient ack。helper 不 mock 这条链上的任何一环。
"""

from __future__ import annotations

from collections.abc import Awaitable, Callable
from typing import Any

import pytest

from office4ai.environment.workspace.base import OfficeAction, OfficeObs
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

SERVER_URL = "http://127.0.0.1:3003"


@pytest.fixture
def excel_roundtrip(
    workspace: OfficeWorkspace,
    mock_word_client_factory: Any,
) -> Callable[..., Awaitable[tuple[OfficeObs, list[tuple[str, dict]]]]]:
    """返回一个异步 helper，执行一次 /excel 事件的完整 RPC round-trip。

    Args (helper 调用时):
        action_name: 不含 ``excel:`` 前缀的动作名（如 ``"get:range"``）。
        params: snake_case 业务参数（无需带 ``document_uri``，helper 自动注入）。
        response_factory: 接收 wire 请求 dict、返回 ack 响应信封的工厂函数。
        document_uri / client_id: 可选，默认每次复用同一测试 URI（顺序执行安全）。

    Returns:
        (result, received_events) —— OfficeObs 与 MockAddInClient 收到的 (event, payload) 列表。
    """

    async def _run(
        action_name: str,
        params: dict[str, Any],
        response_factory: Callable[[dict], dict],
        *,
        document_uri: str = "file:///tmp/contract_test.xlsx",
        client_id: str = "contract_test_excel_client",
    ) -> tuple[OfficeObs, list[tuple[str, dict]]]:
        wire_event = f"excel:{action_name}"
        client = mock_word_client_factory(
            server_url=SERVER_URL,
            namespace="/excel",
            client_id=client_id,
            document_uri=document_uri,
        )
        # 必须在 connect() 之前注册响应（MockAddInClient 约定）。
        client.register_response(wire_event, response_factory)
        await client.connect()
        try:
            action = OfficeAction(
                category="excel",
                action_name=action_name,
                params={**params, "document_uri": document_uri},
            )
            result = await workspace.execute(action)
            return result, client.received_events
        finally:
            await client.disconnect()

    return _run
