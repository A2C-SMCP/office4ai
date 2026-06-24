"""
Excel 手动测试辅助函数

提供 Workspace 上下文管理 + 一个通用的 ``excel_op`` 调用封装（驱动任意 /excel 事件，
统一日志 + 返回 (success, data, error)）。镜像 ``manual_tests/word/test_helpers.py`` 的
约定，但 Excel 37 个事件用一个 DRY 的通用封装而非 37 个逐工具函数。

使用方式:
    from manual_tests.excel.test_helpers import ready_workspace, excel_op

    async with ready_workspace() as (workspace, doc_uri):
        ok, data, err = await excel_op(workspace, doc_uri, "set:range",
                                       address="A1:B1", values=[["x", "y"]])
"""

from __future__ import annotations

import asyncio
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


def _log(emoji: str, message: str) -> None:
    print(f"\n{emoji} {message}")


# ==============================================================================
# Workspace 上下文
# ==============================================================================


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000) -> AsyncIterator[OfficeWorkspace]:
    """Workspace 上下文管理器，自动 start/stop。"""
    workspace = OfficeWorkspace(host=host, port=port)
    try:
        await workspace.start()
        yield workspace
    finally:
        await workspace.stop()


async def wait_for_connection(workspace: OfficeWorkspace, timeout: float = 30.0) -> bool:
    """等待 Excel Add-In 连接。"""
    _log("⏳", "等待 Excel Add-In 连接...")
    connected = await workspace.wait_for_addin_connection(timeout=timeout)
    if not connected:
        _log("❌", "超时：未检测到 Excel Add-In 连接")
    return connected


def get_document_uri(workspace: OfficeWorkspace) -> str | None:
    """获取首个已连接文档 URI。"""
    documents = workspace.get_connected_documents()
    if not documents:
        _log("❌", "未找到已连接文档")
        return None
    return documents[0]


@asynccontextmanager
async def ready_workspace(
    host: str = "127.0.0.1",
    port: int = 3000,
    timeout: float = 30.0,
) -> AsyncIterator[tuple[OfficeWorkspace, str]]:
    """完整上下文：启动 → 等待 Excel Add-In 连接 → 取第一个文档 URI。

    适用于「文档已由用户在 Excel 中打开」的单独场景（如手工触发的错误码验证）。
    """
    async with workspace_context(host, port) as workspace:
        if not await wait_for_connection(workspace, timeout):
            raise RuntimeError("Excel Add-In 连接失败")
        doc_uri = get_document_uri(workspace)
        if not doc_uri:
            raise RuntimeError("未找到已连接文档")
        _log("✅", f"已连接文档: {doc_uri}")
        yield workspace, doc_uri


# ==============================================================================
# 通用 /excel 事件调用封装
# ==============================================================================


async def excel_op(
    workspace: OfficeWorkspace,
    document_uri: str,
    action_name: str,
    *,
    wait: float = 1.0,
    quiet: bool = False,
    **params: Any,
) -> tuple[bool, dict | None, str | None]:
    """驱动任意 /excel 事件。

    Args:
        workspace: Workspace 实例。
        document_uri: 目标工作簿 URI。
        action_name: 不含 ``excel:`` 前缀的动作名（如 ``"set:range"``）。
        wait: 调用后等待秒数（便于在 Excel 中目测过程）。
        quiet: 安静模式，只在失败时打印。
        **params: snake_case 业务参数（无需带 document_uri，自动注入）。DTO 层会按
            populate_by_name 转 camelCase 上线。

    Returns:
        (success, data, error)。
    """
    if not quiet:
        print(f"\n▶️  excel:{action_name}  {params}")

    action = OfficeAction(
        category="excel",
        action_name=action_name,
        params={"document_uri": document_uri, **params},
    )
    result = await workspace.execute(action)

    if result.success:
        if not quiet:
            print(f"   ✅ data={result.data}")
    else:
        print(f"   ❌ excel:{action_name} 失败: {result.error}")

    if wait:
        await asyncio.sleep(wait)
    return result.success, result.data, result.error
