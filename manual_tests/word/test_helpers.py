"""
共享测试辅助函数

提供所有 manual_tests 的通用工具函数，避免代码重复。

使用方式:
    from manual_tests.word.test_helpers import (
        workspace_context,
        ready_workspace,
        wait_for_connection,
        get_document_uri,
        select_text,
        insert_text,
        replace_selection,
        replace_text,
    )
"""

import asyncio
from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from typing import Any, Literal, overload

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

# ==============================================================================
# 日志工具
# ==============================================================================


def _log(emoji: str, message: str) -> None:
    """统一的日志格式"""
    print(f"\n{emoji} {message}")


# ==============================================================================
# Workspace 上下文管理器
# ==============================================================================


@asynccontextmanager
async def workspace_context(host: str = "127.0.0.1", port: int = 3000) -> AsyncIterator[OfficeWorkspace]:
    """
    Workspace 上下文管理器，自动处理启动和停止

    Args:
        host: WebSocket 服务器地址
        port: WebSocket 服务器端口

    Yields:
        OfficeWorkspace: 已启动并连接的 workspace 实例
    """
    workspace = OfficeWorkspace(host=host, port=port)
    try:
        await workspace.start()
        yield workspace
    finally:
        await workspace.stop()


@asynccontextmanager
async def ready_workspace(
    host: str = "127.0.0.1",
    port: int = 3000,
    timeout: float = 30.0,
) -> AsyncIterator[tuple[OfficeWorkspace, str]]:
    """
    完整的 Workspace 上下文管理器：启动 → 等待连接 → 获取文档 URI

    Yields:
        (workspace, document_uri): 已连接的 workspace 和第一个文档 URI

    Raises:
        RuntimeError: 如果连接失败或未找到文档
    """
    async with workspace_context(host, port) as workspace:
        if not await wait_for_connection(workspace, timeout):
            raise RuntimeError("Add-In 连接失败")

        doc_uri = get_document_uri(workspace)
        if not doc_uri:
            raise RuntimeError("未找到已连接文档")

        _log("✅", f"已连接文档: {doc_uri}")
        yield workspace, doc_uri


# ==============================================================================
# 连接和文档工具
# ==============================================================================


async def wait_for_connection(workspace: OfficeWorkspace, timeout: float = 30.0) -> bool:
    """
    等待 Add-In 连接

    Args:
        workspace: Workspace 实例
        timeout: 超时时间（秒）

    Returns:
        bool: 是否成功连接
    """
    _log("⏳", "等待 Word Add-In 连接...")
    connected = await workspace.wait_for_addin_connection(timeout=timeout)
    if not connected:
        _log("❌", "超时：未检测到 Add-In 连接")
        return False
    return True


def get_document_uri(workspace: OfficeWorkspace) -> str | None:
    """
    获取已连接文档的 URI

    Args:
        workspace: Workspace 实例

    Returns:
        Optional[str]: 文档 URI，如果未找到则返回 None
    """
    documents = workspace.get_connected_documents()
    if not documents:
        _log("❌", "未找到已连接文档")
        return None
    return documents[0]


# ==============================================================================
# Word 操作函数
# ==============================================================================


async def select_text(
    workspace: OfficeWorkspace,
    document_uri: str,
    search_text: str,
    selection_mode: str = "select",
    select_index: int = 1,
    search_options: dict | None = None,
    wait_seconds: int = 3,
) -> tuple[bool, dict | None, str | None]:
    """
    执行文本选择动作

    Args:
        workspace: Office Workspace 实例
        document_uri: 目标文档 URI
        search_text: 要搜索的文本
        selection_mode: 选择模式（select/start/end）
        select_index: 要选择的匹配索引（1-based）
        search_options: 搜索选项（matchCase, matchWholeWord, matchWildcards）
        wait_seconds: 等待秒数

    Returns:
        (success, data, error): 操作是否成功、返回数据和错误消息
    """
    print(f"\n📝 搜索文本: '{search_text[:50]}{'...' if len(search_text) > 50 else ''}'")
    print(f"📝 选择模式: {selection_mode}")
    print(f"📝 选择索引: {select_index}")
    if search_options:
        print(f"📝 搜索选项: {search_options}")

    # 使用 snake_case（符合 Python 约定），DTO 系统会自动转换为 camelCase
    params: dict[str, Any] = {
        "document_uri": document_uri,
        "search_text": search_text,
        "selection_mode": selection_mode,
        "select_index": select_index,
    }

    if search_options:
        params["search_options"] = search_options

    action = OfficeAction(
        category="word",
        action_name="select:text",
        params=params,
    )

    result = await workspace.execute(action)

    if result.success:
        print("✅ 选择成功")
        print(f"   返回数据: {result.data}")
        if result.data:
            if "matchCount" in result.data:
                print(f"   匹配数量: {result.data['matchCount']}")
            if "selectedIndex" in result.data:
                print(f"   选中索引: {result.data['selectedIndex']}")
            if "selectedText" in result.data:
                print(f"   选中文本: '{result.data['selectedText']}'")
    else:
        print(f"❌ 选择失败: {result.error}")

    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)

    return result.success, result.data, result.error


async def insert_text(
    workspace: OfficeWorkspace,
    document_uri: str,
    text: str,
    location: str = "Cursor",
    format_: dict | None = None,
    wait_seconds: int = 3,
) -> bool:
    """
    执行文本插入动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        text: 要插入的文本
        location: 插入位置（Cursor/Start/End）
        format_: 文本格式（可选）
        wait_seconds: 执行前等待秒数

    Returns:
        bool: 是否成功
    """
    print(f"\n📝 插入文本: '{text[:50]}{'...' if len(text) > 50 else ''}'")
    print(f"   长度: {len(text)} 字符")
    print(f"   位置: {location}")
    print("   提示: 请将光标放在要插入文本的位置")

    await asyncio.sleep(wait_seconds)

    params: dict[str, Any] = {
        "document_uri": document_uri,
        "text": text,
        "location": location,
    }

    if format_:
        params["format"] = format_

    action = OfficeAction(
        category="word",
        action_name="insert:text",
        params=params,
    )

    result = await workspace.execute(action)

    # 验证结果
    print("\n📊 验证结果:")
    if result.success:
        print("✅ 插入成功")
        print(f"   返回数据: {result.data}")
        return True
    else:
        print(f"❌ 插入失败: {result.error}")
        return False


@overload
async def replace_selection(
    workspace: OfficeWorkspace,
    document_uri: str,
    content: dict,
    wait_seconds: int = 3,
    *,
    return_error: Literal[False] = False,
) -> bool: ...


@overload
async def replace_selection(
    workspace: OfficeWorkspace,
    document_uri: str,
    content: dict,
    wait_seconds: int = 3,
    *,
    return_error: Literal[True],
) -> tuple[bool, str | None]: ...


async def replace_selection(
    workspace: OfficeWorkspace,
    document_uri: str,
    content: dict,
    wait_seconds: int = 3,
    *,
    return_error: bool = False,
) -> bool | tuple[bool, str | None]:
    """
    执行选择替换动作

    Args:
        workspace: Workspace 实例
        document_uri: 目标文档 URI
        content: 替换内容
        wait_seconds: 执行前等待秒数
        return_error: 是否返回错误信息（默认 False，使用 keyword-only 参数）

    Returns:
        bool | tuple[bool, str | None]: 是否成功，或 (是否成功, 错误信息)
    """
    print(f"\n📝 替换选择: {content}")
    print(f"   等待 {wait_seconds} 秒...")

    # Wait for user to select text
    await asyncio.sleep(wait_seconds)

    # Create action
    action = OfficeAction(
        category="word",
        action_name="replace:selection",
        params={
            "document_uri": document_uri,
            "content": content,
        },
    )

    # Execute action
    observation = await workspace.execute(action)

    if return_error:
        if observation:
            return observation.success, observation.error
        return False, "No observation returned"
    else:
        return observation.success if observation else False


async def insert_table(
    workspace: OfficeWorkspace,
    document_uri: str,
    rows: int,
    columns: int,
    data: list[list[str]] | None = None,
    style: str | None = None,
    insert_location: str | None = None,
    wait_seconds: int = 3,
) -> tuple[bool, dict | None, str | None]:
    """插入表格。"""
    print(f"\n📝 插入表格: rows={rows}, columns={columns}, insert_location={insert_location}")
    options: dict[str, Any] = {"rows": rows, "columns": columns}
    if data is not None:
        options["data"] = data
    if style is not None:
        options["style"] = style
    if insert_location is not None:
        options["insertLocation"] = insert_location

    action = OfficeAction(
        category="word",
        action_name="insert:table",
        params={"document_uri": document_uri, "options": options},
    )
    result = await workspace.execute(action)
    if result.success:
        print(f"✅ 插入表格成功: {result.data}")
    else:
        print(f"❌ 插入表格失败: {result.error}")

    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)
    return result.success, result.data, result.error


async def merge_cells(
    workspace: OfficeWorkspace,
    document_uri: str,
    start_row_index: int,
    start_column_index: int,
    end_row_index: int,
    end_column_index: int,
    table_id: str | None = None,
    wait_seconds: int = 2,
    expect_failure: bool = False,
) -> tuple[bool, dict | None, str | None]:
    """合并表格单元格 (OASP /word Draft, v0.2.0)。

    expect_failure=True 表示负例语境：原始操作失败为预期，实时日志不打 ❌
    （✅/❌/⚠️ 保留给用例判定层，不表达正负向——由场景判定行裁决）。
    """
    print(
        f"\n📝 合并单元格: tableId={table_id} "
        f"({start_row_index},{start_column_index}) → ({end_row_index},{end_column_index})"
    )
    params: dict[str, Any] = {
        "document_uri": document_uri,
        "start_row_index": start_row_index,
        "start_column_index": start_column_index,
        "end_row_index": end_row_index,
        "end_column_index": end_column_index,
    }
    if table_id is not None:
        params["table_id"] = table_id

    action = OfficeAction(
        category="word",
        action_name="merge:cells",
        params=params,
    )
    result = await workspace.execute(action)
    if expect_failure:
        if result.success:
            print(f"📋 合并成功（负例预期外，由场景判定裁决）: {result.data}")
        else:
            print(f"📋 合并失败（负例预期内）: {result.error}")
    elif result.success:
        print(f"✅ 合并成功: {result.data}")
    else:
        print(f"❌ 合并失败: {result.error}")

    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)
    return result.success, result.data, result.error


async def update_table_cell(
    workspace: OfficeWorkspace,
    document_uri: str,
    cells: list[dict],
    table_id: str | None = None,
    wait_seconds: int = 2,
    expect_failure: bool = False,
) -> tuple[bool, dict | None, str | None]:
    """更新表格单元格的文本和/或格式 (OASP /word Draft, v0.2.0)。

    expect_failure 语义同 merge_cells：负例语境下原始操作日志不打 ❌。
    """
    print(f"\n📝 更新单元格: tableId={table_id}, cells={len(cells)}")
    params: dict[str, Any] = {"document_uri": document_uri, "cells": cells}
    if table_id is not None:
        params["table_id"] = table_id

    action = OfficeAction(
        category="word",
        action_name="update:tableCell",
        params=params,
    )
    result = await workspace.execute(action)
    if expect_failure:
        if result.success:
            print(f"📋 更新单元格成功（负例预期外，由场景判定裁决）: {result.data}")
        else:
            print(f"📋 更新单元格失败（负例预期内）: {result.error}")
    elif result.success:
        print(f"✅ 更新单元格成功: {result.data}")
    else:
        print(f"❌ 更新单元格失败: {result.error}")

    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)
    return result.success, result.data, result.error


async def update_table_row_column(
    workspace: OfficeWorkspace,
    document_uri: str,
    rows: list[dict] | None = None,
    columns: list[dict] | None = None,
    table_id: str | None = None,
    wait_seconds: int = 2,
) -> tuple[bool, dict | None, str | None]:
    """按行/列批量写入表格文本 (OASP /word Draft, v0.2.0)。"""
    print(f"\n📝 批量行/列写入: tableId={table_id}, rows={len(rows or [])}, columns={len(columns or [])}")
    params: dict[str, Any] = {"document_uri": document_uri}
    if table_id is not None:
        params["table_id"] = table_id
    if rows is not None:
        params["rows"] = rows
    if columns is not None:
        params["columns"] = columns

    action = OfficeAction(
        category="word",
        action_name="update:tableRowColumn",
        params=params,
    )
    result = await workspace.execute(action)
    if result.success:
        print(f"✅ 批量写入成功: {result.data}")
    else:
        print(f"❌ 批量写入失败: {result.error}")

    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)
    return result.success, result.data, result.error


async def update_table_format(
    workspace: OfficeWorkspace,
    document_uri: str,
    table_id: str | None = None,
    style_options: dict | None = None,
    border_options: dict | None = None,
    column_widths: list[float] | None = None,
    alignment: str | None = None,
    wait_seconds: int = 2,
) -> tuple[bool, dict | None, str | None]:
    """更新整表格式 (OASP /word Draft, v0.2.0)。"""
    print(f"\n📝 更新整表格式: tableId={table_id}, alignment={alignment}, columnWidths={column_widths}")
    params: dict[str, Any] = {"document_uri": document_uri}
    if table_id is not None:
        params["table_id"] = table_id
    if style_options is not None:
        params["style_options"] = style_options
    if border_options is not None:
        params["border_options"] = border_options
    if column_widths is not None:
        params["column_widths"] = column_widths
    if alignment is not None:
        params["alignment"] = alignment

    action = OfficeAction(
        category="word",
        action_name="update:tableFormat",
        params=params,
    )
    result = await workspace.execute(action)
    if result.success:
        print(f"✅ 更新表格格式成功: {result.data}")
    else:
        print(f"❌ 更新表格格式失败: {result.error}")

    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)
    return result.success, result.data, result.error


async def replace_text(
    workspace: OfficeWorkspace,
    document_uri: str,
    search_text: str,
    replace_text: str,
    options: dict | None = None,
    wait_seconds: int = 3,
) -> bool:
    """
    执行文本替换动作

    Args:
        workspace: Workspace 实例
        document_uri: 文档 URI
        search_text: 要搜索的文本
        replace_text: 替换文本
        options: 替换选项
        wait_seconds: 等待时间（秒）

    Returns:
        bool: 是否成功
    """
    print(f"\n📝 查找文本: '{search_text[:50]}{'...' if len(search_text) > 50 else ''}'")
    print(f"📝 替换文本: '{replace_text[:50]}{'...' if len(replace_text) > 50 else ''}'")
    if options:
        print(f"📝 选项: {options}")

    params: dict[str, Any] = {
        "document_uri": document_uri,
        "searchText": search_text,
        "replaceText": replace_text,
    }

    if options:
        params["options"] = options

    action = OfficeAction(
        category="word",
        action_name="replace:text",
        params=params,
    )

    result = await workspace.execute(action)

    if result.success:
        print("✅ 替换成功")
        print(f"   返回数据: {result.data}")
        if result.data and "replaceCount" in result.data:
            print(f"   替换次数: {result.data['replaceCount']}")
    else:
        print(f"❌ 替换失败: {result.error}")
        if result.error:
            print(f"   错误码: {result.error}")

    # 等待一段时间让用户观察结果
    print(f"\n⏳ 等待 {wait_seconds} 秒...")
    await asyncio.sleep(wait_seconds)

    return result.success
