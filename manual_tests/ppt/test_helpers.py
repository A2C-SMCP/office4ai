"""
PPT 共享测试辅助函数

提供 PPT manual_tests 的通用工具函数，每个函数返回 tuple[bool, dict | None, str | None]。

使用方式:
    from manual_tests.ppt.test_helpers import (
        ppt_get_slide_info,
        ppt_get_current_slide_elements,
        ppt_insert_text,
        ppt_insert_shape,
        ppt_insert_table,
        ppt_insert_image,
        ppt_add_slide,
        ppt_delete_slide,
        ppt_goto_slide,
        ppt_delete_element,
        ppt_update_text_box,
        ppt_update_element,
    )

NOTE: params dict keys use snake_case matching DTO field names.
      DTOs have populate_by_name=True so both snake_case and camelCase are accepted.
      Only "document_uri" is special-cased by wrap_request() and removed before DTO validation.
"""

from typing import Any

from office4ai.environment.workspace.base import OfficeAction
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

# ==============================================================================
# 返回类型
# ==============================================================================

# (success, data, error)
PptResult = tuple[bool, dict[str, Any] | None, str | None]


# ==============================================================================
# 共享图片载荷
# ==============================================================================

#: 最小合法 1x1 灰度 TIFF (base64)——**可解码但 Office.js 不受支持**的格式，用于
#: `3007 FORMAT_NOT_SUPPORTED` 用例（issue #90 / oasp#23）。
#:
#: 关键：它必须是**真正合法的图片**（PIL 可解码），否则 Add-In 会先在 base64/解码层判
#: `4002 INVALID_PARAM`，就测不到 3007「解得开、只是这一格式不吃」这条判法了——两码的
#: 分界正是「载荷在线缆层是否合法」。放这里供 insert:image / update:image 共用，勿各抄一份。
TIFF_UNSUPPORTED = (
    "SUkqAAgAAAAIAAABAwABAAAAAQAAAAEBAwABAAAAAQAAAAIBAwABAAAACAAA"
    "AAMBAwABAAAAAQAAAAYBAwABAAAAAQAAABEBBAABAAAAbgAAABYBAwABAAAA"
    "AQAAABcBBAABAAAAAQAAAAAAAACA"
)


# ==============================================================================
# Content Retrieval
# ==============================================================================


async def ppt_get_slide_info(
    workspace: OfficeWorkspace,
    doc_uri: str,
    slide_index: int | None = None,
) -> PptResult:
    """获取幻灯片信息"""
    params: dict[str, Any] = {"document_uri": doc_uri}
    if slide_index is not None:
        params["slide_index"] = slide_index

    action = OfficeAction(category="ppt", action_name="get:slideInfo", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_get_slide_layouts(
    workspace: OfficeWorkspace,
    doc_uri: str,
    options: dict[str, Any] | None = None,
) -> PptResult:
    """获取幻灯片布局"""
    params: dict[str, Any] = {"document_uri": doc_uri}
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="get:slideLayouts", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_get_current_slide_elements(
    workspace: OfficeWorkspace,
    doc_uri: str,
) -> PptResult:
    """获取当前幻灯片元素"""
    params: dict[str, Any] = {"document_uri": doc_uri}
    action = OfficeAction(category="ppt", action_name="get:currentSlideElements", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_get_slide_elements(
    workspace: OfficeWorkspace,
    doc_uri: str,
    slide_index: int,
    options: dict[str, Any] | None = None,
) -> PptResult:
    """获取指定幻灯片的元素"""
    params: dict[str, Any] = {"document_uri": doc_uri, "slide_index": slide_index}
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="get:slideElements", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_get_slide_screenshot(
    workspace: OfficeWorkspace,
    doc_uri: str,
    slide_index: int,
    options: dict[str, Any] | None = None,
) -> PptResult:
    """获取幻灯片截图"""
    params: dict[str, Any] = {"document_uri": doc_uri, "slide_index": slide_index}
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="get:slideScreenshot", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


# ==============================================================================
# Content Insertion
# ==============================================================================


async def ppt_insert_text(
    workspace: OfficeWorkspace,
    doc_uri: str,
    text: str,
    options: dict[str, Any] | None = None,
) -> PptResult:
    """插入文本"""
    params: dict[str, Any] = {"document_uri": doc_uri, "text": text}
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="insert:text", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_insert_shape(
    workspace: OfficeWorkspace,
    doc_uri: str,
    shape_type: str,
    options: dict[str, Any] | None = None,
) -> PptResult:
    """插入形状"""
    params: dict[str, Any] = {"document_uri": doc_uri, "shape_type": shape_type}
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="insert:shape", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_insert_table(
    workspace: OfficeWorkspace,
    doc_uri: str,
    options: dict[str, Any],
) -> PptResult:
    """插入表格"""
    params: dict[str, Any] = {"document_uri": doc_uri, "options": options}
    action = OfficeAction(category="ppt", action_name="insert:table", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_insert_image(
    workspace: OfficeWorkspace,
    doc_uri: str,
    image_data: dict[str, Any],
    options: dict[str, Any] | None = None,
) -> PptResult:
    """插入图片"""
    params: dict[str, Any] = {"document_uri": doc_uri, "image": image_data}
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="insert:image", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


# ==============================================================================
# Update Operations
# ==============================================================================


async def ppt_update_text_box(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str,
    updates: dict[str, Any],
) -> PptResult:
    """更新文本框"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "element_id": element_id,
        "updates": updates,
    }
    action = OfficeAction(category="ppt", action_name="update:textBox", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_update_image(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str,
    image_data: dict[str, Any],
    options: dict[str, Any] | None = None,
) -> PptResult:
    """更新图片"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "element_id": element_id,
        "image": image_data,
    }
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="update:image", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_update_table_cell(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str,
    cells: list[dict[str, Any]],
) -> PptResult:
    """更新表格单元格"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "element_id": element_id,
        "cells": cells,
    }
    action = OfficeAction(category="ppt", action_name="update:tableCell", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_update_table_row_column(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str,
    rows: list[dict[str, Any]] | None = None,
    columns: list[dict[str, Any]] | None = None,
) -> PptResult:
    """批量更新表格行/列"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "element_id": element_id,
    }
    if rows:
        params["rows"] = rows
    if columns:
        params["columns"] = columns

    action = OfficeAction(category="ppt", action_name="update:tableRowColumn", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_update_table_format(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str,
    cell_formats: list[dict[str, Any]] | None = None,
    row_formats: list[dict[str, Any]] | None = None,
    column_formats: list[dict[str, Any]] | None = None,
) -> PptResult:
    """更新表格格式"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "element_id": element_id,
    }
    if cell_formats:
        params["cell_formats"] = cell_formats
    if row_formats:
        params["row_formats"] = row_formats
    if column_formats:
        params["column_formats"] = column_formats

    action = OfficeAction(category="ppt", action_name="update:tableFormat", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_update_element(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str,
    updates: dict[str, Any],
    slide_index: int | None = None,
) -> PptResult:
    """更新元素位置/大小/旋转"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "element_id": element_id,
        "updates": updates,
    }
    if slide_index is not None:
        params["slide_index"] = slide_index

    action = OfficeAction(category="ppt", action_name="update:element", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


# ==============================================================================
# Delete & Reorder
# ==============================================================================


async def ppt_delete_element(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str | None = None,
    element_ids: list[str] | None = None,
    slide_index: int | None = None,
) -> PptResult:
    """删除元素"""
    params: dict[str, Any] = {"document_uri": doc_uri}
    if element_id:
        params["element_id"] = element_id
    if element_ids:
        params["element_ids"] = element_ids
    if slide_index is not None:
        params["slide_index"] = slide_index

    action = OfficeAction(category="ppt", action_name="delete:element", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_reorder_element(
    workspace: OfficeWorkspace,
    doc_uri: str,
    element_id: str,
    reorder_action: str,
    slide_index: int | None = None,
) -> PptResult:
    """调整元素 z-order"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "element_id": element_id,
        "action": reorder_action,
    }
    if slide_index is not None:
        params["slide_index"] = slide_index

    action = OfficeAction(category="ppt", action_name="reorder:element", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


# ==============================================================================
# Slide Management
# ==============================================================================


async def ppt_add_slide(
    workspace: OfficeWorkspace,
    doc_uri: str,
    options: dict[str, Any] | None = None,
) -> PptResult:
    """添加幻灯片"""
    params: dict[str, Any] = {"document_uri": doc_uri}
    if options:
        params["options"] = options

    action = OfficeAction(category="ppt", action_name="add:slide", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_delete_slide(
    workspace: OfficeWorkspace,
    doc_uri: str,
    slide_index: int,
) -> PptResult:
    """删除幻灯片"""
    params: dict[str, Any] = {"document_uri": doc_uri, "slide_index": slide_index}
    action = OfficeAction(category="ppt", action_name="delete:slide", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_move_slide(
    workspace: OfficeWorkspace,
    doc_uri: str,
    from_index: int,
    to_index: int,
) -> PptResult:
    """移动幻灯片"""
    params: dict[str, Any] = {
        "document_uri": doc_uri,
        "from_index": from_index,
        "to_index": to_index,
    }
    action = OfficeAction(category="ppt", action_name="move:slide", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error


async def ppt_goto_slide(
    workspace: OfficeWorkspace,
    doc_uri: str,
    slide_index: int,
) -> PptResult:
    """跳转到指定幻灯片"""
    params: dict[str, Any] = {"document_uri": doc_uri, "slide_index": slide_index}
    action = OfficeAction(category="ppt", action_name="goto:slide", params=params)
    result = await workspace.execute(action)
    return result.success, result.data if result.data else None, result.error
