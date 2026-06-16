"""
Excel Socket.IO DTOs

Defines data structures for Excel-specific Socket.IO events (``/excel`` namespace).

约定 (CLAUDE.md §DTO 命名规范 / OASP 0.3.0):
- Python 字段名: snake_case (PEP 8)
- Wire format alias: camelCase (OASP 协议)
- 所有模型继承 ``SocketIOBaseModel`` (``populate_by_name=True``)，输入同时接受
  snake_case 与 camelCase；``model_dump(by_alias=True)`` 始终输出 camelCase。
- Request DTO 继承 ``BaseRequest`` 并声明 ``event_name`` ClassVar，定义时自动
  注册到全局 ``request_registry``。

范围 (OASP 0.3.0 ``events-excel.md``, milestone #3 / issue #18 Foundation):
- 仅状态感知读事件: ``excel:get:workbookInfo`` / ``excel:get:worksheetInfo`` /
  ``excel:get:selectedRange``。
- Range / 格式 / 表格 / 图表等数据结构随各自子 issue (#19–#25) 落地，避免提前过度建模。
"""

from typing import Any, ClassVar

from pydantic import Field

from .common import BaseRequest, SocketIOBaseModel

# ============================================================================
# Shared data models (response payloads)
# ============================================================================


class SheetInfo(SocketIOBaseModel):
    """
    工作表概要条目 | Worksheet summary entry.

    复用于 ``excel:get:workbookInfo`` 与 ``excel:get:worksheets`` (#21)。
    """

    name: str = Field(..., alias="name", description="Worksheet name")
    index: int = Field(..., alias="index", description="Worksheet index (0-based)")
    is_active: bool = Field(..., alias="isActive", description="Whether this is the active worksheet")
    is_hidden: bool = Field(..., alias="isHidden", description="Whether the worksheet is hidden")


class UsedRangeInfo(SocketIOBaseModel):
    """工作表已使用范围概要 | Used-range summary (``excel:get:worksheetInfo``)."""

    address: str = Field(..., alias="address", description="Used range address (e.g. 'Sheet1!A1:D10')")
    row_count: int = Field(..., alias="rowCount", description="Used range row count")
    column_count: int = Field(..., alias="columnCount", description="Used range column count")


class WorkbookInfo(SocketIOBaseModel):
    """工作簿信息 | Workbook info payload (``excel:get:workbookInfo`` response data)."""

    sheets: list[SheetInfo] = Field(..., alias="sheets", description="All worksheets in the workbook")
    active_sheet: str = Field(..., alias="activeSheet", description="Active worksheet name")
    file_name: str = Field(..., alias="fileName", description="Workbook file name")


class WorksheetInfo(SocketIOBaseModel):
    """工作表详细信息 | Worksheet detail payload (``excel:get:worksheetInfo`` response data)."""

    name: str = Field(..., alias="name", description="Worksheet name")
    used_range: UsedRangeInfo = Field(..., alias="usedRange", description="Used range summary")
    table_count: int = Field(..., alias="tableCount", description="Number of structured tables")
    chart_count: int = Field(..., alias="chartCount", description="Number of charts")


class SelectedRangeInfo(SocketIOBaseModel):
    """选中范围信息 | Selected range payload (``excel:get:selectedRange`` response data)."""

    address: str = Field(..., alias="address", description="Selected range address (e.g. 'Sheet1!A1:C3')")
    values: list[list[Any]] = Field(..., alias="values", description="2D value array, row-major")
    row_count: int = Field(..., alias="rowCount", description="Selected range row count")
    column_count: int = Field(..., alias="columnCount", description="Selected range column count")


# ============================================================================
# Request DTOs (Server → AddIn, 自动注册 via event_name)
# ============================================================================


class ExcelGetWorkbookInfoRequest(BaseRequest):
    """
    Request: ``excel:get:workbookInfo`` — 获取工作簿信息。

    无业务参数，仅携带 BaseRequest 三要素 (requestId / documentUri / timestamp)。
    """

    event_name: ClassVar[str] = "excel:get:workbookInfo"


class ExcelGetWorksheetInfoRequest(BaseRequest):
    """Request: ``excel:get:worksheetInfo`` — 获取指定工作表的详细信息。"""

    event_name: ClassVar[str] = "excel:get:worksheetInfo"

    worksheet_name: str | None = Field(
        default=None,
        alias="worksheetName",
        description="Worksheet name; omitted = active worksheet",
    )


class ExcelGetSelectedRangeRequest(BaseRequest):
    """
    Request: ``excel:get:selectedRange`` — 获取当前选中范围及其值。

    无业务参数，仅携带 BaseRequest 三要素。
    """

    event_name: ClassVar[str] = "excel:get:selectedRange"
