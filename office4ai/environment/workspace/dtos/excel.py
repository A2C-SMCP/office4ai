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

范围 (OASP 0.3.0 ``events-excel.md``, milestone #3):
- #18 Foundation — 状态感知读事件: ``excel:get:workbookInfo`` /
  ``excel:get:worksheetInfo`` / ``excel:get:selectedRange``。
- #19 Range — Range CRUD + 公式: ``excel:get:range`` / ``excel:set:range`` /
  ``excel:clear:range`` / ``excel:copy:range`` / ``excel:delete:range`` /
  ``excel:insert:range`` / ``excel:set:formula``；新增共享 ``RangeFormatInfo``
  （#20 格式类复用）。
- 格式 / 表格 / 图表等其余数据结构随各自子 issue (#20–#25) 落地，避免提前过度建模。
"""

from typing import Any, ClassVar, Literal

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


# ---- #19 Range: format models (shared, #20 格式类复用) ----------------------


class RangeFontInfo(SocketIOBaseModel):
    """单元格字体信息 | Cell font info (part of ``RangeFormatInfo``)."""

    name: str = Field(..., alias="name", description="Font name")
    size: float = Field(..., alias="size", description="Font size in points")
    bold: bool = Field(..., alias="bold", description="Bold")
    italic: bool = Field(..., alias="italic", description="Italic")
    color: str = Field(..., alias="color", description="Font color (#RRGGBB)")
    underline: str = Field(..., alias="underline", description="Underline style (e.g. 'None', 'Single')")


class RangeFillInfo(SocketIOBaseModel):
    """单元格填充信息 | Cell fill info (part of ``RangeFormatInfo``)."""

    color: str = Field(..., alias="color", description="Fill color (#RRGGBB)")


class RangeFormatInfo(SocketIOBaseModel):
    """
    范围格式信息 | Range format info.

    ``excel:get:range`` 在 ``includeFormat=true`` 时返回；#20 ``excel:get:rangeFormat`` /
    ``excel:set:rangeFormat`` 复用本模型，故定义在 #19 Range 切片中作为共享类型。
    """

    font: RangeFontInfo = Field(..., alias="font", description="Font info")
    fill: RangeFillInfo = Field(..., alias="fill", description="Fill info")
    horizontal_alignment: str = Field(..., alias="horizontalAlignment", description="Horizontal alignment")
    vertical_alignment: str = Field(..., alias="verticalAlignment", description="Vertical alignment")
    wrap_text: bool = Field(..., alias="wrapText", description="Wrap text")
    number_format: list[list[str]] = Field(..., alias="numberFormat", description="Per-cell number formats (2D array)")


class GetRangeData(SocketIOBaseModel):
    """范围读取结果 | ``excel:get:range`` response data."""

    address: str = Field(..., alias="address", description="Range address")
    values: list[list[Any]] = Field(..., alias="values", description="2D value array, row-major")
    row_count: int = Field(..., alias="rowCount", description="Range row count")
    column_count: int = Field(..., alias="columnCount", description="Range column count")
    format: RangeFormatInfo | None = Field(
        default=None, alias="format", description="Format info; only when includeFormat=true"
    )


class RangeOperationResult(SocketIOBaseModel):
    """
    范围写操作结果 | Shared write-op response data.

    ``excel:set:range`` / ``clear:range`` / ``copy:range`` / ``delete:range`` /
    ``insert:range`` / ``set:formula`` 的响应统一为 ``{address}``。
    """

    address: str = Field(..., alias="address", description="Address operated on")


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


# ---- #19 Range: CRUD + 公式 -------------------------------------------------


class ExcelGetRangeRequest(BaseRequest):
    """Request: ``excel:get:range`` — 获取指定范围的值，可选含格式。"""

    event_name: ClassVar[str] = "excel:get:range"

    address: str = Field(..., alias="address", description="Range address, e.g. 'A1:C3'")
    worksheet_name: str | None = Field(
        default=None, alias="worksheetName", description="Worksheet name; omitted = active worksheet"
    )
    include_format: bool = Field(default=False, alias="includeFormat", description="Whether to return RangeFormatInfo")


class ExcelSetRangeRequest(BaseRequest):
    """Request: ``excel:set:range`` — 设置范围的值（标量填充或二维数组）。"""

    event_name: ClassVar[str] = "excel:set:range"

    address: str = Field(..., alias="address", description="Target range address")
    values: Any = Field(..., alias="values", description="Scalar (fills range) or 2D array")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelClearRangeRequest(BaseRequest):
    """Request: ``excel:clear:range`` — 清除范围的内容/格式/全部。"""

    event_name: ClassVar[str] = "excel:clear:range"

    address: str = Field(..., alias="address", description="Range address")
    clear_type: Literal["contents", "formats", "all"] = Field(
        ..., alias="clearType", description="What to clear: contents | formats | all"
    )
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelCopyRangeRequest(BaseRequest):
    """Request: ``excel:copy:range`` — 复制源范围到目标位置。"""

    event_name: ClassVar[str] = "excel:copy:range"

    source_address: str = Field(..., alias="sourceAddress", description="Source range address")
    target_address: str = Field(..., alias="targetAddress", description="Target range address")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelDeleteRangeRequest(BaseRequest):
    """Request: ``excel:delete:range`` — 删除范围并移动周围单元格。"""

    event_name: ClassVar[str] = "excel:delete:range"

    address: str = Field(..., alias="address", description="Range address to delete")
    shift_direction: Literal["up", "left"] = Field(
        ..., alias="shiftDirection", description="Shift surrounding cells: up | left"
    )
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelInsertRangeRequest(BaseRequest):
    """Request: ``excel:insert:range`` — 插入空白单元格并移动现有单元格。"""

    event_name: ClassVar[str] = "excel:insert:range"

    address: str = Field(..., alias="address", description="Range address to insert at")
    shift_direction: Literal["down", "right"] = Field(
        ..., alias="shiftDirection", description="Shift existing cells: down | right"
    )
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelSetFormulaRequest(BaseRequest):
    """Request: ``excel:set:formula`` — 设置单元格公式（透传字符串）。"""

    event_name: ClassVar[str] = "excel:set:formula"

    address: str = Field(..., alias="address", description="Target cell address")
    formula: str = Field(..., alias="formula", description="Formula string including '=' (e.g. '=SUM(A1:A10)')")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")
