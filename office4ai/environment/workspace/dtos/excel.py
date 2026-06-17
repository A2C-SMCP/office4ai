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
  （#20 ``excel:get:rangeFormat`` 复用）。
- #20 Format — 格式/条件格式/合并: ``excel:get:rangeFormat`` /
  ``excel:set:rangeFormat`` / ``excel:add:conditionalFormat`` /
  ``excel:clear:conditionalFormat`` / ``excel:merge:cells`` /
  ``excel:unmerge:cells``。读侧 ``RangeFormatInfo`` 复用；写侧 ``set:rangeFormat``
  另用一套全可选偏更新模型（协议本身读/写不对称，详见下方注释）。
- #21 Worksheet — 工作表管理: ``excel:get:worksheets`` / ``excel:add:worksheet`` /
  ``excel:delete:worksheet`` / ``excel:rename:worksheet`` / ``excel:activate:worksheet``。
  ``get:worksheets`` 响应复用 #18 ``SheetInfo``；各事件请求字段形态以 spec type 为准
  （非统一的可选 ``worksheetName``，详见各 Request DTO）。
- #22 Table — 表格操作: ``excel:insert:table`` / ``excel:get:table`` /
  ``excel:get:tables`` / ``excel:add:tableRow`` / ``excel:delete:tableRow`` /
  ``excel:sort:table``。读侧两响应形态各异（``get:table`` 富 8 字段含列明细 /
  ``get:tables`` 精简 3 字段每条目），故**不**共用单一 ``TableInfo``（issue 文案
  「共享 TableInfo」与 spec type 不符 → 从 spec type）；4 个写结果亦各自建模。
- #23 Chart — 图表操作: ``excel:insert:chart`` / ``excel:get:charts`` /
  ``excel:update:chart`` / ``excel:delete:chart``。**按 spec type 实装**：Excel 图表
  以 ``sourceAddress``（工作簿范围引用）+ ``chartType``（**开放字符串**，spec 类型为
  ``string`` 而非闭合枚举）构建，**不**复用 PPT 的内联 ``ChartData`` 判别联合，也无
  ``3015 INVALID_CHART_DATA``（issue/北极星文案「复用 ChartData / 3015」描述的是
  ``data-structures.md`` 标注的「未来 Excel」设想，events-excel.md 当前 4 个事件并未
  采纳 → 从 spec type）。无效类型→AddIn 返回 4002，无效范围→5002，图表不存在→5007。
  ``insert`` 与 ``update`` 响应同为 ``{name}`` → 共用 ``ChartOperationResult``（同 #19
  ``RangeOperationResult`` 的「形态相同则共用」原则）；``ChartPosition`` 由 insert 与
  update 复用。
- #24 PivotTable — 透视表操作: ``excel:insert:pivotTable`` / ``excel:get:pivotTables`` /
  ``excel:delete:pivotTable``。**按 spec type 实装**：透视表以 ``sourceAddress`` +
  ``targetAddress`` 放置创建，spec 未定义 rows/columns/values/filters 字段或聚合函数枚举
  （issue/北极星文案「行/列/值/筛选 + 聚合枚举 / 共享 PivotTableInfo」描述的富模型，
  events-excel.md 当前 3 个事件并未采纳 → 从 spec type，避免发明协议外 surface）。
  ``insert`` 响应 ``{name}`` 与 ``get`` 条目 ``{name, id}`` 形态不同 → **不**强行共用单一
  ``PivotTableInfo``（同 #22 TableInfo / #23 ChartSummary 的「形态不同不共用」原则）。
  无效范围→AddIn 返回 5002，透视表不存在→5008，平台不支持→5010。
- 查找与筛选数据结构随子 issue (#25) 落地，避免提前过度建模。
"""

from typing import Any, ClassVar, Literal

from pydantic import ConfigDict, Field

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

    ``excel:get:range`` 在 ``includeFormat=true`` 时返回；#20 ``excel:get:rangeFormat``
    的响应 ``format`` 字段也复用本模型（OASP spec 明示「见 excel:get:range 中的
    RangeFormatInfo 定义」），故定义在 #19 Range 切片中作为共享读侧类型。
    注意：``excel:set:rangeFormat`` 写入侧**不**复用本模型（结构不对称，见
    ``SetRangeFormatOptions``）。
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
    #20 ``set:rangeFormat`` / ``add:conditionalFormat`` / ``clear:conditionalFormat`` /
    ``merge:cells`` / ``unmerge:cells`` 的响应同为 ``{address}``，亦复用本模型。
    """

    address: str = Field(..., alias="address", description="Address operated on")


# ---- #20 Format: get:rangeFormat response data -----------------------------


class GetRangeFormatData(SocketIOBaseModel):
    """
    范围格式读取结果 | ``excel:get:rangeFormat`` response data.

    ``format`` 字段复用 #19 共享读侧 ``RangeFormatInfo``。
    """

    address: str = Field(..., alias="address", description="Range address")
    format: RangeFormatInfo = Field(..., alias="format", description="Range format info")


# ---- #20 Format: set:rangeFormat write-side options (偏更新, 全可选) ---------
#
# 写入侧与读侧 ``RangeFormatInfo`` 结构不对称（协议如此），故不复用读侧模型：
#   - 含 ``borders``（读侧无）
#   - ``alignment`` 为嵌套对象（读侧是扁平 horizontalAlignment/verticalAlignment）
#   - ``numberFormat`` 为标量字符串（读侧是二维数组）
#   - 全字段可选（仅传入的属性会被修改）
# 对齐/下划线/边框枚举为开放词表（协议用「等」），故用 ``str`` 透传并保留大小写。


class BorderFormat(SocketIOBaseModel):
    """单边框样式 | One border edge style (``excel:set:rangeFormat``)."""

    style: str | None = Field(
        default=None, alias="style", description="None|Thin|Medium|Thick|Dashed|Dotted ... (open vocab)"
    )
    color: str | None = Field(default=None, alias="color", description="Border color (#RRGGBB)")
    weight: str | None = Field(default=None, alias="weight", description="Hairline|Thin|Medium|Thick")


class RangeBorders(SocketIOBaseModel):
    """四向边框 | Range border edges (``excel:set:rangeFormat``)."""

    top: BorderFormat | None = Field(default=None, alias="top", description="Top border")
    bottom: BorderFormat | None = Field(default=None, alias="bottom", description="Bottom border")
    left: BorderFormat | None = Field(default=None, alias="left", description="Left border")
    right: BorderFormat | None = Field(default=None, alias="right", description="Right border")


class SetFontFormat(SocketIOBaseModel):
    """字体偏更新 | Optional font fields (``excel:set:rangeFormat``)."""

    name: str | None = Field(default=None, alias="name", description="Font name")
    size: float | None = Field(default=None, alias="size", description="Font size in points")
    bold: bool | None = Field(default=None, alias="bold", description="Bold")
    italic: bool | None = Field(default=None, alias="italic", description="Italic")
    color: str | None = Field(default=None, alias="color", description="Font color (#RRGGBB)")
    underline: str | None = Field(default=None, alias="underline", description="None|Single|Double ... (open vocab)")


class SetFillFormat(SocketIOBaseModel):
    """填充偏更新 | Optional fill fields (``excel:set:rangeFormat``)."""

    color: str | None = Field(default=None, alias="color", description="Fill color (#RRGGBB)")


class SetAlignmentFormat(SocketIOBaseModel):
    """对齐偏更新 | Optional alignment fields (``excel:set:rangeFormat``)."""

    horizontal: str | None = Field(
        default=None, alias="horizontal", description="Left|Center|Right|Fill|Justify|General ... (open vocab)"
    )
    vertical: str | None = Field(
        default=None, alias="vertical", description="Top|Center|Bottom|Justify ... (open vocab)"
    )
    wrap_text: bool | None = Field(default=None, alias="wrapText", description="Wrap text")
    indent_level: int | None = Field(default=None, alias="indentLevel", description="Indent level (0-based)")
    text_orientation: int | None = Field(default=None, alias="textOrientation", description="Text orientation angle")


class SetRangeFormatOptions(SocketIOBaseModel):
    """
    范围格式偏更新载荷 | ``excel:set:rangeFormat`` ``format`` payload (all optional).

    仅传入的属性会被修改（偏更新语义）。结构与读侧 ``RangeFormatInfo`` 不同。
    """

    font: SetFontFormat | None = Field(default=None, alias="font", description="Font fields")
    fill: SetFillFormat | None = Field(default=None, alias="fill", description="Fill fields")
    borders: RangeBorders | None = Field(default=None, alias="borders", description="Border edges")
    alignment: SetAlignmentFormat | None = Field(default=None, alias="alignment", description="Alignment fields")
    number_format: str | None = Field(
        default=None, alias="numberFormat", description="Number format string, e.g. '0.00', 'yyyy-mm-dd'"
    )


class ConditionalFormatRule(SocketIOBaseModel):
    """
    条件格式规则（透传）| Conditional-format rule, passthrough.

    协议为透传模式：必填 ``type`` + 任意附加参数键（按 type 不同而不同）。
    ``extra="allow"`` 保留附加键，``model_dump(by_alias=True, exclude_none=True)``
    原样透传。``type`` 为开放字符串（如 cellValue/colorScale/dataBar/iconSet）。
    """

    model_config: ClassVar[ConfigDict] = ConfigDict(populate_by_name=True, extra="allow")

    type: str = Field(..., alias="type", description="Rule type, e.g. cellValue|colorScale|dataBar|iconSet")


# ---- #21 Worksheet: response data models -----------------------------------
#
# 各写事件响应形态互不相同（与 #19/#20 的统一 ``{address}`` 不同），故不复用
# ``RangeOperationResult``，按 spec 分别建模。


class GetWorksheetsData(SocketIOBaseModel):
    """工作表列表 | ``excel:get:worksheets`` response data."""

    worksheets: list[SheetInfo] = Field(..., alias="worksheets", description="All worksheets in the workbook")


class AddWorksheetData(SocketIOBaseModel):
    """新增工作表结果 | ``excel:add:worksheet`` response data."""

    name: str = Field(..., alias="name", description="Actual name of the new worksheet")
    index: int = Field(..., alias="index", description="Index of the new worksheet")


class DeleteWorksheetData(SocketIOBaseModel):
    """删除工作表结果 | ``excel:delete:worksheet`` response data."""

    deleted: bool = Field(..., alias="deleted", description="Whether the worksheet was deleted")


class RenameWorksheetData(SocketIOBaseModel):
    """重命名工作表结果 | ``excel:rename:worksheet`` response data."""

    name: str = Field(..., alias="name", description="Name after rename")


class ActivateWorksheetData(SocketIOBaseModel):
    """激活工作表结果 | ``excel:activate:worksheet`` response data."""

    activated: bool = Field(..., alias="activated", description="Whether the worksheet was activated")


# ---- #22 Table: nested + response data models ------------------------------
#
# 读侧两形态各异，故不共用单一 TableInfo（spec type 优先于 issue 文案）：
#   get:table  — 富信息（含 columns 列明细 / styleName / showHeaders 等 8 字段）
#   get:tables — 每条目仅 {name, id, address}（精简）
# 4 个写结果（insert/add/delete/sort）形态互不相同，各自建模。


class TableColumnInfo(SocketIOBaseModel):
    """表格列条目 | Table column entry (part of ``excel:get:table`` response)."""

    name: str = Field(..., alias="name", description="Column name")
    index: int = Field(..., alias="index", description="Column index (0-based)")


class TableSummary(SocketIOBaseModel):
    """表格概要条目 | Table summary entry (``excel:get:tables`` response)."""

    name: str = Field(..., alias="name", description="Table name")
    id: str = Field(..., alias="id", description="Table ID")
    address: str = Field(..., alias="address", description="Table range address")


class InsertTableData(SocketIOBaseModel):
    """新建表格结果 | ``excel:insert:table`` response data."""

    name: str = Field(..., alias="name", description="Auto-generated table name (e.g. 'Table1')")
    address: str = Field(..., alias="address", description="Actual table range address")


class GetTableData(SocketIOBaseModel):
    """表格详细信息 | ``excel:get:table`` response data (rich)."""

    name: str = Field(..., alias="name", description="Table name")
    id: str = Field(..., alias="id", description="Table ID")
    address: str = Field(..., alias="address", description="Table range address")
    row_count: int = Field(..., alias="rowCount", description="Data row count (excluding header)")
    column_count: int = Field(..., alias="columnCount", description="Column count")
    columns: list[TableColumnInfo] = Field(..., alias="columns", description="Column entries")
    style_name: str = Field(..., alias="styleName", description="Table style name")
    show_headers: bool = Field(..., alias="showHeaders", description="Whether the header row is shown")


class GetTablesData(SocketIOBaseModel):
    """表格列表 | ``excel:get:tables`` response data (slim entries)."""

    tables: list[TableSummary] = Field(..., alias="tables", description="Table summaries in the worksheet")


class AddTableRowData(SocketIOBaseModel):
    """追加表格行结果 | ``excel:add:tableRow`` response data."""

    table_id: str = Field(..., alias="tableId", description="Table the row was appended to")


class DeleteTableRowData(SocketIOBaseModel):
    """删除表格行结果 | ``excel:delete:tableRow`` response data."""

    deleted: bool = Field(..., alias="deleted", description="Whether the row was deleted")


class SortTableData(SocketIOBaseModel):
    """表格排序结果 | ``excel:sort:table`` response data."""

    sorted: bool = Field(..., alias="sorted", description="Whether the table was sorted")


# ---- #23 Chart: nested + response data models ------------------------------
#
# 按 spec type 实装（events-excel.md §2023-2333）：Excel 图表以 sourceAddress +
# chartType(开放字符串) 构建，不内联 ChartData，无 3015。详见模块 docstring §#23。
#   insert:chart  — {name}           ┐ 形态相同 → 共用 ChartOperationResult
#   update:chart  — {name}           ┘ （同 #19 RangeOperationResult 原则）
#   get:charts    — {charts: [{name, chartType, title, top, left, width, height}]}
#   delete:chart  — {deleted}
# ChartPosition 由 insert 请求与 update properties 复用。
# 错误码 4002 INVALID_PARAM / 5001 WORKSHEET_NOT_FOUND / 5002 RANGE_INVALID /
# 5007 CHART_NOT_FOUND 均由 AddIn 产生，office4ai 透传 obs.error，不在此定义。

# Excel 常见图表类型（chartType 开放字符串的常见取值，与 Office.js Excel.ChartType 对齐）。
# 注意：Excel 散点图为 'XYScatter'（≠ PPT ChartType 的 'Scatter'），故各命名空间枚举不同，
# 不复用 PPT 的闭合 ChartType Literal——保持开放 str 以接纳 AddIn 支持的其余类型。
COMMON_CHART_TYPES = (
    "ColumnClustered | ColumnStacked | BarClustered | Line | LineMarkers | Pie | Doughnut | Area | XYScatter | Radar"
)


class ChartPosition(SocketIOBaseModel):
    """图表位置与尺寸 | Chart position/size in points (insert/update, all optional)."""

    top: float | None = Field(default=None, alias="top", description="Top position in points")
    left: float | None = Field(default=None, alias="left", description="Left position in points")
    width: float | None = Field(default=None, alias="width", description="Width in points")
    height: float | None = Field(default=None, alias="height", description="Height in points")


class ChartUpdateProperties(SocketIOBaseModel):
    """
    图表属性偏更新 | ``excel:update:chart`` ``properties`` payload (all optional).

    仅传入的属性会被修改（偏更新语义）。``chart_type`` 为开放字符串（见 COMMON_CHART_TYPES）。
    """

    title: str | None = Field(default=None, alias="title", description="New chart title")
    chart_type: str | None = Field(
        default=None, alias="chartType", description=f"New chart type (open string; common: {COMMON_CHART_TYPES})"
    )
    source_address: str | None = Field(
        default=None, alias="sourceAddress", description="New data source range, e.g. 'A1:C4'"
    )
    position: ChartPosition | None = Field(default=None, alias="position", description="New position/size")


class ChartSummary(SocketIOBaseModel):
    """图表概要条目 | Chart summary entry (``excel:get:charts`` response)."""

    name: str = Field(..., alias="name", description="Chart name")
    chart_type: str = Field(..., alias="chartType", description="Chart type")
    title: str = Field(..., alias="title", description="Chart title")
    top: float = Field(..., alias="top", description="Top position in points")
    left: float = Field(..., alias="left", description="Left position in points")
    width: float = Field(..., alias="width", description="Width in points")
    height: float = Field(..., alias="height", description="Height in points")


class ChartOperationResult(SocketIOBaseModel):
    """
    图表写操作结果 | Shared write-op response data.

    ``excel:insert:chart`` 与 ``excel:update:chart`` 的响应同为 ``{name}``，共用本模型
    （同 #19 ``RangeOperationResult`` 的「形态相同则共用」原则）。
    """

    name: str = Field(..., alias="name", description="Chart name")


class GetChartsData(SocketIOBaseModel):
    """图表列表 | ``excel:get:charts`` response data."""

    charts: list[ChartSummary] = Field(..., alias="charts", description="Chart summaries in the worksheet")


class DeleteChartData(SocketIOBaseModel):
    """删除图表结果 | ``excel:delete:chart`` response data."""

    deleted: bool = Field(..., alias="deleted", description="Whether the chart was deleted")


# ---- #24 PivotTable: response data models ----------------------------------
#
# 按 spec type 实装（events-excel.md §2335-2533）：透视表以 sourceAddress +
# targetAddress 放置创建，spec 未定义 rows/columns/values/filters 字段或聚合枚举。
#   insert:pivotTable  — {name}
#   get:pivotTables    — {pivotTables: [{name, id}]}
#   delete:pivotTable  — {deleted}
# insert 响应 {name} 与 get 条目 {name, id} 形态不同 → 不强行共用单一 PivotTableInfo
# （同 #22 TableInfo / #23 ChartSummary 的「形态不同不共用」原则）。
# 错误码 5001 WORKSHEET_NOT_FOUND / 5002 RANGE_INVALID / 5008 PIVOT_NOT_FOUND /
# 5010 NOT_SUPPORTED 均由 AddIn 产生，office4ai 透传 obs.error，不在此定义。


class PivotTableOperationResult(SocketIOBaseModel):
    """透视表写操作结果 | ``excel:insert:pivotTable`` response data (``{name}``)."""

    name: str = Field(..., alias="name", description="Pivot table name")


class PivotTableSummary(SocketIOBaseModel):
    """透视表概要条目 | Pivot-table summary entry (``excel:get:pivotTables`` response)."""

    name: str = Field(..., alias="name", description="Pivot table name")
    id: str = Field(..., alias="id", description="Pivot table ID")


class GetPivotTablesData(SocketIOBaseModel):
    """透视表列表 | ``excel:get:pivotTables`` response data."""

    pivot_tables: list[PivotTableSummary] = Field(
        ..., alias="pivotTables", description="Pivot-table summaries in the worksheet"
    )


class DeletePivotTableData(SocketIOBaseModel):
    """删除透视表结果 | ``excel:delete:pivotTable`` response data."""

    deleted: bool = Field(..., alias="deleted", description="Whether the pivot table was deleted")


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


# ---- #20 Format: 格式 / 条件格式 / 合并单元格 -------------------------------


class ExcelGetRangeFormatRequest(BaseRequest):
    """Request: ``excel:get:rangeFormat`` — 获取范围完整格式信息。"""

    event_name: ClassVar[str] = "excel:get:rangeFormat"

    address: str = Field(..., alias="address", description="Range address, e.g. 'A1:C3'")
    worksheet_name: str | None = Field(
        default=None, alias="worksheetName", description="Worksheet name; omitted = active worksheet"
    )


class ExcelSetRangeFormatRequest(BaseRequest):
    """Request: ``excel:set:rangeFormat`` — 设置范围格式（偏更新，仅传入属性被改）。"""

    event_name: ClassVar[str] = "excel:set:rangeFormat"

    address: str = Field(..., alias="address", description="Range address")
    format: SetRangeFormatOptions = Field(..., alias="format", description="Optional format fields to apply")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelAddConditionalFormatRequest(BaseRequest):
    """Request: ``excel:add:conditionalFormat`` — 添加条件格式规则（规则透传）。"""

    event_name: ClassVar[str] = "excel:add:conditionalFormat"

    address: str = Field(..., alias="address", description="Range address, e.g. 'B2:B100'")
    rule: ConditionalFormatRule = Field(..., alias="rule", description="Conditional-format rule (passthrough)")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelClearConditionalFormatRequest(BaseRequest):
    """Request: ``excel:clear:conditionalFormat`` — 清除范围上的所有条件格式。"""

    event_name: ClassVar[str] = "excel:clear:conditionalFormat"

    address: str = Field(..., alias="address", description="Range address")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelMergeCellsRequest(BaseRequest):
    """Request: ``excel:merge:cells`` — 合并单元格（保留左上角值）。"""

    event_name: ClassVar[str] = "excel:merge:cells"

    address: str = Field(..., alias="address", description="Range address to merge, e.g. 'A1:C1'")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelUnmergeCellsRequest(BaseRequest):
    """Request: ``excel:unmerge:cells`` — 取消单元格合并。"""

    event_name: ClassVar[str] = "excel:unmerge:cells"

    address: str = Field(..., alias="address", description="Range address to unmerge, e.g. 'A1:C1'")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


# ---- #21 Worksheet: 工作表管理 ---------------------------------------------
#
# 字段形态以 spec type 为准，各事件不同（非 #19/#20 的统一可选 worksheetName）：
#   get:worksheets   — 无业务参数
#   add:worksheet    — name 可选（省略时 Excel 自动命名）
#   delete/activate  — worksheetName 必填
#   rename           — currentName + newName 均必填


class ExcelGetWorksheetsRequest(BaseRequest):
    """
    Request: ``excel:get:worksheets`` — 获取工作簿中所有工作表列表。

    无业务参数，仅携带 BaseRequest 三要素 (requestId / documentUri / timestamp)。
    """

    event_name: ClassVar[str] = "excel:get:worksheets"


class ExcelAddWorksheetRequest(BaseRequest):
    """Request: ``excel:add:worksheet`` — 添加新工作表（name 省略时由 Excel 自动命名）。"""

    event_name: ClassVar[str] = "excel:add:worksheet"

    name: str | None = Field(default=None, alias="name", description="New worksheet name; omitted = auto-named")


class ExcelDeleteWorksheetRequest(BaseRequest):
    """Request: ``excel:delete:worksheet`` — 删除指定工作表。"""

    event_name: ClassVar[str] = "excel:delete:worksheet"

    worksheet_name: str = Field(..., alias="worksheetName", description="Name of the worksheet to delete")


class ExcelRenameWorksheetRequest(BaseRequest):
    """Request: ``excel:rename:worksheet`` — 重命名工作表。"""

    event_name: ClassVar[str] = "excel:rename:worksheet"

    current_name: str = Field(..., alias="currentName", description="Current worksheet name")
    new_name: str = Field(..., alias="newName", description="New worksheet name")


class ExcelActivateWorksheetRequest(BaseRequest):
    """Request: ``excel:activate:worksheet`` — 激活（切换到）指定工作表。"""

    event_name: ClassVar[str] = "excel:activate:worksheet"

    worksheet_name: str = Field(..., alias="worksheetName", description="Name of the worksheet to activate")


# ---- #22 Table: 表格操作 ----------------------------------------------------
#
# 字段形态以 spec type 为准（各事件不同）：
#   insert:table     — address + hasHeaders 必填；data / styleName / worksheetName 可选
#   get:table        — tableId 必填；worksheetName 可选
#   get:tables       — 仅 worksheetName 可选
#   add:tableRow     — tableId + values 必填；worksheetName 可选
#   delete:tableRow  — tableId + rowIndex 必填；worksheetName 可选
#   sort:table       — tableId + sortFields 必填；worksheetName 可选
# 错误码 5006 TABLE_NOT_FOUND / 5009 DATA_TYPE_MISMATCH / 4004 PARAM_OUT_OF_RANGE
# 均由 AddIn 产生，office4ai 透传 obs.error，不在此定义。


class SortField(SocketIOBaseModel):
    """排序键 | One sort key (``excel:sort:table`` request, priority-ordered)."""

    column_index: int = Field(..., alias="columnIndex", description="Column index (0-based)")
    ascending: bool | None = Field(
        default=None, alias="ascending", description="Ascending order; omitted = true (AddIn default)"
    )


class ExcelInsertTableRequest(BaseRequest):
    """Request: ``excel:insert:table`` — 在指定范围创建结构化表格。"""

    event_name: ClassVar[str] = "excel:insert:table"

    address: str = Field(..., alias="address", description="Table range address, e.g. 'A1:C4'")
    has_headers: bool = Field(..., alias="hasHeaders", description="Whether the first row is a header row")
    data: list[list[Any]] | None = Field(default=None, alias="data", description="Initial data (2D array)")
    style_name: str | None = Field(
        default=None, alias="styleName", description="Table style name, e.g. 'TableStyleMedium2'"
    )
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelGetTableRequest(BaseRequest):
    """Request: ``excel:get:table`` — 获取指定表格的详细信息。"""

    event_name: ClassVar[str] = "excel:get:table"

    table_id: str = Field(..., alias="tableId", description="Table name or ID")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelGetTablesRequest(BaseRequest):
    """Request: ``excel:get:tables`` — 获取工作表中所有表格的概要列表。"""

    event_name: ClassVar[str] = "excel:get:tables"

    worksheet_name: str | None = Field(
        default=None, alias="worksheetName", description="Worksheet name; omitted = active worksheet"
    )


class ExcelAddTableRowRequest(BaseRequest):
    """Request: ``excel:add:tableRow`` — 向表格末尾追加一行数据。"""

    event_name: ClassVar[str] = "excel:add:tableRow"

    table_id: str = Field(..., alias="tableId", description="Table name or ID")
    values: list[Any] = Field(..., alias="values", description="Row values (1D array, in column order)")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelDeleteTableRowRequest(BaseRequest):
    """Request: ``excel:delete:tableRow`` — 删除表格中指定索引的行。"""

    event_name: ClassVar[str] = "excel:delete:tableRow"

    table_id: str = Field(..., alias="tableId", description="Table name or ID")
    row_index: int = Field(..., alias="rowIndex", description="Row index to delete (0-based, excluding header)")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelSortTableRequest(BaseRequest):
    """Request: ``excel:sort:table`` — 对表格按指定列多级排序。"""

    event_name: ClassVar[str] = "excel:sort:table"

    table_id: str = Field(..., alias="tableId", description="Table name or ID")
    sort_fields: list[SortField] = Field(..., alias="sortFields", description="Sort keys in priority order")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


# ---- #23 Chart: 图表操作 ----------------------------------------------------
#
# 字段形态以 spec type 为准（events-excel.md §2023-2333）：
#   insert:chart  — sourceAddress + chartType 必填；title / position / worksheetName 可选
#   get:charts    — 仅 worksheetName 可选
#   update:chart  — chartName + properties 必填（properties 内全可选，偏更新）；worksheetName 可选
#   delete:chart  — chartName 必填；worksheetName 可选
# chartType 为开放字符串（spec 类型 string，非闭合枚举）→ 用 str 透传，不锁定 Literal，
# 以免拒绝 AddIn 支持的合法 Office.js 类型；无效类型由 AddIn 返回 4002。


class ExcelInsertChartRequest(BaseRequest):
    """Request: ``excel:insert:chart`` — 根据数据范围创建图表。"""

    event_name: ClassVar[str] = "excel:insert:chart"

    source_address: str = Field(..., alias="sourceAddress", description="Data source range, e.g. 'A1:C4'")
    chart_type: str = Field(
        ..., alias="chartType", description=f"Chart type (open string; common: {COMMON_CHART_TYPES})"
    )
    title: str | None = Field(default=None, alias="title", description="Chart title")
    position: ChartPosition | None = Field(default=None, alias="position", description="Position/size in points")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelGetChartsRequest(BaseRequest):
    """Request: ``excel:get:charts`` — 获取工作表中所有图表的信息。"""

    event_name: ClassVar[str] = "excel:get:charts"

    worksheet_name: str | None = Field(
        default=None, alias="worksheetName", description="Worksheet name; omitted = active worksheet"
    )


class ExcelUpdateChartRequest(BaseRequest):
    """Request: ``excel:update:chart`` — 更新图表属性（偏更新，仅传入属性被改）。"""

    event_name: ClassVar[str] = "excel:update:chart"

    chart_name: str = Field(..., alias="chartName", description="Chart name")
    properties: ChartUpdateProperties = Field(
        ..., alias="properties", description="Properties to update (all optional)"
    )
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelDeleteChartRequest(BaseRequest):
    """Request: ``excel:delete:chart`` — 删除指定图表。"""

    event_name: ClassVar[str] = "excel:delete:chart"

    chart_name: str = Field(..., alias="chartName", description="Chart name")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


# ---- #24 PivotTable: 透视表操作 --------------------------------------------
#
# 字段形态以 spec type 为准（events-excel.md §2335-2533）：
#   insert:pivotTable  — sourceAddress + targetAddress 必填；name / worksheetName 可选
#   get:pivotTables    — 仅 worksheetName 可选
#   delete:pivotTable  — pivotTableName 必填；worksheetName 可选
# spec 未定义 rows/columns/values/filters 与聚合枚举 → 不建相关字段，避免发明协议外
# surface（参见 #18 namespace 越界教训）。无效范围→5002 / 透视表不存在→5008 /
# 平台不支持→5010 由 AddIn 产生，office4ai 透传。


class ExcelInsertPivotTableRequest(BaseRequest):
    """Request: ``excel:insert:pivotTable`` — 基于数据源范围创建透视表。"""

    event_name: ClassVar[str] = "excel:insert:pivotTable"

    source_address: str = Field(..., alias="sourceAddress", description="Data source range, e.g. 'A1:D100'")
    target_address: str = Field(
        ..., alias="targetAddress", description="Top-left cell to place the pivot table, e.g. 'F1'"
    )
    name: str | None = Field(default=None, alias="name", description="Pivot table name; omitted = auto-generated")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")


class ExcelGetPivotTablesRequest(BaseRequest):
    """Request: ``excel:get:pivotTables`` — 获取工作表中所有透视表列表。"""

    event_name: ClassVar[str] = "excel:get:pivotTables"

    worksheet_name: str | None = Field(
        default=None, alias="worksheetName", description="Worksheet name; omitted = active worksheet"
    )


class ExcelDeletePivotTableRequest(BaseRequest):
    """Request: ``excel:delete:pivotTable`` — 删除指定透视表。"""

    event_name: ClassVar[str] = "excel:delete:pivotTable"

    pivot_table_name: str = Field(..., alias="pivotTableName", description="Pivot table name")
    worksheet_name: str | None = Field(default=None, alias="worksheetName", description="Worksheet name")
