"""
Excel Data Factory

生成 Excel 事件测试数据的工厂类（契约测试用）。

所有响应方法输出 **camelCase** 的 ``data`` 载荷（wire format），与 OASP 0.3.0
``events-excel.md`` 的响应类型对齐，对应 ``dtos/excel.py`` 的响应数据模型。
契约测试在 ``response_factory`` 中把这些 data 包进 ``{requestId, success, data, timestamp}``
信封后回给服务器（模拟真实 Add-In 的 ack 响应）。
"""

from __future__ import annotations

from typing import Any


class ExcelDataFactory:
    """
    Excel 事件测试数据工厂。

    覆盖 OASP 0.3.0 ``/excel`` 命名空间的 37 个事件的响应 ``data`` 载荷。
    方法按 10 个功能类分组，命名与事件对应（snake_case 方法名，camelCase 输出）。
    """

    # ------------------------------------------------------------------
    # #18 Foundation — 状态感知读
    # ------------------------------------------------------------------

    def sheet_info(
        self,
        name: str = "Sheet1",
        index: int = 0,
        is_active: bool = True,
        is_hidden: bool = False,
    ) -> dict[str, Any]:
        """生成单个工作表概要条目（``SheetInfo``，复用于 workbookInfo / get:worksheets）。"""
        return {"name": name, "index": index, "isActive": is_active, "isHidden": is_hidden}

    def workbook_info_response(
        self,
        sheets: list[dict[str, Any]] | None = None,
        active_sheet: str = "Sheet1",
        file_name: str = "test.xlsx",
    ) -> dict[str, Any]:
        """生成 ``excel:get:workbookInfo`` 响应 data。"""
        if sheets is None:
            sheets = [
                self.sheet_info("Sheet1", 0, is_active=True),
                self.sheet_info("Sheet2", 1, is_active=False),
            ]
        return {"sheets": sheets, "activeSheet": active_sheet, "fileName": file_name}

    def worksheet_info_response(
        self,
        name: str = "Sheet1",
        used_range_address: str = "Sheet1!A1:D10",
        row_count: int = 10,
        column_count: int = 4,
        table_count: int = 1,
        chart_count: int = 0,
    ) -> dict[str, Any]:
        """生成 ``excel:get:worksheetInfo`` 响应 data。"""
        return {
            "name": name,
            "usedRange": {"address": used_range_address, "rowCount": row_count, "columnCount": column_count},
            "tableCount": table_count,
            "chartCount": chart_count,
        }

    def selected_range_response(
        self,
        address: str = "Sheet1!A1:B2",
        values: list[list[Any]] | None = None,
        row_count: int | None = None,
        column_count: int | None = None,
    ) -> dict[str, Any]:
        """生成 ``excel:get:selectedRange`` 响应 data。"""
        if values is None:
            values = [["A", "B"], [1, 2]]
        if row_count is None:
            row_count = len(values)
        if column_count is None:
            column_count = len(values[0]) if values else 0
        return {"address": address, "values": values, "rowCount": row_count, "columnCount": column_count}

    # ------------------------------------------------------------------
    # #19 Range — CRUD + 公式
    # ------------------------------------------------------------------

    def get_range_response(
        self,
        address: str = "Sheet1!A1:B2",
        values: list[list[Any]] | None = None,
        row_count: int | None = None,
        column_count: int | None = None,
        include_format: bool = False,
    ) -> dict[str, Any]:
        """生成 ``excel:get:range`` 响应 data；include_format 时附带 ``RangeFormatInfo``。"""
        if values is None:
            values = [["A", "B"], [1, 2]]
        if row_count is None:
            row_count = len(values)
        if column_count is None:
            column_count = len(values[0]) if values else 0
        data: dict[str, Any] = {
            "address": address,
            "values": values,
            "rowCount": row_count,
            "columnCount": column_count,
        }
        if include_format:
            data["format"] = self.range_format_info()
        return data

    def range_format_info(self) -> dict[str, Any]:
        """生成 ``RangeFormatInfo``（读侧，get:range / get:rangeFormat 复用）。"""
        return {
            "font": {
                "name": "Calibri",
                "size": 11.0,
                "bold": False,
                "italic": False,
                "color": "#000000",
                "underline": "None",
            },
            "fill": {"color": "#FFFFFF"},
            "horizontalAlignment": "General",
            "verticalAlignment": "Bottom",
            "wrapText": False,
            "numberFormat": [["General"]],
        }

    def range_operation_response(self, address: str = "Sheet1!A1:B2") -> dict[str, Any]:
        """生成统一写操作结果 ``{address}``（set/clear/copy/delete/insert:range, set:formula,
        set:rangeFormat, add/clear:conditionalFormat, merge/unmerge:cells 共用）。"""
        return {"address": address}

    # ------------------------------------------------------------------
    # #20 Format — get:rangeFormat（其余写操作复用 range_operation_response）
    # ------------------------------------------------------------------

    def get_range_format_response(self, address: str = "Sheet1!A1:C3") -> dict[str, Any]:
        """生成 ``excel:get:rangeFormat`` 响应 data。"""
        return {"address": address, "format": self.range_format_info()}

    # ------------------------------------------------------------------
    # #21 Worksheet — 工作表管理
    # ------------------------------------------------------------------

    def get_worksheets_response(self, worksheets: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        """生成 ``excel:get:worksheets`` 响应 data。"""
        if worksheets is None:
            worksheets = [
                self.sheet_info("Sheet1", 0, is_active=True),
                self.sheet_info("Sheet2", 1, is_active=False),
            ]
        return {"worksheets": worksheets}

    def add_worksheet_response(self, name: str = "Sheet3", index: int = 2) -> dict[str, Any]:
        """生成 ``excel:add:worksheet`` 响应 data。"""
        return {"name": name, "index": index}

    def delete_worksheet_response(self, deleted: bool = True) -> dict[str, Any]:
        """生成 ``excel:delete:worksheet`` 响应 data。"""
        return {"deleted": deleted}

    def rename_worksheet_response(self, name: str = "Renamed") -> dict[str, Any]:
        """生成 ``excel:rename:worksheet`` 响应 data。"""
        return {"name": name}

    def activate_worksheet_response(self, activated: bool = True) -> dict[str, Any]:
        """生成 ``excel:activate:worksheet`` 响应 data。"""
        return {"activated": activated}

    # ------------------------------------------------------------------
    # #22 Table — 表格操作
    # ------------------------------------------------------------------

    def insert_table_response(self, name: str = "Table1", address: str = "Sheet1!A1:C4") -> dict[str, Any]:
        """生成 ``excel:insert:table`` 响应 data。"""
        return {"name": name, "address": address}

    def get_table_response(
        self,
        name: str = "Table1",
        table_id: str = "{00000000-0001}",
        address: str = "Sheet1!A1:C4",
        row_count: int = 3,
        column_count: int = 3,
        columns: list[dict[str, Any]] | None = None,
        style_name: str = "TableStyleMedium2",
        show_headers: bool = True,
    ) -> dict[str, Any]:
        """生成 ``excel:get:table`` 响应 data（富信息，含列明细）。"""
        if columns is None:
            columns = [
                {"name": "Col1", "index": 0},
                {"name": "Col2", "index": 1},
                {"name": "Col3", "index": 2},
            ]
        return {
            "name": name,
            "id": table_id,
            "address": address,
            "rowCount": row_count,
            "columnCount": column_count,
            "columns": columns,
            "styleName": style_name,
            "showHeaders": show_headers,
        }

    def get_tables_response(self, tables: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        """生成 ``excel:get:tables`` 响应 data（精简条目）。"""
        if tables is None:
            tables = [{"name": "Table1", "id": "{00000000-0001}", "address": "Sheet1!A1:C4"}]
        return {"tables": tables}

    def add_table_row_response(self, table_id: str = "Table1") -> dict[str, Any]:
        """生成 ``excel:add:tableRow`` 响应 data。"""
        return {"tableId": table_id}

    def delete_table_row_response(self, deleted: bool = True) -> dict[str, Any]:
        """生成 ``excel:delete:tableRow`` 响应 data。"""
        return {"deleted": deleted}

    def sort_table_response(self, sorted_: bool = True) -> dict[str, Any]:
        """生成 ``excel:sort:table`` 响应 data。"""
        return {"sorted": sorted_}

    # ------------------------------------------------------------------
    # #23 Chart — 图表操作
    # ------------------------------------------------------------------

    def chart_operation_response(self, name: str = "Chart1") -> dict[str, Any]:
        """生成 ``excel:insert:chart`` / ``excel:update:chart`` 响应 data（共用 ``{name}``）。"""
        return {"name": name}

    def get_charts_response(self, charts: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        """生成 ``excel:get:charts`` 响应 data。"""
        if charts is None:
            charts = [
                {
                    "name": "Chart1",
                    "chartType": "ColumnClustered",
                    "title": "Sales",
                    "top": 10.0,
                    "left": 20.0,
                    "width": 360.0,
                    "height": 240.0,
                }
            ]
        return {"charts": charts}

    def delete_chart_response(self, deleted: bool = True) -> dict[str, Any]:
        """生成 ``excel:delete:chart`` 响应 data。"""
        return {"deleted": deleted}

    # ------------------------------------------------------------------
    # #24 PivotTable — 透视表操作
    # ------------------------------------------------------------------

    def pivot_table_operation_response(self, name: str = "PivotTable1") -> dict[str, Any]:
        """生成 ``excel:insert:pivotTable`` 响应 data（``{name}``）。"""
        return {"name": name}

    def get_pivot_tables_response(self, pivot_tables: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        """生成 ``excel:get:pivotTables`` 响应 data。"""
        if pivot_tables is None:
            pivot_tables = [{"name": "PivotTable1", "id": "{00000000-0002}"}]
        return {"pivotTables": pivot_tables}

    def delete_pivot_table_response(self, deleted: bool = True) -> dict[str, Any]:
        """生成 ``excel:delete:pivotTable`` 响应 data。"""
        return {"deleted": deleted}

    # ------------------------------------------------------------------
    # #25 Find&Filter — 查找与筛选
    # ------------------------------------------------------------------

    def find_values_response(self, matches: list[dict[str, Any]] | None = None) -> dict[str, Any]:
        """生成 ``excel:find:values`` 响应 data。"""
        if matches is None:
            matches = [
                {"address": "Sheet1!A2", "value": "foo"},
                {"address": "Sheet1!C5", "value": 42},
            ]
        return {"matches": matches}

    def set_auto_filter_response(self, address: str = "Sheet1!A1:C10") -> dict[str, Any]:
        """生成 ``excel:set:autoFilter`` 响应 data。"""
        return {"address": address}

    def clear_auto_filter_response(self, cleared: bool = True) -> dict[str, Any]:
        """生成 ``excel:clear:autoFilter`` 响应 data。"""
        return {"cleared": cleared}

    # ------------------------------------------------------------------
    # 错误响应（error 信封 portion；code 取 events-excel.md 错误码 5001–5010）
    # ------------------------------------------------------------------

    def error(self, code: str = "5002", message: str = "Invalid range address") -> dict[str, Any]:
        """生成错误信封的 ``error`` 部分 ``{code, message}``。"""
        return {"code": code, "message": message}
