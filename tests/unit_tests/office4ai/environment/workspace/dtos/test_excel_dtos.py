"""
Test Excel DTOs

测试 Excel 事件的数据传输对象 (OASP 0.3.0 events-excel.md, issue #18 Foundation)。

覆盖:
- #18 读事件 + #19 Range/公式 + #20 Format/条件格式/合并 + #21 Worksheet 管理
  + #22 Table 操作 Request DTO 的构造、必填校验、event_name、camelCase 序列化、
  枚举约束、注册。
- 共享数据模型 (SheetInfo / UsedRangeInfo / WorkbookInfo / WorksheetInfo /
  SelectedRangeInfo / RangeFormatInfo / GetRangeData / RangeOperationResult /
  GetRangeFormatData / GetWorksheetsData / AddWorksheetData / DeleteWorksheetData /
  RenameWorksheetData / ActivateWorksheetData / TableColumnInfo / TableSummary /
  InsertTableData / GetTableData / GetTablesData / AddTableRowData /
  DeleteTableRowData / SortTableData) 的 snake_case ↔ camelCase 双向兼容。
- #20 写侧偏更新模型 (SetRangeFormatOptions) 与条件格式透传模型
  (ConditionalFormatRule) —— 读/写不对称、None 剔除、extra 透传。
- #21 工作表事件字段形态各异 (无统一 worksheetName)，按 spec type 钉死必填/可选。
- #22 Table 读侧两形态各异 (get:table 富 / get:tables 精简)，不共用单一 TableInfo；
  写结果各自建模；sort:table 含嵌套 SortField 列表 (ascending 可选)。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import request_registry
from office4ai.environment.workspace.dtos.excel import (
    ActivateWorksheetData,
    AddTableRowData,
    AddWorksheetData,
    ChartOperationResult,
    ChartPosition,
    ChartSummary,
    ChartUpdateProperties,
    ConditionalFormatRule,
    DeleteChartData,
    DeletePivotTableData,
    DeleteTableRowData,
    DeleteWorksheetData,
    ExcelActivateWorksheetRequest,
    ExcelAddConditionalFormatRequest,
    ExcelAddTableRowRequest,
    ExcelAddWorksheetRequest,
    ExcelClearConditionalFormatRequest,
    ExcelClearRangeRequest,
    ExcelCopyRangeRequest,
    ExcelDeleteChartRequest,
    ExcelDeletePivotTableRequest,
    ExcelDeleteRangeRequest,
    ExcelDeleteTableRowRequest,
    ExcelDeleteWorksheetRequest,
    ExcelGetChartsRequest,
    ExcelGetPivotTablesRequest,
    ExcelGetRangeFormatRequest,
    ExcelGetRangeRequest,
    ExcelGetSelectedRangeRequest,
    ExcelGetTableRequest,
    ExcelGetTablesRequest,
    ExcelGetWorkbookInfoRequest,
    ExcelGetWorksheetInfoRequest,
    ExcelGetWorksheetsRequest,
    ExcelInsertChartRequest,
    ExcelInsertPivotTableRequest,
    ExcelInsertRangeRequest,
    ExcelInsertTableRequest,
    ExcelMergeCellsRequest,
    ExcelRenameWorksheetRequest,
    ExcelSetFormulaRequest,
    ExcelSetRangeFormatRequest,
    ExcelSetRangeRequest,
    ExcelSortTableRequest,
    ExcelUnmergeCellsRequest,
    ExcelUpdateChartRequest,
    GetChartsData,
    GetPivotTablesData,
    GetRangeData,
    GetRangeFormatData,
    GetTableData,
    GetTablesData,
    GetWorksheetsData,
    InsertTableData,
    PivotTableOperationResult,
    PivotTableSummary,
    RangeFormatInfo,
    RangeOperationResult,
    RenameWorksheetData,
    SelectedRangeInfo,
    SetRangeFormatOptions,
    SheetInfo,
    SortField,
    SortTableData,
    TableColumnInfo,
    TableSummary,
    UsedRangeInfo,
    WorkbookInfo,
    WorksheetInfo,
)

# ============================================================================
# Request DTOs
# ============================================================================


class TestExcelGetWorkbookInfoRequest:
    """Test ExcelGetWorkbookInfoRequest DTO (excel:get:workbookInfo)"""

    def test_valid_request_with_defaults(self) -> None:
        request = ExcelGetWorkbookInfoRequest(
            requestId="req_001",
            documentUri="file:///data.xlsx",
        )
        assert request.request_id == "req_001"
        assert request.document_uri == "file:///data.xlsx"
        assert isinstance(request.timestamp, int)

    def test_missing_required_fields(self) -> None:
        with pytest.raises(ValidationError) as exc_info:
            ExcelGetWorkbookInfoRequest(documentUri="file:///data.xlsx")
        assert "requestId" in str(exc_info.value)

        with pytest.raises(ValidationError) as exc_info:
            ExcelGetWorkbookInfoRequest(requestId="req_001")
        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        assert ExcelGetWorkbookInfoRequest.event_name == "excel:get:workbookInfo"

    def test_to_payload_camel_case(self) -> None:
        payload = ExcelGetWorkbookInfoRequest(
            requestId="req_002",
            documentUri="file:///data.xlsx",
        ).to_payload()
        assert payload["requestId"] == "req_002"
        assert payload["documentUri"] == "file:///data.xlsx"
        assert isinstance(payload["timestamp"], int)

    def test_build_class_method(self) -> None:
        request = ExcelGetWorkbookInfoRequest.build(document_uri="file:///data.xlsx")
        assert isinstance(request.request_id, str)
        assert request.document_uri == "file:///data.xlsx"


class TestExcelGetWorksheetInfoRequest:
    """Test ExcelGetWorksheetInfoRequest DTO (excel:get:worksheetInfo)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetWorksheetInfoRequest.event_name == "excel:get:worksheetInfo"

    def test_worksheet_name_optional(self) -> None:
        request = ExcelGetWorksheetInfoRequest(
            requestId="req_010",
            documentUri="file:///data.xlsx",
        )
        assert request.worksheet_name is None

    def test_accepts_snake_case_field_name(self) -> None:
        """populate_by_name: 输入接受 snake_case (Python 内部链路传入)。"""
        request = ExcelGetWorksheetInfoRequest(
            requestId="req_011",
            documentUri="file:///data.xlsx",
            worksheet_name="Sheet2",
        )
        assert request.worksheet_name == "Sheet2"

    def test_accepts_camel_case_alias(self) -> None:
        """wire format: 输入接受 camelCase alias。"""
        request = ExcelGetWorksheetInfoRequest(
            requestId="req_012",
            documentUri="file:///data.xlsx",
            worksheetName="Sheet3",
        )
        assert request.worksheet_name == "Sheet3"

    def test_payload_emits_camel_case_and_omits_none(self) -> None:
        # worksheet_name 省略时, to_payload (exclude_none) 不输出该键
        payload_omitted = ExcelGetWorksheetInfoRequest.build(document_uri="file:///data.xlsx").to_payload()
        assert "worksheetName" not in payload_omitted
        assert "worksheet_name" not in payload_omitted

        payload = ExcelGetWorksheetInfoRequest.build(
            document_uri="file:///data.xlsx",
            worksheet_name="Sheet1",
        ).to_payload()
        assert payload["worksheetName"] == "Sheet1"
        assert "worksheet_name" not in payload


class TestExcelGetSelectedRangeRequest:
    """Test ExcelGetSelectedRangeRequest DTO (excel:get:selectedRange)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetSelectedRangeRequest.event_name == "excel:get:selectedRange"

    def test_valid_request(self) -> None:
        request = ExcelGetSelectedRangeRequest.build(document_uri="file:///data.xlsx")
        assert request.document_uri == "file:///data.xlsx"
        assert isinstance(request.request_id, str)


class TestRequestRegistration:
    """读事件 (#18) + Range/公式 (#19) + Format/条件格式/合并 (#20) + Worksheet (#21) + Table (#22) + Chart (#23) + PivotTable (#24) 应自动注册。"""

    @pytest.mark.parametrize(
        "event,dto_cls",
        [
            ("excel:get:workbookInfo", ExcelGetWorkbookInfoRequest),
            ("excel:get:worksheetInfo", ExcelGetWorksheetInfoRequest),
            ("excel:get:selectedRange", ExcelGetSelectedRangeRequest),
            ("excel:get:range", ExcelGetRangeRequest),
            ("excel:set:range", ExcelSetRangeRequest),
            ("excel:clear:range", ExcelClearRangeRequest),
            ("excel:copy:range", ExcelCopyRangeRequest),
            ("excel:delete:range", ExcelDeleteRangeRequest),
            ("excel:insert:range", ExcelInsertRangeRequest),
            ("excel:set:formula", ExcelSetFormulaRequest),
            ("excel:get:rangeFormat", ExcelGetRangeFormatRequest),
            ("excel:set:rangeFormat", ExcelSetRangeFormatRequest),
            ("excel:add:conditionalFormat", ExcelAddConditionalFormatRequest),
            ("excel:clear:conditionalFormat", ExcelClearConditionalFormatRequest),
            ("excel:merge:cells", ExcelMergeCellsRequest),
            ("excel:unmerge:cells", ExcelUnmergeCellsRequest),
            ("excel:get:worksheets", ExcelGetWorksheetsRequest),
            ("excel:add:worksheet", ExcelAddWorksheetRequest),
            ("excel:delete:worksheet", ExcelDeleteWorksheetRequest),
            ("excel:rename:worksheet", ExcelRenameWorksheetRequest),
            ("excel:activate:worksheet", ExcelActivateWorksheetRequest),
            ("excel:insert:table", ExcelInsertTableRequest),
            ("excel:get:table", ExcelGetTableRequest),
            ("excel:get:tables", ExcelGetTablesRequest),
            ("excel:add:tableRow", ExcelAddTableRowRequest),
            ("excel:delete:tableRow", ExcelDeleteTableRowRequest),
            ("excel:sort:table", ExcelSortTableRequest),
            ("excel:insert:chart", ExcelInsertChartRequest),
            ("excel:get:charts", ExcelGetChartsRequest),
            ("excel:update:chart", ExcelUpdateChartRequest),
            ("excel:delete:chart", ExcelDeleteChartRequest),
            ("excel:insert:pivotTable", ExcelInsertPivotTableRequest),
            ("excel:get:pivotTables", ExcelGetPivotTablesRequest),
            ("excel:delete:pivotTable", ExcelDeletePivotTableRequest),
        ],
    )
    def test_event_registered(self, event: str, dto_cls: type) -> None:
        assert request_registry.contains(event)
        assert request_registry.get(event) is dto_cls


# ============================================================================
# Shared data models
# ============================================================================


class TestSheetInfo:
    """Test SheetInfo data model"""

    def test_from_camel_case(self) -> None:
        sheet = SheetInfo(name="Sheet1", index=0, isActive=True, isHidden=False)
        assert sheet.name == "Sheet1"
        assert sheet.index == 0
        assert sheet.is_active is True
        assert sheet.is_hidden is False

    def test_from_snake_case(self) -> None:
        sheet = SheetInfo(name="Sheet1", index=0, is_active=True, is_hidden=False)
        assert sheet.is_active is True

    def test_serialization_by_alias(self) -> None:
        dumped = SheetInfo(name="Hidden", index=2, is_active=False, is_hidden=True).model_dump(by_alias=True)
        assert dumped == {"name": "Hidden", "index": 2, "isActive": False, "isHidden": True}

    def test_missing_required_fields(self) -> None:
        with pytest.raises(ValidationError):
            SheetInfo(name="Sheet1", index=0)  # type: ignore[call-arg]


class TestWorkbookInfo:
    """Test WorkbookInfo data model (excel:get:workbookInfo response data)"""

    def test_round_trip_with_sheets(self) -> None:
        wire = {
            "sheets": [
                {"name": "Sheet1", "index": 0, "isActive": True, "isHidden": False},
                {"name": "Sheet2", "index": 1, "isActive": False, "isHidden": False},
            ],
            "activeSheet": "Sheet1",
            "fileName": "data.xlsx",
        }
        info = WorkbookInfo.model_validate(wire)
        assert len(info.sheets) == 2
        assert info.sheets[0].name == "Sheet1"
        assert info.active_sheet == "Sheet1"
        assert info.file_name == "data.xlsx"
        assert info.model_dump(by_alias=True) == wire


class TestWorksheetInfo:
    """Test WorksheetInfo data model (excel:get:worksheetInfo response data)"""

    def test_round_trip(self) -> None:
        wire = {
            "name": "Sheet1",
            "usedRange": {"address": "Sheet1!A1:D10", "rowCount": 10, "columnCount": 4},
            "tableCount": 1,
            "chartCount": 2,
        }
        info = WorksheetInfo.model_validate(wire)
        assert info.name == "Sheet1"
        assert isinstance(info.used_range, UsedRangeInfo)
        assert info.used_range.address == "Sheet1!A1:D10"
        assert info.used_range.row_count == 10
        assert info.used_range.column_count == 4
        assert info.table_count == 1
        assert info.chart_count == 2
        assert info.model_dump(by_alias=True) == wire


class TestSelectedRangeInfo:
    """Test SelectedRangeInfo data model (excel:get:selectedRange response data)"""

    def test_round_trip_preserves_values(self) -> None:
        wire = {
            "address": "Sheet1!A1:C3",
            "values": [["姓名", "年龄", "城市"], ["张三", 25, "北京"], ["李四", 30, "上海"]],
            "rowCount": 3,
            "columnCount": 3,
        }
        info = SelectedRangeInfo.model_validate(wire)
        assert info.address == "Sheet1!A1:C3"
        assert info.values[1] == ["张三", 25, "北京"]
        assert info.row_count == 3
        assert info.column_count == 3
        assert info.model_dump(by_alias=True) == wire


# ============================================================================
# #19 Range: Request DTOs (CRUD + 公式)
# ============================================================================


class TestExcelGetRangeRequest:
    """Test ExcelGetRangeRequest DTO (excel:get:range)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetRangeRequest.event_name == "excel:get:range"

    def test_required_address(self) -> None:
        with pytest.raises(ValidationError):
            ExcelGetRangeRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_include_format_defaults_false(self) -> None:
        req = ExcelGetRangeRequest.build(document_uri="file:///d.xlsx", address="A1:C3")
        assert req.include_format is False

    def test_payload_camel_case_and_omits_none_worksheet(self) -> None:
        payload = ExcelGetRangeRequest.build(
            document_uri="file:///d.xlsx", address="A1:C3", include_format=True
        ).to_payload()
        assert payload["address"] == "A1:C3"
        assert payload["includeFormat"] is True
        assert "worksheetName" not in payload

    def test_accepts_snake_and_camel(self) -> None:
        snake = ExcelGetRangeRequest(requestId="r", documentUri="file:///d.xlsx", address="A1", worksheet_name="Sheet2")
        camel = ExcelGetRangeRequest(requestId="r", documentUri="file:///d.xlsx", address="A1", worksheetName="Sheet2")
        assert snake.worksheet_name == camel.worksheet_name == "Sheet2"


class TestExcelSetRangeRequest:
    """Test ExcelSetRangeRequest DTO (excel:set:range)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelSetRangeRequest.event_name == "excel:set:range"

    def test_values_accepts_scalar(self) -> None:
        payload = ExcelSetRangeRequest.build(document_uri="file:///d.xlsx", address="A1:C3", values=0).to_payload()
        assert payload["values"] == 0
        assert payload["address"] == "A1:C3"

    def test_values_accepts_2d_array(self) -> None:
        values = [["a", "b"], ["c", "d"]]
        payload = ExcelSetRangeRequest.build(document_uri="file:///d.xlsx", address="A1:B2", values=values).to_payload()
        assert payload["values"] == values

    def test_required_values(self) -> None:
        with pytest.raises(ValidationError):
            ExcelSetRangeRequest(requestId="r", documentUri="file:///d.xlsx", address="A1")  # type: ignore[call-arg]


class TestExcelClearRangeRequest:
    """Test ExcelClearRangeRequest DTO (excel:clear:range)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelClearRangeRequest.event_name == "excel:clear:range"

    @pytest.mark.parametrize("clear_type", ["contents", "formats", "all"])
    def test_valid_clear_types(self, clear_type: str) -> None:
        payload = ExcelClearRangeRequest.build(
            document_uri="file:///d.xlsx", address="A1:C3", clear_type=clear_type
        ).to_payload()
        assert payload["clearType"] == clear_type

    def test_invalid_clear_type_rejected(self) -> None:
        with pytest.raises(ValidationError):
            ExcelClearRangeRequest.build(document_uri="file:///d.xlsx", address="A1:C3", clear_type="everything")


class TestExcelCopyRangeRequest:
    """Test ExcelCopyRangeRequest DTO (excel:copy:range)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelCopyRangeRequest.event_name == "excel:copy:range"

    def test_payload_uses_source_target_aliases(self) -> None:
        payload = ExcelCopyRangeRequest.build(
            document_uri="file:///d.xlsx", source_address="A1:C3", target_address="E1:G3"
        ).to_payload()
        assert payload["sourceAddress"] == "A1:C3"
        assert payload["targetAddress"] == "E1:G3"
        assert "source_address" not in payload


class TestExcelDeleteRangeRequest:
    """Test ExcelDeleteRangeRequest DTO (excel:delete:range)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelDeleteRangeRequest.event_name == "excel:delete:range"

    @pytest.mark.parametrize("direction", ["up", "left"])
    def test_valid_shift_directions(self, direction: str) -> None:
        payload = ExcelDeleteRangeRequest.build(
            document_uri="file:///d.xlsx", address="B2:B5", shift_direction=direction
        ).to_payload()
        assert payload["shiftDirection"] == direction

    @pytest.mark.parametrize("direction", ["down", "right"])
    def test_insert_directions_rejected_for_delete(self, direction: str) -> None:
        """delete 仅接受 up/left；插入方向 down/right 必须被枚举拒绝。"""
        with pytest.raises(ValidationError):
            ExcelDeleteRangeRequest.build(document_uri="file:///d.xlsx", address="B2:B5", shift_direction=direction)


class TestExcelInsertRangeRequest:
    """Test ExcelInsertRangeRequest DTO (excel:insert:range)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelInsertRangeRequest.event_name == "excel:insert:range"

    @pytest.mark.parametrize("direction", ["down", "right"])
    def test_valid_shift_directions(self, direction: str) -> None:
        payload = ExcelInsertRangeRequest.build(
            document_uri="file:///d.xlsx", address="B2:B5", shift_direction=direction
        ).to_payload()
        assert payload["shiftDirection"] == direction

    @pytest.mark.parametrize("direction", ["up", "left"])
    def test_delete_directions_rejected_for_insert(self, direction: str) -> None:
        """insert 仅接受 down/right；删除方向 up/left 必须被枚举拒绝。"""
        with pytest.raises(ValidationError):
            ExcelInsertRangeRequest.build(document_uri="file:///d.xlsx", address="B2:B5", shift_direction=direction)


class TestExcelSetFormulaRequest:
    """Test ExcelSetFormulaRequest DTO (excel:set:formula)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelSetFormulaRequest.event_name == "excel:set:formula"

    def test_payload_preserves_formula(self) -> None:
        payload = ExcelSetFormulaRequest.build(
            document_uri="file:///d.xlsx", address="D1", formula="=SUM(A1:C1)"
        ).to_payload()
        assert payload["address"] == "D1"
        assert payload["formula"] == "=SUM(A1:C1)"

    def test_required_formula(self) -> None:
        with pytest.raises(ValidationError):
            ExcelSetFormulaRequest(requestId="r", documentUri="file:///d.xlsx", address="D1")  # type: ignore[call-arg]


# ============================================================================
# #19 Range: shared data models
# ============================================================================


class TestRangeFormatInfo:
    """Test RangeFormatInfo data model (excel:get:range format payload, #20 复用)"""

    def test_round_trip(self) -> None:
        wire = {
            "font": {
                "name": "等线",
                "size": 11,
                "bold": False,
                "italic": False,
                "color": "#000000",
                "underline": "None",
            },
            "fill": {"color": "#FFFFFF"},
            "horizontalAlignment": "General",
            "verticalAlignment": "Bottom",
            "wrapText": False,
            "numberFormat": [["General", "0"], ["General", "0"]],
        }
        info = RangeFormatInfo.model_validate(wire)
        assert info.font.name == "等线"
        assert info.font.size == 11
        assert info.fill.color == "#FFFFFF"
        assert info.horizontal_alignment == "General"
        assert info.vertical_alignment == "Bottom"
        assert info.wrap_text is False
        assert info.number_format[1] == ["General", "0"]
        assert info.model_dump(by_alias=True) == wire


class TestGetRangeData:
    """Test GetRangeData data model (excel:get:range response data)"""

    def test_round_trip_without_format(self) -> None:
        wire = {
            "address": "Sheet1!A1:C3",
            "values": [["姓名", "年龄", "城市"], ["张三", 25, "北京"]],
            "rowCount": 2,
            "columnCount": 3,
        }
        data = GetRangeData.model_validate(wire)
        assert data.address == "Sheet1!A1:C3"
        assert data.values[1] == ["张三", 25, "北京"]
        assert data.format is None
        # exclude_none=True 下 format=None 被剔除, payload 等于 wire
        assert data.model_dump(by_alias=True, exclude_none=True) == wire
        # 不带 exclude_none 时, format=None 显式出现在 dump 中
        assert data.model_dump(by_alias=True)["format"] is None

    def test_round_trip_with_format(self) -> None:
        wire = {
            "address": "Sheet1!A1",
            "values": [[1]],
            "rowCount": 1,
            "columnCount": 1,
            "format": {
                "font": {
                    "name": "Calibri",
                    "size": 12,
                    "bold": True,
                    "italic": False,
                    "color": "#FF0000",
                    "underline": "None",
                },
                "fill": {"color": "#FFFF00"},
                "horizontalAlignment": "Center",
                "verticalAlignment": "Center",
                "wrapText": True,
                "numberFormat": [["0.00"]],
            },
        }
        data = GetRangeData.model_validate(wire)
        assert data.format is not None
        assert data.format.font.bold is True
        assert data.model_dump(by_alias=True, exclude_none=True) == wire


class TestRangeOperationResult:
    """Test RangeOperationResult data model (shared write-op response)"""

    def test_round_trip(self) -> None:
        wire = {"address": "Sheet1!A1:C3"}
        result = RangeOperationResult.model_validate(wire)
        assert result.address == "Sheet1!A1:C3"
        assert result.model_dump(by_alias=True) == wire


# ============================================================================
# #20 Format: Request DTOs (格式 / 条件格式 / 合并单元格)
# ============================================================================


class TestExcelGetRangeFormatRequest:
    """Test ExcelGetRangeFormatRequest DTO (excel:get:rangeFormat)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetRangeFormatRequest.event_name == "excel:get:rangeFormat"

    def test_required_address(self) -> None:
        with pytest.raises(ValidationError):
            ExcelGetRangeFormatRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_payload_camel_case_and_omits_none_worksheet(self) -> None:
        payload = ExcelGetRangeFormatRequest.build(document_uri="file:///d.xlsx", address="A1:C3").to_payload()
        assert payload["address"] == "A1:C3"
        assert "worksheetName" not in payload


class TestExcelSetRangeFormatRequest:
    """Test ExcelSetRangeFormatRequest DTO (excel:set:rangeFormat)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelSetRangeFormatRequest.event_name == "excel:set:rangeFormat"

    def test_required_format(self) -> None:
        with pytest.raises(ValidationError):
            ExcelSetRangeFormatRequest(requestId="r", documentUri="file:///d.xlsx", address="A1")  # type: ignore[call-arg]

    def test_partial_format_payload_camel_case_and_drops_none(self) -> None:
        """写侧偏更新: 仅传入字段出现, 未传字段 (exclude_none) 不出现, alias → camelCase。"""
        payload = ExcelSetRangeFormatRequest.build(
            document_uri="file:///d.xlsx",
            address="A1:C1",
            format={
                "font": {"bold": True, "size": 14},
                "alignment": {"horizontal": "Center", "wrap_text": True},
                "number_format": "0.00",
            },
        ).to_payload()
        assert payload["address"] == "A1:C1"
        assert payload["format"]["font"]["bold"] is True
        assert payload["format"]["font"]["size"] == 14
        # snake_case 输入 → camelCase wire
        assert payload["format"]["alignment"]["wrapText"] is True
        assert payload["format"]["numberFormat"] == "0.00"
        # 未传入的可选字段被剔除
        assert "italic" not in payload["format"]["font"]
        assert "fill" not in payload["format"]
        assert "borders" not in payload["format"]

    def test_borders_passthrough(self) -> None:
        payload = ExcelSetRangeFormatRequest.build(
            document_uri="file:///d.xlsx",
            address="A1:C1",
            format={"borders": {"top": {"style": "Thin", "color": "#000000"}}},
        ).to_payload()
        assert payload["format"]["borders"]["top"]["style"] == "Thin"
        assert payload["format"]["borders"]["top"]["color"] == "#000000"


class TestExcelAddConditionalFormatRequest:
    """Test ExcelAddConditionalFormatRequest DTO (excel:add:conditionalFormat)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelAddConditionalFormatRequest.event_name == "excel:add:conditionalFormat"

    def test_required_rule(self) -> None:
        with pytest.raises(ValidationError):
            ExcelAddConditionalFormatRequest(requestId="r", documentUri="file:///d.xlsx", address="B2:B100")  # type: ignore[call-arg]

    def test_rule_requires_type(self) -> None:
        """规则透传但 type 必填。"""
        with pytest.raises(ValidationError):
            ExcelAddConditionalFormatRequest.build(
                document_uri="file:///d.xlsx", address="B2:B100", rule={"operator": "greaterThan"}
            )

    def test_rule_passthrough_preserves_extra_keys(self) -> None:
        payload = ExcelAddConditionalFormatRequest.build(
            document_uri="file:///d.xlsx",
            address="B2:B100",
            rule={
                "type": "cellValue",
                "operator": "greaterThan",
                "value": 90,
                "format": {"fill": {"color": "#C6EFCE"}},
            },
        ).to_payload()
        assert payload["rule"]["type"] == "cellValue"
        assert payload["rule"]["operator"] == "greaterThan"
        assert payload["rule"]["value"] == 90
        assert payload["rule"]["format"]["fill"]["color"] == "#C6EFCE"


class TestExcelClearConditionalFormatRequest:
    """Test ExcelClearConditionalFormatRequest DTO (excel:clear:conditionalFormat)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelClearConditionalFormatRequest.event_name == "excel:clear:conditionalFormat"

    def test_payload_minimal_address(self) -> None:
        payload = ExcelClearConditionalFormatRequest.build(
            document_uri="file:///d.xlsx", address="B2:B100"
        ).to_payload()
        assert payload["address"] == "B2:B100"
        assert "worksheetName" not in payload


class TestExcelMergeCellsRequest:
    """Test ExcelMergeCellsRequest DTO (excel:merge:cells)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelMergeCellsRequest.event_name == "excel:merge:cells"

    def test_required_address(self) -> None:
        with pytest.raises(ValidationError):
            ExcelMergeCellsRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_payload_with_worksheet(self) -> None:
        payload = ExcelMergeCellsRequest.build(
            document_uri="file:///d.xlsx", address="A1:C1", worksheet_name="Sheet1"
        ).to_payload()
        assert payload["address"] == "A1:C1"
        assert payload["worksheetName"] == "Sheet1"
        assert "worksheet_name" not in payload


class TestExcelUnmergeCellsRequest:
    """Test ExcelUnmergeCellsRequest DTO (excel:unmerge:cells)"""

    def test_event_name_attribute(self) -> None:
        assert ExcelUnmergeCellsRequest.event_name == "excel:unmerge:cells"

    def test_payload_minimal_address(self) -> None:
        payload = ExcelUnmergeCellsRequest.build(document_uri="file:///d.xlsx", address="A1:C1").to_payload()
        assert payload["address"] == "A1:C1"
        assert "worksheetName" not in payload


# ============================================================================
# #20 Format: data + write-side option models
# ============================================================================


class TestGetRangeFormatData:
    """Test GetRangeFormatData data model (excel:get:rangeFormat response data)"""

    def test_round_trip_reuses_range_format_info(self) -> None:
        wire = {
            "address": "Sheet1!A1:C3",
            "format": {
                "font": {
                    "name": "等线",
                    "size": 11,
                    "bold": True,
                    "italic": False,
                    "color": "#000000",
                    "underline": "None",
                },
                "fill": {"color": "#FFFF00"},
                "horizontalAlignment": "Center",
                "verticalAlignment": "Center",
                "wrapText": True,
                "numberFormat": [["General", "0.00", "#,##0"]],
            },
        }
        data = GetRangeFormatData.model_validate(wire)
        assert data.address == "Sheet1!A1:C3"
        # format 字段复用 #19 读侧 RangeFormatInfo
        assert isinstance(data.format, RangeFormatInfo)
        assert data.format.font.bold is True
        assert data.format.horizontal_alignment == "Center"
        assert data.model_dump(by_alias=True) == wire

    def test_format_required(self) -> None:
        with pytest.raises(ValidationError):
            GetRangeFormatData.model_validate({"address": "A1"})


class TestSetRangeFormatOptions:
    """Test SetRangeFormatOptions write-side partial-update model (read/write asymmetry)"""

    def test_all_fields_optional_empty_payload(self) -> None:
        """全字段可选: 空对象 dump (exclude_none) 为空 dict。"""
        opts = SetRangeFormatOptions()
        assert opts.model_dump(by_alias=True, exclude_none=True) == {}

    def test_accepts_snake_and_camel_alignment(self) -> None:
        snake = SetRangeFormatOptions(alignment={"wrap_text": True, "indent_level": 2})
        camel = SetRangeFormatOptions(alignment={"wrapText": True, "indentLevel": 2})
        assert snake.alignment is not None and camel.alignment is not None
        assert snake.alignment.wrap_text is camel.alignment.wrap_text is True
        assert snake.alignment.indent_level == camel.alignment.indent_level == 2

    def test_number_format_is_scalar_string(self) -> None:
        """写侧 numberFormat 为标量字符串 (区别于读侧 RangeFormatInfo 的二维数组)。"""
        opts = SetRangeFormatOptions(number_format="yyyy-mm-dd")
        dumped = opts.model_dump(by_alias=True, exclude_none=True)
        assert dumped == {"numberFormat": "yyyy-mm-dd"}

    def test_alignment_preserves_case(self) -> None:
        """对齐枚举为开放词表, str 透传保留大小写 (如 'Center')。"""
        opts = SetRangeFormatOptions(alignment={"horizontal": "Center", "vertical": "Justify"})
        dumped = opts.model_dump(by_alias=True, exclude_none=True)
        assert dumped["alignment"]["horizontal"] == "Center"
        assert dumped["alignment"]["vertical"] == "Justify"


class TestConditionalFormatRule:
    """Test ConditionalFormatRule passthrough model (extra='allow')"""

    def test_type_required(self) -> None:
        with pytest.raises(ValidationError):
            ConditionalFormatRule.model_validate({"operator": "greaterThan"})

    def test_extra_keys_preserved_round_trip(self) -> None:
        wire = {"type": "colorScale", "minColor": "#FF0000", "maxColor": "#00FF00"}
        rule = ConditionalFormatRule.model_validate(wire)
        assert rule.type == "colorScale"
        assert rule.model_dump(by_alias=True, exclude_none=True) == wire


# ============================================================================
# #21 Worksheet: Request DTOs (字段形态以 spec type 为准, 各事件不同)
# ============================================================================


class TestExcelGetWorksheetsRequest:
    """Test ExcelGetWorksheetsRequest DTO (excel:get:worksheets) — 无业务参数"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetWorksheetsRequest.event_name == "excel:get:worksheets"

    def test_payload_carries_only_base_fields(self) -> None:
        payload = ExcelGetWorksheetsRequest.build(document_uri="file:///d.xlsx").to_payload()
        assert payload["documentUri"] == "file:///d.xlsx"
        assert "requestId" in payload
        # 无业务字段
        assert "worksheetName" not in payload
        assert "name" not in payload


class TestExcelAddWorksheetRequest:
    """Test ExcelAddWorksheetRequest DTO (excel:add:worksheet) — name 可选"""

    def test_event_name_attribute(self) -> None:
        assert ExcelAddWorksheetRequest.event_name == "excel:add:worksheet"

    def test_payload_with_name(self) -> None:
        payload = ExcelAddWorksheetRequest.build(document_uri="file:///d.xlsx", name="数据分析").to_payload()
        assert payload["name"] == "数据分析"

    def test_payload_without_name_drops_field(self) -> None:
        """name 省略时由 Excel 自动命名 → exclude_none 剔除该键。"""
        payload = ExcelAddWorksheetRequest.build(document_uri="file:///d.xlsx").to_payload()
        assert "name" not in payload


class TestExcelDeleteWorksheetRequest:
    """Test ExcelDeleteWorksheetRequest DTO (excel:delete:worksheet) — worksheet_name 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelDeleteWorksheetRequest.event_name == "excel:delete:worksheet"

    def test_required_worksheet_name(self) -> None:
        with pytest.raises(ValidationError):
            ExcelDeleteWorksheetRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_payload_snake_to_camel(self) -> None:
        payload = ExcelDeleteWorksheetRequest.build(document_uri="file:///d.xlsx", worksheet_name="Sheet3").to_payload()
        assert payload["worksheetName"] == "Sheet3"
        assert "worksheet_name" not in payload


class TestExcelRenameWorksheetRequest:
    """Test ExcelRenameWorksheetRequest DTO (excel:rename:worksheet) — current_name + new_name 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelRenameWorksheetRequest.event_name == "excel:rename:worksheet"

    def test_required_both_names(self) -> None:
        # 缺 new_name
        with pytest.raises(ValidationError):
            ExcelRenameWorksheetRequest(requestId="r", documentUri="file:///d.xlsx", currentName="Sheet1")  # type: ignore[call-arg]
        # 缺 current_name
        with pytest.raises(ValidationError):
            ExcelRenameWorksheetRequest(requestId="r", documentUri="file:///d.xlsx", newName="销售数据")  # type: ignore[call-arg]

    def test_payload_snake_to_camel(self) -> None:
        payload = ExcelRenameWorksheetRequest.build(
            document_uri="file:///d.xlsx", current_name="Sheet1", new_name="销售数据"
        ).to_payload()
        assert payload["currentName"] == "Sheet1"
        assert payload["newName"] == "销售数据"
        assert "current_name" not in payload
        assert "new_name" not in payload


class TestExcelActivateWorksheetRequest:
    """Test ExcelActivateWorksheetRequest DTO (excel:activate:worksheet) — worksheet_name 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelActivateWorksheetRequest.event_name == "excel:activate:worksheet"

    def test_required_worksheet_name(self) -> None:
        with pytest.raises(ValidationError):
            ExcelActivateWorksheetRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_payload_snake_to_camel(self) -> None:
        payload = ExcelActivateWorksheetRequest.build(
            document_uri="file:///d.xlsx", worksheet_name="Sheet2"
        ).to_payload()
        assert payload["worksheetName"] == "Sheet2"


# ============================================================================
# #21 Worksheet: response data models
# ============================================================================


class TestGetWorksheetsData:
    """Test GetWorksheetsData data model (excel:get:worksheets response data)"""

    def test_round_trip_reuses_sheet_info(self) -> None:
        wire = {
            "worksheets": [
                {"name": "Sheet1", "index": 0, "isActive": True, "isHidden": False},
                {"name": "Sheet2", "index": 1, "isActive": False, "isHidden": False},
            ]
        }
        data = GetWorksheetsData.model_validate(wire)
        assert len(data.worksheets) == 2
        # 列表项复用 #18 SheetInfo
        assert isinstance(data.worksheets[0], SheetInfo)
        assert data.worksheets[0].is_active is True
        assert data.worksheets[1].name == "Sheet2"
        assert data.model_dump(by_alias=True) == wire

    def test_empty_list(self) -> None:
        data = GetWorksheetsData.model_validate({"worksheets": []})
        assert data.worksheets == []


class TestAddWorksheetData:
    """Test AddWorksheetData data model (excel:add:worksheet response data)"""

    def test_round_trip(self) -> None:
        wire = {"name": "数据分析", "index": 2}
        data = AddWorksheetData.model_validate(wire)
        assert data.name == "数据分析"
        assert data.index == 2
        assert data.model_dump(by_alias=True) == wire

    def test_both_fields_required(self) -> None:
        with pytest.raises(ValidationError):
            AddWorksheetData.model_validate({"name": "X"})


class TestDeleteWorksheetData:
    """Test DeleteWorksheetData data model (excel:delete:worksheet response data)"""

    def test_round_trip(self) -> None:
        data = DeleteWorksheetData.model_validate({"deleted": True})
        assert data.deleted is True
        assert data.model_dump(by_alias=True) == {"deleted": True}


class TestRenameWorksheetData:
    """Test RenameWorksheetData data model (excel:rename:worksheet response data)"""

    def test_round_trip(self) -> None:
        data = RenameWorksheetData.model_validate({"name": "销售数据"})
        assert data.name == "销售数据"
        assert data.model_dump(by_alias=True) == {"name": "销售数据"}


class TestActivateWorksheetData:
    """Test ActivateWorksheetData data model (excel:activate:worksheet response data)"""

    def test_round_trip(self) -> None:
        data = ActivateWorksheetData.model_validate({"activated": True})
        assert data.activated is True
        assert data.model_dump(by_alias=True) == {"activated": True}


# ============================================================================
# #22 Table: Request DTOs (字段形态以 spec type 为准, 各事件不同)
# ============================================================================


class TestExcelInsertTableRequest:
    """Test ExcelInsertTableRequest DTO (excel:insert:table) — address + has_headers 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelInsertTableRequest.event_name == "excel:insert:table"

    def test_required_address_and_has_headers(self) -> None:
        # 缺 has_headers
        with pytest.raises(ValidationError):
            ExcelInsertTableRequest(requestId="r", documentUri="file:///d.xlsx", address="A1:C4")  # type: ignore[call-arg]
        # 缺 address
        with pytest.raises(ValidationError):
            ExcelInsertTableRequest(requestId="r", documentUri="file:///d.xlsx", hasHeaders=True)  # type: ignore[call-arg]

    def test_payload_minimal_drops_optionals(self) -> None:
        """data / styleName / worksheetName 省略时 exclude_none 剔除。"""
        payload = ExcelInsertTableRequest.build(
            document_uri="file:///d.xlsx", address="A1:C4", has_headers=True
        ).to_payload()
        assert payload["address"] == "A1:C4"
        assert payload["hasHeaders"] is True
        assert "data" not in payload
        assert "styleName" not in payload
        assert "worksheetName" not in payload
        assert "has_headers" not in payload

    def test_payload_with_data_and_style(self) -> None:
        payload = ExcelInsertTableRequest.build(
            document_uri="file:///d.xlsx",
            address="A1:C2",
            has_headers=True,
            data=[["姓名", "年龄", "城市"], ["张三", 25, "北京"]],
            style_name="TableStyleMedium2",
        ).to_payload()
        assert payload["data"] == [["姓名", "年龄", "城市"], ["张三", 25, "北京"]]
        assert payload["styleName"] == "TableStyleMedium2"
        assert "style_name" not in payload


class TestExcelGetTableRequest:
    """Test ExcelGetTableRequest DTO (excel:get:table) — table_id 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetTableRequest.event_name == "excel:get:table"

    def test_required_table_id(self) -> None:
        with pytest.raises(ValidationError):
            ExcelGetTableRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_payload_snake_to_camel(self) -> None:
        payload = ExcelGetTableRequest.build(document_uri="file:///d.xlsx", table_id="Table1").to_payload()
        assert payload["tableId"] == "Table1"
        assert "table_id" not in payload
        assert "worksheetName" not in payload


class TestExcelGetTablesRequest:
    """Test ExcelGetTablesRequest DTO (excel:get:tables) — 仅 worksheet_name 可选"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetTablesRequest.event_name == "excel:get:tables"

    def test_payload_minimal_omits_worksheet(self) -> None:
        payload = ExcelGetTablesRequest.build(document_uri="file:///d.xlsx").to_payload()
        assert payload["documentUri"] == "file:///d.xlsx"
        assert "worksheetName" not in payload

    def test_payload_with_worksheet(self) -> None:
        payload = ExcelGetTablesRequest.build(document_uri="file:///d.xlsx", worksheet_name="Sheet2").to_payload()
        assert payload["worksheetName"] == "Sheet2"


class TestExcelAddTableRowRequest:
    """Test ExcelAddTableRowRequest DTO (excel:add:tableRow) — table_id + values 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelAddTableRowRequest.event_name == "excel:add:tableRow"

    def test_required_table_id_and_values(self) -> None:
        with pytest.raises(ValidationError):
            ExcelAddTableRowRequest(requestId="r", documentUri="file:///d.xlsx", tableId="Table1")  # type: ignore[call-arg]
        with pytest.raises(ValidationError):
            ExcelAddTableRowRequest(requestId="r", documentUri="file:///d.xlsx", values=["赵六", 35])  # type: ignore[call-arg]

    def test_payload_preserves_values(self) -> None:
        payload = ExcelAddTableRowRequest.build(
            document_uri="file:///d.xlsx", table_id="Table1", values=["赵六", 35, "深圳"]
        ).to_payload()
        assert payload["tableId"] == "Table1"
        assert payload["values"] == ["赵六", 35, "深圳"]
        assert "table_id" not in payload


class TestExcelDeleteTableRowRequest:
    """Test ExcelDeleteTableRowRequest DTO (excel:delete:tableRow) — table_id + row_index 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelDeleteTableRowRequest.event_name == "excel:delete:tableRow"

    def test_required_table_id_and_row_index(self) -> None:
        with pytest.raises(ValidationError):
            ExcelDeleteTableRowRequest(requestId="r", documentUri="file:///d.xlsx", tableId="Table1")  # type: ignore[call-arg]
        with pytest.raises(ValidationError):
            ExcelDeleteTableRowRequest(requestId="r", documentUri="file:///d.xlsx", rowIndex=0)  # type: ignore[call-arg]

    def test_payload_snake_to_camel(self) -> None:
        payload = ExcelDeleteTableRowRequest.build(
            document_uri="file:///d.xlsx", table_id="Table1", row_index=2
        ).to_payload()
        assert payload["tableId"] == "Table1"
        assert payload["rowIndex"] == 2
        assert "row_index" not in payload

    def test_payload_row_index_zero_retained(self) -> None:
        """rowIndex=0 是合法首行索引，exclude_none 不得剔除 (0 ≠ None)。"""
        payload = ExcelDeleteTableRowRequest.build(
            document_uri="file:///d.xlsx", table_id="Table1", row_index=0
        ).to_payload()
        assert payload["rowIndex"] == 0


class TestExcelSortTableRequest:
    """Test ExcelSortTableRequest DTO (excel:sort:table) — table_id + sort_fields 必填, 含嵌套 SortField"""

    def test_event_name_attribute(self) -> None:
        assert ExcelSortTableRequest.event_name == "excel:sort:table"

    def test_required_table_id_and_sort_fields(self) -> None:
        with pytest.raises(ValidationError):
            ExcelSortTableRequest(requestId="r", documentUri="file:///d.xlsx", tableId="Table1")  # type: ignore[call-arg]

    def test_payload_nested_sort_fields_camel_case(self) -> None:
        """sortFields 内嵌 columnIndex (camel)；ascending 显式 False 保留。"""
        payload = ExcelSortTableRequest.build(
            document_uri="file:///d.xlsx",
            table_id="Table1",
            sort_fields=[SortField(column_index=1, ascending=False), SortField(column_index=0)],
        ).to_payload()
        assert payload["tableId"] == "Table1"
        assert payload["sortFields"][0] == {"columnIndex": 1, "ascending": False}
        # 第二个键省略 ascending → exclude_none 剔除, 仅余 columnIndex
        assert payload["sortFields"][1] == {"columnIndex": 0}
        assert "sort_fields" not in payload

    def test_sort_field_accepts_snake_and_camel(self) -> None:
        """populate_by_name: SortField 内部链路接受 snake_case column_index。"""
        snake = SortField(column_index=2)
        camel = SortField.model_validate({"columnIndex": 2})
        assert snake.column_index == camel.column_index == 2
        assert snake.ascending is None


# ============================================================================
# #22 Table: nested + response data models
# ============================================================================


class TestTableColumnInfo:
    """Test TableColumnInfo nested model (part of excel:get:table response)"""

    def test_round_trip(self) -> None:
        info = TableColumnInfo.model_validate({"name": "年龄", "index": 1})
        assert info.name == "年龄"
        assert info.index == 1
        assert info.model_dump(by_alias=True) == {"name": "年龄", "index": 1}


class TestTableSummary:
    """Test TableSummary nested model (excel:get:tables entry)"""

    def test_round_trip(self) -> None:
        wire = {"name": "Table1", "id": "{12345}", "address": "Sheet1!A1:C4"}
        summary = TableSummary.model_validate(wire)
        assert summary.name == "Table1"
        assert summary.id == "{12345}"
        assert summary.address == "Sheet1!A1:C4"
        assert summary.model_dump(by_alias=True) == wire


class TestInsertTableData:
    """Test InsertTableData data model (excel:insert:table response data)"""

    def test_round_trip(self) -> None:
        wire = {"name": "Table1", "address": "Sheet1!A1:C4"}
        data = InsertTableData.model_validate(wire)
        assert data.name == "Table1"
        assert data.address == "Sheet1!A1:C4"
        assert data.model_dump(by_alias=True) == wire

    def test_both_fields_required(self) -> None:
        with pytest.raises(ValidationError):
            InsertTableData.model_validate({"name": "Table1"})


class TestGetTableData:
    """Test GetTableData data model (excel:get:table response data, rich)"""

    def test_round_trip_with_columns(self) -> None:
        wire = {
            "name": "Table1",
            "id": "{12345}",
            "address": "Sheet1!A1:C4",
            "rowCount": 3,
            "columnCount": 3,
            "columns": [
                {"name": "姓名", "index": 0},
                {"name": "年龄", "index": 1},
                {"name": "城市", "index": 2},
            ],
            "styleName": "TableStyleMedium2",
            "showHeaders": True,
        }
        data = GetTableData.model_validate(wire)
        assert data.row_count == 3
        assert data.column_count == 3
        assert data.style_name == "TableStyleMedium2"
        assert data.show_headers is True
        # columns 列表项为 TableColumnInfo
        assert len(data.columns) == 3
        assert isinstance(data.columns[0], TableColumnInfo)
        assert data.columns[1].name == "年龄"
        assert data.model_dump(by_alias=True) == wire

    def test_required_fields(self) -> None:
        with pytest.raises(ValidationError):
            GetTableData.model_validate({"name": "Table1", "id": "{1}", "address": "A1:C4"})


class TestGetTablesData:
    """Test GetTablesData data model (excel:get:tables response data, slim entries)"""

    def test_round_trip_reuses_table_summary(self) -> None:
        wire = {
            "tables": [
                {"name": "Table1", "id": "{1}", "address": "Sheet1!A1:C4"},
                {"name": "Table2", "id": "{2}", "address": "Sheet1!E1:G10"},
            ]
        }
        data = GetTablesData.model_validate(wire)
        assert len(data.tables) == 2
        assert isinstance(data.tables[0], TableSummary)
        assert data.tables[1].name == "Table2"
        assert data.model_dump(by_alias=True) == wire

    def test_empty_list(self) -> None:
        data = GetTablesData.model_validate({"tables": []})
        assert data.tables == []


class TestAddTableRowData:
    """Test AddTableRowData data model (excel:add:tableRow response data)"""

    def test_round_trip(self) -> None:
        data = AddTableRowData.model_validate({"tableId": "Table1"})
        assert data.table_id == "Table1"
        assert data.model_dump(by_alias=True) == {"tableId": "Table1"}


class TestDeleteTableRowData:
    """Test DeleteTableRowData data model (excel:delete:tableRow response data)"""

    def test_round_trip(self) -> None:
        data = DeleteTableRowData.model_validate({"deleted": True})
        assert data.deleted is True
        assert data.model_dump(by_alias=True) == {"deleted": True}


class TestSortTableData:
    """Test SortTableData data model (excel:sort:table response data)"""

    def test_round_trip(self) -> None:
        data = SortTableData.model_validate({"sorted": True})
        assert data.sorted is True
        assert data.model_dump(by_alias=True) == {"sorted": True}


# ============================================================================
# #23 Chart: Request DTOs (按 spec type: sourceAddress + chartType:str 开放, 无 ChartData/3015)
# ============================================================================


class TestExcelInsertChartRequest:
    """Test ExcelInsertChartRequest DTO (excel:insert:chart) — source_address + chart_type 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelInsertChartRequest.event_name == "excel:insert:chart"

    def test_required_source_address_and_chart_type(self) -> None:
        # 缺 chart_type
        with pytest.raises(ValidationError):
            ExcelInsertChartRequest(requestId="r", documentUri="file:///d.xlsx", sourceAddress="A1:C4")  # type: ignore[call-arg]
        # 缺 source_address
        with pytest.raises(ValidationError):
            ExcelInsertChartRequest(requestId="r", documentUri="file:///d.xlsx", chartType="Line")  # type: ignore[call-arg]

    def test_chart_type_accepts_open_string(self) -> None:
        """chartType 为开放字符串 (spec type string)，非 spec「常见」列表的取值亦合法。"""
        payload = ExcelInsertChartRequest.build(
            document_uri="file:///d.xlsx", source_address="A1:C4", chart_type="Sunburst"
        ).to_payload()
        assert payload["chartType"] == "Sunburst"

    def test_payload_minimal_drops_optionals(self) -> None:
        """title / position / worksheetName 省略时 exclude_none 剔除。"""
        payload = ExcelInsertChartRequest.build(
            document_uri="file:///d.xlsx", source_address="A1:C4", chart_type="ColumnClustered"
        ).to_payload()
        assert payload["sourceAddress"] == "A1:C4"
        assert payload["chartType"] == "ColumnClustered"
        assert "title" not in payload
        assert "position" not in payload
        assert "worksheetName" not in payload
        assert "source_address" not in payload
        assert "chart_type" not in payload

    def test_payload_nested_position_drops_none_keeps_zero(self) -> None:
        """position 内 top=0 保留 (0 ≠ None)，省略的 height 被 exclude_none 剔除。"""
        payload = ExcelInsertChartRequest.build(
            document_uri="file:///d.xlsx",
            source_address="A1:C4",
            chart_type="Pie",
            title="占比",
            position=ChartPosition(top=0, left=300, width=400),
        ).to_payload()
        assert payload["title"] == "占比"
        assert payload["position"] == {"top": 0, "left": 300, "width": 400}
        assert "height" not in payload["position"]


class TestExcelGetChartsRequest:
    """Test ExcelGetChartsRequest DTO (excel:get:charts) — 仅 worksheet_name 可选"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetChartsRequest.event_name == "excel:get:charts"

    def test_payload_minimal_omits_worksheet(self) -> None:
        payload = ExcelGetChartsRequest.build(document_uri="file:///d.xlsx").to_payload()
        assert payload["documentUri"] == "file:///d.xlsx"
        assert "worksheetName" not in payload

    def test_payload_with_worksheet(self) -> None:
        payload = ExcelGetChartsRequest.build(document_uri="file:///d.xlsx", worksheet_name="Sheet2").to_payload()
        assert payload["worksheetName"] == "Sheet2"


class TestExcelUpdateChartRequest:
    """Test ExcelUpdateChartRequest DTO (excel:update:chart) — chart_name + properties 必填, 偏更新"""

    def test_event_name_attribute(self) -> None:
        assert ExcelUpdateChartRequest.event_name == "excel:update:chart"

    def test_required_chart_name_and_properties(self) -> None:
        # 缺 properties
        with pytest.raises(ValidationError):
            ExcelUpdateChartRequest(requestId="r", documentUri="file:///d.xlsx", chartName="Chart 1")  # type: ignore[call-arg]
        # 缺 chart_name
        with pytest.raises(ValidationError):
            ExcelUpdateChartRequest(  # type: ignore[call-arg]
                requestId="r", documentUri="file:///d.xlsx", properties=ChartUpdateProperties(title="t")
            )

    def test_payload_partial_properties_camel_case(self) -> None:
        """偏更新: 仅传入 chartType/sourceAddress; title/position 省略 → exclude_none 剔除。"""
        payload = ExcelUpdateChartRequest.build(
            document_uri="file:///d.xlsx",
            chart_name="Chart 1",
            properties=ChartUpdateProperties(chart_type="Line", source_address="A1:D9"),
        ).to_payload()
        assert payload["chartName"] == "Chart 1"
        assert payload["properties"] == {"chartType": "Line", "sourceAddress": "A1:D9"}
        assert "chart_name" not in payload
        assert "chart_type" not in payload["properties"]

    def test_payload_properties_nested_position(self) -> None:
        payload = ExcelUpdateChartRequest.build(
            document_uri="file:///d.xlsx",
            chart_name="Chart 1",
            properties=ChartUpdateProperties(title="新标题", position=ChartPosition(width=500, height=300)),
        ).to_payload()
        assert payload["properties"]["title"] == "新标题"
        assert payload["properties"]["position"] == {"width": 500, "height": 300}

    def test_empty_properties_payload(self) -> None:
        """properties 全省略 → 空对象 (偏更新允许无字段)。"""
        payload = ExcelUpdateChartRequest.build(
            document_uri="file:///d.xlsx", chart_name="Chart 1", properties=ChartUpdateProperties()
        ).to_payload()
        assert payload["properties"] == {}

    def test_properties_accepts_snake_and_camel(self) -> None:
        """populate_by_name: ChartUpdateProperties 接受 snake_case chart_type。"""
        snake = ChartUpdateProperties(chart_type="Line")
        camel = ChartUpdateProperties.model_validate({"chartType": "Line"})
        assert snake.chart_type == camel.chart_type == "Line"


class TestExcelDeleteChartRequest:
    """Test ExcelDeleteChartRequest DTO (excel:delete:chart) — chart_name 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelDeleteChartRequest.event_name == "excel:delete:chart"

    def test_required_chart_name(self) -> None:
        with pytest.raises(ValidationError):
            ExcelDeleteChartRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_payload_snake_to_camel(self) -> None:
        payload = ExcelDeleteChartRequest.build(document_uri="file:///d.xlsx", chart_name="Chart 1").to_payload()
        assert payload["chartName"] == "Chart 1"
        assert "chart_name" not in payload
        assert "worksheetName" not in payload


# ============================================================================
# #23 Chart: nested + response data models
# ============================================================================


class TestChartPosition:
    """Test ChartPosition nested model (insert/update, all optional)"""

    def test_round_trip_partial(self) -> None:
        pos = ChartPosition.model_validate({"top": 200, "left": 300})
        assert pos.top == 200
        assert pos.left == 300
        assert pos.width is None
        # exclude_none: 省略字段不出现在 wire payload
        assert pos.model_dump(by_alias=True, exclude_none=True) == {"top": 200, "left": 300}

    def test_all_optional_empty(self) -> None:
        pos = ChartPosition()
        assert pos.model_dump(by_alias=True, exclude_none=True) == {}


class TestChartUpdateProperties:
    """Test ChartUpdateProperties nested model (excel:update:chart properties, all optional)"""

    def test_round_trip_camel_case(self) -> None:
        wire = {"title": "T", "chartType": "Line", "sourceAddress": "A1:D9"}
        props = ChartUpdateProperties.model_validate(wire)
        assert props.title == "T"
        assert props.chart_type == "Line"
        assert props.source_address == "A1:D9"
        assert props.model_dump(by_alias=True, exclude_none=True) == wire

    def test_nested_position(self) -> None:
        props = ChartUpdateProperties.model_validate({"position": {"top": 10, "left": 20}})
        assert isinstance(props.position, ChartPosition)
        assert props.position.top == 10


class TestChartSummary:
    """Test ChartSummary nested model (excel:get:charts entry)"""

    def test_round_trip(self) -> None:
        wire = {
            "name": "Chart 1",
            "chartType": "ColumnClustered",
            "title": "销售数据",
            "top": 200.0,
            "left": 300.0,
            "width": 400.0,
            "height": 300.0,
        }
        summary = ChartSummary.model_validate(wire)
        assert summary.name == "Chart 1"
        assert summary.chart_type == "ColumnClustered"
        assert summary.title == "销售数据"
        assert summary.model_dump(by_alias=True) == wire


class TestChartOperationResult:
    """Test ChartOperationResult data model (excel:insert/update:chart response, shared {name})"""

    def test_round_trip(self) -> None:
        data = ChartOperationResult.model_validate({"name": "Chart 1"})
        assert data.name == "Chart 1"
        assert data.model_dump(by_alias=True) == {"name": "Chart 1"}

    def test_name_required(self) -> None:
        with pytest.raises(ValidationError):
            ChartOperationResult.model_validate({})


class TestGetChartsData:
    """Test GetChartsData data model (excel:get:charts response data)"""

    def test_round_trip_reuses_chart_summary(self) -> None:
        wire = {
            "charts": [
                {
                    "name": "Chart 1",
                    "chartType": "ColumnClustered",
                    "title": "销售",
                    "top": 200.0,
                    "left": 300.0,
                    "width": 400.0,
                    "height": 300.0,
                }
            ]
        }
        data = GetChartsData.model_validate(wire)
        assert len(data.charts) == 1
        assert isinstance(data.charts[0], ChartSummary)
        assert data.charts[0].name == "Chart 1"
        assert data.model_dump(by_alias=True) == wire

    def test_empty_list(self) -> None:
        data = GetChartsData.model_validate({"charts": []})
        assert data.charts == []


class TestDeleteChartData:
    """Test DeleteChartData data model (excel:delete:chart response data)"""

    def test_round_trip(self) -> None:
        data = DeleteChartData.model_validate({"deleted": True})
        assert data.deleted is True
        assert data.model_dump(by_alias=True) == {"deleted": True}


# ============================================================================
# #24 PivotTable: Request DTOs (按 spec type: sourceAddress + targetAddress, 无 行/列/值/筛选)
# ============================================================================


class TestExcelInsertPivotTableRequest:
    """Test ExcelInsertPivotTableRequest DTO (excel:insert:pivotTable) — source_address + target_address 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelInsertPivotTableRequest.event_name == "excel:insert:pivotTable"

    def test_required_source_and_target_address(self) -> None:
        # 缺 target_address
        with pytest.raises(ValidationError):
            ExcelInsertPivotTableRequest(requestId="r", documentUri="file:///d.xlsx", sourceAddress="A1:D100")  # type: ignore[call-arg]
        # 缺 source_address
        with pytest.raises(ValidationError):
            ExcelInsertPivotTableRequest(requestId="r", documentUri="file:///d.xlsx", targetAddress="F1")  # type: ignore[call-arg]

    def test_payload_minimal_drops_optionals(self) -> None:
        """name / worksheetName 省略时 exclude_none 剔除；snake_case 不泄漏到 wire。"""
        payload = ExcelInsertPivotTableRequest.build(
            document_uri="file:///d.xlsx", source_address="A1:D100", target_address="F1"
        ).to_payload()
        assert payload["sourceAddress"] == "A1:D100"
        assert payload["targetAddress"] == "F1"
        assert "name" not in payload
        assert "worksheetName" not in payload
        assert "source_address" not in payload
        assert "target_address" not in payload

    def test_payload_with_name_and_worksheet(self) -> None:
        payload = ExcelInsertPivotTableRequest.build(
            document_uri="file:///d.xlsx",
            source_address="A1:D100",
            target_address="F1",
            name="销售汇总",
            worksheet_name="Sheet2",
        ).to_payload()
        assert payload["name"] == "销售汇总"
        assert payload["worksheetName"] == "Sheet2"


class TestExcelGetPivotTablesRequest:
    """Test ExcelGetPivotTablesRequest DTO (excel:get:pivotTables) — 仅 worksheet_name 可选"""

    def test_event_name_attribute(self) -> None:
        assert ExcelGetPivotTablesRequest.event_name == "excel:get:pivotTables"

    def test_payload_minimal_omits_worksheet(self) -> None:
        payload = ExcelGetPivotTablesRequest.build(document_uri="file:///d.xlsx").to_payload()
        assert payload["documentUri"] == "file:///d.xlsx"
        assert "worksheetName" not in payload

    def test_payload_with_worksheet(self) -> None:
        payload = ExcelGetPivotTablesRequest.build(document_uri="file:///d.xlsx", worksheet_name="Sheet2").to_payload()
        assert payload["worksheetName"] == "Sheet2"


class TestExcelDeletePivotTableRequest:
    """Test ExcelDeletePivotTableRequest DTO (excel:delete:pivotTable) — pivot_table_name 必填"""

    def test_event_name_attribute(self) -> None:
        assert ExcelDeletePivotTableRequest.event_name == "excel:delete:pivotTable"

    def test_required_pivot_table_name(self) -> None:
        with pytest.raises(ValidationError):
            ExcelDeletePivotTableRequest(requestId="r", documentUri="file:///d.xlsx")  # type: ignore[call-arg]

    def test_payload_snake_to_camel(self) -> None:
        payload = ExcelDeletePivotTableRequest.build(
            document_uri="file:///d.xlsx", pivot_table_name="销售汇总"
        ).to_payload()
        assert payload["pivotTableName"] == "销售汇总"
        assert "pivot_table_name" not in payload
        assert "worksheetName" not in payload


# ============================================================================
# #24 PivotTable: response data models
# ============================================================================


class TestPivotTableOperationResult:
    """Test PivotTableOperationResult data model (excel:insert:pivotTable response, {name})"""

    def test_round_trip(self) -> None:
        data = PivotTableOperationResult.model_validate({"name": "PivotTable1"})
        assert data.name == "PivotTable1"
        assert data.model_dump(by_alias=True) == {"name": "PivotTable1"}

    def test_name_required(self) -> None:
        with pytest.raises(ValidationError):
            PivotTableOperationResult.model_validate({})


class TestPivotTableSummary:
    """Test PivotTableSummary nested model (excel:get:pivotTables entry, {name, id})"""

    def test_round_trip(self) -> None:
        wire = {"name": "销售汇总", "id": "{abcd-1234}"}
        summary = PivotTableSummary.model_validate(wire)
        assert summary.name == "销售汇总"
        assert summary.id == "{abcd-1234}"
        assert summary.model_dump(by_alias=True) == wire

    def test_both_fields_required(self) -> None:
        # 缺 id
        with pytest.raises(ValidationError):
            PivotTableSummary.model_validate({"name": "销售汇总"})


class TestGetPivotTablesData:
    """Test GetPivotTablesData data model (excel:get:pivotTables response data)"""

    def test_round_trip_reuses_summary(self) -> None:
        wire = {
            "pivotTables": [
                {"name": "销售汇总", "id": "{abcd-1234}"},
                {"name": "PivotTable2", "id": "{efgh-5678}"},
            ]
        }
        data = GetPivotTablesData.model_validate(wire)
        assert len(data.pivot_tables) == 2
        assert isinstance(data.pivot_tables[0], PivotTableSummary)
        assert data.pivot_tables[0].name == "销售汇总"
        assert data.model_dump(by_alias=True) == wire

    def test_empty_list(self) -> None:
        data = GetPivotTablesData.model_validate({"pivotTables": []})
        assert data.pivot_tables == []


class TestDeletePivotTableData:
    """Test DeletePivotTableData data model (excel:delete:pivotTable response data)"""

    def test_round_trip(self) -> None:
        data = DeletePivotTableData.model_validate({"deleted": True})
        assert data.deleted is True
        assert data.model_dump(by_alias=True) == {"deleted": True}
