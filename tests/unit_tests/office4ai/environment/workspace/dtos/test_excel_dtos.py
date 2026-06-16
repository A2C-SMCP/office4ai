"""
Test Excel DTOs

测试 Excel 事件的数据传输对象 (OASP 0.3.0 events-excel.md, issue #18 Foundation)。

覆盖:
- #18 读事件 + #19 Range/公式 + #20 Format/条件格式/合并 Request DTO 的构造、必填
  校验、event_name、camelCase 序列化、枚举约束、注册。
- 共享数据模型 (SheetInfo / UsedRangeInfo / WorkbookInfo / WorksheetInfo /
  SelectedRangeInfo / RangeFormatInfo / GetRangeData / RangeOperationResult /
  GetRangeFormatData) 的 snake_case ↔ camelCase 双向兼容。
- #20 写侧偏更新模型 (SetRangeFormatOptions) 与条件格式透传模型
  (ConditionalFormatRule) —— 读/写不对称、None 剔除、extra 透传。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import request_registry
from office4ai.environment.workspace.dtos.excel import (
    ConditionalFormatRule,
    ExcelAddConditionalFormatRequest,
    ExcelClearConditionalFormatRequest,
    ExcelClearRangeRequest,
    ExcelCopyRangeRequest,
    ExcelDeleteRangeRequest,
    ExcelGetRangeFormatRequest,
    ExcelGetRangeRequest,
    ExcelGetSelectedRangeRequest,
    ExcelGetWorkbookInfoRequest,
    ExcelGetWorksheetInfoRequest,
    ExcelInsertRangeRequest,
    ExcelMergeCellsRequest,
    ExcelSetFormulaRequest,
    ExcelSetRangeFormatRequest,
    ExcelSetRangeRequest,
    ExcelUnmergeCellsRequest,
    GetRangeData,
    GetRangeFormatData,
    RangeFormatInfo,
    RangeOperationResult,
    SelectedRangeInfo,
    SetRangeFormatOptions,
    SheetInfo,
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
    """读事件 (#18) + Range/公式 (#19) + Format/条件格式/合并 (#20) 应自动注册。"""

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
