"""
Test Excel DTOs

测试 Excel 事件的数据传输对象 (OASP 0.3.0 events-excel.md, issue #18 Foundation)。

覆盖:
- 3 个状态感知读事件 Request DTO 的构造、必填校验、event_name、camelCase 序列化、注册。
- 共享数据模型 (SheetInfo / UsedRangeInfo / WorkbookInfo / WorksheetInfo /
  SelectedRangeInfo) 的 snake_case ↔ camelCase 双向兼容。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import request_registry
from office4ai.environment.workspace.dtos.excel import (
    ExcelGetSelectedRangeRequest,
    ExcelGetWorkbookInfoRequest,
    ExcelGetWorksheetInfoRequest,
    SelectedRangeInfo,
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
    """3 个读事件应自动注册到全局 request_registry。"""

    @pytest.mark.parametrize(
        "event,dto_cls",
        [
            ("excel:get:workbookInfo", ExcelGetWorkbookInfoRequest),
            ("excel:get:worksheetInfo", ExcelGetWorksheetInfoRequest),
            ("excel:get:selectedRange", ExcelGetSelectedRangeRequest),
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
