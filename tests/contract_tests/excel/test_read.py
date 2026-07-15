"""
Contract tests — #18 Foundation 状态感知读 (OASP 0.3.0 /excel Draft)

事件：excel:get:workbookInfo / excel:get:worksheetInfo / excel:get:selectedRange
验证 OfficeWorkspace.execute() → wrap_request → AsyncServer.call() → MockAddInClient
的端到端 RPC round-trip，重点核对 snake_case 参数在 wire 上转为 camelCase。

OASP Spec: https://doc.turingfocus.cn/oasp/0.3.0/specification/events-excel/
"""

from __future__ import annotations

import asyncio

import pytest

pytestmark = [pytest.mark.asyncio, pytest.mark.contract]


def _ok(request: dict, data: dict) -> dict:
    """构造成功 ack 信封。"""
    return {
        "requestId": request["requestId"],
        "success": True,
        "data": data,
        "timestamp": int(asyncio.get_event_loop().time() * 1000),
    }


async def test_get_workbook_info_success(excel_roundtrip, excel_factory):
    """excel:get:workbookInfo —— 无业务参数，仅三要素。"""

    def factory(request: dict) -> dict:
        assert "requestId" in request
        assert "documentUri" in request
        # 无业务参数：除三要素外不应出现其他键
        assert set(request.keys()) <= {"requestId", "documentUri", "timestamp"}
        return _ok(request, excel_factory.workbook_info_response(active_sheet="Sheet1"))

    result, events = await excel_roundtrip("get:workbookInfo", {}, factory)

    assert result.success is True, result.error
    assert result.data["activeSheet"] == "Sheet1"
    assert result.data["fileName"] == "test.xlsx"
    assert result.data["sheets"][0]["isActive"] is True
    assert events[0][0] == "excel:get:workbookInfo"


async def test_get_worksheet_info_with_worksheet_name(excel_roundtrip, excel_factory):
    """excel:get:worksheetInfo —— worksheet_name 在 wire 上转为 worksheetName。"""

    def factory(request: dict) -> dict:
        assert request["worksheetName"] == "Data"
        assert "worksheet_name" not in request  # snake_case 不应泄漏到 wire
        return _ok(request, excel_factory.worksheet_info_response(name="Data", table_count=2, chart_count=1))

    result, _ = await excel_roundtrip("get:worksheetInfo", {"worksheet_name": "Data"}, factory)

    assert result.success is True
    assert result.data["name"] == "Data"
    assert result.data["usedRange"]["rowCount"] == 10
    assert result.data["tableCount"] == 2


async def test_get_worksheet_info_omits_worksheet_name(excel_roundtrip, excel_factory):
    """缺省 worksheet_name → wire 上不应出现 worksheetName 键（exclude_none）。"""

    def factory(request: dict) -> dict:
        assert "worksheetName" not in request
        return _ok(request, excel_factory.worksheet_info_response())

    result, _ = await excel_roundtrip("get:worksheetInfo", {}, factory)
    assert result.success is True


async def test_get_selected_range_success(excel_roundtrip, excel_factory):
    """excel:get:selectedRange —— 返回 2D 值数组与行列计数。"""

    def factory(request: dict) -> dict:
        return _ok(request, excel_factory.selected_range_response(values=[["x", "y"], [1, 2]]))

    result, _ = await excel_roundtrip("get:selectedRange", {}, factory)

    assert result.success is True
    assert result.data["values"] == [["x", "y"], [1, 2]]
    assert result.data["rowCount"] == 2
    assert result.data["columnCount"] == 2


async def test_get_worksheet_info_worksheet_not_found(excel_roundtrip, excel_factory):
    """worksheet 不存在 → AddIn 返回 3010 + details.kind（oasp#17），execute() 摊平后 details 存活。"""

    def factory(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": excel_factory.error(
                "3010", "Worksheet 'Ghost' not found", details={"kind": "worksheet", "name": "Ghost"}
            ),
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    result, _ = await excel_roundtrip("get:worksheetInfo", {"worksheet_name": "Ghost"}, factory)

    assert result.success is False
    assert "3010" in (result.error or "")
    # details 经 format_wire_error 存活于错误字符串（issue #82），供 e2e/AI 消费者定位对象类型
    assert "worksheet" in (result.error or "")
