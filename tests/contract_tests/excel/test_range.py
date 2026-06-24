"""
Contract tests — #19 Range CRUD + 公式 (OASP 0.3.0 /excel Draft)

事件：excel:get:range / set:range / clear:range / copy:range / delete:range /
      insert:range / set:formula
写操作统一返回 {address}（RangeOperationResult）。

OASP Spec: https://doc.turingfocus.cn/oasp/0.3.0/specification/events-excel/
"""

from __future__ import annotations

import asyncio

import pytest

pytestmark = [pytest.mark.asyncio, pytest.mark.contract]


def _ok(request: dict, data: dict) -> dict:
    return {
        "requestId": request["requestId"],
        "success": True,
        "data": data,
        "timestamp": int(asyncio.get_event_loop().time() * 1000),
    }


def _err(request: dict, code: str, message: str) -> dict:
    return {
        "requestId": request["requestId"],
        "success": False,
        "error": {"code": code, "message": message},
        "timestamp": int(asyncio.get_event_loop().time() * 1000),
    }


async def test_get_range_without_format(excel_roundtrip, excel_factory):
    """excel:get:range —— includeFormat 默认 false 仍出现在 wire（非 None）。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "A1:B2"
        assert request["includeFormat"] is False  # default False 不被 exclude_none 丢弃
        return _ok(request, excel_factory.get_range_response(values=[["a", "b"], [1, 2]]))

    result, events = await excel_roundtrip("get:range", {"address": "A1:B2"}, factory)

    assert result.success is True
    assert result.data["values"] == [["a", "b"], [1, 2]]
    assert "format" not in result.data
    assert events[0][0] == "excel:get:range"


async def test_get_range_with_format_and_worksheet(excel_roundtrip, excel_factory):
    """include_format=true + worksheet_name → wire camelCase；响应含 RangeFormatInfo。"""

    def factory(request: dict) -> dict:
        assert request["includeFormat"] is True
        assert request["worksheetName"] == "Sheet2"
        return _ok(request, excel_factory.get_range_response(include_format=True))

    result, _ = await excel_roundtrip(
        "get:range", {"address": "A1:C3", "include_format": True, "worksheet_name": "Sheet2"}, factory
    )

    assert result.success is True
    assert result.data["format"]["font"]["name"] == "Calibri"


async def test_set_range_2d_values(excel_roundtrip, excel_factory):
    """excel:set:range —— 二维数组写入，最小返回 {address}。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "A1:B2"
        assert request["values"] == [[1, 2], [3, 4]]
        return _ok(request, excel_factory.range_operation_response(address="Sheet1!A1:B2"))

    result, _ = await excel_roundtrip("set:range", {"address": "A1:B2", "values": [[1, 2], [3, 4]]}, factory)

    assert result.success is True
    assert result.data == {"address": "Sheet1!A1:B2"}


async def test_set_range_scalar_fill(excel_roundtrip, excel_factory):
    """标量值填充整个范围（values 为标量，透传不被改写）。"""

    def factory(request: dict) -> dict:
        assert request["values"] == 0  # falsy scalar 必须存活（非 None）
        return _ok(request, excel_factory.range_operation_response())

    result, _ = await excel_roundtrip("set:range", {"address": "A1:Z100", "values": 0}, factory)
    assert result.success is True


async def test_clear_range_contents(excel_roundtrip, excel_factory):
    """excel:clear:range —— clearType 字面量在 wire 上保持。"""

    def factory(request: dict) -> dict:
        assert request["clearType"] == "contents"
        return _ok(request, excel_factory.range_operation_response())

    result, _ = await excel_roundtrip("clear:range", {"address": "A1:B2", "clear_type": "contents"}, factory)
    assert result.success is True


async def test_copy_range_cross_sheet(excel_roundtrip, excel_factory):
    """excel:copy:range —— source/target 地址转 camelCase，支持跨 sheet 引用。"""

    def factory(request: dict) -> dict:
        assert request["sourceAddress"] == "Sheet1!A1:B2"
        assert request["targetAddress"] == "Sheet2!A1"
        return _ok(request, excel_factory.range_operation_response(address="Sheet2!A1:B2"))

    result, _ = await excel_roundtrip(
        "copy:range", {"source_address": "Sheet1!A1:B2", "target_address": "Sheet2!A1"}, factory
    )

    assert result.success is True
    assert result.data["address"] == "Sheet2!A1:B2"


async def test_delete_range_shift_up(excel_roundtrip, excel_factory):
    """excel:delete:range —— shiftDirection 'up'。"""

    def factory(request: dict) -> dict:
        assert request["shiftDirection"] == "up"
        return _ok(request, excel_factory.range_operation_response())

    result, _ = await excel_roundtrip("delete:range", {"address": "A1:A5", "shift_direction": "up"}, factory)
    assert result.success is True


async def test_insert_range_shift_right(excel_roundtrip, excel_factory):
    """excel:insert:range —— shiftDirection 'right'。"""

    def factory(request: dict) -> dict:
        assert request["shiftDirection"] == "right"
        return _ok(request, excel_factory.range_operation_response())

    result, _ = await excel_roundtrip("insert:range", {"address": "B1:B5", "shift_direction": "right"}, factory)
    assert result.success is True


async def test_set_formula_success(excel_roundtrip, excel_factory):
    """excel:set:formula —— 公式字符串透传。"""

    def factory(request: dict) -> dict:
        assert request["formula"] == "=SUM(A1:A10)"
        return _ok(request, excel_factory.range_operation_response(address="Sheet1!B1"))

    result, _ = await excel_roundtrip("set:formula", {"address": "B1", "formula": "=SUM(A1:A10)"}, factory)

    assert result.success is True
    assert result.data["address"] == "Sheet1!B1"


async def test_get_range_invalid_address(excel_roundtrip):
    """无效区域地址 → 5002 RANGE_INVALID。"""

    def factory(request: dict) -> dict:
        return _err(request, "5002", "Invalid range address 'ZZZ'")

    result, _ = await excel_roundtrip("get:range", {"address": "ZZZ"}, factory)
    assert result.success is False
    assert "5002" in (result.error or "")


async def test_set_formula_error(excel_roundtrip):
    """公式语法错误 → 5005 FORMULA_ERROR。"""

    def factory(request: dict) -> dict:
        return _err(request, "5005", "Formula syntax error")

    result, _ = await excel_roundtrip("set:formula", {"address": "B1", "formula": "=BADFUNC("}, factory)
    assert result.success is False
    assert "5005" in (result.error or "")


async def test_set_range_protected_sheet(excel_roundtrip):
    """受保护工作表写入 → 5004 PROTECTED_SHEET。"""

    def factory(request: dict) -> dict:
        return _err(request, "5004", "Worksheet is protected")

    result, _ = await excel_roundtrip("set:range", {"address": "A1", "values": [["x"]]}, factory)
    assert result.success is False
    assert "5004" in (result.error or "")
