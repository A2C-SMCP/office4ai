"""
Contract tests — #20 Format / 条件格式 / 合并单元格 (OASP 0.3.0 /excel Draft)

事件：excel:get:rangeFormat / set:rangeFormat / add:conditionalFormat /
      clear:conditionalFormat / merge:cells / unmerge:cells
五个写操作统一返回 {address}；get:rangeFormat 返回 {address, format}。
重点：set:rangeFormat 的嵌套偏更新载荷（font/fill/borders/alignment/numberFormat）
与 add:conditionalFormat 的透传 rule（extra=allow）的 wire 序列化。

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


async def test_get_range_format_success(excel_roundtrip, excel_factory):
    """excel:get:rangeFormat —— 返回 {address, format}（复用 RangeFormatInfo）。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "A1:C3"
        return _ok(request, excel_factory.get_range_format_response(address="Sheet1!A1:C3"))

    result, _ = await excel_roundtrip("get:rangeFormat", {"address": "A1:C3"}, factory)

    assert result.success is True
    assert result.data["address"] == "Sheet1!A1:C3"
    assert result.data["format"]["wrapText"] is False


async def test_set_range_format_nested_partial_update(excel_roundtrip, excel_factory):
    """excel:set:rangeFormat —— 嵌套偏更新载荷全程 camelCase（font/alignment/numberFormat）。"""

    def factory(request: dict) -> dict:
        fmt = request["format"]
        assert fmt["font"]["bold"] is True
        assert fmt["alignment"]["horizontal"] == "Center"
        assert fmt["alignment"]["wrapText"] is True  # snake wrap_text → wrapText
        assert fmt["numberFormat"] == "0.00"
        # 未传入的属性不应出现（exclude_none 偏更新语义）
        assert "fill" not in fmt
        assert "borders" not in fmt
        return _ok(request, excel_factory.range_operation_response(address="Sheet1!A1:C3"))

    result, _ = await excel_roundtrip(
        "set:rangeFormat",
        {
            "address": "A1:C3",
            "format": {
                "font": {"bold": True},
                "alignment": {"horizontal": "Center", "wrap_text": True},
                "number_format": "0.00",
            },
        },
        factory,
    )

    assert result.success is True
    assert result.data == {"address": "Sheet1!A1:C3"}


async def test_add_conditional_format_passthrough_rule(excel_roundtrip, excel_factory):
    """excel:add:conditionalFormat —— rule 透传（extra=allow），附加键原样到 wire。"""

    def factory(request: dict) -> dict:
        rule = request["rule"]
        assert rule["type"] == "cellValue"
        # 透传的附加键应原样保留
        assert rule["operator"] == "greaterThan"
        assert rule["formula1"] == "100"
        return _ok(request, excel_factory.range_operation_response(address="Sheet1!B2:B100"))

    result, _ = await excel_roundtrip(
        "add:conditionalFormat",
        {
            "address": "B2:B100",
            "rule": {"type": "cellValue", "operator": "greaterThan", "formula1": "100"},
        },
        factory,
    )

    assert result.success is True


async def test_clear_conditional_format_success(excel_roundtrip, excel_factory):
    """excel:clear:conditionalFormat —— 清除范围条件格式。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "B2:B100"
        return _ok(request, excel_factory.range_operation_response(address="Sheet1!B2:B100"))

    result, _ = await excel_roundtrip("clear:conditionalFormat", {"address": "B2:B100"}, factory)
    assert result.success is True


async def test_merge_cells_success(excel_roundtrip, excel_factory):
    """excel:merge:cells —— 合并区域，返回 {address}。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "A1:C1"
        return _ok(request, excel_factory.range_operation_response(address="Sheet1!A1:C1"))

    result, _ = await excel_roundtrip("merge:cells", {"address": "A1:C1"}, factory)

    assert result.success is True
    assert result.data["address"] == "Sheet1!A1:C1"


async def test_unmerge_cells_success(excel_roundtrip, excel_factory):
    """excel:unmerge:cells —— 取消合并。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "A1:C1"
        return _ok(request, excel_factory.range_operation_response(address="Sheet1!A1:C1"))

    result, _ = await excel_roundtrip("unmerge:cells", {"address": "A1:C1"}, factory)
    assert result.success is True


async def test_merge_cells_conflict(excel_roundtrip):
    """合并与现有合并区域冲突 → 5003 MERGE_CONFLICT。"""

    def factory(request: dict) -> dict:
        return _err(request, "5003", "Merge conflicts with existing merged region")

    result, _ = await excel_roundtrip("merge:cells", {"address": "A1:C1"}, factory)
    assert result.success is False
    assert "5003" in (result.error or "")
