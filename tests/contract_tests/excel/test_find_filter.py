"""
Contract tests — #25 Find&Filter 查找与筛选 (OASP 0.3.0 /excel Draft)

事件：excel:find:values / set:autoFilter / clear:autoFilter
重点：find:values 的 matchCase/matchEntireCell falsy 标志存活；set:autoFilter 的
criteria 嵌套列表（columnIndex + filterOn 开放字符串 + 可选 values）的 wire 序列化。

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


async def test_find_values_minimal(excel_roundtrip, excel_factory):
    """excel:find:values —— 仅 searchText，optionals 不出现在 wire。"""

    def factory(request: dict) -> dict:
        assert request["searchText"] == "foo"
        assert "address" not in request
        assert "matchCase" not in request
        assert "matchEntireCell" not in request
        return _ok(request, excel_factory.find_values_response())

    result, _ = await excel_roundtrip("find:values", {"search_text": "foo"}, factory)

    assert result.success is True
    assert len(result.data["matches"]) == 2
    assert result.data["matches"][0]["address"] == "Sheet1!A2"


async def test_find_values_keeps_false_flags(excel_roundtrip, excel_factory):
    """match_case / match_entire_cell = False 必须存活到 wire（falsy 非 None）。"""

    def factory(request: dict) -> dict:
        assert request["matchCase"] is False
        assert request["matchEntireCell"] is False
        assert request["address"] == "A1:Z100"
        return _ok(request, excel_factory.find_values_response(matches=[]))

    result, _ = await excel_roundtrip(
        "find:values",
        {"search_text": "bar", "address": "A1:Z100", "match_case": False, "match_entire_cell": False},
        factory,
    )

    assert result.success is True
    assert result.data["matches"] == []


async def test_find_values_any_value_type(excel_roundtrip, excel_factory):
    """命中 value 为任意类型（数字 / 字符串 / 布尔）。"""

    def factory(request: dict) -> dict:
        return _ok(
            request,
            excel_factory.find_values_response(
                matches=[
                    {"address": "Sheet1!A1", "value": 3.14},
                    {"address": "Sheet1!A2", "value": True},
                    {"address": "Sheet1!A3", "value": "text"},
                ]
            ),
        )

    result, _ = await excel_roundtrip("find:values", {"search_text": "x"}, factory)

    assert result.success is True
    values = [m["value"] for m in result.data["matches"]]
    assert values == [3.14, True, "text"]


async def test_set_auto_filter_nested_criteria(excel_roundtrip, excel_factory):
    """excel:set:autoFilter —— criteria 嵌套列表 camelCase；filterOn 开放字符串。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "A1:C10"
        criteria = request["criteria"]
        assert criteria[0] == {"columnIndex": 0, "filterOn": "Values", "values": ["x", "y"]}
        # 第二个 criterion 省略 values（按颜色筛选）→ 不应出现 values 键
        assert criteria[1] == {"columnIndex": 2, "filterOn": "CellColor"}
        return _ok(request, excel_factory.set_auto_filter_response(address="Sheet1!A1:C10"))

    result, _ = await excel_roundtrip(
        "set:autoFilter",
        {
            "address": "A1:C10",
            "criteria": [
                {"column_index": 0, "filter_on": "Values", "values": ["x", "y"]},
                {"column_index": 2, "filter_on": "CellColor"},
            ],
        },
        factory,
    )

    assert result.success is True
    assert result.data["address"] == "Sheet1!A1:C10"


async def test_clear_auto_filter_success(excel_roundtrip, excel_factory):
    """excel:clear:autoFilter —— 无业务参数（worksheet_name 可选缺省）。"""

    def factory(request: dict) -> dict:
        assert "worksheetName" not in request
        return _ok(request, excel_factory.clear_auto_filter_response(cleared=True))

    result, _ = await excel_roundtrip("clear:autoFilter", {}, factory)

    assert result.success is True
    assert result.data["cleared"] is True


async def test_set_auto_filter_column_out_of_range(excel_roundtrip):
    """列索引越界 → 4004 PARAM_OUT_OF_RANGE。"""

    def factory(request: dict) -> dict:
        return _err(request, "4004", "Column index 99 out of range")

    result, _ = await excel_roundtrip(
        "set:autoFilter",
        {"address": "A1:C10", "criteria": [{"column_index": 99, "filter_on": "Values"}]},
        factory,
    )
    assert result.success is False
    assert "4004" in (result.error or "")


async def test_find_values_worksheet_not_found(excel_roundtrip):
    """工作表不存在 → 5001 WORKSHEET_NOT_FOUND。"""

    def factory(request: dict) -> dict:
        return _err(request, "5001", "Worksheet 'Ghost' not found")

    result, _ = await excel_roundtrip("find:values", {"search_text": "x", "worksheet_name": "Ghost"}, factory)
    assert result.success is False
    assert "5001" in (result.error or "")
