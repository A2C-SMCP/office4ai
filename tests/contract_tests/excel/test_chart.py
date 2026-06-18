"""
Contract tests — #23 Chart 操作 (OASP 0.3.0 /excel Draft)

事件：excel:insert:chart / get:charts / update:chart / delete:chart
重点：insert 的 position 嵌套对象；update 的 properties 偏更新对象；chartType
为开放字符串（无效类型由 AddIn 返回 4002）。insert/update 响应同为 {name}。

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


async def test_insert_chart_with_position(excel_roundtrip, excel_factory):
    """excel:insert:chart —— sourceAddress + chartType + 嵌套 position。"""

    def factory(request: dict) -> dict:
        assert request["sourceAddress"] == "A1:C4"
        assert request["chartType"] == "ColumnClustered"
        assert request["title"] == "Sales"
        assert request["position"] == {"top": 10.0, "left": 20.0, "width": 360.0, "height": 240.0}
        return _ok(request, excel_factory.chart_operation_response(name="Chart1"))

    result, _ = await excel_roundtrip(
        "insert:chart",
        {
            "source_address": "A1:C4",
            "chart_type": "ColumnClustered",
            "title": "Sales",
            "position": {"top": 10.0, "left": 20.0, "width": 360.0, "height": 240.0},
        },
        factory,
    )

    assert result.success is True
    assert result.data == {"name": "Chart1"}


async def test_insert_chart_omits_optional(excel_roundtrip, excel_factory):
    """缺省 title/position → wire 不含这些键。"""

    def factory(request: dict) -> dict:
        assert "title" not in request
        assert "position" not in request
        return _ok(request, excel_factory.chart_operation_response())

    result, _ = await excel_roundtrip("insert:chart", {"source_address": "A1:C4", "chart_type": "Pie"}, factory)
    assert result.success is True


async def test_get_charts_success(excel_roundtrip, excel_factory):
    """excel:get:charts —— 返回 {charts:[ChartSummary]}。"""

    def factory(request: dict) -> dict:
        return _ok(request, excel_factory.get_charts_response())

    result, _ = await excel_roundtrip("get:charts", {}, factory)

    assert result.success is True
    assert result.data["charts"][0]["chartType"] == "ColumnClustered"


async def test_update_chart_partial_properties(excel_roundtrip, excel_factory):
    """excel:update:chart —— properties 偏更新，仅传入键到 wire。"""

    def factory(request: dict) -> dict:
        assert request["chartName"] == "Chart1"
        props = request["properties"]
        assert props["title"] == "Revenue"
        assert props["chartType"] == "Line"
        # 未传 sourceAddress / position → 不应出现
        assert "sourceAddress" not in props
        assert "position" not in props
        return _ok(request, excel_factory.chart_operation_response(name="Chart1"))

    result, _ = await excel_roundtrip(
        "update:chart",
        {"chart_name": "Chart1", "properties": {"title": "Revenue", "chart_type": "Line"}},
        factory,
    )

    assert result.success is True
    assert result.data == {"name": "Chart1"}


async def test_delete_chart_success(excel_roundtrip, excel_factory):
    """excel:delete:chart —— chartName 必填。"""

    def factory(request: dict) -> dict:
        assert request["chartName"] == "Chart1"
        return _ok(request, excel_factory.delete_chart_response(deleted=True))

    result, _ = await excel_roundtrip("delete:chart", {"chart_name": "Chart1"}, factory)

    assert result.success is True
    assert result.data["deleted"] is True


async def test_insert_chart_invalid_type(excel_roundtrip):
    """无效图表类型 → 4002 INVALID_PARAM（由 AddIn 产生，透传）。"""

    def factory(request: dict) -> dict:
        return _err(request, "4002", "Invalid chartType 'Nope'")

    result, _ = await excel_roundtrip("insert:chart", {"source_address": "A1:C4", "chart_type": "Nope"}, factory)
    assert result.success is False
    assert "4002" in (result.error or "")


async def test_update_chart_not_found(excel_roundtrip):
    """图表不存在 → 5007 CHART_NOT_FOUND。"""

    def factory(request: dict) -> dict:
        return _err(request, "5007", "Chart 'Ghost' not found")

    result, _ = await excel_roundtrip("update:chart", {"chart_name": "Ghost", "properties": {"title": "x"}}, factory)
    assert result.success is False
    assert "5007" in (result.error or "")
