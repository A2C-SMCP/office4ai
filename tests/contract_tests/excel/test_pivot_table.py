"""
Contract tests — #24 PivotTable 操作 (OASP 0.3.0 /excel Draft)

事件：excel:insert:pivotTable / get:pivotTables / delete:pivotTable
透视表以 sourceAddress + targetAddress 放置创建（spec 未定义行/列/值/筛选字段）。

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


async def test_insert_pivot_table_with_name(excel_roundtrip, excel_factory):
    """excel:insert:pivotTable —— sourceAddress + targetAddress + 显式 name。"""

    def factory(request: dict) -> dict:
        assert request["sourceAddress"] == "A1:D100"
        assert request["targetAddress"] == "F1"
        assert request["name"] == "Summary"
        return _ok(request, excel_factory.pivot_table_operation_response(name="Summary"))

    result, _ = await excel_roundtrip(
        "insert:pivotTable",
        {"source_address": "A1:D100", "target_address": "F1", "name": "Summary"},
        factory,
    )

    assert result.success is True
    assert result.data == {"name": "Summary"}


async def test_insert_pivot_table_omits_name(excel_roundtrip, excel_factory):
    """缺省 name → wire 不含 name 键（Excel 自动生成）。"""

    def factory(request: dict) -> dict:
        assert "name" not in request
        return _ok(request, excel_factory.pivot_table_operation_response(name="PivotTable1"))

    result, _ = await excel_roundtrip(
        "insert:pivotTable", {"source_address": "A1:D100", "target_address": "F1"}, factory
    )
    assert result.success is True
    assert result.data["name"] == "PivotTable1"


async def test_get_pivot_tables_success(excel_roundtrip, excel_factory):
    """excel:get:pivotTables —— 返回 {pivotTables:[{name, id}]}。"""

    def factory(request: dict) -> dict:
        return _ok(request, excel_factory.get_pivot_tables_response())

    result, _ = await excel_roundtrip("get:pivotTables", {}, factory)

    assert result.success is True
    assert result.data["pivotTables"][0]["name"] == "PivotTable1"
    assert "id" in result.data["pivotTables"][0]


async def test_delete_pivot_table_success(excel_roundtrip, excel_factory):
    """excel:delete:pivotTable —— pivotTableName 必填。"""

    def factory(request: dict) -> dict:
        assert request["pivotTableName"] == "PivotTable1"
        return _ok(request, excel_factory.delete_pivot_table_response(deleted=True))

    result, _ = await excel_roundtrip("delete:pivotTable", {"pivot_table_name": "PivotTable1"}, factory)

    assert result.success is True
    assert result.data["deleted"] is True


async def test_delete_pivot_table_not_found(excel_roundtrip):
    """透视表不存在 → 5008 PIVOT_NOT_FOUND。"""

    def factory(request: dict) -> dict:
        return _err(request, "5008", "Pivot table 'Ghost' not found")

    result, _ = await excel_roundtrip("delete:pivotTable", {"pivot_table_name": "Ghost"}, factory)
    assert result.success is False
    assert "5008" in (result.error or "")


async def test_insert_pivot_table_not_supported(excel_roundtrip):
    """平台不支持透视表 → 5010 NOT_SUPPORTED。"""

    def factory(request: dict) -> dict:
        return _err(request, "5010", "Pivot tables are not supported on this platform")

    result, _ = await excel_roundtrip(
        "insert:pivotTable", {"source_address": "A1:D100", "target_address": "F1"}, factory
    )
    assert result.success is False
    assert "5010" in (result.error or "")
