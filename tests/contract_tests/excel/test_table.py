"""
Contract tests — #22 Table 操作 (OASP 0.3.0 /excel Draft)

事件：excel:insert:table / get:table / get:tables / add:tableRow /
      delete:tableRow / sort:table
重点：insert:table 的 hasHeaders + 2D data；sort:table 的 sortFields 嵌套列表
（columnIndex + 可选 ascending）的 wire 序列化。

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


async def test_insert_table_with_data(excel_roundtrip, excel_factory):
    """excel:insert:table —— hasHeaders + 2D data + styleName 全程 camelCase。"""

    def factory(request: dict) -> dict:
        assert request["address"] == "A1:C4"
        assert request["hasHeaders"] is True
        assert request["data"] == [["H1", "H2", "H3"], [1, 2, 3]]
        assert request["styleName"] == "TableStyleMedium2"
        return _ok(request, excel_factory.insert_table_response(name="Table1", address="Sheet1!A1:C4"))

    result, _ = await excel_roundtrip(
        "insert:table",
        {
            "address": "A1:C4",
            "has_headers": True,
            "data": [["H1", "H2", "H3"], [1, 2, 3]],
            "style_name": "TableStyleMedium2",
        },
        factory,
    )

    assert result.success is True
    assert result.data == {"name": "Table1", "address": "Sheet1!A1:C4"}


async def test_insert_table_omits_optional(excel_roundtrip, excel_factory):
    """缺省 data/style_name → wire 不含这些键；hasHeaders=False 仍存活（falsy 非 None）。"""

    def factory(request: dict) -> dict:
        assert request["hasHeaders"] is False
        assert "data" not in request
        assert "styleName" not in request
        return _ok(request, excel_factory.insert_table_response())

    result, _ = await excel_roundtrip("insert:table", {"address": "A1:C4", "has_headers": False}, factory)
    assert result.success is True


async def test_get_table_rich_info(excel_roundtrip, excel_factory):
    """excel:get:table —— tableId 必填，返回富信息（含 columns 列明细）。"""

    def factory(request: dict) -> dict:
        assert request["tableId"] == "Table1"
        return _ok(request, excel_factory.get_table_response(name="Table1"))

    result, _ = await excel_roundtrip("get:table", {"table_id": "Table1"}, factory)

    assert result.success is True
    assert result.data["showHeaders"] is True
    assert len(result.data["columns"]) == 3
    assert result.data["columns"][0] == {"name": "Col1", "index": 0}


async def test_get_tables_slim_list(excel_roundtrip, excel_factory):
    """excel:get:tables —— 精简条目 {name, id, address}。"""

    def factory(request: dict) -> dict:
        return _ok(request, excel_factory.get_tables_response())

    result, _ = await excel_roundtrip("get:tables", {"worksheet_name": "Sheet1"}, factory)

    assert result.success is True
    assert result.data["tables"][0]["name"] == "Table1"


async def test_add_table_row_success(excel_roundtrip, excel_factory):
    """excel:add:tableRow —— values 一维数组追加。"""

    def factory(request: dict) -> dict:
        assert request["tableId"] == "Table1"
        assert request["values"] == ["a", "b", "c"]
        return _ok(request, excel_factory.add_table_row_response(table_id="Table1"))

    result, _ = await excel_roundtrip("add:tableRow", {"table_id": "Table1", "values": ["a", "b", "c"]}, factory)

    assert result.success is True
    assert result.data["tableId"] == "Table1"


async def test_delete_table_row_success(excel_roundtrip, excel_factory):
    """excel:delete:tableRow —— rowIndex=0 falsy 必须存活。"""

    def factory(request: dict) -> dict:
        assert request["rowIndex"] == 0
        return _ok(request, excel_factory.delete_table_row_response(deleted=True))

    result, _ = await excel_roundtrip("delete:tableRow", {"table_id": "Table1", "row_index": 0}, factory)

    assert result.success is True
    assert result.data["deleted"] is True


async def test_sort_table_nested_sort_fields(excel_roundtrip, excel_factory):
    """excel:sort:table —— sortFields 嵌套列表 camelCase；ascending 缺省被丢弃。"""

    def factory(request: dict) -> dict:
        fields = request["sortFields"]
        assert fields[0] == {"columnIndex": 1, "ascending": False}
        # 第二个排序键省略 ascending → 不应出现该键（exclude_none）
        assert fields[1] == {"columnIndex": 0}
        return _ok(request, excel_factory.sort_table_response(sorted_=True))

    result, _ = await excel_roundtrip(
        "sort:table",
        {
            "table_id": "Table1",
            "sort_fields": [{"column_index": 1, "ascending": False}, {"column_index": 0}],
        },
        factory,
    )

    assert result.success is True
    assert result.data["sorted"] is True


async def test_get_table_not_found(excel_roundtrip):
    """表格不存在 → 5006 TABLE_NOT_FOUND。"""

    def factory(request: dict) -> dict:
        return _err(request, "5006", "Table 'Ghost' not found")

    result, _ = await excel_roundtrip("get:table", {"table_id": "Ghost"}, factory)
    assert result.success is False
    assert "5006" in (result.error or "")


async def test_add_table_row_type_mismatch(excel_roundtrip):
    """写入值类型不匹配 → 5009 DATA_TYPE_MISMATCH。"""

    def factory(request: dict) -> dict:
        return _err(request, "5009", "Data type mismatch in column 2")

    result, _ = await excel_roundtrip("add:tableRow", {"table_id": "Table1", "values": ["x", "not-a-number"]}, factory)
    assert result.success is False
    assert "5009" in (result.error or "")
