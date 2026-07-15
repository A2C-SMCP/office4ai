"""
Contract tests — #21 Worksheet 管理 (OASP 0.3.0 /excel Draft)

事件：excel:get:worksheets / add:worksheet / delete:worksheet /
      rename:worksheet / activate:worksheet
字段形态各事件不同（add.name 可选；delete/activate.worksheetName 必填；
rename.currentName+newName 必填）。

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


async def test_get_worksheets_success(excel_roundtrip, excel_factory):
    """excel:get:worksheets —— 无业务参数，返回 {worksheets:[SheetInfo]}。"""

    def factory(request: dict) -> dict:
        assert set(request.keys()) <= {"requestId", "documentUri", "timestamp"}
        return _ok(request, excel_factory.get_worksheets_response())

    result, _ = await excel_roundtrip("get:worksheets", {}, factory)

    assert result.success is True
    assert len(result.data["worksheets"]) == 2
    assert result.data["worksheets"][0]["isActive"] is True


async def test_add_worksheet_with_name(excel_roundtrip, excel_factory):
    """excel:add:worksheet —— 显式 name 在 wire 上保持。"""

    def factory(request: dict) -> dict:
        assert request["name"] == "Report"
        return _ok(request, excel_factory.add_worksheet_response(name="Report", index=2))

    result, _ = await excel_roundtrip("add:worksheet", {"name": "Report"}, factory)

    assert result.success is True
    assert result.data == {"name": "Report", "index": 2}


async def test_add_worksheet_omits_name(excel_roundtrip, excel_factory):
    """缺省 name → wire 不含 name 键（Excel 自动命名）。"""

    def factory(request: dict) -> dict:
        assert "name" not in request
        return _ok(request, excel_factory.add_worksheet_response(name="Sheet3", index=2))

    result, _ = await excel_roundtrip("add:worksheet", {}, factory)
    assert result.success is True
    assert result.data["name"] == "Sheet3"


async def test_delete_worksheet_success(excel_roundtrip, excel_factory):
    """excel:delete:worksheet —— worksheetName 必填。"""

    def factory(request: dict) -> dict:
        assert request["worksheetName"] == "Sheet2"
        return _ok(request, excel_factory.delete_worksheet_response(deleted=True))

    result, _ = await excel_roundtrip("delete:worksheet", {"worksheet_name": "Sheet2"}, factory)

    assert result.success is True
    assert result.data["deleted"] is True


async def test_rename_worksheet_success(excel_roundtrip, excel_factory):
    """excel:rename:worksheet —— currentName + newName 必填，均转 camelCase。"""

    def factory(request: dict) -> dict:
        assert request["currentName"] == "Sheet1"
        assert request["newName"] == "Summary"
        return _ok(request, excel_factory.rename_worksheet_response(name="Summary"))

    result, _ = await excel_roundtrip("rename:worksheet", {"current_name": "Sheet1", "new_name": "Summary"}, factory)

    assert result.success is True
    assert result.data["name"] == "Summary"


async def test_activate_worksheet_success(excel_roundtrip, excel_factory):
    """excel:activate:worksheet —— worksheetName 必填。"""

    def factory(request: dict) -> dict:
        assert request["worksheetName"] == "Sheet2"
        return _ok(request, excel_factory.activate_worksheet_response(activated=True))

    result, _ = await excel_roundtrip("activate:worksheet", {"worksheet_name": "Sheet2"}, factory)

    assert result.success is True
    assert result.data["activated"] is True


async def test_delete_worksheet_not_found(excel_roundtrip):
    """删除不存在的工作表 → 3010 ELEMENT_NOT_FOUND(kind:worksheet)（oasp#17）。"""

    def factory(request: dict) -> dict:
        return _err(request, "3010", "Worksheet 'Ghost' not found")

    result, _ = await excel_roundtrip("delete:worksheet", {"worksheet_name": "Ghost"}, factory)
    assert result.success is False
    assert "3010" in (result.error or "")
