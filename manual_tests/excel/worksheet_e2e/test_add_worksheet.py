"""
Excel Worksheet E2E — add:worksheet + get:worksheets（新增 / 列举）

覆盖 ``excel:add:worksheet``（命名 / 自动命名）与 ``excel:get:worksheets``。

wire 形态（AddIn worksheets.ts 确认）:
- add:worksheet → ``{name, index}``（name 省略时 Excel 自动命名 "SheetN"）
- get:worksheets → ``{worksheets: [{name, index, isActive, isHidden}]}``

双重验证：协议返回 + openpyxl ``sheet_names`` 读盘核对新表已落盘。
含错误码用例：add 重名 → 3004 OPERATION_FAILED（名称已存在，oasp#17 / events-excel.md
add:worksheet；Add-In 接线前 XFAIL）。

运行方式:
    uv run python manual_tests/excel/worksheet_e2e/test_add_worksheet.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import PENDING_ADDIN_80, ExcelCase, run_main
from manual_tests.excel.worksheet_e2e._fixtures import ensure_fixtures

BOOK = "worksheet_e2e/book.xlsx"

# 夹具固有 4 张表（_fixtures._SHEETS 的 key 顺序）。
FIXTURE_SHEETS = ["Alpha", "Beta", "Gamma", "Delta"]


def validate_add_named(data: dict[str, Any], reader: WorkbookReader) -> bool:
    name, index = data.get("name"), data.get("index")
    print(f"   add 返回: name={name!r} index={index}")
    if name != "Summary":
        print(f"   ❌ 新表名应为 'Summary'，实得 {name!r}")
        return False
    if not isinstance(index, int):
        print(f"   ❌ index 应为整数，实得 {index!r}")
        return False
    reader.reload()
    print(f"   读盘 sheet_names={reader.sheet_names}")
    if "Summary" not in reader.sheet_names:
        print("   ❌ 新表 'Summary' 未落盘")
        return False
    print("   ✅ 命名新表已创建并落盘")
    return True


def validate_add_auto(data: dict[str, Any], reader: WorkbookReader) -> bool:
    name = data.get("name")
    print(f"   add（无 name）返回: name={name!r} index={data.get('index')}")
    if not name:
        print("   ❌ 自动命名应返回非空 name")
        return False
    reader.reload()
    print(f"   读盘 sheet_names={reader.sheet_names}")
    if name not in reader.sheet_names:
        print(f"   ❌ 自动命名表 {name!r} 未落盘")
        return False
    if len(reader.sheet_names) != len(FIXTURE_SHEETS) + 1:
        print(f"   ❌ 表数应为 {len(FIXTURE_SHEETS) + 1}，实得 {len(reader.sheet_names)}")
        return False
    print(f"   ✅ 自动命名表 {name!r} 已创建并落盘")
    return True


def validate_get_worksheets(data: dict[str, Any]) -> bool:
    sheets = data.get("worksheets")
    if not isinstance(sheets, list) or not sheets:
        print(f"   ❌ worksheets 非非空列表: {sheets!r}")
        return False
    names = [s.get("name") for s in sheets]
    print(f"   worksheets={names}")
    required = {"name", "index", "isActive", "isHidden"}
    for s in sheets:
        missing = required - set(s.keys())
        if missing:
            print(f"   ❌ 条目缺字段 {missing}: {s}")
            return False
    for fx in FIXTURE_SHEETS:
        if fx not in names:
            print(f"   ❌ 缺夹具表 {fx!r}")
            return False
    actives = [s["name"] for s in sheets if s.get("isActive")]
    print(f"   isActive 的表: {actives}")
    if actives != ["Alpha"]:
        print(f"   ❌ 应恰有 Alpha 为 active，实得 {actives}")
        return False
    print("   ✅ get:worksheets 字段完整，Alpha 为唯一 active")
    return True


def validate_get_after_add(data: dict[str, Any]) -> bool:
    sheets = data.get("worksheets") or []
    names = [s.get("name") for s in sheets]
    print(f"   add 后 worksheets={names}")
    if "Extra" not in names:
        print("   ❌ 新增的 'Extra' 未出现在 get:worksheets 结果")
        return False
    if len(sheets) != len(FIXTURE_SHEETS) + 1:
        print(f"   ❌ 表数应为 {len(FIXTURE_SHEETS) + 1}，实得 {len(sheets)}")
        return False
    print("   ✅ get:worksheets 实时反映新增的 'Extra'")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="add 命名工作表",
        fixture_name=BOOK,
        description="add:worksheet name='Summary' → {name,index} + 落盘",
        action="add:worksheet",
        params={"name": "Summary"},
        validator=validate_add_named,
        tags=["add"],
    ),
    ExcelCase(
        name="add 自动命名工作表",
        fixture_name=BOOK,
        description="add:worksheet（省略 name）→ Excel 自动命名并落盘，表数 +1",
        action="add:worksheet",
        params={},
        validator=validate_add_auto,
        tags=["add", "auto"],
    ),
    ExcelCase(
        name="get:worksheets 列举",
        fixture_name=BOOK,
        description="get:worksheets → 4 张夹具表，字段完整，Alpha 唯一 active",
        action="get:worksheets",
        params={},
        validator=validate_get_worksheets,
        tags=["get"],
    ),
    ExcelCase(
        name="get:worksheets 反映新增",
        fixture_name=BOOK,
        description="先 add 'Extra' → get:worksheets 实时含 'Extra'（表数 5）",
        action="get:worksheets",
        params={},
        pre_ops=[("add:worksheet", {"name": "Extra"})],
        validator=validate_get_after_add,
        tags=["get", "add"],
    ),
    ExcelCase(
        name="错误码 3004 — add 重名工作表（OPERATION_FAILED）",
        fixture_name=BOOK,
        description="add:worksheet name='Alpha'（已存在）→ 3004（oasp#17，events-excel.md add:worksheet「名称已存在」）",
        action="add:worksheet",
        params={"name": "Alpha"},
        expect_error_code="3004",
        xfail_reason=PENDING_ADDIN_80,
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Worksheet E2E — add:worksheet + get:worksheets", TEST_CASES)
