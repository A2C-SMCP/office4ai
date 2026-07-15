"""
Excel Worksheet E2E — delete:worksheet（删除）

覆盖 ``excel:delete:worksheet``。

wire 形态（AddIn worksheets.ts 确认）:
- delete:worksheet（worksheetName）→ **void**（无 data；office4ai DTO 声明的 ``{deleted}``
  真机不出现，故不能断言响应体）。

双重验证以 openpyxl ``sheet_names`` 为唯一依据：删除后目标表消失、其它表保留、表数 -1。
``get:worksheets``（pre_op 先 delete）做协议层二次确认。

含错误码用例：delete 不存在的表 → 3010 ELEMENT_NOT_FOUND(kind:worksheet)（oasp#17 定案；
Add-In 接线前 XFAIL）。

运行方式:
    uv run python manual_tests/excel/worksheet_e2e/test_delete_worksheet.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.worksheet_e2e._fixtures import ensure_fixtures

BOOK = "worksheet_e2e/book.xlsx"

FIXTURE_SHEETS = ["Alpha", "Beta", "Gamma", "Delta"]


def validate_delete(data: dict[str, Any], reader: WorkbookReader) -> bool:
    # delete 返回 void：data 无可断言字段，只读盘核对。
    reader.reload()
    names = reader.sheet_names
    print(f"   删除 Beta 后 sheet_names={names}")
    if "Beta" in names:
        print("   ❌ 目标表 'Beta' 仍在")
        return False
    for keep in ["Alpha", "Gamma", "Delta"]:
        if keep not in names:
            print(f"   ❌ 非目标表 {keep!r} 被误删")
            return False
    if len(names) != len(FIXTURE_SHEETS) - 1:
        print(f"   ❌ 表数应为 {len(FIXTURE_SHEETS) - 1}，实得 {len(names)}")
        return False
    print("   ✅ Beta 已删除，其余 3 表保留")
    return True


def validate_get_after_delete(data: dict[str, Any]) -> bool:
    sheets = data.get("worksheets") or []
    names = [s.get("name") for s in sheets]
    print(f"   delete 后 worksheets={names}")
    if "Gamma" in names:
        print("   ❌ 已删除的 'Gamma' 仍出现在 get:worksheets")
        return False
    if len(sheets) != len(FIXTURE_SHEETS) - 1:
        print(f"   ❌ 表数应为 {len(FIXTURE_SHEETS) - 1}，实得 {len(sheets)}")
        return False
    print("   ✅ get:worksheets 实时反映 Gamma 已删除")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="delete 工作表",
        fixture_name=BOOK,
        description="delete:worksheet Beta → 读盘 Beta 消失，Alpha/Gamma/Delta 保留",
        action="delete:worksheet",
        params={"worksheet_name": "Beta"},
        validator=validate_delete,
        tags=["delete"],
    ),
    ExcelCase(
        name="delete 后 get:worksheets 反映",
        fixture_name=BOOK,
        description="先 delete Gamma → get:worksheets 不再含 Gamma（表数 3）",
        action="get:worksheets",
        params={},
        pre_ops=[("delete:worksheet", {"worksheet_name": "Gamma"})],
        validator=validate_get_after_delete,
        tags=["delete", "get"],
    ),
    ExcelCase(
        name="错误码 3010 — delete 不存在的表（ELEMENT_NOT_FOUND kind:worksheet）",
        fixture_name=BOOK,
        description="delete:worksheet worksheetName='NoSuch' → 3010 + details.kind=worksheet（oasp#17）",
        action="delete:worksheet",
        params={"worksheet_name": "NoSuchSheet"},
        expect_error_code="3010",
        expect_error_details={"kind": "worksheet"},
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Worksheet E2E — delete:worksheet", TEST_CASES)
