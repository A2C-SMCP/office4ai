"""
Excel Worksheet E2E — rename:worksheet + activate:worksheet（改名 / 激活）

覆盖 ``excel:rename:worksheet`` 与 ``excel:activate:worksheet``。

wire 形态（AddIn worksheets.ts 确认）:
- rename:worksheet（currentName + newName）→ ``{name}``（新名）
- activate:worksheet（worksheetName）→ **void**（无 data）

双重验证：
- rename：协议返回 + openpyxl ``sheet_names`` 含新名、不含旧名，隔离表 Delta 保留。
- activate：因 handler 返回 void，无响应体可断言——改用 ``get:worksheets``（pre_op 先 activate）
  断言目标表 isActive，并 best-effort 用 openpyxl ``wb.active`` 佐证。

含错误码用例：rename / activate 不存在的表 → 3000 DOCUMENT_ERROR。

运行方式:
    uv run python manual_tests/excel/worksheet_e2e/test_rename_activate.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.worksheet_e2e._fixtures import ensure_fixtures

BOOK = "worksheet_e2e/book.xlsx"


def validate_rename(data: dict[str, Any], reader: WorkbookReader) -> bool:
    name = data.get("name")
    print(f"   rename 返回: name={name!r}")
    if name != "BetaRenamed":
        print(f"   ❌ 返回新名应为 'BetaRenamed'，实得 {name!r}")
        return False
    reader.reload()
    print(f"   读盘 sheet_names={reader.sheet_names}")
    if "BetaRenamed" not in reader.sheet_names:
        print("   ❌ 新名 'BetaRenamed' 未落盘")
        return False
    if "Beta" in reader.sheet_names:
        print("   ❌ 旧名 'Beta' 仍在（改名应原地替换）")
        return False
    if "Delta" not in reader.sheet_names:
        print("   ❌ 隔离表 'Delta' 丢失（改名不应影响其它表）")
        return False
    print("   ✅ Beta → BetaRenamed 已落盘，旧名消失，Delta 保留")
    return True


def validate_activate_via_get(data: dict[str, Any], reader: WorkbookReader) -> bool:
    sheets = data.get("worksheets") or []
    actives = [s["name"] for s in sheets if s.get("isActive")]
    print(f"   activate 后 isActive 的表: {actives}")
    if actives != ["Gamma"]:
        print(f"   ❌ 激活后应恰有 Gamma 为 active，实得 {actives}")
        return False
    # best-effort：openpyxl 读盘的活动表（存盘视图状态，可能因 Excel 差异不一致，不判失败）
    reader.reload()
    try:
        active_title = reader.wb.active.title
        print(f"   openpyxl wb.active={active_title!r}")
        if active_title != "Gamma":
            print("   ⚠️  openpyxl 读到的活动表非 Gamma（存盘视图差异；以协议 isActive 为准）")
    except Exception as exc:  # noqa: BLE001
        print(f"   ⚠️  openpyxl 活动表读取失败（{exc}）；以协议 isActive 为准")
    print("   ✅ Gamma 已激活（get:worksheets isActive）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="rename 工作表",
        fixture_name=BOOK,
        description="rename:worksheet Beta → BetaRenamed → {name} + 落盘替换，Delta 保留",
        action="rename:worksheet",
        params={"current_name": "Beta", "new_name": "BetaRenamed"},
        validator=validate_rename,
        tags=["rename"],
    ),
    ExcelCase(
        name="activate 工作表",
        fixture_name=BOOK,
        description="先 activate Gamma → get:worksheets 断言 Gamma isActive",
        action="get:worksheets",
        params={},
        pre_ops=[("activate:worksheet", {"worksheet_name": "Gamma"})],
        validator=validate_activate_via_get,
        tags=["activate"],
    ),
    ExcelCase(
        name="错误码 3000 — rename 不存在的表（DOCUMENT_ERROR）",
        fixture_name=BOOK,
        description="rename:worksheet currentName='NoSuch' → 3000",
        action="rename:worksheet",
        params={"current_name": "NoSuchSheet", "new_name": "X"},
        expect_error_code="3000",
        tags=["error"],
    ),
    ExcelCase(
        name="错误码 3000 — activate 不存在的表（DOCUMENT_ERROR）",
        fixture_name=BOOK,
        description="activate:worksheet worksheetName='NoSuch' → 3000",
        action="activate:worksheet",
        params={"worksheet_name": "NoSuchSheet"},
        expect_error_code="3000",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Worksheet E2E — rename:worksheet + activate:worksheet", TEST_CASES)
