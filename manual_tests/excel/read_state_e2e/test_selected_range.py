"""
Excel Read State E2E — get:selectedRange（当前选中范围）

覆盖 ``excel:get:selectedRange``（无业务参数）。读取 Excel **实时选区**，故每个用例
需在运行前**在 Excel 中手动选好对应区域**（README 有指引）。验证返回 data 的结构
不变量（address / 2D values / rowCount×columnCount）。

运行方式（建议 --no-auto-open，便于手动选区后回车继续）:
    uv run python manual_tests/excel/read_state_e2e/test_selected_range.py --test 2 --no-auto-open
    uv run python manual_tests/excel/read_state_e2e/test_selected_range.py --test all --no-auto-open
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.read_state_e2e._fixtures import ensure_fixtures

PREFILLED = "read_state_e2e/prefilled.xlsx"


def _shape_ok(data: dict[str, Any], exp_rows: int | None, exp_cols: int | None) -> bool:
    values = data.get("values")
    rc, cc = data.get("rowCount"), data.get("columnCount")
    print(f"   address={data.get('address')!r} rowCount={rc} columnCount={cc}")
    print(f"   values={values}")
    if not isinstance(values, list) or (values and not isinstance(values[0], list)):
        print("   ❌ values 不是 2D 数组")
        return False
    if rc != len(values):
        print(f"   ❌ rowCount {rc} ≠ len(values) {len(values)}")
        return False
    if values and cc != len(values[0]):
        print(f"   ❌ columnCount {cc} ≠ len(values[0]) {len(values[0])}")
        return False
    if exp_rows is not None and rc != exp_rows:
        print(f"   ⚠️  预期 {exp_rows} 行，实际 {rc}（请确认已按 README 选好区域）")
    if exp_cols is not None and cc != exp_cols:
        print(f"   ⚠️  预期 {exp_cols} 列，实际 {cc}（请确认已按 README 选好区域）")
    return True


def validate_single_cell(data: dict[str, Any]) -> bool:
    ok = _shape_ok(data, 1, 1)
    if ok:
        print("   ✅ 单元格选区结构正确")
    return ok


def validate_2d_values(data: dict[str, Any]) -> bool:
    ok = _shape_ok(data, 2, 3)
    if ok:
        print("   ✅ 多单元格 2D values 结构正确")
    return ok


def validate_empty_selection(data: dict[str, Any]) -> bool:
    ok = _shape_ok(data, 1, 1)
    if not ok:
        return False
    cell = (data.get("values") or [[None]])[0][0]
    print(f"   空选区首格值={cell!r}（预期 None / 空串）")
    print("   ✅ 空选区返回结构正确")
    return True


def validate_mixed_types(data: dict[str, Any]) -> bool:
    ok = _shape_ok(data, 1, 3)
    if not ok:
        return False
    flat = [v for row in (data.get("values") or []) for v in row]
    kinds = {type(v).__name__ for v in flat}
    print(f"   选区值={flat}，类型集={kinds}")
    if len(kinds) < 2:
        print(f"   ⚠️  类型种类少（{kinds}）；请选含 字符串/数字/布尔 的 A1:C1")
    print("   ✅ 混合类型选区已读取")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="单元格选区",
        fixture_name=PREFILLED,
        description="请先在 Excel 选中 A1（单格），验证 1x1 结构",
        action="get:selectedRange",
        validator=validate_single_cell,
        tags=["basic"],
    ),
    ExcelCase(
        name="多单元格 2D values",
        fixture_name=PREFILLED,
        description="请先选中 A1:C2，验证 2 行 3 列 2D values",
        action="get:selectedRange",
        validator=validate_2d_values,
        tags=["2d"],
    ),
    ExcelCase(
        name="空选区",
        fixture_name=PREFILLED,
        description="请先选中空白单元格 F10，验证空值结构",
        action="get:selectedRange",
        validator=validate_empty_selection,
        tags=["empty"],
    ),
    ExcelCase(
        name="混合类型值",
        fixture_name=PREFILLED,
        description="请先选中 A1:C1（字符串/数字/布尔），验证混合类型",
        action="get:selectedRange",
        validator=validate_mixed_types,
        tags=["mixed"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Read State E2E — get:selectedRange", TEST_CASES)
