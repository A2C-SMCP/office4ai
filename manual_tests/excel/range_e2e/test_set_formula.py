"""
Excel Range E2E — set:formula（设置单元格公式）

覆盖 ``excel:set:formula``（透传公式字符串，含前导 '='）。在 Data 表数值区域写入
SUM / 算术 / 引用类公式，openpyxl（data_only=False）读盘核对单元格保存的**公式字符串**。
含错误码用例：非法 address → ``3009 RANGE_INVALID``（真实码，旧 DoD 5002 已过时）。

注意：openpyxl 不计算公式，读到的是公式串本身；Excel 可能做轻度规范化（大小写/空格），
故断言对「去空格 + 大写」后的子串做包含匹配，避免被格式差异误伤。

运行方式:
    uv run python manual_tests/excel/range_e2e/test_set_formula.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.range_e2e._fixtures import ensure_fixtures

GRID = "range_e2e/grid.xlsx"


def _norm(s: Any) -> str:
    """去空格 + 大写，便于容忍 Excel 公式规范化差异。"""
    return str(s).replace(" ", "").upper()


def _formula_at(reader: WorkbookReader, cell: str) -> Any:
    reader.reload()
    return reader.cell_value("Data", cell)


def _check(reader: WorkbookReader, cell: str, expect_fragment: str) -> bool:
    got = _formula_at(reader, cell)
    print(f"   openpyxl Data!{cell} = {got!r}")
    if got is None:
        print("   ❌ 单元格未写入公式（None）")
        return False
    if _norm(expect_fragment) not in _norm(got):
        print(f"   ❌ 公式不含预期片段 {expect_fragment!r}")
        return False
    print(f"   ✅ 公式已落盘（含 {expect_fragment!r}）")
    return True


def validate_sum(data: dict[str, Any], reader: WorkbookReader) -> bool:
    return _check(reader, "E2", "SUM(B2:C2)")


def validate_arith(data: dict[str, Any], reader: WorkbookReader) -> bool:
    return _check(reader, "E3", "B3+C3")


def validate_ref(data: dict[str, Any], reader: WorkbookReader) -> bool:
    return _check(reader, "E4", "B4*2")


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="SUM 公式",
        fixture_name=GRID,
        description="set:formula Data!E2 = '=SUM(B2:C2)' → 公式落盘",
        action="set:formula",
        params={"address": "E2", "formula": "=SUM(B2:C2)", "worksheet_name": "Data"},
        validator=validate_sum,
        tags=["sum"],
    ),
    ExcelCase(
        name="算术公式",
        fixture_name=GRID,
        description="set:formula Data!E3 = '=B3+C3' → 公式落盘",
        action="set:formula",
        params={"address": "E3", "formula": "=B3+C3", "worksheet_name": "Data"},
        validator=validate_arith,
        tags=["arith"],
    ),
    ExcelCase(
        name="引用/乘法公式",
        fixture_name=GRID,
        description="set:formula Data!E4 = '=B4*2' → 公式落盘",
        action="set:formula",
        params={"address": "E4", "formula": "=B4*2", "worksheet_name": "Data"},
        validator=validate_ref,
        tags=["ref"],
    ),
    ExcelCase(
        # 真机实测：非法 address → 3000 DOCUMENT_ERROR（详见 test_set_get_range 用例 6 注释 /
        # README「错误码现实」）。3009 RANGE_INVALID 为 dead code（无 handler 发射）。
        name="错误码 3000 — 非法 address（DOCUMENT_ERROR）",
        fixture_name=GRID,
        description="set:formula 传非法地址 → 3000（旧 DoD 5002 / 假设的 3009 均不成立，真机为 3000）",
        action="set:formula",
        params={"address": "ZZZZ99999999", "formula": "=1+1", "worksheet_name": "Data"},
        expect_error_code="3000",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Range E2E — set:formula", TEST_CASES)
