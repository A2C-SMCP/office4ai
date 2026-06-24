"""
Excel Find&Filter E2E — find:values（查找匹配单元格）

覆盖 ``excel:find:values`` 的 matchCase / matchEntireCell（默认均 false）与无命中。

wire 形态（AddIn findValues.ts 确认）:
- find:values（searchText, address?, worksheetName?, matchCase?=false, matchEntireCell?=false）
  → ``{matches:[{address, value}]}``。address 省略 = 整个 usedRange。matches[].address 为
  含表名全地址（如 'Data!B2'）。

夹具 Data Product 列大小写混合（Apple/apple）以区分 matchCase。含错误码：searchText 空串
（Zod min(1)）→ 4000 VALIDATION_ERROR。

运行方式:
    uv run python manual_tests/excel/find_filter_e2e/test_find_values.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.find_filter_e2e._fixtures import ensure_fixtures

FILT = "find_filter_e2e/filt.xlsx"


def _matches(data: dict[str, Any]) -> list[dict[str, Any]]:
    return data.get("matches") or []


def _cells(data: dict[str, Any]) -> list[str]:
    # 取末段单元格坐标（去掉 'Data!' 前缀），排序便于断言。
    out = []
    for m in _matches(data):
        addr = str(m.get("address", ""))
        out.append(addr.split("!")[-1])
    return sorted(out)


def _make_count_validator(label: str, expect_cells: list[str]):
    def _v(data: dict[str, Any]) -> bool:
        cells = _cells(data)
        vals = [m.get("value") for m in _matches(data)]
        print(f"   {label}: matches={cells} values={vals}")
        if cells != sorted(expect_cells):
            print(f"   ❌ 命中单元格应为 {sorted(expect_cells)}，实得 {cells}")
            return False
        print(f"   ✅ {label}：命中 {len(cells)} 个，地址符合预期")
        return True

    return _v


def validate_no_hit(data: dict[str, Any]) -> bool:
    matches = _matches(data)
    print(f"   matchEntireCell 'App' → matches={matches}")
    if matches != []:
        print("   ❌ 整单元格匹配 'App' 应无命中（无单元格整体等于 'App'）")
        return False
    print("   ✅ 无命中 → matches=[]（整单元格匹配排除子串）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="默认查找（不区分大小写·子串）",
        fixture_name=FILT,
        description="find:values 'apple'（默认）→ 命中 Apple/apple/apple 共 3（B2/B3/B5）",
        action="find:values",
        params={"search_text": "apple", "worksheet_name": "Data"},
        validator=_make_count_validator("默认子串", ["B2", "B3", "B5"]),
        tags=["find"],
    ),
    ExcelCase(
        name="matchCase 区分大小写",
        fixture_name=FILT,
        description="find:values 'apple' matchCase=true → 仅小写 B3/B5 共 2",
        action="find:values",
        params={"search_text": "apple", "match_case": True, "worksheet_name": "Data"},
        validator=_make_count_validator("matchCase", ["B3", "B5"]),
        tags=["find", "case"],
    ),
    ExcelCase(
        name="matchEntireCell 无命中",
        fixture_name=FILT,
        description="find:values 'App' matchEntireCell=true → 无单元格整体等于 'App' → 0 命中",
        action="find:values",
        params={"search_text": "App", "match_entire_cell": True, "worksheet_name": "Data"},
        validator=validate_no_hit,
        tags=["find", "entire", "nohit"],
    ),
    ExcelCase(
        name="限定 address 范围查找",
        fixture_name=FILT,
        description="find:values 'East' address=A1:A5 → 仅 Region 列 A2/A4 共 2",
        action="find:values",
        params={"search_text": "East", "address": "A1:A5", "worksheet_name": "Data"},
        validator=_make_count_validator("限定范围", ["A2", "A4"]),
        tags=["find", "scoped"],
    ),
    ExcelCase(
        name="错误码 4000 — searchText 空串（VALIDATION_ERROR）",
        fixture_name=FILT,
        description="find:values searchText='' → 4000（Zod min(1) 失败）",
        action="find:values",
        params={"search_text": "", "worksheet_name": "Data"},
        expect_error_code="4000",
        tags=["error", "zod"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Find&Filter E2E — find:values", TEST_CASES)
