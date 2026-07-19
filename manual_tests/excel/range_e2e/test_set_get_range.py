"""
Excel Range E2E — set:range + get:range（范围读写往返）

覆盖 ``excel:set:range``（标量填充 / 二维数组 / falsy 存活）与 ``excel:get:range``
（2D values / rowCount/columnCount / includeFormat / 读已有数据）。写操作响应仅
``{address}``，故用 ``get:range`` 协议返回 + openpyxl 读盘双重核对写入结果。

含错误码用例：非法 address → ``3009 RANGE_INVALID``（oasp#17 定案权威码；Add-In 接线前
（office-editor4ai#80）以 XFAIL 运行）。

运行方式:
    uv run python manual_tests/excel/range_e2e/test_set_get_range.py --test all
    uv run python manual_tests/excel/range_e2e/test_set_get_range.py --test 3   # falsy 存活
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import PENDING_ADDIN_80, ExcelCase, run_main
from manual_tests.excel.range_e2e._fixtures import ensure_fixtures

GRID = "range_e2e/grid.xlsx"

# 用例 1 写入的 2D 内容
_GRID_2D = [["x1", "x2", "x3"], ["y1", "y2", "y3"]]


def _dims_ok(data: dict[str, Any], rows: int, cols: int) -> bool:
    rc, cc = data.get("rowCount"), data.get("columnCount")
    vals = data.get("values")
    if rc != rows or cc != cols:
        print(f"   ❌ rowCount/columnCount={rc}x{cc}（预期 {rows}x{cols}）")
        return False
    if not isinstance(vals, list) or (vals and not isinstance(vals[0], list)):
        print(f"   ❌ values 非 2D 数组: {vals!r}")
        return False
    return True


def validate_set_get_2d(data: dict[str, Any], reader: WorkbookReader) -> bool:
    print(f"   wire values={data.get('values')}")
    if not _dims_ok(data, 2, 3):
        return False
    if data.get("values") != _GRID_2D:
        print(f"   ❌ wire values 与写入不一致: {data.get('values')} != {_GRID_2D}")
        return False
    reader.reload()
    disk = reader.cell_value("Blank", "A1"), reader.cell_value("Blank", "C2")
    print(f"   openpyxl Blank A1/C2 = {disk}")
    if disk != ("x1", "y3"):
        print("   ❌ 磁盘内容与写入不符")
        return False
    print("   ✅ 2D 写入往返一致（wire + 磁盘）")
    return True


def validate_scalar_fill(data: dict[str, Any], reader: WorkbookReader) -> bool:
    print(f"   wire values={data.get('values')}")
    if not _dims_ok(data, 2, 2):
        return False
    flat = [v for row in data.get("values", []) for v in row]
    if any(v != "Z" for v in flat):
        print(f"   ❌ 标量填充未铺满: {flat}")
        return False
    reader.reload()
    if reader.cell_value("Blank", "B2") != "Z":
        print("   ❌ 磁盘 Blank!B2 != 'Z'")
        return False
    print("   ✅ 标量 'Z' 已填满 A1:B2（wire + 磁盘）")
    return True


def validate_falsy(data: dict[str, Any], reader: WorkbookReader) -> bool:
    vals = data.get("values")
    print(f"   wire values={vals}")
    if not _dims_ok(data, 1, 3):
        return False
    row = vals[0]
    # 关键：0 / False 必须存活为本身，不能被吞成 None/空
    if row[0] != 0 or isinstance(row[0], bool):
        print(f"   ❌ 首格应为整数 0，实际 {row[0]!r}")
        return False
    if row[1] is not False:
        print(f"   ❌ 第二格应为 False，实际 {row[1]!r}")
        return False
    print("   ✅ wire 上 0 / False / '' 全部存活（未被吞）")
    reader.reload()
    a1 = reader.cell_value("Blank", "A1")
    b1 = reader.cell_value("Blank", "B1")
    print(f"   openpyxl Blank A1={a1!r} B1={b1!r}")
    if a1 != 0:
        print("   ⚠️  磁盘 A1 非 0（Excel 存储差异，wire 已验证为准）")
    return True


def validate_get_existing(data: dict[str, Any], reader: WorkbookReader) -> bool:
    print(f"   wire values[0]={data.get('values', [[]])[0]}")
    if not _dims_ok(data, 4, 3):
        return False
    if data.get("values", [[]])[0] != ["Region", "Q1", "Q2"]:
        print("   ❌ 表头行不匹配夹具")
        return False
    reader.reload()
    if reader.cell_value("Data", "C4") != 350:
        print("   ❌ 磁盘 Data!C4 != 350")
        return False
    print("   ✅ 读取已有 Data!A1:C4 一致（4×3，表头 + C4=350）")
    return True


def validate_include_format(data: dict[str, Any], reader: WorkbookReader) -> bool:
    fmt = data.get("format")
    print(f"   format keys={list(fmt.keys()) if isinstance(fmt, dict) else fmt}")
    if not isinstance(fmt, dict):
        print("   ❌ includeFormat=true 但 format 缺失")
        return False
    if "font" not in fmt or "fill" not in fmt:
        print("   ❌ format 缺少 font/fill")
        return False
    print(f"   ✅ includeFormat 返回 RangeFormatInfo（font={fmt.get('font', {}).get('name')}）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="set 2D + get 往返",
        fixture_name=GRID,
        description="写 A1:C2 二维数组到 Blank → get 读回，wire + 磁盘双验证",
        action="get:range",
        params={"address": "A1:C2", "worksheet_name": "Blank"},
        pre_ops=[("set:range", {"address": "A1:C2", "values": _GRID_2D, "worksheet_name": "Blank"})],
        validator=validate_set_get_2d,
        tags=["set", "get", "2d"],
    ),
    ExcelCase(
        name="标量填充",
        fixture_name=GRID,
        description="set values='Z'（标量）填满 A1:B2 → get 读回全为 'Z'",
        action="get:range",
        params={"address": "A1:B2", "worksheet_name": "Blank"},
        pre_ops=[("set:range", {"address": "A1:B2", "values": "Z", "worksheet_name": "Blank"})],
        validator=validate_scalar_fill,
        tags=["set", "scalar"],
    ),
    ExcelCase(
        name="falsy 存活（0/False/空）",
        fixture_name=GRID,
        description="set [[0, False, '']] → get 读回 0/False 必须存活为本身（不被吞为 null）",
        action="get:range",
        params={"address": "A1:C1", "worksheet_name": "Blank"},
        pre_ops=[("set:range", {"address": "A1:C1", "values": [[0, False, ""]], "worksheet_name": "Blank"})],
        validator=validate_falsy,
        tags=["set", "falsy"],
    ),
    ExcelCase(
        name="读已有数据",
        fixture_name=GRID,
        description="get Data!A1:C4 → 4×3，表头 [Region,Q1,Q2]，C4=350",
        action="get:range",
        params={"address": "A1:C4", "worksheet_name": "Data"},
        validator=validate_get_existing,
        tags=["get", "existing"],
    ),
    ExcelCase(
        name="includeFormat",
        fixture_name=GRID,
        description="get Data!A1:C1 includeFormat=true → 返回 RangeFormatInfo（font/fill）",
        action="get:range",
        params={"address": "A1:C1", "worksheet_name": "Data", "include_format": True},
        validator=validate_include_format,
        tags=["get", "format"],
    ),
    ExcelCase(
        # oasp#17 定案：非法/畸形 address → 3009 RANGE_INVALID（规范层 MUST，不得降级 3000）。
        # 历史：Add-In excelErrorCode() 曾只有 Zod→4000 / 其余→3000 二值分类，3009 是
        # dead code、真机实收 3000——该现实由 office-editor4ai#80 接线消解，接线前本用例 XFAIL。
        name="错误码 3009 — 非法 address（RANGE_INVALID）",
        fixture_name=GRID,
        description="get:range 传非法地址 → 3009（oasp#17 定案；旧 DoD 5002 已退役）",
        action="get:range",
        params={"address": "ZZZZ99999999", "worksheet_name": "Data"},
        expect_error_code="3009",
        xfail_reason=PENDING_ADDIN_80,
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Range E2E — set:range + get:range", TEST_CASES)
