"""
Excel Range E2E — insert:range + delete:range（插入/删除单元格并位移）

覆盖 ``excel:insert:range``（shiftDirection down/right，现有单元格让位）与
``excel:delete:range``（shiftDirection up/left，周围单元格补位）。在自描述的 Shift 表
（每格值=自身地址 "A1".."C3"）上操作，openpyxl 读盘核对「谁落到了哪」，从而确定
位移方向正确。每个用例独立打开一份夹具副本，互不污染。

运行方式:
    uv run python manual_tests/excel/range_e2e/test_shift_range.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.range_e2e._fixtures import ensure_fixtures

GRID = "range_e2e/grid.xlsx"


def _addr_ok(data: dict[str, Any]) -> bool:
    if not data.get("address"):
        print(f"   ❌ 响应缺少 address: {data}")
        return False
    return True


def validate_insert_down(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    a2 = reader.cell_value("Shift", "A2")  # 新空白
    a3 = reader.cell_value("Shift", "A3")  # 旧 A2 下移到此
    print(f"   insert down @A2 → A2={a2!r}(空) A3={a3!r}(旧 A2)")
    if a2 not in (None, ""):
        print("   ❌ 插入点 A2 未变空白")
        return False
    if a3 != "A2":
        print("   ❌ 旧 A2 未下移到 A3")
        return False
    print("   ✅ insert down：A2 空出，旧内容下移一行")
    return True


def validate_insert_right(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    a1 = reader.cell_value("Shift", "A1")  # 不动
    c1 = reader.cell_value("Shift", "C1")  # 旧 B1 右移到此
    print(f"   insert right @B1 → A1={a1!r}(不动) C1={c1!r}(旧 B1)")
    if a1 != "A1":
        print("   ❌ A1 不应受影响")
        return False
    if c1 != "B1":
        print("   ❌ 旧 B1 未右移到 C1")
        return False
    print("   ✅ insert right：B1 空出，旧内容右移一列")
    return True


def validate_delete_up(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    a2 = reader.cell_value("Shift", "A2")  # 旧 A3 上移补位
    print(f"   delete up @A2 → A2={a2!r}(旧 A3)")
    if a2 != "A3":
        print("   ❌ 删除 A2 后旧 A3 未上移补位")
        return False
    print("   ✅ delete up：A2 删除后下方上移补位")
    return True


def validate_delete_left(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    reader.reload()
    b1 = reader.cell_value("Shift", "B1")  # 旧 C1 左移补位
    print(f"   delete left @B1 → B1={b1!r}(旧 C1)")
    if b1 != "C1":
        print("   ❌ 删除 B1 后旧 C1 未左移补位")
        return False
    print("   ✅ delete left：B1 删除后右侧左移补位")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="insert down",
        fixture_name=GRID,
        description="insert Shift!A2 shiftDirection=down → A2 空出，旧 A2 落到 A3",
        action="insert:range",
        params={"address": "A2", "shift_direction": "down", "worksheet_name": "Shift"},
        validator=validate_insert_down,
        tags=["insert", "down"],
    ),
    ExcelCase(
        name="insert right",
        fixture_name=GRID,
        description="insert Shift!B1 shiftDirection=right → B1 空出，旧 B1 落到 C1",
        action="insert:range",
        params={"address": "B1", "shift_direction": "right", "worksheet_name": "Shift"},
        validator=validate_insert_right,
        tags=["insert", "right"],
    ),
    ExcelCase(
        name="delete up",
        fixture_name=GRID,
        description="delete Shift!A2 shiftDirection=up → 旧 A3 上移到 A2",
        action="delete:range",
        params={"address": "A2", "shift_direction": "up", "worksheet_name": "Shift"},
        validator=validate_delete_up,
        tags=["delete", "up"],
    ),
    ExcelCase(
        name="delete left",
        fixture_name=GRID,
        description="delete Shift!B1 shiftDirection=left → 旧 C1 左移到 B1",
        action="delete:range",
        params={"address": "B1", "shift_direction": "left", "worksheet_name": "Shift"},
        validator=validate_delete_left,
        tags=["delete", "left"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Range E2E — insert:range + delete:range", TEST_CASES)
