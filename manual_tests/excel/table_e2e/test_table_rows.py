"""
Excel Table E2E — add:tableRow + delete:tableRow（行增删）

覆盖 ``excel:add:tableRow``（末尾追加）与 ``excel:delete:tableRow``（按 0-based 索引删，
含 **rowIndex=0 falsy** 必须落 wire）。

wire 形态（AddIn table.ts 确认）:
- add:tableRow（tableId, values, worksheetName?）→ ``{tableId}``
- delete:tableRow（tableId, rowIndex, worksheetName?）→ **void**（无 data）

双重验证：openpyxl 读 Roster 表正文单元格核对「谁还在 / 谁上移」。RosterTable 正文 A2:B4 =
R0row/R1row/R2row（自描述）。错误码两路径：
- 表不存在 → ``3010``(kind:table)；rowIndex 越界（合法非负）→ ``4004``（oasp#17，接线前 XFAIL）；
- rowIndex 负数（Zod nonnegative 失败）→ ``4000``。

运行方式:
    uv run python manual_tests/excel/table_e2e/test_table_rows.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.table_e2e._fixtures import ensure_fixtures

TBL = "table_e2e/tbl.xlsx"


def _items(reader: WorkbookReader, n: int) -> list[Any]:
    """读 Roster Item 列前 n 个正文单元格（A2 起）。"""
    reader.reload()
    return [reader.cell_value("Roster", f"A{r}") for r in range(2, 2 + n)]


def validate_add_row(data: dict[str, Any], reader: WorkbookReader) -> bool:
    print(f"   add 返回: {data}")
    if data.get("tableId") != "RosterTable":
        print(f"   ❌ 响应应回显 tableId='RosterTable'，实得 {data.get('tableId')!r}")
        return False
    items = _items(reader, 4)
    new_qty = reader.cell_value("Roster", "B5")
    print(f"   追加后 Item 列={items} B5(新行 Qty)={new_qty!r}")
    if items != ["R0row", "R1row", "R2row", "R3row"]:
        print("   ❌ 新行未追加到表尾（A5 应为 'R3row'）")
        return False
    if new_qty != 40:
        print(f"   ❌ 新行 Qty 应为 40，实得 {new_qty!r}")
        return False
    print("   ✅ add:tableRow 追加 [R3row,40] 到表尾成功")
    return True


def validate_delete_first(data: dict[str, Any], reader: WorkbookReader) -> bool:
    # delete 返回 void → 只读盘核对。rowIndex=0 删首行，R1row 上移到 A2。
    items = _items(reader, 3)
    print(f"   删 rowIndex=0 后 Item 列={items}")
    if "R0row" in items:
        print("   ❌ 首行 R0row 应已删除")
        return False
    if items[:2] != ["R1row", "R2row"]:
        print(f"   ❌ 删首行后应为 [R1row,R2row,...]，实得 {items}")
        return False
    print("   ✅ delete:tableRow rowIndex=0（falsy）正确删除首行，R1row 上移")
    return True


def validate_delete_middle(data: dict[str, Any], reader: WorkbookReader) -> bool:
    items = _items(reader, 2)
    print(f"   删 rowIndex=1 后 Item 列={items}")
    if "R1row" in items:
        print("   ❌ 中间行 R1row 应已删除")
        return False
    if items != ["R0row", "R2row"]:
        print(f"   ❌ 删中间行后应为 [R0row,R2row]，实得 {items}")
        return False
    print("   ✅ delete:tableRow rowIndex=1 正确删除中间行")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="add:tableRow 末尾追加",
        fixture_name=TBL,
        description="add:tableRow RosterTable values=['R3row',40] → 表尾新增行",
        action="add:tableRow",
        params={"table_id": "RosterTable", "values": ["R3row", 40], "worksheet_name": "Roster"},
        validator=validate_add_row,
        tags=["add"],
    ),
    ExcelCase(
        name="delete:tableRow rowIndex=0（falsy）",
        fixture_name=TBL,
        description="delete:tableRow rowIndex=0 → 删首行 R0row，R1row 上移",
        action="delete:tableRow",
        params={"table_id": "RosterTable", "row_index": 0, "worksheet_name": "Roster"},
        validator=validate_delete_first,
        tags=["delete", "falsy"],
    ),
    ExcelCase(
        name="delete:tableRow rowIndex=1（中间）",
        fixture_name=TBL,
        description="delete:tableRow rowIndex=1 → 删中间行 R1row",
        action="delete:tableRow",
        params={"table_id": "RosterTable", "row_index": 1, "worksheet_name": "Roster"},
        validator=validate_delete_middle,
        tags=["delete"],
    ),
    ExcelCase(
        name="错误码 4004 — rowIndex 越界（PARAM_OUT_OF_RANGE）",
        fixture_name=TBL,
        description="delete:tableRow rowIndex=99（合法非负但越界）→ 4004（oasp#17，events-excel.md 逐事件表）",
        action="delete:tableRow",
        params={"table_id": "RosterTable", "row_index": 99, "worksheet_name": "Roster"},
        expect_error_code="4004",
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
    ExcelCase(
        name="错误码 4000 — rowIndex 负数（VALIDATION_ERROR）",
        fixture_name=TBL,
        description="delete:tableRow rowIndex=-1 → 4000（Zod nonnegative 失败）",
        action="delete:tableRow",
        params={"table_id": "RosterTable", "row_index": -1, "worksheet_name": "Roster"},
        expect_error_code="4000",
        tags=["error", "zod"],
    ),
    ExcelCase(
        name="错误码 3010 — add 表不存在（ELEMENT_NOT_FOUND kind:table）",
        fixture_name=TBL,
        description="add:tableRow tableId='NoSuchTable' → 3010 + details.kind=table（oasp#17）",
        action="add:tableRow",
        params={"table_id": "NoSuchTable", "values": ["x", 1], "worksheet_name": "Roster"},
        expect_error_code="3010",
        expect_error_details={"kind": "table"},
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Table E2E — add/delete:tableRow", TEST_CASES)
