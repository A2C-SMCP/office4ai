"""
Excel Table E2E — insert:table（创建结构化表格）

覆盖 ``excel:insert:table``：在已有数据上建表 / 带 data 写入空白区 / 指定 styleName。

wire 形态（AddIn table.ts 确认）:
- insert:table（address, hasHeaders, data?, styleName?, worksheetName?）→ ``{name, address}``
  （name 由 Excel 自动命名 "TableN"）。hasHeaders=true 时 data 第一行为表头、覆盖整张表范围。

双重验证：协议返回 + openpyxl ``table_names`` 读盘核对表已落盘（带 data 时再读首行/正文单元格）。
含错误码用例：非法 address → 3009 RANGE_INVALID（oasp#17 定案；Add-In 接线前 XFAIL）。

运行方式:
    uv run python manual_tests/excel/table_e2e/test_insert_table.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.table_e2e._fixtures import ensure_fixtures

TBL = "table_e2e/tbl.xlsx"


def _name_ok(data: dict[str, Any]) -> str | None:
    name, addr = data.get("name"), data.get("address")
    if not name or not addr:
        print(f"   ❌ 响应缺少 name/address: {data}")
        return None
    print(f"   insert 返回: name={name!r} address={addr!r}")
    return name


def validate_insert_over_data(data: dict[str, Any], reader: WorkbookReader) -> bool:
    name = _name_ok(data)
    if not name:
        return False
    reader.reload()
    tables = reader.table_names("Raw")
    print(f"   Raw 表列表={tables}")
    if name not in tables:
        print(f"   ❌ 新表 {name!r} 未在 Raw 落盘")
        return False
    print(f"   ✅ 在 Raw!A1:C4 已有数据上建表 {name!r}")
    return True


def validate_insert_with_data(data: dict[str, Any], reader: WorkbookReader) -> bool:
    name = _name_ok(data)
    if not name:
        return False
    reader.reload()
    tables = reader.table_names("Blank")
    print(f"   Blank 表列表={tables}")
    if name not in tables:
        print(f"   ❌ 新表 {name!r} 未在 Blank 落盘")
        return False
    h1, body = reader.cell_value("Blank", "A1"), reader.cell_value("Blank", "A2")
    print(f"   Blank A1(表头)={h1!r} A2(正文)={body!r}")
    if h1 != "Fruit" or body != "Apple":
        print("   ❌ data 未按 hasHeaders=true 写入（A1 应为表头 'Fruit'，A2 应为 'Apple'）")
        return False
    print(f"   ✅ 带 data 建表 {name!r}，表头+正文写入正确")
    return True


def validate_insert_styled(data: dict[str, Any], reader: WorkbookReader) -> bool:
    name = _name_ok(data)
    if not name:
        return False
    reader.reload()
    if name not in reader.table_names("Blank"):
        print(f"   ❌ 带样式新表 {name!r} 未落盘")
        return False
    print(f"   ✅ 指定 styleName 建表 {name!r}（样式视觉验收见 docs，落盘已确认）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="insert 在已有数据上建表",
        fixture_name=TBL,
        description="insert:table Raw!A1:C4 hasHeaders=true → {name,address} + 落盘",
        action="insert:table",
        params={"address": "A1:C4", "has_headers": True, "worksheet_name": "Raw"},
        validator=validate_insert_over_data,
        tags=["insert"],
    ),
    ExcelCase(
        name="insert 带 data 写入空白区",
        fixture_name=TBL,
        description="insert:table Blank!A1:B3 hasHeaders=true data=[表头,正文×2] → 表+数据落盘",
        action="insert:table",
        params={
            "address": "A1:B3",
            "has_headers": True,
            "data": [["Fruit", "Qty"], ["Apple", 5], ["Pear", 8]],
            "worksheet_name": "Blank",
        },
        validator=validate_insert_with_data,
        tags=["insert", "data"],
    ),
    ExcelCase(
        name="insert 指定 styleName",
        fixture_name=TBL,
        description="insert:table Blank!A1:B2 styleName='TableStyleLight9' → 表落盘",
        action="insert:table",
        params={
            "address": "A1:B2",
            "has_headers": True,
            "data": [["K", "V"], ["x", 1]],
            "style_name": "TableStyleLight9",
            "worksheet_name": "Blank",
        },
        validator=validate_insert_styled,
        tags=["insert", "style"],
    ),
    ExcelCase(
        name="错误码 3009 — 非法 address（RANGE_INVALID）",
        fixture_name=TBL,
        description="insert:table 传非法地址 → 3009（oasp#17）",
        action="insert:table",
        params={"address": "ZZZZ99999999", "has_headers": True, "worksheet_name": "Raw"},
        expect_error_code="3009",
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Table E2E — insert:table", TEST_CASES)
