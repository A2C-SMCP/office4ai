"""
Excel Format E2E — add/clear:conditionalFormat（条件格式）

覆盖 ``excel:add:conditionalFormat``（规则透传，cellValue 类型需 operator/formula1/format）
与 ``excel:clear:conditionalFormat``。条件格式以**协议成功**为主验证，并尽力（best-effort）
用 openpyxl ``ws.conditional_formatting`` 核对规则条数（读不到时打印 ⚠️ 不判失败——
不同 Excel 存盘对 CF 的序列化差异较大，视觉验收见 docs）。

cellValue 规则形态（AddIn conditionalFormat.ts 确认）:
    {type:"cellValue", operator:"GreaterThan", formula1:"150",
     format:{font:{color,bold,italic}, fill:{color}}}

运行方式:
    uv run python manual_tests/excel/format_e2e/test_conditional_format.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_base import WorkbookReader
from manual_tests.excel.e2e_case import PENDING_ADDIN_80, ExcelCase, run_main
from manual_tests.excel.format_e2e._fixtures import ensure_fixtures

FMT = "format_e2e/fmt.xlsx"


def _cf_count(reader: WorkbookReader) -> int | None:
    """best-effort 读 Data 表条件格式规则条数；读不到返回 None。"""
    reader.reload()
    try:
        cf = reader.wb["Data"].conditional_formatting
        return sum(1 for _ in cf)
    except Exception:  # noqa: BLE001
        return None


def _addr_ok(data: dict[str, Any]) -> bool:
    if not data.get("address"):
        print(f"   ❌ 响应缺少 address: {data}")
        return False
    return True


def validate_add_cellvalue(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    n = _cf_count(reader)
    print(f"   协议成功；openpyxl CF 规则数={n}")
    if n is not None and n < 1:
        print("   ⚠️  openpyxl 未读到 CF 规则（存盘序列化差异，协议已成功）")
    print("   ✅ cellValue 条件格式已添加（协议成功）")
    return True


def validate_add_colorscale(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    print(f"   协议成功，address={data.get('address')}")
    print("   ✅ colorScale 条件格式已添加（协议成功）")
    return True


def validate_clear(data: dict[str, Any], reader: WorkbookReader) -> bool:
    if not _addr_ok(data):
        return False
    n = _cf_count(reader)
    print(f"   清除后 openpyxl CF 规则数={n}")
    if n is not None and n > 0:
        print("   ⚠️  openpyxl 仍读到 CF（可能为其它区域规则；clear 仅作用目标区域）")
    print("   ✅ 条件格式已清除（协议成功）")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="add cellValue 规则",
        fixture_name=FMT,
        description="add:conditionalFormat B2:B4 cellValue>150 红底 → 协议成功 + best-effort CF",
        action="add:conditionalFormat",
        params={
            "address": "B2:B4",
            "rule": {
                "type": "cellValue",
                "operator": "GreaterThan",
                "formula1": "150",
                "format": {"fill": {"color": "#FF0000"}},
            },
            "worksheet_name": "Data",
        },
        validator=validate_add_cellvalue,
        tags=["cellvalue"],
    ),
    ExcelCase(
        name="add colorScale 规则",
        fixture_name=FMT,
        description="add:conditionalFormat B2:B4 colorScale → 协议成功",
        action="add:conditionalFormat",
        params={"address": "B2:B4", "rule": {"type": "colorScale"}, "worksheet_name": "Data"},
        validator=validate_add_colorscale,
        tags=["colorscale"],
    ),
    ExcelCase(
        name="clear 条件格式",
        fixture_name=FMT,
        description="先 add cellValue → clear:conditionalFormat B2:B4 → 协议成功",
        action="clear:conditionalFormat",
        params={"address": "B2:B4", "worksheet_name": "Data"},
        pre_ops=[
            (
                "add:conditionalFormat",
                {
                    "address": "B2:B4",
                    "rule": {
                        "type": "cellValue",
                        "operator": "GreaterThan",
                        "formula1": "150",
                        "format": {"fill": {"color": "#FF0000"}},
                    },
                    "worksheet_name": "Data",
                },
            )
        ],
        validator=validate_clear,
        tags=["clear"],
    ),
    ExcelCase(
        name="错误码 3009 — 非法 address（RANGE_INVALID）",
        fixture_name=FMT,
        description="add:conditionalFormat 传非法地址 → 3009（oasp#17 定案）",
        action="add:conditionalFormat",
        params={
            "address": "ZZZZ99999999",
            "rule": {"type": "cellValue", "operator": "GreaterThan", "formula1": "1"},
            "worksheet_name": "Data",
        },
        expect_error_code="3009",
        xfail_reason=PENDING_ADDIN_80,
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Format E2E — add/clear:conditionalFormat", TEST_CASES)
