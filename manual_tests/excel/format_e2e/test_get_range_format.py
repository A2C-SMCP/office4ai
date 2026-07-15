"""
Excel Format E2E — get:rangeFormat（读取区域格式）

覆盖 ``excel:get:rangeFormat``。

⚠️ **wire 实测与 DTO 不一致**：office4ai ``GetRangeFormatData`` 声明 ``{address, format}``，
但 AddIn ``rangeFormat.ts::getRangeFormat()`` 实际**直接返回扁平 RangeFormatInfo**——即
``{font, fill, horizontalAlignment, verticalAlignment, wrapText, numberFormat}`` 在**顶层**，
既无 ``address`` 也无 ``format`` 包裹。本套件按真机实际形态（顶层扁平）断言。
``numberFormat`` 为 **2D 数组**（每格一个格式串）。

含错误码用例：非法 address → 3009 RANGE_INVALID（oasp#17 定案；Add-In 接线前
（office-editor4ai#80）以 XFAIL 运行）。

运行方式:
    uv run python manual_tests/excel/format_e2e/test_get_range_format.py --test all
"""

from __future__ import annotations

from typing import Any

from manual_tests.excel.e2e_case import ExcelCase, run_main
from manual_tests.excel.format_e2e._fixtures import ensure_fixtures

FMT = "format_e2e/fmt.xlsx"


def _fmt(data: dict[str, Any]) -> dict[str, Any]:
    # AddIn 直接返回扁平 RangeFormatInfo（顶层即格式字段，无 {format} 包裹）。
    # 兼容两种形态：若未来加了 {format} 包裹则取之，否则用顶层 data。
    return data.get("format") or data


def validate_structure(data: dict[str, Any]) -> bool:
    fmt = _fmt(data)
    print(f"   format keys={list(fmt.keys())}")
    required = ["font", "fill", "horizontalAlignment", "verticalAlignment", "wrapText", "numberFormat"]
    missing = [k for k in required if k not in fmt]
    if missing:
        print(f"   ❌ 缺字段: {missing}")
        return False
    if not isinstance(fmt.get("font"), dict) or "name" not in fmt["font"]:
        print("   ❌ font 结构异常")
        return False
    nf = fmt.get("numberFormat")
    if not isinstance(nf, list) or (nf and not isinstance(nf[0], list)):
        print(f"   ❌ numberFormat 非 2D 数组: {nf!r}")
        return False
    print(f"   ✅ RangeFormatInfo 结构完整（font.name={fmt['font'].get('name')}，numberFormat 为 2D）")
    return True


def validate_set_then_get(data: dict[str, Any]) -> bool:
    fmt = _fmt(data)
    font = fmt.get("font") or {}
    fill = fmt.get("fill") or {}
    nf = fmt.get("numberFormat") or [[]]
    print(
        f"   font.bold={font.get('bold')} fill.color={fill.get('color')} numberFormat[0][0]={nf[0][0] if nf and nf[0] else None}"
    )
    if not font.get("bold"):
        print("   ❌ 先前 set 的 bold 未在 get 中反映")
        return False
    if str(fill.get("color", "")).upper() not in ("#FFFF00", "FFFF00"):
        print(f"   ⚠️  fill.color={fill.get('color')}（Excel 归一化差异；bold/numberFormat 为准）")
    if nf and nf[0] and nf[0][0] != "0.00":
        print(f"   ❌ numberFormat 未反映 '0.00'：{nf[0][0]!r}")
        return False
    print("   ✅ set→get 往返：bold + numberFormat '0.00' 已反映")
    return True


def validate_2d_dims(data: dict[str, Any]) -> bool:
    nf = _fmt(data).get("numberFormat")
    if not isinstance(nf, list) or not nf or not isinstance(nf[0], list):
        print(f"   ❌ numberFormat 非 2D 数组: {nf!r}")
        return False
    print(f"   numberFormat dims={len(nf)}x{len(nf[0])}")
    if len(nf) != 2 or len(nf[0]) != 3:
        print("   ❌ A1:C2 的 numberFormat 应为 2×3 二维数组")
        return False
    print("   ✅ numberFormat 维度 2×3 与范围一致")
    return True


TEST_CASES: list[ExcelCase] = [
    ExcelCase(
        name="读默认格式结构",
        fixture_name=FMT,
        description="get:rangeFormat Data!A1 → RangeFormatInfo 字段完整（含 2D numberFormat）",
        action="get:rangeFormat",
        params={"address": "A1", "worksheet_name": "Data"},
        validator=validate_structure,
        tags=["structure"],
    ),
    ExcelCase(
        name="set 后 get 往返",
        fixture_name=FMT,
        description="先 set B2 {bold,fill,numberFormat='0.00'} → get 反映",
        action="get:rangeFormat",
        params={"address": "B2", "worksheet_name": "Data"},
        pre_ops=[
            (
                "set:rangeFormat",
                {
                    "address": "B2",
                    "format": {"font": {"bold": True}, "fill": {"color": "#FFFF00"}, "numberFormat": "0.00"},
                    "worksheet_name": "Data",
                },
            )
        ],
        validator=validate_set_then_get,
        tags=["roundtrip"],
    ),
    ExcelCase(
        name="多格 numberFormat 2D 维度",
        fixture_name=FMT,
        description="get:rangeFormat Data!A1:C2 → numberFormat 为 2×3 二维数组",
        action="get:rangeFormat",
        params={"address": "A1:C2", "worksheet_name": "Data"},
        validator=validate_2d_dims,
        tags=["2d"],
    ),
    ExcelCase(
        name="错误码 3009 — 非法 address（RANGE_INVALID）",
        fixture_name=FMT,
        description="get:rangeFormat 传非法地址 → 3009（oasp#17 定案）",
        action="get:rangeFormat",
        params={"address": "ZZZZ99999999", "worksheet_name": "Data"},
        expect_error_code="3009",
        xfail_reason="待 Add-In 接线 office-editor4ai#80",
        tags=["error"],
    ),
]


if __name__ == "__main__":
    ensure_fixtures()
    run_main("Excel Format E2E — get:rangeFormat", TEST_CASES)
