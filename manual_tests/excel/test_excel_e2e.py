"""
Excel End-to-End Test (OASP /excel Draft 0.3.0, milestone #3 / issue #26)

在真实 Excel + office-editor4ai Add-In 环境下，依次驱动全部 37 个 Excel MCP 事件
（10 类），并用 openpyxl 双重验证写盘后的 .xlsx 结构。镜像 word/test_word_table_e2e.py
与 ppt/test_chart_e2e.py 的结构。

两种运行模式：

────────────────────────────────────────────────────────────────────
  --mode health  （默认，不需要 Add-In）
────────────────────────────────────────────────────────────────────
启动 Workspace，校验全部 37 个 excel:* 事件已注册到 request_registry。
用于 PR / CI 阶段的接线回归保护。

    uv run python manual_tests/excel/test_excel_e2e.py --mode health

────────────────────────────────────────────────────────────────────
  --mode full  （真机验收，需要 macOS + Excel + Add-In）
────────────────────────────────────────────────────────────────────
打开一个空工作簿，把 37 个事件编排成一条「构建销售报表」工作流逐个跑通，
**continue-on-error**（单步失败不中断，最后汇总每个事件 ✅/❌），再用 openpyxl
读盘双重验证 + 触发权威错误码场景（3010+kind / 3009，oasp#17 定案；Add-In 接线前
（office-editor4ai#80）以 XFAIL 运行）。成功后保留工作副本供目测。

    uv run python manual_tests/excel/test_excel_e2e.py --mode full

成功后，工作副本保留在 manual_tests/excel/.test_working/ 下，用 Excel 打开人眼检查
「表头蓝底白字 + 表格 + 图表 + 透视表」等视觉效果。需人工验证项见
docs/manual_tests/excel_v0.3.0.md。
"""

from __future__ import annotations

import asyncio
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

# Make this script runnable both as `python -m` and as a path
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from manual_tests.excel.e2e_base import (  # noqa: E402
    ExcelTestRunner,
    WorkbookReader,
    ensure_empty_fixture,
)
from manual_tests.excel.e2e_case import evaluate_error_case  # noqa: E402
from manual_tests.excel.test_helpers import excel_op  # noqa: E402

# ============================================================================
# 全部 37 个 Excel 事件（按 10 类分组）—— health 模式校验注册用
# ============================================================================

ALL_EXCEL_EVENTS: list[str] = [
    # #18 Foundation (3)
    "excel:get:workbookInfo",
    "excel:get:worksheetInfo",
    "excel:get:selectedRange",
    # #19 Range (7)
    "excel:get:range",
    "excel:set:range",
    "excel:clear:range",
    "excel:copy:range",
    "excel:delete:range",
    "excel:insert:range",
    "excel:set:formula",
    # #20 Format (6)
    "excel:get:rangeFormat",
    "excel:set:rangeFormat",
    "excel:add:conditionalFormat",
    "excel:clear:conditionalFormat",
    "excel:merge:cells",
    "excel:unmerge:cells",
    # #21 Worksheet (5)
    "excel:get:worksheets",
    "excel:add:worksheet",
    "excel:delete:worksheet",
    "excel:rename:worksheet",
    "excel:activate:worksheet",
    # #22 Table (6)
    "excel:insert:table",
    "excel:get:table",
    "excel:get:tables",
    "excel:add:tableRow",
    "excel:delete:tableRow",
    "excel:sort:table",
    # #23 Chart (4)
    "excel:insert:chart",
    "excel:get:charts",
    "excel:update:chart",
    "excel:delete:chart",
    # #24 PivotTable (3)
    "excel:insert:pivotTable",
    "excel:get:pivotTables",
    "excel:delete:pivotTable",
    # #25 Find&Filter (3)
    "excel:find:values",
    "excel:set:autoFilter",
    "excel:clear:autoFilter",
]


# ============================================================================
# 结果收集器（continue-on-error）
# ============================================================================


@dataclass
class ResultCollector:
    """逐事件记录 (event, ok, note)，最后汇总。"""

    rows: list[tuple[str, bool, str]] = field(default_factory=list)

    def record(self, event: str, ok: bool, note: str = "") -> bool:
        self.rows.append((event, ok, note))
        return ok

    @property
    def passed(self) -> int:
        return sum(1 for _, ok, _ in self.rows if ok)

    @property
    def total(self) -> int:
        return len(self.rows)

    def summary(self) -> str:
        lines = ["", "=" * 70, f"事件执行汇总  {self.passed}/{self.total} ✅", "=" * 70]
        for event, ok, note in self.rows:
            mark = "✅" if ok else "❌"
            suffix = f"  — {note}" if note else ""
            lines.append(f"  {mark} {event}{suffix}")
        return "\n".join(lines)


# ============================================================================
# health 模式
# ============================================================================


async def test_workspace_health() -> bool:
    """快速健康检查：37 个 excel:* 事件均已注册到 request_registry（无需 Add-In）。"""
    from office4ai.environment.workspace.dtos.common import request_registry
    from office4ai.environment.workspace.office_workspace import OfficeWorkspace

    print("\n" + "=" * 70)
    print("🏥 Excel Workspace Health Check (37 events)")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)
    try:
        await workspace.start()
        print("✅ Workspace 运行正常")
        events = set(request_registry.all_events())
        missing = [e for e in ALL_EXCEL_EVENTS if e not in events]
        if missing:
            print(f"❌ 以下事件未注册 ({len(missing)}): {missing}")
            return False
        print(f"✅ 全部 {len(ALL_EXCEL_EVENTS)} 个 excel:* 事件均已注册到 request_registry")
        return True
    except Exception as exc:  # noqa: BLE001
        print(f"❌ 健康检查失败: {exc}")
        return False
    finally:
        await workspace.stop()


# ============================================================================
# full 模式：按 10 类编排的工作流（continue-on-error）
# ============================================================================

HEADER_FILL = "#1F4E79"
HEADER_FONT = "#FFFFFF"


async def run_full_workflow(workspace: Any, uri: str, rc: ResultCollector) -> None:
    """把 37 个事件编排成一条「销售报表」工作流；单步失败仅记录、不抛出。"""

    async def step(event: str, note: str = "", **params: Any) -> dict | None:
        ok, data, err = await excel_op(workspace, uri, event, **params)
        rc.record(f"excel:{event}", ok, note if ok else (err or "失败"))
        return data

    # ── #18 Foundation: 状态感知读 ────────────────────────────────────────
    print("\n" + "─" * 70 + "\n#18 Foundation 状态感知读\n" + "─" * 70)
    await step("get:workbookInfo", "读工作簿信息")
    await step("get:worksheetInfo", "读 Sheet1 信息")
    await step("get:selectedRange", "读当前选区")

    # ── #19 Range: CRUD + 公式 ───────────────────────────────────────────
    print("\n" + "─" * 70 + "\n#19 Range CRUD + 公式\n" + "─" * 70)
    await step(
        "set:range",
        "写表头 + 数据",
        address="A1:C4",
        values=[["Region", "Q1", "Q2"], ["East", 10, 20], ["West", 30, 40], ["North", 50, 60]],
    )
    await step("get:range", "回读 A1:C4", address="A1:C4")
    await step("set:formula", "D1 求和公式", address="D1", formula="=SUM(B2:C4)")
    await step("copy:range", "复制表头到 A6", source_address="A1:C1", target_address="A6")
    await step("insert:range", "A10 下移插入", address="A10:C10", shift_direction="down")
    await step("delete:range", "A10 上移删除", address="A10:C10", shift_direction="up")
    await step("clear:range", "清空 A6:C6 内容", address="A6:C6", clear_type="contents")

    # ── #20 Format: 格式 / 条件格式 / 合并 ───────────────────────────────
    print("\n" + "─" * 70 + "\n#20 Format / 条件格式 / 合并\n" + "─" * 70)
    await step("get:rangeFormat", "读表头格式", address="A1:C1")
    await step(
        "set:rangeFormat",
        "表头蓝底白字加粗",
        address="A1:C1",
        format={"font": {"bold": True, "color": HEADER_FONT}, "fill": {"color": HEADER_FILL}},
    )
    await step(
        "add:conditionalFormat",
        "B2:C4 > 25 高亮",
        address="B2:C4",
        rule={"type": "cellValue", "operator": "greaterThan", "formula1": "25"},
    )
    await step("clear:conditionalFormat", "清条件格式", address="B2:C4")
    await step("merge:cells", "合并 E1:F1", address="E1:F1")
    await step("unmerge:cells", "取消合并 E1:F1", address="E1:F1")

    # ── #21 Worksheet: 工作表管理 ────────────────────────────────────────
    print("\n" + "─" * 70 + "\n#21 Worksheet 管理\n" + "─" * 70)
    await step("get:worksheets", "列工作表")
    await step("add:worksheet", "新增 Report 表", name="Report")
    await step("rename:worksheet", "Report → Summary", current_name="Report", new_name="Summary")
    await step("activate:worksheet", "切到 Summary", worksheet_name="Summary")
    await step("activate:worksheet", "切回 Sheet1", worksheet_name="Sheet1")
    await step("delete:worksheet", "删除 Summary", worksheet_name="Summary")

    # ── #22 Table: 表格操作 ──────────────────────────────────────────────
    print("\n" + "─" * 70 + "\n#22 Table 操作\n" + "─" * 70)
    tbl = await step("insert:table", "A1:C4 建表", address="A1:C4", has_headers=True)
    table_id = (tbl or {}).get("name", "Table1")
    await step("get:tables", "列表格")
    await step("get:table", f"读表 {table_id}", table_id=table_id)
    await step("add:tableRow", "追加一行", table_id=table_id, values=["South", 70, 80])
    await step("sort:table", "按 Q1 降序", table_id=table_id, sort_fields=[{"column_index": 1, "ascending": False}])
    await step("delete:tableRow", "删首数据行", table_id=table_id, row_index=0)

    # ── #23 Chart: 图表操作 ──────────────────────────────────────────────
    print("\n" + "─" * 70 + "\n#23 Chart 操作\n" + "─" * 70)
    ch = await step(
        "insert:chart",
        "柱形图",
        source_address="A1:C4",
        chart_type="ColumnClustered",
        title="季度业绩",
        position={"top": 20, "left": 360, "width": 360, "height": 240},
    )
    chart_name = (ch or {}).get("name", "Chart 1")
    await step("get:charts", "列图表")
    await step("update:chart", "改标题", chart_name=chart_name, properties={"title": "季度业绩（修订）"})
    await step("delete:chart", "删图表", chart_name=chart_name)

    # ── #24 PivotTable: 透视表操作 ───────────────────────────────────────
    print("\n" + "─" * 70 + "\n#24 PivotTable 操作\n" + "─" * 70)
    pv = await step("insert:pivotTable", "建透视表", source_address="A1:C4", target_address="H1", name="PivotSales")
    pivot_name = (pv or {}).get("name", "PivotSales")
    await step("get:pivotTables", "列透视表")
    await step("delete:pivotTable", "删透视表", pivot_table_name=pivot_name)

    # ── #25 Find&Filter: 查找与筛选 ──────────────────────────────────────
    print("\n" + "─" * 70 + "\n#25 Find&Filter 查找与筛选\n" + "─" * 70)
    await step("find:values", "查找 East", search_text="East")
    # 夹具：set:autoFilter 必须落在「干净、无表」区域。#22 已把 A1:C4 变成 Table「表1」，
    # 而 worksheet 级 autoFilter 不能套在 Table range 上（Office.js 抛 InvalidArgument →
    # 通用 3000，office-editor4ai 复盘确认）。故在未被前序触及的 A20:C23 写一份独立数据供筛选；
    # 此写入仅作夹具、不经 step() 计入事件汇总。（对表 range 设筛选另属 Add-In 易用性增强，见 #78 后续。）
    await excel_op(
        workspace,
        uri,
        "set:range",
        quiet=True,
        address="A20:C23",
        values=[["Region", "Q1", "Q2"], ["East", 10, 20], ["West", 30, 40], ["North", 50, 60]],
    )
    await step(
        "set:autoFilter",
        "按 Region 筛选（独立区域 A20:C23，避开表1）",
        address="A20:C23",
        criteria=[{"column_index": 0, "filter_on": "Values", "values": ["East", "West"]}],
    )
    await step("clear:autoFilter", "清筛选")


async def run_error_scenarios(workspace: Any, uri: str, rc: ResultCollector) -> None:
    """错误码场景：每步**预期失败**，并校验权威错误码 + details.kind（oasp#17）。

    Add-In 接线前（office-editor4ai#80）真机实收 3000 兜底 → 记 ⚠️ XFAIL（计通过
    不红）；接线后严格命中 → ✅。XFAIL 仍要求响应必须失败，防掩盖回归。
    """
    print("\n" + "─" * 70 + "\n错误码场景（预期失败 + 校验权威码；Add-In 接线前 XFAIL）\n" + "─" * 70)

    async def expect_error(event: str, code: str, details: dict[str, Any] | None = None, **params: Any) -> None:
        ok, _, err = await excel_op(workspace, uri, event, quiet=True, **params)
        expected = code + (f" details⊇{details}" if details else "")
        if evaluate_error_case(ok, err, code, details):
            rc.record(f"[err {code}] excel:{event}", True, "")
            print(f"   ✅ [{expected}] excel:{event} → {err}")
        elif not ok:
            rc.record(f"[err {code}] excel:{event}", True, f"XFAIL 实际: {err}")
            print(f"   ⚠️  XFAIL（待 Add-In 接线 office-editor4ai#80）[{expected}] excel:{event} → {err}")
        else:
            rc.record(f"[err {code}] excel:{event}", False, f"实际: ok={ok} err={err}")
            print(f"   ❌ [{expected}] excel:{event} → ok={ok} err={err}")

    await expect_error("get:worksheetInfo", "3010", {"kind": "worksheet"}, worksheet_name="GhostSheet")
    await expect_error("get:range", "3009", address="!!!bad!!!")
    await expect_error("get:table", "3010", {"kind": "table"}, table_id="GhostTable")
    await expect_error("delete:chart", "3010", {"kind": "chart"}, chart_name="GhostChart")
    await expect_error("delete:pivotTable", "3010", {"kind": "pivotTable"}, pivot_table_name="GhostPivot")


def verify_with_openpyxl(working_path: Path) -> tuple[bool, list[str]]:
    """openpyxl 双重验证：读盘核对稳定不变量。"""
    msgs: list[str] = []
    ok = True
    reader = WorkbookReader(working_path)
    reader.reload()  # 先 AppleScript 让 Excel 存盘

    print(f"   📄 工作表: {reader.sheet_names}")
    if "Sheet1" in reader.sheet_names:
        msgs.append("✅ Sheet1 存在")
    else:
        msgs.append(f"❌ Sheet1 不在 {reader.sheet_names}")
        ok = False

    if "Summary" not in reader.sheet_names:
        msgs.append("✅ 临时表 Summary 已删除")
    else:
        msgs.append("❌ Summary 未删除（delete:worksheet 未生效）")
        ok = False

    a1 = reader.cell_value("Sheet1", "A1")
    if a1 == "Region":
        msgs.append(f"✅ Sheet1!A1 = {a1!r}")
    else:
        msgs.append(f"⚠️  Sheet1!A1 = {a1!r}（预期 'Region'；若已转表格表头可能不同）")

    tables = reader.table_names("Sheet1")
    msgs.append(f"{'✅' if tables else '⚠️ '} Sheet1 表格 = {tables}")

    merged = reader.merged_ranges("Sheet1")
    if "E1:F1" not in merged:
        msgs.append("✅ E1:F1 合并已取消")
    else:
        msgs.append("❌ E1:F1 仍处于合并状态（unmerge 未生效）")
        ok = False

    return ok, msgs


async def test_excel_full_e2e() -> bool:
    """完整 37 事件真机工作流 + openpyxl 验证 + 错误码场景。"""
    print("\n" + "=" * 70)
    print("🧪 Excel Full E2E (OASP /excel Draft 0.3.0, 37 events)")
    print("=" * 70)

    ensure_empty_fixture()
    rc = ResultCollector()

    runner = ExcelTestRunner(
        fixtures_dir=_PROJECT_ROOT / "manual_tests" / "excel" / "fixtures",
        host="127.0.0.1",
        port=3000,
        connection_timeout=30.0,
        auto_open=True,
        auto_close=False,  # 保留工作副本供视觉验收
        auto_activate=True,
        cleanup_on_success=False,
    )

    try:
        async with runner.run_with_workspace("empty.xlsx", open_delay=3.0) as (workspace, fixture):
            print(f"\n📂 工作副本: {fixture.working_path}")

            await run_full_workflow(workspace, fixture.document_uri, rc)
            await run_error_scenarios(workspace, fixture.document_uri, rc)

            print("\n" + "─" * 70 + "\nopenpyxl 双重验证（读盘）\n" + "─" * 70)
            await asyncio.sleep(1.5)
            verified, vmsgs = verify_with_openpyxl(fixture.working_path)
            for m in vmsgs:
                print(f"  {m}")

            print(rc.summary())

            all_ok = rc.passed == rc.total and verified
            print("\n" + "=" * 70)
            if all_ok:
                print("✅ 全部 37 事件 + 错误码 + openpyxl 验证通过。请用 Excel 打开做视觉验收：")
            else:
                print("⚠️  部分检查未通过（见上方汇总）；保留工作副本供调试：")
            print(f"   {fixture.working_path}")
            print("   视觉验收清单见 docs/manual_tests/excel_v0.3.0.md")
            print("=" * 70)
            return all_ok
    except Exception as exc:  # noqa: BLE001
        print(f"\n❌ 测试中断: {exc}")
        import traceback

        traceback.print_exc()
        print(rc.summary())
        return False


# ============================================================================
# CLI entry
# ============================================================================


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Excel E2E Tests (OASP /excel Draft 0.3.0)")
    parser.add_argument(
        "--mode",
        choices=["health", "full"],
        default="health",
        help=(
            "health: 校验 37 个事件已注册（无需 Add-In）；"
            "full: 完整 37 事件真机工作流 + openpyxl 验证 + 错误码（需 macOS + Excel + Add-In）"
        ),
    )
    args = parser.parse_args()

    try:
        if args.mode == "health":
            success = asyncio.run(test_workspace_health())
        else:
            success = asyncio.run(test_excel_full_e2e())
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
