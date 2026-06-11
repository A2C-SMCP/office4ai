"""
PPT Chart End-to-End Test (OASP /ppt Draft, v0.2.0 — Server OOXML)

Issue #9 验收：依次驱动 3 个新 chart MCP 工具并用 python-pptx 双重验证 OOXML 结构。
Chart 类事件由 OASP Server 端通过 python-pptx 直接修改 .pptx。

三种运行模式：

────────────────────────────────────────────────────────────────────
  --mode offline (默认)
────────────────────────────────────────────────────────────────────
不依赖 Add-In，纯离线跑通三件工具 + 错误码 + OOXML 双重验证。
主要用作 PR / CI 阶段的回归保护。

    uv run python manual_tests/ppt/test_chart_e2e.py
    # or
    uv run python manual_tests/ppt/test_chart_e2e.py --mode offline

覆盖：
- ✅ ppt_insert_chart   插入分类型图 + 散点图
- ✅ ppt_get_chart      回读数据
- ✅ ppt_update_chart   仅改标题（同 variant）+ 跨 variant 切换
- ✅ 3015 INVALID_CHART_DATA — categorical 维度不匹配 / scatter 非有限值
- ✅ 3010 ELEMENT_NOT_FOUND — 不存在的 chart-N
- ✅ requiresReload 标志 + notify_resource_updated 调用次数

────────────────────────────────────────────────────────────────────
  --mode pathb  (#16 — 打开态路径 B，模拟 Add-In)
────────────────────────────────────────────────────────────────────
忠实模拟 office-editor4ai 的两条 OASP 0.3.0 搬运事件（进程内 FakeAddIn），
把 #15 双路径路由器在 CONNECTED 下的整条 path-B 编排跑通，并用 python-pptx
验证「live 单页包」里的图表。不依赖真实设备。

    uv run python manual_tests/ppt/test_chart_e2e.py --mode pathb

覆盖：
- ✅ insert/get/update 三工具走 CONNECTED → 客户端 round-trip
     （事件序列 ppt:get:slideOoxml → ppt:insert:slidesOoxml）
- ✅ requiresReload=False（live 已就地更新，无需重开）
- ✅ path B 不碰盘（磁盘副本图表数恒为 0）
- ✅ 反应式降级：写超时 → 翻转后的 3003 文案 + 磁盘字节不变；读超时 → 回退读盘
⚠️  真机全链路联调 / Web·Windows 边界 / masterLeak 累积仍需真实 Add-In
    （office-editor4ai Task 2，#38/#39，尚 OPEN）——属阻塞项，不在本模式内。

────────────────────────────────────────────────────────────────────
  --mode conflict
────────────────────────────────────────────────────────────────────
**真实环境冲突实验**：把 Server-OOXML 路径（chart）与 Add-In Office.js 路径
（image）交替运行，观察两条写入路径在同一 .pptx 上的实际冲突表现。

需要：
  - macOS（依赖 AppleScript 强制保存）
  - PowerPoint 已启动并允许加载项
  - office-editor4ai Add-In 可用

    uv run python manual_tests/ppt/test_chart_e2e.py --mode conflict

实验目的：验证 README / docs/manual_tests/ppt_chart_v0.2.0.md 中描述的
「Server OOXML 写入 vs PowerPoint 内存模型」冲突——预期会观察到 chart
在 Add-In 触发 save() 后被 PowerPoint 内存覆盖的现象，从而为后续
「ppt:notify:reload」OASP 扩展提供数据依据。
"""

from __future__ import annotations

import asyncio
import shutil
import sys
import time
from pathlib import Path
from typing import Any
from unittest.mock import AsyncMock, MagicMock

# Make this script runnable both as `python -m` and as a path
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from pptx import Presentation  # noqa: E402

from office4ai.a2c_smcp.tools.ppt.get_chart import PptGetChartTool  # noqa: E402
from office4ai.a2c_smcp.tools.ppt.insert_chart import PptInsertChartTool  # noqa: E402
from office4ai.a2c_smcp.tools.ppt.update_chart import PptUpdateChartTool  # noqa: E402

WORKING_ROOT = _PROJECT_ROOT / "manual_tests" / ".test_working" / "ppt_chart_v0_2_0"


# ============================================================================
# Helpers
# ============================================================================


def _build_workspace_mock() -> tuple[Any, list[list[str]]]:
    """Mock the workspace surface the chart tools touch (notify + last activity)."""
    notified: list[list[str]] = []
    ws = MagicMock()
    ws.notify_resource_updated = lambda uris: notified.append(list(uris))
    ws.update_last_activity = MagicMock()
    return ws, notified


def _create_blank_deck(path: Path) -> None:
    """Two blank slides; chart tests will populate them."""
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[5])
    prs.slides.add_slide(prs.slide_layouts[5])
    prs.save(str(path))


def _count_charts(path: Path) -> int:
    prs = Presentation(str(path))
    n = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            if getattr(shape, "has_chart", False):
                n += 1
    return n


def _chart_titles(path: Path) -> list[str]:
    prs = Presentation(str(path))
    titles: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_chart", False):
                continue
            chart = shape.chart
            if chart.has_title:
                titles.append(chart.chart_title.text_frame.text)
            else:
                titles.append("")
    return titles


# ============================================================================
# Path B simulation (#16) — in-process stand-in for office-editor4ai Task 2
# ============================================================================


def _encode_b64(prs: Any) -> str:
    """Serialize a ``Presentation`` to a base64 .pptx package (mirrors chart_engine)."""
    import base64
    import io

    buf = io.BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _decode_b64_to_prs(slide_b64: str) -> Any:
    import base64
    import io

    return Presentation(io.BytesIO(base64.b64decode(slide_b64)))


def _chart_titles_in_b64(slide_b64: str) -> list[str]:
    """Chart titles found inside a base64 single-slide package."""
    prs = _decode_b64_to_prs(slide_b64)
    titles: list[str] = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_chart", False):
                continue
            chart = shape.chart
            titles.append(chart.chart_title.text_frame.text if chart.has_title else "")
    return titles


class FakeAddIn:
    """In-process stand-in for office-editor4ai's path-B carrier-event handlers (#16).

    The real Add-In (office-editor4ai Task 2 — issues #38/#39, still OPEN) exports a live
    slide via Office.js ``Slide.exportAsBase64`` and re-applies a rebuilt slide via
    ``insertSlidesFromBase64``. Until it ships we model that contract faithfully in-process:
    a *live* (open, possibly-unsaved) deck = a dict of single-slide base64 packages keyed by
    slide index. This lets the Server's #15 router run its full path-B orchestration
    (``ppt:get:slideOoxml`` → chart_engine base64 work → ``ppt:insert:slidesOoxml``) against
    real OOXML, with python-pptx verification of the resulting live slide — no real device.

    NOT a substitute for the real-device joint debug / Web·Windows boundary / masterLeak
    spot-checks (those need a real Add-In and stay blocked on office-editor4ai Task 2).
    """

    def __init__(self) -> None:
        self._slides: dict[int, str] = {}
        self.events: list[str] = []

    def live_slide_b64(self, slide_index: int) -> str:
        """Current base64 package for a slide; a fresh blank slide if never written."""
        if slide_index not in self._slides:
            prs = Presentation()
            prs.slides.add_slide(prs.slide_layouts[6])  # blank
            self._slides[slide_index] = _encode_b64(prs)
        return self._slides[slide_index]

    def emit(self, document_uri: str, event: str, data: dict[str, Any]) -> dict[str, Any]:
        """Sync side_effect for ``AsyncMock`` — the two OASP 0.3.0 carrier events."""
        self.events.append(event)
        if event == "ppt:get:slideOoxml":
            idx = data["slideIndex"]
            return {
                "success": True,
                "data": {"slideIndex": idx, "slideId": f"sid-{idx}", "base64": self.live_slide_b64(idx)},
            }
        if event == "ppt:insert:slidesOoxml":
            idx = data["targetSlideIndex"]
            # #15 in-place replace contract: the router MUST echo the exported opaque slideId
            # (else the old slide is left behind → duplicate page). Assert it round-trips.
            assert data.get("replaceSlideId") == f"sid-{idx}", (
                f"router must echo the exported opaque slideId; got {data.get('replaceSlideId')!r}"
            )
            self._slides[idx] = data["base64"]  # apply the rebuilt slide in place
            final = data.get("finalSlideIndex", idx)
            return {"success": True, "data": {"insertedSlideIndices": [final], "insertedSlideIds": [f"sid-{idx}"]}}
        raise AssertionError(f"unexpected path-B event: {event!r}")


def _build_connected_workspace_mock(emit_side_effect: Any) -> tuple[Any, list[list[str]]]:
    """A CONNECTED workspace whose ``emit_to_document`` runs ``emit_side_effect``.

    Mirrors the #15 unit-test rig (``_connected_ws`` + ``_make_path_b_emit``) but the fake
    is *stateful*, so insert → get → update form a coherent live round-trip. Pass a
    :class:`FakeAddIn`'s ``emit`` for the happy path, or an exception to drive degradation.
    """
    from office4ai.environment.workspace.base import DocumentStatus

    notified: list[list[str]] = []
    ws = MagicMock()
    ws.notify_resource_updated = lambda uris: notified.append(list(uris))
    ws.update_last_activity = MagicMock()
    ws.get_document_status = MagicMock(return_value=DocumentStatus.CONNECTED)
    ws.emit_to_document = AsyncMock(side_effect=emit_side_effect)
    return ws, notified


# ============================================================================
# Test scenarios
# ============================================================================


async def scenario_insert_categorical_chart(workspace: Any, deck_uri: str) -> str:
    """A.1 插入柱形图（带标题、图例）"""
    print("\n📊 A.1  ppt_insert_chart  柱形图")
    tool = PptInsertChartTool(workspace)
    result = await tool.execute(
        {
            "document_uri": deck_uri,
            "chart": {
                "chartType": "ColumnClustered",
                "categories": ["Q1", "Q2", "Q3", "Q4"],
                "series": [
                    {"name": "营收", "values": [120, 135, 158, 180]},
                    {"name": "成本", "values": [80, 90, 102, 115]},
                ],
                "title": "2026 年度业绩",
                "showLegend": True,
                "showDataLabels": False,
            },
            "options": {"slideIndex": 0, "left": 60, "top": 60, "width": 540, "height": 360},
        }
    )
    assert result["success"], f"插入失败: {result}"
    assert result["data"]["requiresReload"] is True
    eid = result["data"]["elementId"]
    print(f"   ✅ elementId={eid} chartType={result['data']['chartType']} series={result['data']['seriesCount']}")
    return eid


async def scenario_get_chart(workspace: Any, deck_uri: str, eid: str) -> None:
    """A.2 读回柱形图数据"""
    print("\n🔍 A.2  ppt_get_chart  读回数据")
    tool = PptGetChartTool(workspace)
    result = await tool.execute({"document_uri": deck_uri, "elementId": eid})
    assert result["success"], f"读取失败: {result}"
    chart = result["data"]["chart"]
    assert chart["chartType"] == "ColumnClustered"
    assert chart["categories"] == ["Q1", "Q2", "Q3", "Q4"]
    assert chart["title"] == "2026 年度业绩"
    assert len(chart["series"]) == 2
    print(f"   ✅ chartType={chart['chartType']} categories={chart['categories']} title='{chart['title']}'")


async def scenario_update_title(workspace: Any, deck_uri: str, eid: str) -> None:
    """A.3 仅修改标题（同 variant，部分更新）"""
    print("\n✏️  A.3  ppt_update_chart  仅修改标题")
    tool = PptUpdateChartTool(workspace)
    result = await tool.execute(
        {
            "document_uri": deck_uri,
            "elementId": eid,
            "chart": {"chartType": "ColumnClustered", "title": "2026 年度业绩（修订版）"},
        }
    )
    assert result["success"], f"更新失败: {result}"
    assert "title" in result["data"]["updatedFields"]
    assert result["data"]["requiresReload"] is True
    print(f"   ✅ updatedFields={result['data']['updatedFields']}")


async def scenario_insert_scatter(workspace: Any, deck_uri: str) -> str:
    """A.4 插入散点图（不同 variant，验证 discriminated union）"""
    print("\n📈 A.4  ppt_insert_chart  散点图")
    tool = PptInsertChartTool(workspace)
    result = await tool.execute(
        {
            "document_uri": deck_uri,
            "chart": {
                "chartType": "Scatter",
                "series": [
                    {
                        "name": "广告 vs 销售",
                        "points": [
                            {"x": 1000, "y": 50},
                            {"x": 1500, "y": 80},
                            {"x": 3000, "y": 200},
                            {"x": 5000, "y": 350},
                            {"x": 8000, "y": 600},
                        ],
                    }
                ],
                "title": "广告投入 vs 销售额",
            },
            "options": {"slideIndex": 1, "left": 60, "top": 60, "width": 540, "height": 360},
        }
    )
    assert result["success"], f"插入散点失败: {result}"
    eid = result["data"]["elementId"]
    print(f"   ✅ elementId={eid}")
    return eid


async def scenario_cross_variant(workspace: Any, deck_uri: str, eid: str) -> str:
    """A.5 跨 variant 切换：散点图 → 折线图（必须补 categories+series）"""
    print("\n🔄 A.5  ppt_update_chart  跨 variant：Scatter → Line")
    tool = PptUpdateChartTool(workspace)
    result = await tool.execute(
        {
            "document_uri": deck_uri,
            "elementId": eid,
            "chart": {
                "chartType": "Line",
                "categories": ["Jan", "Feb", "Mar", "Apr"],
                "series": [{"name": "趋势", "values": [10, 25, 30, 28]}],
            },
        }
    )
    assert result["success"], f"跨 variant 失败: {result}"
    assert result["data"]["chartType"] == "Line"
    new_eid = result["data"]["elementId"]
    print(f"   ✅ 新 elementId={new_eid} updatedFields={result['data']['updatedFields']}")
    return new_eid


async def scenario_3015_categorical_dimension(workspace: Any, deck_uri: str) -> None:
    """B.1 错误码 3015 INVALID_CHART_DATA — categorical 维度不匹配"""
    print("\n⚠️  B.1  3015 INVALID_CHART_DATA  categorical 维度不匹配")
    tool = PptInsertChartTool(workspace)
    result = await tool.execute(
        {
            "document_uri": deck_uri,
            "chart": {
                "chartType": "Pie",
                "categories": ["A", "B", "C"],
                "series": [{"name": "x", "values": [1, 2]}],  # 2 != 3
            },
        }
    )
    assert result["success"] is False, f"应失败但成功: {result}"
    assert "3015" in result["error"], f"错误码非 3015: {result['error']}"
    print(f"   ✅ 错误返回: {result['error']}")


async def scenario_3015_scatter_non_finite(workspace: Any, deck_uri: str) -> None:
    """B.2 错误码 3015 INVALID_CHART_DATA — scatter 含 NaN/Infinity"""
    print("\n⚠️  B.2  3015 INVALID_CHART_DATA  scatter 含非有限值")
    tool = PptInsertChartTool(workspace)
    result = await tool.execute(
        {
            "document_uri": deck_uri,
            "chart": {
                "chartType": "Scatter",
                "series": [{"name": "x", "points": [{"x": float("inf"), "y": 1}]}],
            },
        }
    )
    assert result["success"] is False
    assert "3015" in result["error"]
    print(f"   ✅ 错误返回: {result['error']}")


async def scenario_3010_element_not_found(workspace: Any, deck_uri: str) -> None:
    """B.3 错误码 3010 ELEMENT_NOT_FOUND — 不存在的 chart-N"""
    print("\n⚠️  B.3  3010 ELEMENT_NOT_FOUND  不存在的 elementId")
    tool = PptGetChartTool(workspace)
    result = await tool.execute({"document_uri": deck_uri, "elementId": "chart-99999"})
    assert result["success"] is False
    assert "3010" in result["error"]
    print(f"   ✅ 错误返回: {result['error']}")


# ============================================================================
# Runner
# ============================================================================


async def offline_main() -> int:
    print("=" * 70)
    print("PPT Chart MCP Tools — End-to-End OFFLINE (Server OOXML, OASP /ppt Draft v0.2.0)")
    print("=" * 70)

    # Prepare working dir
    WORKING_ROOT.mkdir(parents=True, exist_ok=True)
    deck_path = WORKING_ROOT / f"charts_demo_{int(time.time())}.pptx"
    _create_blank_deck(deck_path)
    deck_uri = deck_path.as_uri()
    print(f"\n📄 工作副本: {deck_path}")

    workspace, notified = _build_workspace_mock()

    # A. Happy path
    eid_col = await scenario_insert_categorical_chart(workspace, deck_uri)
    await scenario_get_chart(workspace, deck_uri, eid_col)
    await scenario_update_title(workspace, deck_uri, eid_col)

    eid_sc = await scenario_insert_scatter(workspace, deck_uri)
    await scenario_cross_variant(workspace, deck_uri, eid_sc)

    # B. Error paths (must NOT fire reload notification)
    notified_before_b = len(notified)
    await scenario_3015_categorical_dimension(workspace, deck_uri)
    await scenario_3015_scatter_non_finite(workspace, deck_uri)
    await scenario_3010_element_not_found(workspace, deck_uri)
    assert len(notified) == notified_before_b, "错误路径不应触发 notify_resource_updated"

    # OOXML-level structural checks
    print("\n🧪 C  OOXML 双重验证（python-pptx 直接读取 .pptx）")
    chart_count = _count_charts(deck_path)
    assert chart_count == 2, f"预期 2 个图表，实际 {chart_count}"
    print(f"   ✅ 文档中图表数 = {chart_count}")

    titles = _chart_titles(deck_path)
    print(f"   ✅ 图表标题列表 = {titles}")
    assert any("修订版" in t for t in titles), "应找到修订版标题"

    # Notify accounting
    write_ops = 4  # 1 insert col + 1 update title + 1 insert scatter + 1 cross-variant update
    print(f"\n🔔 D  notify_resource_updated 被调用次数 = {len(notified)} (预期 {write_ops})")
    assert len(notified) == write_ops
    for uris in notified:
        assert uris == ["window://office4ai/ppt", "window://office4ai"]

    print("\n" + "=" * 70)
    print(f"✅ 全部自动场景通过。请用 PowerPoint 打开 {deck_path.name} 做视觉验收：")
    print("   - 第 1 张幻灯片：柱形图（标题『2026 年度业绩（修订版）』，营收/成本两条 series）")
    print("   - 第 2 张幻灯片：折线图（标题『广告投入 vs 销售额』，由散点图跨 variant 切换而来）")
    print("=" * 70)
    return 0


# ============================================================================
# PATH-B MODE (#16) — open-document client round-trip against a simulated Add-In
# ============================================================================
# Drives the #15 dual-path router (insert / get / update) end-to-end through the
# real OASP 0.3.0 wire DTOs against an in-process FakeAddIn that synthesizes and
# applies real single-slide OOXML. Fully unblocked (no real device); the real
# Add-In joint debug / Web·Win / masterLeak spot-checks remain blocked on
# office-editor4ai Task 2 and are listed in docs/manual_tests/ppt_chart_v0.3.0.md.


async def pathb_main() -> int:
    # Lazy imports — keep offline mode lightweight (mirrors conflict_main).
    from office4ai.environment.workspace.dtos.ppt import (
        CategoricalChartData,
        CategoricalSeries,
        ChartInsertOptions,
    )
    from office4ai.environment.workspace.services import chart_engine

    print("=" * 70)
    print("PPT Chart MCP Tools — OPEN-DOCUMENT PATH B (simulated Add-In, OASP 0.3.0)")
    print("=" * 70)
    print()
    print("路径 B = 打开态客户端 round-trip。真实 Add-In（office-editor4ai Task 2，")
    print("issues #38/#39）尚未发布——这里用进程内 FakeAddIn 忠实模拟两条 OASP 0.3.0")
    print("搬运事件，把 #15 路由器整条 path-B 编排跑通并用 python-pptx 验证 live slide。")
    print("⚠️  真机联调 / Web·Windows 边界 / masterLeak 累积仍需真实 Add-In（阻塞项）。")

    WORKING_ROOT.mkdir(parents=True, exist_ok=True)
    open_deck = WORKING_ROOT / f"pathb_open_{int(time.time())}.pptx"
    _create_blank_deck(open_deck)  # disk stays blank — path B is entirely in-memory
    open_uri = open_deck.as_uri()
    print(f"\n📄 打开态工作副本（磁盘保持空，用于证明 path B 不碰盘）: {open_deck}")

    fake = FakeAddIn()
    workspace, notified = _build_connected_workspace_mock(fake.emit)

    # ── P.1  insert path B — existing-page live round-trip ────────────────────
    print("\n📊 P.1  ppt_insert_chart  路径 B（CONNECTED → 整页 round-trip）")
    ev0 = len(fake.events)
    insert_tool = PptInsertChartTool(workspace)
    r1 = await insert_tool.execute(
        {
            "document_uri": open_uri,
            "chart": {
                "chartType": "ColumnClustered",
                "categories": ["Q1", "Q2", "Q3", "Q4"],
                "series": [{"name": "营收", "values": [120, 135, 158, 180]}],
                "title": "打开态业绩",
                "showLegend": True,
            },
            "options": {"slideIndex": 0, "left": 60, "top": 60, "width": 540, "height": 360},
        }
    )
    assert r1["success"], f"path B 插入失败: {r1}"
    assert r1["data"]["requiresReload"] is False, f"CONNECTED 应 requiresReload=False: {r1['data']}"
    assert fake.events[ev0:] == ["ppt:get:slideOoxml", "ppt:insert:slidesOoxml"], fake.events[ev0:]
    eid = r1["data"]["elementId"]
    live_titles = _chart_titles_in_b64(fake.live_slide_b64(0))
    assert live_titles == ["打开态业绩"], f"live 单页图表标题异常: {live_titles}"
    assert _count_charts(open_deck) == 0, "path B 不应写磁盘（磁盘副本应仍为 0 图表）"
    print(f"   ✅ elementId={eid} 事件序列={fake.events[ev0:]} requiresReload=False")
    print(f"   ✅ live 单页图表标题={live_titles}；磁盘副本图表数=0（path B 未碰盘）")

    # ── P.2  get path B — read the live (unsaved) slide ───────────────────────
    print("\n🔍 P.2  ppt_get_chart  路径 B（读 live 单页，非磁盘）")
    ev1 = len(fake.events)
    get_tool = PptGetChartTool(workspace)
    r2 = await get_tool.execute({"document_uri": open_uri, "elementId": eid, "slideIndex": 0})
    assert r2["success"], f"path B 读取失败: {r2}"
    chart = r2["data"]["chart"]
    assert chart["chartType"] == "ColumnClustered", chart
    assert chart["categories"] == ["Q1", "Q2", "Q3", "Q4"], chart
    assert chart["title"] == "打开态业绩", chart
    assert fake.events[ev1:] == ["ppt:get:slideOoxml"], fake.events[ev1:]
    print(f"   ✅ 回读 chartType={chart['chartType']} categories={chart['categories']} title='{chart['title']}'")

    # ── P.3  update path B — mutate the chart in place on the live slide ──────
    print("\n✏️  P.3  ppt_update_chart  路径 B（就地改 live 单页标题）")
    ev2 = len(fake.events)
    update_tool = PptUpdateChartTool(workspace)
    r3 = await update_tool.execute(
        {
            "document_uri": open_uri,
            "elementId": eid,
            "slideIndex": 0,
            "chart": {"chartType": "ColumnClustered", "title": "打开态业绩（已更新）"},
        }
    )
    assert r3["success"], f"path B 更新失败: {r3}"
    assert "title" in r3["data"]["updatedFields"], r3["data"]
    assert r3["data"]["requiresReload"] is False, r3["data"]
    assert fake.events[ev2:] == ["ppt:get:slideOoxml", "ppt:insert:slidesOoxml"], fake.events[ev2:]
    live_titles = _chart_titles_in_b64(fake.live_slide_b64(0))
    assert live_titles == ["打开态业绩（已更新）"], f"更新后 live 标题异常: {live_titles}"
    print(f"   ✅ updatedFields={r3['data']['updatedFields']}；live 标题={live_titles}")

    # Writes notify /ppt subscribers on both paths; reads do not.
    assert len(notified) == 2, f"预期 2 次写通知（P.1 insert + P.3 update），实际 {len(notified)}"
    for uris in notified:
        assert uris == ["window://office4ai/ppt", "window://office4ai"], uris
    print(f"   ✅ notify_resource_updated 调用 {len(notified)} 次（insert + update；get 只读不通知）")

    # ── P.4  reactive WRITE degrade — Add-In carrier handler absent → 3003 ───
    print("\n⚠️  P.4  反应式写降级（搬运事件超时 → 翻转后的 3003 + 磁盘字节不变）")
    degrade_deck = WORKING_ROOT / f"pathb_degrade_{int(time.time())}.pptx"
    _create_blank_deck(degrade_deck)
    degrade_uri = degrade_deck.as_uri()
    before_bytes = degrade_deck.read_bytes()
    ws_timeout, _ = _build_connected_workspace_mock(TimeoutError("no ack from Add-In"))
    r4 = await PptInsertChartTool(ws_timeout).execute(
        {
            "document_uri": degrade_uri,
            "chart": {
                "chartType": "Pie",
                "categories": ["a", "b"],
                "series": [{"name": "x", "values": [1, 2]}],
                "title": "SHOULD-DEGRADE",
            },
            "options": {"slideIndex": 0},
        }
    )
    assert r4["success"] is False, f"path B 不可用时写应降级: {r4}"
    assert "3003" in r4["error"], f"降级应带 3003 前缀: {r4['error']}"
    assert degrade_deck.read_bytes() == before_bytes, "写降级不得改动磁盘字节"
    print(f"   ✅ 降级返回: {r4['error'][:72]}...")
    print("   ✅ 磁盘字节零变化（降级未触发任何盘写）")

    # ── P.4b reactive WRITE degrade — Add-In acks 3016 API_NOT_SUPPORTED ─────
    # Completes the degrade matrix: timeout (transport gone) AND 3016 (handler
    # present but the platform's Office.js requirement set is unsupported) both
    # classify as "path B unavailable" → flipped 3003, disk untouched.
    print("\n⚠️  P.4b 反应式写降级（Add-In 回 3016 不支持 → 同样翻转 3003 + 磁盘字节不变）")
    deck_3016 = WORKING_ROOT / f"pathb_3016_{int(time.time())}.pptx"
    _create_blank_deck(deck_3016)
    uri_3016 = deck_3016.as_uri()
    before_3016 = deck_3016.read_bytes()

    def _emit_3016(document_uri: str, event: str, data: dict[str, Any]) -> dict[str, Any]:
        return {"success": False, "error": {"code": "3016", "message": "PowerPointApi 1.8 unavailable"}}

    ws_3016, _ = _build_connected_workspace_mock(_emit_3016)
    r4b = await PptInsertChartTool(ws_3016).execute(
        {
            "document_uri": uri_3016,
            "chart": {"chartType": "Pie", "categories": ["a", "b"], "series": [{"name": "x", "values": [1, 2]}]},
            "options": {"slideIndex": 0},
        }
    )
    assert r4b["success"] is False, f"3016 应触发降级: {r4b}"
    assert "3003" in r4b["error"], f"3016 降级应回退 3003 文案: {r4b['error']}"
    assert deck_3016.read_bytes() == before_3016, "3016 降级不得改动磁盘字节"
    print(f"   ✅ 3016 降级返回: {r4b['error'][:72]}...")

    # ── P.5  reactive READ fallback — get degrades to the on-disk read ───────
    print("\n🔁 P.5  反应式读降级（搬运事件超时 → 回退读盘）")
    disk_deck = WORKING_ROOT / f"pathb_diskread_{int(time.time())}.pptx"
    _create_blank_deck(disk_deck)
    disk_uri = disk_deck.as_uri()
    seeded = await chart_engine.insert_chart(
        disk_uri,
        CategoricalChartData(
            chartType="Line",
            categories=["Jan", "Feb"],
            series=[CategoricalSeries(name="趋势", values=[10, 20])],
            title="DISK-CHART",
        ),
        ChartInsertOptions(slideIndex=0),
    )
    ws_timeout_read, _ = _build_connected_workspace_mock(TimeoutError("no ack from Add-In"))
    r5 = await PptGetChartTool(ws_timeout_read).execute(
        {"document_uri": disk_uri, "elementId": seeded["elementId"], "slideIndex": 0}
    )
    assert r5["success"], f"读应回退读盘并成功: {r5}"
    assert r5["data"]["chart"]["title"] == "DISK-CHART", r5["data"]["chart"]
    print(f"   ✅ 回退读盘成功，回读 title='{r5['data']['chart']['title']}'")

    print("\n" + "=" * 70)
    print("✅ 路径 B（模拟 Add-In）全部场景通过。")
    print("   真机全链路联调 / Web·Windows 边界 / masterLeak 累积 → 待 office-editor4ai")
    print("   Task 2（#38/#39）发布后执行，清单见 docs/manual_tests/ppt_chart_v0.3.0.md。")
    print("=" * 70)
    return 0


# ============================================================================
# CONFLICT-MODE EXPERIMENT (real PowerPoint + real Add-In)
# ============================================================================
# Demonstrates the conflict between two write paths on the same .pptx:
#   - Server-OOXML path (chart_engine writes the file directly)
#   - Add-In Office.js path (PowerPoint mutates its in-memory model;
#     write to disk only happens on save)
#
# When PowerPoint flushes its in-memory model to disk, it does NOT merge with
# external changes — it overwrites them. So a Server-inserted chart can be
# wiped out by a subsequent Add-In save.

# 1×1 transparent PNG (smallest valid baseline image)
_MINIMAL_PNG_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8BQDwAEgAF/pooBPQAAAABJRU5ErkJggg=="


def _observe_disk(deck: Path, label: str) -> dict[str, Any]:
    """Read the on-disk .pptx, return per-slide chart / picture counts."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    prs = Presentation(str(deck))
    per_slide: list[dict[str, int]] = []
    total_charts = 0
    total_pictures = 0
    for slide in prs.slides:
        ch = sum(1 for s in slide.shapes if getattr(s, "has_chart", False))
        pic = sum(1 for s in slide.shapes if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
        per_slide.append({"charts": ch, "pictures": pic})
        total_charts += ch
        total_pictures += pic
    snapshot = {
        "label": label,
        "slides": len(prs.slides),
        "charts": total_charts,
        "pictures": total_pictures,
        "per_slide": per_slide,
    }
    print(f"   📋 [{label}] charts={total_charts}, pictures={total_pictures}, per_slide={per_slide}")
    return snapshot


async def conflict_main() -> int:
    """Run the interleaved chart vs image experiment against a real Add-In."""
    import platform

    if platform.system() != "Darwin":
        print("❌ conflict 模式仅支持 macOS（依赖 AppleScript 强制保存 PowerPoint）", file=sys.stderr)
        return 2

    # Lazy imports — these touch a lot of test infra; keep offline mode lightweight.
    from manual_tests.ppt.e2e_base import (  # noqa: E402
        PPTTestRunner,
        save_ppt_document,
    )
    from manual_tests.ppt.test_helpers import ppt_insert_image  # noqa: E402

    print("=" * 70)
    print("PPT Chart × Image CONFLICT EXPERIMENT")
    print("(Server-OOXML chart writes vs Add-In Office.js image writes)")
    print("=" * 70)
    print()
    print("⚠️  本实验包含三个阶段：")
    print("    Phase A/B  — 直接调用 chart_engine 绕过工具防御，复现冲突场景")
    print("                 （证明「没有防御时，PowerPoint save 会覆盖 Server 写入」）")
    print("    Phase C    — 走完整工具路径，验证 #15 翻转守卫后的反应式降级（CONNECTED →")
    print("                 路由 path B；真实 Add-In 搬运事件未发布 → 回退 3003），")
    print("                 引导用户关闭文档，关闭后工具放行，最后重新打开验证持久化")
    print()

    # Use the existing PPT fixture infra so an empty deck is opened in PowerPoint
    # and the office-editor4ai Add-In is auto-activated.
    runner = PPTTestRunner(
        fixtures_dir=_PROJECT_ROOT / "manual_tests" / "ppt" / "fixtures" / "ppt_e2e",
        cleanup_on_success=False,  # keep the deck for post-mortem viewing in PowerPoint
    )

    snapshots: list[dict[str, Any]] = []

    # Use multi_slide.pptx — it has 5 slides so Phase A (slide 0) and Phase B
    # (slide 1) operate on independent slides. empty.pptx only has 1 slide so
    # Phase B's slideIndex=1 would be out-of-range.
    async with runner.run_with_workspace("multi_slide.pptx") as (workspace, fixture):
        deck_path: Path = fixture.working_path
        deck_uri: str = fixture.document_uri
        print(f"\n📄 工作副本: {deck_path}")

        snapshots.append(_observe_disk(deck_path, "0  起点：空 .pptx"))

        # IMPORTANT: Phase A/B 故意绕过工具层的 CONNECTED 防御，直接调用
        # chart_engine 来「无视防御」地写盘——这样才能复现没有防御时 PowerPoint
        # 内存模型覆盖磁盘的冲突场景。Phase C 才走完整的工具路径来验证防御 +
        # 用户关闭后的恢复流程。
        from office4ai.environment.workspace.dtos.ppt import (
            CategoricalChartData,
            CategoricalSeries,
            ChartInsertOptions,
        )
        from office4ai.environment.workspace.services import chart_engine

        # Tool instance is reused in Phase C (defense path).
        insert_chart_tool = PptInsertChartTool(workspace)

        # ────────────────────────────────────────────────────────────────
        # Phase A — Server first, Add-In second (no save in between).
        # Disk gets the chart from chart_engine (we bypass the tool's
        # CONNECTED defense on purpose to demonstrate the underlying
        # conflict). PowerPoint memory diverges when Add-In adds an image;
        # the save AFTER both reveals what PowerPoint flushes to disk.
        # ────────────────────────────────────────────────────────────────
        print("\n" + "─" * 70)
        print("Phase A — Server inserts chart → Add-In inserts image → save")
        print("        （Server 路径绕过 CONNECTED 防御，仅用于复现冲突）")
        print("─" * 70)

        print("\n🅰️  A.1  Server: chart_engine.insert_chart on slide 0（绕过工具防御）")
        r1 = await chart_engine.insert_chart(
            deck_uri,
            CategoricalChartData(
                chartType="ColumnClustered",
                categories=["Q1", "Q2", "Q3"],
                series=[CategoricalSeries(name="Rev", values=[100, 150, 200])],
                title="PHASE-A-CHART",
            ),
            ChartInsertOptions(slideIndex=0, left=60, top=60, width=360, height=240),
        )
        print(f"   ✅ Server 写盘成功，elementId={r1['elementId']}")
        snapshots.append(_observe_disk(deck_path, "A.1 after Server chart write"))

        print("\n🅰️  A.2  Add-In: ppt_insert_image on slide 0 (内存写入，未 save)")
        ok, data, err = await ppt_insert_image(
            workspace,
            deck_uri,
            {"base64": _MINIMAL_PNG_BASE64},
            options={"slideIndex": 0, "left": 450, "top": 60, "width": 100, "height": 100},
        )
        assert ok, f"Add-In image insert 失败: {err}"
        print(f"   ✅ Add-In 内存改动成功: {data}")
        # 此时磁盘还未感知 image — PowerPoint 还没保存。
        snapshots.append(_observe_disk(deck_path, "A.2 right after Add-In insert (BEFORE save)"))

        print("\n🅰️  A.3  AppleScript 强制 PowerPoint 保存（用户按 Cmd+S 的等价操作）")
        if not save_ppt_document(deck_path):
            print("   ⚠️  保存失败，无法继续 Phase A")
            return 3
        await asyncio.sleep(1.0)  # let the save finish flushing
        snap_a3 = _observe_disk(deck_path, "A.3 AFTER save — 关键观察点")
        snapshots.append(snap_a3)

        # Did the chart survive?
        a_chart_lost = snap_a3["charts"] == 0
        a_image_present = snap_a3["pictures"] >= 1
        print()
        if a_chart_lost and a_image_present:
            print("   ❗ 观察：保存后 chart 已被 PowerPoint 内存模型覆盖（消失），仅剩 image。")
            print("       这正是 OASP /ppt 文档警告的冲突场景。")
        elif not a_chart_lost and a_image_present:
            print("   🤔 观察：保存后 chart + image 共存——PowerPoint 可能感知到了外部修改。")
            print("       值得细看：是 PowerPoint 自动 merge 还是文件锁阻止了我们的 OOXML 写入？")
        else:
            print(f"   ⚠️  非预期状态：charts={snap_a3['charts']}, pictures={snap_a3['pictures']}")

        # ────────────────────────────────────────────────────────────────
        # Phase B — Add-In first (with save), then Server, then Add-In op,
        # then save. Tests whether the Server-inserted chart can survive a
        # subsequent Add-In save flow.
        # ────────────────────────────────────────────────────────────────
        print("\n" + "─" * 70)
        print("Phase B — Add-In op → save → Server inserts chart → Add-In op → save")
        print("        （Server 路径同样绕过 CONNECTED 防御）")
        print("─" * 70)

        print("\n🅱️  B.1  Add-In: ppt_insert_image on slide 1（内存）+ 立即保存")
        ok, data, err = await ppt_insert_image(
            workspace,
            deck_uri,
            {"base64": _MINIMAL_PNG_BASE64},
            options={"slideIndex": 1, "left": 60, "top": 60, "width": 100, "height": 100},
        )
        assert ok, f"Add-In image insert 失败: {err}"
        if not save_ppt_document(deck_path):
            print("   ⚠️  保存失败")
            return 3
        await asyncio.sleep(1.0)
        snapshots.append(_observe_disk(deck_path, "B.1 after Add-In img + save"))

        print("\n🅱️  B.2  Server: chart_engine.insert_chart on slide 1（绕过工具防御，直接写盘）")
        r2 = await chart_engine.insert_chart(
            deck_uri,
            CategoricalChartData(
                chartType="Pie",
                categories=["alpha", "beta", "gamma"],
                series=[CategoricalSeries(name="share", values=[30, 50, 20])],
                title="PHASE-B-CHART",
            ),
            ChartInsertOptions(slideIndex=1, left=200, top=60, width=360, height=240),
        )
        print(f"   ✅ Server 写盘成功，elementId={r2['elementId']}")
        snap_b2 = _observe_disk(deck_path, "B.2 right after Server chart write")
        snapshots.append(snap_b2)

        print("\n🅱️  B.3  Add-In: 再插一张 image on slide 1（内存）— 模拟用户继续编辑")
        ok, data, err = await ppt_insert_image(
            workspace,
            deck_uri,
            {"base64": _MINIMAL_PNG_BASE64},
            options={"slideIndex": 1, "left": 460, "top": 200, "width": 80, "height": 80},
        )
        assert ok, f"Add-In second image insert 失败: {err}"
        snapshots.append(_observe_disk(deck_path, "B.3 after Add-In img (BEFORE save) — 关键观察点"))

        print("\n🅱️  B.4  AppleScript 强制保存 — 关键时刻")
        if not save_ppt_document(deck_path):
            print("   ⚠️  保存失败")
            return 3
        await asyncio.sleep(1.0)
        snap_b4 = _observe_disk(deck_path, "B.4 AFTER save — 关键观察点")
        snapshots.append(snap_b4)

        # In the worst-case prediction: PowerPoint memory has 2 images on slide 1
        # but no chart. Save flushes that → chart on slide 1 disappears.
        b_chart_lost = snap_b4["per_slide"][1]["charts"] == 0 and snap_b2["per_slide"][1]["charts"] == 1
        b_images_present = snap_b4["per_slide"][1]["pictures"] >= 2
        print()
        if b_chart_lost and b_images_present:
            print("   ❗ 观察：B.2 后存在的 chart 被 B.4 的 save 覆盖丢失。")
            print("       后续 Add-In 操作 + save 会持续覆盖任何 Server-OOXML 写入。")
        elif not b_chart_lost:
            print("   🤔 观察：B.4 后 chart 仍然存在——PowerPoint 可能会感知到外部修改并保留。")

        # ────────────────────────────────────────────────────────────────
        # Phase C — Reactive degradation + recovery.
        # 1) #15 翻转守卫后 CONNECTED 不再「上抛 3003 拒绝」，而是路由到 path B。真实
        #    Add-In 的 path-B 搬运事件（office-editor4ai Task 2，#38/#39）尚未发布 →
        #    path B 不可用 → 反应式降级回退到带 "3003" 前缀的「先关闭文档」文案。
        #    ⚠️  一旦 Add-In 发布搬运事件，本步将经 path B 直接成功——届时把 C.0 的
        #        「应被拒」预期翻转为「应成功（live round-trip）」。
        # 2) 引导用户关闭文档 → poll 直到 Add-In 断开
        # 3) 文档关闭后，再次调用 chart 工具应该成功（path A 写盘）
        # 4) AppleScript 重新打开文档
        # 5) 用 python-pptx 验证 chart 真的留在磁盘上（PowerPoint 读取的版本）
        # ────────────────────────────────────────────────────────────────
        print("\n" + "─" * 70)
        print("Phase C — 反应式降级 + 恢复路径（关闭后再插，验证 chart 真正持久化）")
        print("─" * 70)

        # C.0 — #15 翻转守卫：CONNECTED 路由 path B，但真实 Add-In 搬运事件未发布 →
        # 反应式降级回退带 "3003" 前缀的文案（Add-In 发布后此步会改为经 path B 成功）。
        print("\n🅲  C.0  反应式降级验证：CONNECTED → 路由 path B，Add-In 搬运事件未发布 → 回退 3003")
        degraded = await insert_chart_tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "Pie",
                    "categories": ["a"],
                    "series": [{"name": "x", "values": [1]}],
                    "title": "SHOULD-DEGRADE",
                },
                "options": {"slideIndex": 2},
            }
        )
        if degraded["success"]:
            print("   ℹ️  insert 直接成功——真实 Add-In 似乎已实装 path-B 搬运事件（office-editor4ai Task 2）。")
            print("       这正是 #16 的目标终态；届时请把本步「应被拒」预期翻转为「应成功（live round-trip）」。")
            return 4
        if "3003" not in degraded["error"]:
            print(f"   ⚠️  预期降级回退 3003，实际：{degraded['error']}")
            return 4
        print(f"   ✅ 反应式降级按预期回退 3003：{degraded['error'][:80]}...")
        snapshots.append(_observe_disk(deck_path, "C.0 after degraded insert (no disk change)"))

        # C.1 — Console 提示用户关闭文档
        print()
        print("=" * 70)
        print("🛑  请在 PowerPoint 中关闭这个文档：")
        print(f"      文件名：{deck_path.name}")
        print("      操作：点击文档窗口左上角红色关闭按钮，或菜单 File → Close")
        print("            （如系统提示是否保存，请选择「不保存」以保留 Phase B 写盘的 chart）")
        print("=" * 70)

        # C.2 — Poll 等待 Add-In disconnect
        print("\n   ⏳ 等待 Add-In 断开（Socket.IO 连接消失）...")
        poll_timeout = 90.0
        poll_interval = 1.0
        from office4ai.environment.workspace.base import DocumentStatus

        start = asyncio.get_event_loop().time()
        disconnected = False
        while asyncio.get_event_loop().time() - start < poll_timeout:
            status = workspace.get_document_status(deck_uri)
            if status == DocumentStatus.DISCONNECTED:
                disconnected = True
                elapsed = asyncio.get_event_loop().time() - start
                print(f"   ✅ Add-In 已断开（耗时 {elapsed:.1f}s）")
                break
            await asyncio.sleep(poll_interval)
        if not disconnected:
            print(f"   ⚠️  {poll_timeout}s 内未检测到断开，跳过 Phase C 后续步骤")
            return 5

        # 文档关闭后，给 PowerPoint 一秒让它真正释放文件句柄。
        await asyncio.sleep(1.0)
        snapshots.append(_observe_disk(deck_path, "C.2 after Add-In disconnect"))

        # C.3 — 现在 insert_chart 应该成功
        print("\n🅲  C.3  Server: ppt_insert_chart on slide 2（已断开，应成功）")
        r3 = await insert_chart_tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "ColumnClustered",
                    "categories": ["X", "Y", "Z"],
                    "series": [{"name": "delta", "values": [10, 20, 30]}],
                    "title": "PHASE-C-CHART",
                },
                "options": {"slideIndex": 2, "left": 60, "top": 60, "width": 480, "height": 320},
            }
        )
        if not r3["success"]:
            print(f"   ❌ 关闭后 insert_chart 仍失败：{r3}")
            return 6
        new_eid = r3["data"]["elementId"]
        print(f"   ✅ Server 写盘成功，elementId={new_eid}")
        snap_c3 = _observe_disk(deck_path, "C.3 after Server chart write (closed state)")
        snapshots.append(snap_c3)
        if snap_c3["per_slide"][2]["charts"] != 1:
            print(f"   ⚠️  slide 2 chart 数预期 1，实际 {snap_c3['per_slide'][2]['charts']}")
            return 7

        # C.4 — 重新打开文档（在 PowerPoint 里）
        print("\n🅲  C.4  AppleScript 重新打开文档（PowerPoint 加载磁盘最新版本）")
        from manual_tests.e2e_base import open_document  # local import to keep offline-mode lightweight

        if not open_document(deck_path):
            print(f"   ⚠️  自动打开失败，请手动打开：{deck_path}")
        else:
            print(f"   ✅ 已请求 PowerPoint 打开 {deck_path.name}")
        # PowerPoint 打开 + 渲染 chart 通常需要几秒；不等待 Add-In 重连即可验证。
        await asyncio.sleep(3.0)

        # C.5 — 视觉验收提示
        print("\n🅲  C.5  请在 PowerPoint 中视觉确认：")
        print("       第 1 张幻灯片（slide 0）：含一张 image，无 chart（Phase A 残留）")
        print("       第 2 张幻灯片（slide 1）：含 2 张 image，无 chart（Phase B 残留）")
        print("       第 3 张幻灯片（slide 2）：✨ 含柱形图 'PHASE-C-CHART'（Phase C 新增）")

        # C.6 — Disk verify (independent of PowerPoint visual)
        snap_c6 = _observe_disk(deck_path, "C.6 after reopen (disk truth)")
        snapshots.append(snap_c6)
        c_chart_persisted = snap_c6["per_slide"][2]["charts"] == 1
        if c_chart_persisted:
            print()
            print("   ✅ Phase C chart 已持久化到磁盘——Server-OOXML 路径在 Add-In 关闭后完全可用。")
        else:
            print(f"   ⚠️  Phase C chart 在 reopen 后丢失：{snap_c6['per_slide'][2]}")

    # Done — summarize.
    print("\n" + "=" * 70)
    print("📊 实验总结（按时间顺序）")
    print("=" * 70)
    for i, snap in enumerate(snapshots):
        print(
            f"  [{i}] {snap['label']:50s}  charts={snap['charts']}  "
            f"pictures={snap['pictures']}  per_slide={snap['per_slide']}"
        )

    print()
    print("可在 PowerPoint 中重新打开工作副本以视觉复核：")
    print(f"   {fixture.working_path}")
    print()
    print("根据上面 [A.3] / [B.4] 的关键观察行可以进一步讨论冲突的根因与缓解方案。")
    return 0


def run(mode: str) -> None:
    try:
        if mode == "conflict":
            rc = asyncio.run(conflict_main())
        elif mode == "pathb":
            rc = asyncio.run(pathb_main())
        else:
            rc = asyncio.run(offline_main())
    except AssertionError as e:
        print(f"\n❌ 断言失败: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\n💥 未预期异常: {e}", file=sys.stderr)
        raise
    sys.exit(rc)


def _parse_args(argv: list[str]) -> tuple[str, bool]:
    """Tiny arg parser — avoid argparse to keep the script reasonably one-shot."""
    mode = "offline"
    clean = False
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--clean":
            clean = True
        elif a == "--mode":
            if i + 1 >= len(argv):
                print("❌ --mode requires a value (offline|pathb|conflict)", file=sys.stderr)
                sys.exit(2)
            mode = argv[i + 1]
            i += 1
        elif a.startswith("--mode="):
            mode = a.split("=", 1)[1]
        elif a in ("-h", "--help"):
            print(__doc__)
            sys.exit(0)
        i += 1
    if mode not in ("offline", "pathb", "conflict"):
        print(f"❌ unknown --mode {mode!r} (expected offline|pathb|conflict)", file=sys.stderr)
        sys.exit(2)
    return mode, clean


if __name__ == "__main__":
    mode, clean = _parse_args(sys.argv[1:])
    if clean and WORKING_ROOT.exists():
        shutil.rmtree(WORKING_ROOT)
        print(f"🧹 cleaned: {WORKING_ROOT}")
    run(mode)
