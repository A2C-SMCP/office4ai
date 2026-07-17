"""
Word Table Operations End-to-End Test (OASP /word Draft, v0.2.0)

针对 Issue #8 验收：在真实 Word 文档 + office-editor4ai Add-In 环境下，
依次驱动 4 个新表格 MCP 工具，并用 python-docx 双重验证 OOXML 结构。

覆盖的 Issue #8 手工测试清单项（自动化部分）：
- ✅ 端到端连通性（list_tools 已在 unit + integration 测试覆盖；本脚本验证 4 个事件实际能 emit）
- ✅ word_merge_cells 视觉（合并 5×4 表首行，OOXML 验证合并 span）
- ✅ word_update_table_cell 视觉（设置首行蓝底白字加粗居中表头）
- ✅ word_update_table_row_column 视觉（按行批量写入 4 行数据）
- ✅ word_update_table_format 视觉（边框 + 列宽 + 整表对齐 + 内边距）
- ✅ tableId 缺省命中（执行不传 tableId 的合并，由 Add-In 解析光标所在表格）

需要人工验证（保留在 docs/manual_tests/word_table_v0.2.0.md 清单中）：
- ❌ 3013 NO_TABLE_AT_CURSOR（需手工把光标移到普通段落）
- ❌ 3010 ELEMENT_NOT_FOUND（需手工用超界 tableId）
- ❌ 3014 ALREADY_MERGED（相交合并真·冲突已自动化为 B.3a，中文冲突报文已采样；
  实收 3014 依赖 Add-In 接线 editor4ai#88，接线前 B.3a 暂容忍 3000）
- ❌ AI 业务闭环（合同表头场景，需对接真实 LLM 工具调用）

运行方式：
    # 健康检查（不需要 Add-In）
    uv run python manual_tests/word/test_word_table_e2e.py --mode health

    # 端到端表格场景（需要 Add-In + macOS Word，文件保留供你目测）
    uv run python manual_tests/word/test_word_table_e2e.py --mode tables

成功后，工作副本保留在 manual_tests/.test_working/ 下，用 Word 打开它人眼检查"蓝底
表头 + 灰底标签 + 列宽合理 + 边框样式"等视觉效果。
"""

from __future__ import annotations

import asyncio
import sys
from pathlib import Path
from typing import Any

# Make this script runnable both as `python -m` and as a path
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from manual_tests.e2e_base import (  # noqa: E402
    DocumentReader,
    E2ETestRunner,
    create_empty_docx,
)
from manual_tests.word.test_helpers import (  # noqa: E402
    insert_table,
    merge_cells,
    update_table_cell,
    update_table_format,
    update_table_row_column,
)

# ============================================================================
# Fixture preparation
# ============================================================================

FIXTURES_DIR = _PROJECT_ROOT / "manual_tests" / "fixtures"


def ensure_empty_fixture() -> Path:
    """确保 empty.docx 夹具存在。"""
    FIXTURES_DIR.mkdir(parents=True, exist_ok=True)
    empty_path = FIXTURES_DIR / "empty.docx"
    if not empty_path.exists():
        print(f"📝 创建夹具: {empty_path}")
        create_empty_docx(empty_path)
    return empty_path


# ============================================================================
# OOXML-level verification via python-docx
# ============================================================================


def _table_grid_widths_emu(reader: DocumentReader, table_index: int = 0) -> list[int]:
    """读取表格的 tblGrid 列宽（EMU 单位，1 pt = 20 EMU dxa twips ≈ 1/20 of a point in dxa）。

    python-docx 暴露 tblGrid 子元素，宽度以 dxa（twentieths of a point）保存。
    """
    table = reader.doc.tables[table_index]
    grid = table._tbl.tblGrid  # type: ignore[attr-defined]
    return [int(col.w) for col in grid.gridCol_lst]


def _cell_shading_hex(reader: DocumentReader, table_index: int, row: int, col: int) -> str | None:
    """读取单元格背景色 (w:shd@w:fill)，返回如 '#1F4E79' 或 None。"""
    table = reader.doc.tables[table_index]
    cell = table.cell(row, col)
    tc_pr = cell._tc.tcPr  # type: ignore[attr-defined]
    if tc_pr is None:
        return None
    shd = tc_pr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd")
    if shd is None:
        return None
    fill = shd.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill")
    if not fill or fill in ("auto", ""):
        return None
    return f"#{fill.upper()}"


def _is_first_row_merged(reader: DocumentReader, table_index: int = 0) -> bool:
    """判断首行是否被合并（首行所有单元格的底层 _tc 是否指向同一个 tc）。"""
    table = reader.doc.tables[table_index]
    if not table.rows:
        return False
    first_row_cells = table.rows[0].cells
    if len(first_row_cells) <= 1:
        return False
    first_tc = first_row_cells[0]._tc  # type: ignore[attr-defined]
    return all(c._tc is first_tc for c in first_row_cells[1:])  # type: ignore[attr-defined]


def verify_table_state_after_pipeline(reader: DocumentReader) -> tuple[bool, list[str]]:
    """
    验证经过 4 个表格工具流水线后的 OOXML 结构。

    流水线后预期：
    - 文档至少 1 张表
    - 首行被合并（merge_cells 效果）
    - 首行（合并后）单元格背景色 = #1F4E79（update_table_cell 效果）
    - 第 2 行第 0 列含"甲方" / 第 2 行第 1 列含"ACME Corp"（update_table_row_column 效果）
    - 列宽数量 = 4（update_table_format columnWidths=[120,80,80,80] 效果）
    """
    messages: list[str] = []
    success = True

    if reader.table_count < 1:
        messages.append(f"❌ 表格数 {reader.table_count}，至少应有 1 张")
        return False, messages
    messages.append(f"✅ 表格数 = {reader.table_count}")

    if _is_first_row_merged(reader, table_index=0):
        messages.append("✅ 首行已合并 (merge_cells)")
    else:
        messages.append("❌ 首行未合并，merge_cells 效果未生效")
        success = False

    bg = _cell_shading_hex(reader, table_index=0, row=0, col=0)
    if bg == "#1F4E79":
        messages.append(f"✅ 首行单元格背景色 = {bg} (update_table_cell)")
    else:
        messages.append(f"❌ 首行单元格背景色 = {bg!r}, 预期 #1F4E79")
        success = False

    table = reader.doc.tables[0]
    # 第 2 行 (rowIndex=1) 文本应包含 "甲方" / "ACME Corp"
    try:
        row1_texts = [c.text.strip() for c in table.rows[1].cells]
        if "甲方" in row1_texts[0] and "ACME" in row1_texts[1]:
            messages.append(f"✅ 第 2 行文本 = {row1_texts[:2]} (update_table_row_column)")
        else:
            messages.append(f"❌ 第 2 行文本 = {row1_texts[:2]}，预期含 '甲方' + 'ACME'")
            success = False
    except IndexError:
        messages.append("❌ 表格行数不足，无法读取第 2 行")
        success = False

    widths = _table_grid_widths_emu(reader, table_index=0)
    if len(widths) == 4:
        messages.append(f"✅ tblGrid 列数 = 4，宽度 (dxa) = {widths} (update_table_format)")
    else:
        messages.append(f"❌ tblGrid 列数 = {len(widths)}，预期 4")
        success = False

    return success, messages


# ============================================================================
# Pipelines
# ============================================================================


# OASP 0.4.0: Word 单元格字体收敛到 font（WordFont，fontColor→color）；对齐/背景色仍在顶层
HEADER_FORMAT: dict[str, Any] = {
    "horizontalAlignment": "Centered",
    "verticalAlignment": "Center",
    "backgroundColor": "#1F4E79",
    "font": {"color": "#FFFFFF", "bold": True},
}

LABEL_FORMAT: dict[str, Any] = {
    "backgroundColor": "#EEEEEE",
    "font": {"bold": True},
}


async def run_table_pipeline(workspace: Any, document_uri: str) -> tuple[bool, list[str]]:
    """4-tool 流水线（已根据 Word.js API 限制重排序）：

    顺序：insert → format(columnWidths) → merge → update_table_cell(header)
        → update_table_row_column(rows) → update_table_cell(labels) → format(borders+alignment+padding)

    关键：`columnWidths` 必须在 `merge_cells` 之前应用 —— Word.js `TableColumnCollection`
    在含合并单元格的表上不可访问，是 API 级硬限制
    （见 office-editor4ai PR #31 / office-editor4ai#29 修复说明）。
    其它格式（borders / alignment / cellPadding / styleOptions）在合并表上仍可用。
    """
    log: list[str] = []

    # 1. insert_table 5×4
    ok, _, err = await insert_table(
        workspace,
        document_uri,
        rows=5,
        columns=4,
        insert_location="End",
    )
    if not ok:
        log.append(f"insert_table 失败: {err}")
        return False, log
    log.append("insert_table 5×4 OK")

    # 2. update_table_format(columnWidths) —— 必须在 merge 之前
    ok, _, err = await update_table_format(
        workspace,
        document_uri,
        table_id="table-0",
        column_widths=[120, 80, 80, 80],
    )
    if not ok:
        log.append(f"update_table_format(columnWidths) 失败: {err}")
        return False, log
    log.append("update_table_format(columnWidths) OK (在 merge 之前)")

    # 3. merge_cells: 首行 4 列 → 1 个合并单元格
    ok, _, err = await merge_cells(
        workspace,
        document_uri,
        table_id="table-0",
        start_row_index=0,
        start_column_index=0,
        end_row_index=0,
        end_column_index=3,
    )
    if not ok:
        log.append(f"merge_cells 失败: {err}")
        return False, log
    log.append("merge_cells 首行 OK")

    # 4. update_table_cell: 设置首行表头样式 + 文本
    ok, _, err = await update_table_cell(
        workspace,
        document_uri,
        table_id="table-0",
        cells=[
            {
                "rowIndex": 0,
                "columnIndex": 0,
                "text": "甲方信息",
                "format": HEADER_FORMAT,
            }
        ],
    )
    if not ok:
        log.append(f"update_table_cell 表头 失败: {err}")
        return False, log
    log.append("update_table_cell 表头 OK")

    # 5. update_table_row_column: 4 行标签/值
    ok, _, err = await update_table_row_column(
        workspace,
        document_uri,
        table_id="table-0",
        rows=[
            {"rowIndex": 1, "values": ["甲方", "ACME Corp", "", ""]},
            {"rowIndex": 2, "values": ["地址", "上海市浦东新区", "", ""]},
            {"rowIndex": 3, "values": ["联系人", "张三", "", ""]},
            {"rowIndex": 4, "values": ["日期", "2026-04-30", "", ""]},
        ],
    )
    if not ok:
        log.append(f"update_table_row_column 失败: {err}")
        return False, log
    log.append("update_table_row_column 4 rows OK")

    # 6. 给标签列加灰底加粗
    ok, _, err = await update_table_cell(
        workspace,
        document_uri,
        table_id="table-0",
        cells=[
            {"rowIndex": 1, "columnIndex": 0, "format": LABEL_FORMAT},
            {"rowIndex": 2, "columnIndex": 0, "format": LABEL_FORMAT},
            {"rowIndex": 3, "columnIndex": 0, "format": LABEL_FORMAT},
            {"rowIndex": 4, "columnIndex": 0, "format": LABEL_FORMAT},
        ],
    )
    if not ok:
        log.append(f"update_table_cell 标签列 失败: {err}")
        return False, log
    log.append("update_table_cell 标签列 OK")

    # 7. update_table_format: 边框 + 内边距 + 整表居中（不含 columnWidths，已在第 2 步设置）
    ok, _, err = await update_table_format(
        workspace,
        document_uri,
        table_id="table-0",
        style_options={
            "cellPadding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
        },
        border_options={
            "location": "all",
            "style": "Single",
            "width": 0.5,
        },
        alignment="Centered",
    )
    if not ok:
        log.append(f"update_table_format(borders/alignment/padding) 失败: {err}")
        return False, log
    log.append("update_table_format(borders+alignment+padding) OK")

    return True, log


async def run_error_code_scenarios(workspace: Any, document_uri: str) -> tuple[bool, list[str]]:
    """B.2 + B.3 错误码场景自动触发。

    需要先跑过 `run_table_pipeline`，文档当前状态：5×4 表，首行已合并。

    自动场景：
    - B.2: 不存在的 tableId → 3010 ELEMENT_NOT_FOUND
    - B.3a: 相交合并（真·冲突）→ 失败，应 3014（Add-In 接线前暂容忍 3000，editor4ai#88）
    - B.3b: 混合列宽表不相交合并 → 成功（office-editor4ai#85 修复回归：
      mergeCells 不再访问 table.columns，合法矩形合并直接成功）

    手工场景（仍需在 Word 中点击表格外的段落）：
    - B.1: 缺省 tableId + 光标不在表格内 → 3013 NO_TABLE_AT_CURSOR
    """
    log: list[str] = []
    all_ok = True

    def note(line: str) -> None:
        """实时打印场景标题/判定，让负例的原始日志始终紧跟其场景说明出现。

        符号约定（✅/❌/⚠️ 表达用例判定，不表达正负向）：
        ✅ 用例通过（负例按预期失败且错误码正确也是 ✅）；❌ 用例未通过；
        ⚠️ 部分通过、有异常但可容忍（须显式挂账跟踪 issue）。
        """
        print(f"  {line}")
        log.append(line)

    # B.2: tableId not found → 3010
    note("--- B.2 不存在的 tableId → 3010 ELEMENT_NOT_FOUND（负例） ---")
    ok, _, err = await update_table_cell(
        workspace,
        document_uri,
        table_id="table-99",
        cells=[{"rowIndex": 0, "columnIndex": 0, "text": "x"}],
        wait_seconds=1,
        expect_failure=True,
    )
    if ok:
        note("  ❌ B.2 预期失败但成功了")
        all_ok = False
    elif err and "3010" in err:
        note(f"  ✅ B.2 负例通过，错误码 3010: {err[:80]}")
    else:
        note(f"  ❌ B.2 失败但错误码不是 3010: {err}")
        all_ok = False

    # B.3a: 真·合并冲突（相交合并）→ 失败
    # 当前状态：首行 (0,0)-(0,3) 已合并；再发相交合并 (0,0)-(1,2)，Word 语义拒绝
    # （真机实收中文报文「此方法或属性无效，因为 尚未选定要合并的多个单元格」，
    #   officeCode=GeneralException, errorLocation=TableCell.merge）。
    # OASP 归宿是 3014 ALREADY_MERGED；Add-In 中文/信号接线由 editor4ai#88 跟踪，
    # 接线后删除对 3000 的容忍、收紧为仅 3014。
    note("--- B.3a 相交合并（真·冲突，负例）→ 3014（暂容忍 3000，editor4ai#88） ---")
    ok, _, err = await merge_cells(
        workspace,
        document_uri,
        table_id="table-0",
        start_row_index=0,
        start_column_index=0,
        end_row_index=1,
        end_column_index=2,
        wait_seconds=1,
        expect_failure=True,
    )
    if ok:
        note("  ❌ B.3a 预期失败但成功了")
        all_ok = False
    elif err and "3014" in err:
        note(f"  ✅ B.3a 负例通过，错误码 3014: {err[:80]}")
    elif err and "3000" in err:
        note(f"  ⚠️  B.3a 实收 3000（已知缺口 editor4ai#88 接线后应 3014）: {err[:80]}")
    else:
        note(f"  ❌ B.3a 失败但错误码非 3014/3000: {err}")
        all_ok = False

    # B.3b: 混合列宽表不相交合并 → 成功（editor4ai#85 回归）
    # 第 2 行 (1,0)-(1,2)，与已合并首行无交集。editor4ai#85 修复前：mergeCells 在执行前
    # 访问 table.columns（混合列宽表上 Word.js 禁止）抛裸 3000；修复后不再触碰列集合，
    # startCell.merge(endCell) 对合法矩形应直接成功。
    note("--- B.3b 混合列宽表不相交合并 (1,0)-(1,2) → 成功（editor4ai#85 回归） ---")
    ok, _, err = await merge_cells(
        workspace,
        document_uri,
        table_id="table-0",
        start_row_index=1,
        start_column_index=0,
        end_row_index=1,
        end_column_index=2,
        wait_seconds=1,
    )
    if ok:
        note("  ✅ B.3b 混合列宽表 merge 成功（.merge() 路径可用）")
    else:
        note(f"  ❌ B.3b 预期成功但失败: {err}")
        all_ok = False

    note("--- B.1 (3013 NO_TABLE_AT_CURSOR) 需在 Word 中手工触发 ---")
    note("  手工步骤：")
    note("    1) 在 Word 中点击表格之外的某个段落，让光标离开表格")
    note("    2) 重跑：uv run python manual_tests/word/test_word_table_e2e.py --mode b1")
    note("    3) 预期：错误码包含 3013")

    return all_ok, log


async def test_b1_cursor_outside_table() -> bool:
    """B.1 单独场景：缺省 tableId + 光标不在表格内 → 3013 NO_TABLE_AT_CURSOR.

    前置：用户已在 Word 中打开文档（含至少 1 张表）+ 光标点在表格之外的段落。
    """
    print("\n" + "=" * 70)
    print("🧪 B.1 — 缺省 tableId + 光标不在表格内 → 3013")
    print("=" * 70)
    print("\n📌 前置：请确保 Word 文档已开 + 光标点在表格之外的段落。\n")

    from manual_tests.word.test_helpers import ready_workspace

    try:
        async with ready_workspace(host="127.0.0.1", port=3000) as (workspace, doc_uri):
            print(f"📂 当前文档: {doc_uri}\n")
            ok, _, err = await merge_cells(
                workspace,
                doc_uri,
                start_row_index=0,
                start_column_index=0,
                end_row_index=0,
                end_column_index=2,
                wait_seconds=1,
                expect_failure=True,
            )
            if ok:
                print("\n❌ B.1 预期失败但成功了 — 光标可能仍在表格内")
                return False
            if err and "3013" in err:
                print(f"\n✅ B.1 负例通过，错误码 3013: {err[:120]}")
                return True
            print(f"\n❌ B.1 失败但错误码不是 3013: {err}")
            return False
    except Exception as exc:
        print(f"\n❌ B.1 测试中断: {exc}")
        return False


async def run_omitted_table_id_smoke(workspace: Any, document_uri: str) -> tuple[bool, list[str]]:
    """缺省 tableId 路径冒烟：依赖光标当前位于上一步插入的表格内。

    Add-In 在 PR #27 实装后，缺省 tableId 时应解析光标所在表格。本步骤
    在已经修改完表格的情况下再发一次合并请求验证缺省路径不报错。
    """
    log: list[str] = []
    print("\n📌 提示：请确保光标停留在表格内（脚本不操控光标）。\n")
    await asyncio.sleep(2)

    # 让 Add-In 自行决定要操作的表 — 这里我们用一个安全的 no-op-ish 操作：
    # 在已合并的首行单独发一次 update_table_cell（不带 table_id），
    # 仅用于验证 wire payload 中 tableId 缺失，且 Add-In 能正常 ack。
    ok, data, err = await update_table_cell(
        workspace,
        document_uri,
        cells=[
            {
                "rowIndex": 0,
                "columnIndex": 0,
                "format": {"font": {"bold": True}},  # OASP 0.4.0: 嵌套 font（WordFont）
            }
        ],
    )
    if ok:
        log.append(f"omitted-tableId update_table_cell OK, data={data}")
    else:
        log.append(f"omitted-tableId update_table_cell 失败: {err}")
        # 这一步可能因为光标不在表内而报 3013，视为提示而非硬失败
    return ok, log


# ============================================================================
# Top-level scenarios
# ============================================================================


async def test_workspace_health() -> bool:
    """快速健康检查：验证 Workspace 启动 / 工具注册。"""
    from office4ai.environment.workspace.office_workspace import OfficeWorkspace

    print("\n" + "=" * 70)
    print("🏥 Workspace Health Check (4 new table tools)")
    print("=" * 70)

    workspace = OfficeWorkspace(host="127.0.0.1", port=3000)
    try:
        await workspace.start()
        print("✅ Workspace 运行正常")
        # 列出 Workspace 自身知道的事件，确认 4 个新事件已注册到 DTO 注册表
        from office4ai.environment.workspace.dtos.common import request_registry

        events = request_registry.all_events()
        new_events = [
            "word:merge:cells",
            "word:update:tableCell",
            "word:update:tableRowColumn",
            "word:update:tableFormat",
        ]
        missing = [e for e in new_events if e not in events]
        if missing:
            print(f"❌ 以下事件未注册: {missing}")
            return False
        print(f"✅ 4 个新表格事件均已注册到 request_registry: {new_events}")
        return True
    except Exception as exc:
        print(f"❌ 健康检查失败: {exc}")
        return False
    finally:
        await workspace.stop()


async def test_word_table_e2e() -> bool:
    """端到端表格流水线：插表 → 合并 → 表头 → 行写入 → 整表格式 → OOXML 验证。"""
    print("\n" + "=" * 70)
    print("🧪 Word Table E2E Test (OASP /word Draft, v0.2.0)")
    print("=" * 70)

    ensure_empty_fixture()

    runner = E2ETestRunner(
        fixtures_dir=FIXTURES_DIR,
        host="127.0.0.1",
        port=3000,
        connection_timeout=30.0,
        auto_open=True,
        auto_close=False,  # 让用户视觉验证后再手动关闭
        auto_activate=True,
        cleanup_on_success=False,  # 保留工作副本，便于 Word 中目测
    )

    try:
        async with runner.run_with_workspace("empty.docx", open_delay=3.0) as (workspace, fixture):
            print(f"\n📂 测试工作副本: {fixture.working_path}")

            # Phase 1: 4-tool pipeline
            ok, pipe_log = await run_table_pipeline(workspace, fixture.document_uri)
            for line in pipe_log:
                print(f"  • {line}")
            if not ok:
                print("\n❌ Pipeline 中断；保留工作副本供调试。")
                return False

            # Phase 2: omitted tableId smoke (软性，不计入硬失败 — 取决于光标位置)
            _, omit_log = await run_omitted_table_id_smoke(workspace, fixture.document_uri)
            for line in omit_log:
                print(f"  • {line}")

            # Phase 3: OOXML verification
            print("\n📊 OOXML 结构验证 (需要 Word 已保存到磁盘)...")
            await asyncio.sleep(1.5)
            reader = DocumentReader(fixture.working_path)
            reader.reload()  # 触发 AppleScript save
            verified, messages = verify_table_state_after_pipeline(reader)
            for msg in messages:
                print(f"  {msg}")

            # Phase 4: B.2 错误码 + B.3a 冲突负例 + B.3b merge 回归（B.1 因依赖光标位置另起 --mode b1）
            print("\n🚨 B.2 错误码 + B.3a 冲突负例 + B.3b 混合列宽 merge 回归（自动触发）...")
            err_ok, _ = await run_error_code_scenarios(workspace, fixture.document_uri)

            all_ok = verified and err_ok
            print("\n" + "=" * 70)
            if all_ok:
                print("✅ E2E 表格流水线 + B.2 错误码 + B.3a/B.3b merge 场景全部通过；请打开下面文件目测视觉效果：")
            else:
                print("❌ 部分检查未通过；保留工作副本供调试：")
            print(f"   {fixture.working_path}")
            print("=" * 70)
            return all_ok
    except Exception as exc:
        print(f"\n❌ 测试中断: {exc}")
        import traceback

        traceback.print_exc()
        return False


# ============================================================================
# CLI entry
# ============================================================================


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Word Table E2E Tests (OASP /word Draft, v0.2.0)")
    parser.add_argument(
        "--mode",
        choices=["health", "tables", "b1"],
        default="health",
        help=(
            "health: 快速校验 4 个事件已注册；"
            "tables: 完整端到端表格流水线 (含 B.2 错误码 + B.3a 冲突负例 + B.3b merge 回归，需要 Add-In + Word)；"
            "b1: B.1 单独场景 (光标不在表格内 → 3013，需先在 Word 中点击表格外段落)"
        ),
    )
    args = parser.parse_args()

    try:
        if args.mode == "health":
            success = asyncio.run(test_workspace_health())
        elif args.mode == "b1":
            success = asyncio.run(test_b1_cursor_outside_table())
        else:
            success = asyncio.run(test_word_table_e2e())
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n⏸️  测试被用户中断")
        sys.exit(130)


if __name__ == "__main__":
    main()
