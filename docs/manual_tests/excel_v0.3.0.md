# Excel 37 个 MCP 事件手工验收清单 (Issue #26 / OASP /excel Draft 0.3.0)

> **范围**：`/excel` 命名空间全部 **37 个事件 / 10 类**（#18 Foundation · #19 Range · #20 Format · #21 Worksheet · #22 Table · #23 Chart · #24 PivotTable · #25 Find&Filter）
> **协议引用**：https://doc.turingfocus.cn/oasp/0.3.0/specification/events-excel/
> **关联**：JIAQIA/office4ai#26 · #27 (Tracking) · office-editor4ai（Excel Add-In）

本清单镜像 `word_table_v0.2.0.md` / `ppt_chart_v0.3.0.md`，划分为：

- **A. 自动化部分**：跑 `test_excel_e2e.py`，脚本驱动 37 个事件 + openpyxl 双重验证；你目测视觉
- **B. 错误码部分**：5001/5002/5006/5007/5008 随 `--mode full` 自动触发校验
- **C. 需人工/视觉验证 + 跨仓注意**

> ℹ️ 所有 37 个 Excel 事件均为 **Server → Add-In 请求-响应**（Office.js），不存在 Server 侧
> 文件直写路径（与 PPT chart 的 Server-OOXML 不同）。因此本验收的「写后状态」由 Add-In
> 在 Excel 内存模型生效，openpyxl 双重验证读取的是 **Excel 存盘后**的 .xlsx。

---

## 前置准备

```bash
# 1. office4ai 测试套件通过
uv run poe check                 # ✅ ruff + mypy
uv run poe test                  # ✅ 全量（含 contract_tests/excel/ 60 项）

# 2. 证书已安装到系统信任库（仅首次）
uv run office4ai-mcp setup

# 3. 启动 office-editor4ai 的 Excel Add-In（另一个终端）
cd <office-editor4ai>/excel-editor4ai
pnpm start                       # 或 npm start
```

确认：
- [ ] office4ai 单元 + 契约测试全部通过（`poe test`）
- [ ] office-editor4ai 的 Excel Add-In 已实装 37 事件 Handler + Zod 校验 + 错误码 5001–5010
- [ ] Add-In dev server 在跑
- [ ] macOS + Excel 桌面版（Microsoft 365 / Office 2021+）

---

## A. 自动化场景（跑脚本 + 你目测）

### A.0 端到端连通性（37 事件注册校验，无需 Add-In）

```bash
uv run python manual_tests/excel/test_excel_e2e.py --mode health
```

**预期**：
```
✅ Workspace 运行正常
✅ 全部 37 个 excel:* 事件均已注册到 request_registry
```

- [ ] **A.0** Workspace 启动成功，37 个事件注册到 `request_registry`

---

### A.1 完整 37 事件真机工作流（需 Add-In + Excel）

```bash
uv run python manual_tests/excel/test_excel_e2e.py --mode full
```

脚本将：

1. 启动 Workspace + 尝试自动激活 Excel Add-In（点不到「加载项」按钮时提示手动激活）
2. 复制 `manual_tests/excel/fixtures/empty.xlsx` 为工作副本并打开
3. 把 37 个事件编排成一条「销售报表」工作流**逐个跑通**（continue-on-error，单步失败仅记录）：
   - **#18 Foundation**：`get:workbookInfo` / `get:worksheetInfo` / `get:selectedRange`
   - **#19 Range**：`set:range`（写表头+数据）/ `get:range` / `set:formula`(D1=SUM) / `copy:range` / `insert:range` / `delete:range` / `clear:range`
   - **#20 Format**：`get:rangeFormat` / `set:rangeFormat`(表头蓝底白字) / `add:conditionalFormat`(>25 高亮) / `clear:conditionalFormat` / `merge:cells`(E1:F1) / `unmerge:cells`
   - **#21 Worksheet**：`get:worksheets` / `add:worksheet`(Report) / `rename:worksheet`(→Summary) / `activate:worksheet` / `delete:worksheet`
   - **#22 Table**：`insert:table`(A1:C4) / `get:tables` / `get:table` / `add:tableRow`(South) / `sort:table`(Q1↓) / `delete:tableRow`
   - **#23 Chart**：`insert:chart`(柱形图) / `get:charts` / `update:chart`(改标题) / `delete:chart`
   - **#24 PivotTable**：`insert:pivotTable`(H1) / `get:pivotTables` / `delete:pivotTable`
   - **#25 Find&Filter**：`find:values`(East) / `set:autoFilter`(Region) / `clear:autoFilter`
4. 触发 5001/5002/5006/5007/5008 错误码场景（见 B）
5. 触发 Excel 保存，用 openpyxl 读盘双重验证稳定不变量
6. 打印「事件执行汇总 N/37 ✅」+ openpyxl 验证结果

**预期脚本输出尾部**：
```
======================================================================
事件执行汇总  37/37 ✅
======================================================================
  ✅ excel:get:workbookInfo  — 读工作簿信息
  ...（37 行）
  ✅ Sheet1 存在
  ✅ 临时表 Summary 已删除
  ✅ Sheet1!A1 = 'Region'
  ✅ E1:F1 合并已取消
======================================================================
✅ 全部 37 事件 + 错误码 + openpyxl 验证通过。请用 Excel 打开做视觉验收：
   manual_tests/excel/.test_working/empty_<timestamp>.xlsx
```

**目测验收项**（打开脚本输出的 `.xlsx`）：

- [ ] **A.1 Range**：`A1:C1` = Region/Q1/Q2 表头，`A2:C4` 三行数据；`D1` 显示求和结果（=SUM(B2:C4)）
- [ ] **A.2 Format**：表头行 `A1:C1` = 蓝底 (`#1F4E79`) + 白字 + 加粗；`E1:F1` 最终**未**合并
- [ ] **A.3 Worksheet**：底部仅 `Sheet1`（Report→Summary 已被删除，无残留临时表）
- [ ] **A.4 Table**：`A1:C4` 区域为结构化表格（带筛选下拉/条纹样式），含追加的 `South` 行，按 Q1 降序
- [ ] **A.5 Chart**：工作表上有一张柱形图（脚本最后 `delete:chart` 会删掉；若想留图请注释该步）
- [ ] **A.6 PivotTable**：`insert:pivotTable` 落在 H1（脚本最后 `delete:pivotTable` 会删掉）
- [ ] **A.7 Find&Filter**：`find:values('East')` 命中 A2；`set:autoFilter` 应用后 `clear:autoFilter` 清除

> 💡 想保留图表/透视表做视觉验收，可临时注释 `run_full_workflow` 里的 `delete:chart` /
> `delete:pivotTable` 两步后重跑。

---

## B. 错误码场景（随 `--mode full` 自动触发）

脚本在工作流之后自动跑错误码场景，每步**预期失败**并校验返回码：

```
错误码场景（预期失败 + 校验码）
  ✅ [5001] excel:get:worksheetInfo → 5001 ...（GhostSheet 不存在）
  ✅ [5002] excel:get:range → 5002 ...（无效区域地址）
  ✅ [5006] excel:get:table → 5006 ...（GhostTable 不存在）
  ✅ [5007] excel:delete:chart → 5007 ...（GhostChart 不存在）
  ✅ [5008] excel:delete:pivotTable → 5008 ...（GhostPivot 不存在）
```

- [ ] **B.1** 5001 WORKSHEET_NOT_FOUND
- [ ] **B.2** 5002 RANGE_INVALID
- [ ] **B.3** 5006 TABLE_NOT_FOUND
- [ ] **B.4** 5007 CHART_NOT_FOUND
- [ ] **B.5** 5008 PIVOT_NOT_FOUND

> 其余错误码（5003 MERGE_CONFLICT / 5004 PROTECTED_SHEET / 5005 FORMULA_ERROR /
> 5009 DATA_TYPE_MISMATCH / 5010 NOT_SUPPORTED）已在 `tests/contract_tests/excel/` +
> 单测错误码全表矩阵中覆盖；真机触发需特定文档状态（受保护表/平台不支持等），可按需手工补验。

---

## C. 需人工/视觉验证 + 跨仓注意

- [ ] **C.1 视觉保真**：表头配色、表格样式、图表类型、透视表布局在 Excel 中渲染正确（脚本只验结构，不验渲染）
- [ ] **C.2 跨平台**：在 **Excel for Web** 与 **Windows Excel** 各重跑一次 `--mode full`（Office.js 平台差异）
- [ ] **C.3 AI 业务闭环**：让接入 86 个 MCP 工具的真实 LLM Agent 执行「读销售数据 → 建表 → 加汇总公式 → 配色表头 → 插柱形图 → 建透视表」，验证组合可用性（含「写后再读确认」模式，见 `docs/discussions/excel-minimal-return-decision.md`）
- [ ] **C.4 最小返回**：确认写操作返回最小锚点（如 `set:range` 只回 `{address}`），Agent 需写后再调 `get:range` 等读工具确认——这是 #26 已决策的 **Plan A**（书面结论见上方决策文档）

---

## D. per-feature E2E 套件（#28 子问题逐个落地）

在 A/B「一条 smoke 工作流」之外，#28 把每类做深为独立 `*_e2e/` 目录（README + 多场景脚本
+ per-feature 夹具 + 编号用例 `--test N/all`），完整度对齐 word/ppt。以下按子问题勾验。

### D.1 Read/Foundation — `read_state_e2e/`（#29）

目录：`manual_tests/excel/read_state_e2e/`（`test_workbook_info.py` 4 例 /
`test_worksheet_info.py` 6 例含 5001 / `test_selected_range.py` 4 例需手动选区）。

```bash
uv run python manual_tests/excel/read_state_e2e/test_workbook_info.py --test all
uv run python manual_tests/excel/read_state_e2e/test_worksheet_info.py --test all
uv run python manual_tests/excel/read_state_e2e/test_selected_range.py --test all --no-auto-open
```

- [ ] **D.1.1 workbookInfo**：多 sheet 列表完整；隐藏表 `isHidden=true`（openpyxl `sheet_state=hidden` 核对）；`activeSheet` 与 `isActive` 一致；`fileName` 为 `.xlsx`
- [ ] **D.1.2 worksheetInfo**：默认活动表 / 指定 `worksheetName`；`usedRange` 4×3 与 openpyxl `max_row/max_column` 一致；空表 usedRange 极小；`tableCount/chartCount` 字段就绪
- [x] **D.1.3 selectedRange**：单格 1×1；A1:C2 的 2D values（2 行 3 列，AppleScript 自动选区）；空选区 F10；A1:C1 混合类型——**真机实测 falsy `0`/`False`/`''` 已落 wire**：`[['Hello', 42, True], [0, False, '']]`
- [x] **D.1.4 错误码 3000**：`get:worksheetInfo` 传 ghost 表名 → **3000 DOCUMENT_ERROR**（真机实测；旧 DoD 写的 5001 已过时，见下方 ⚠️）
- [ ] **D.1.5 视觉**：底部三个标签页（Sheet1/Data/Report）；隐藏表夹具仅见 Visible 标签

> ✅ **D.1 真机实测（2026-06-18）**：14/14 全过（workbook_info 4/4 · worksheet_info 6/6 · selectedRange 4/4），openpyxl 双重验证通过。
>
> ⚠️ **错误码现实修正（影响全 8 子问题 + B 节）**：Issue DoD / 旧注释里的 `5001–5010` Excel 错误码**实现里不存在**。Add-In 按 OASP 0.3.0 用 `3xxx`/`4xxx`（#26 已删 5xxx）。真机映射：5001→**3000** DOCUMENT_ERROR、5002→3009 RANGE_INVALID、5006→3010/3013、5007→3015、4002/4004 不变。**B 节与 `test_excel_e2e.py` foundation 冒烟里的 5xxx 断言需回头按真实码修订**。

---

## 验收完成后

1. 全部 A/B 项打勾 → 把勾选状态复制到 Issue #26 评论
2. 逐片确认 #18–#25 写工具真机 round-trip（建议从栈底 #18 往上逐片验，#18 返工会牵动其上 rebase）
3. 开 PR 链把 9 深栈（#18→…→#26）并入 main
4. PR 合并后关闭 Issue #26 + 触发 OASP `events-excel.md` 37 事件 **📋 Draft → ✅ Stable**（#27 协议侧 checklist）

---

**最后更新**：2026-06-18
**维护者**：JQQ &lt;jqq1716@gmail.com&gt;
