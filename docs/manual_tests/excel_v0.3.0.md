# Excel 37 个 MCP 事件手工验收清单 (Issue #26 / OASP /excel Draft 0.3.0)

> **范围**：`/excel` 命名空间全部 **37 个事件 / 10 类**（#18 Foundation · #19 Range · #20 Format · #21 Worksheet · #22 Table · #23 Chart · #24 PivotTable · #25 Find&Filter）
> **协议引用**：https://doc.turingfocus.cn/oasp/0.3.0/specification/events-excel/
> **关联**：JIAQIA/office4ai#26 · #27 (Tracking) · office-editor4ai（Excel Add-In）

本清单镜像 `word_table_v0.2.0.md` / `ppt_chart_v0.3.0.md`，划分为：

- **A. 自动化部分**：跑 `test_excel_e2e.py`，脚本驱动 37 个事件 + openpyxl 双重验证；你目测视觉
- **B. 错误码部分**：权威 3xxx（3010+kind / 3009，oasp#17 定案）随 `--mode full` 自动触发校验（Add-In 接线前 XFAIL）
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
- [ ] office-editor4ai 的 Excel Add-In 已实装 37 事件 Handler + Zod 校验（错误码权威 3xxx 接线见 office-editor4ai#80，未接线前 B 部分为 XFAIL）
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
4. 触发权威错误码场景 3010(kind:worksheet/table/chart/pivotTable)/3009（见 B）
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

脚本在工作流之后自动跑错误码场景，每步**预期失败**并校验权威码 + `details.kind`
（oasp#17 定案；Add-In 接线（office-editor4ai#80）前实收 `3000` 兜底 → ⚠️ XFAIL 计通过，
接线后 ✅ 严格命中）：

```
错误码场景（预期失败 + 校验权威码；Add-In 接线前 XFAIL）
  ✅ [3010 details⊇{'kind': 'worksheet'}] excel:get:worksheetInfo → ...（GhostSheet 不存在）
  ✅ [3009] excel:get:range → ...（无效区域地址）
  ✅ [3010 details⊇{'kind': 'table'}] excel:get:table → ...（GhostTable 不存在）
  ✅ [3010 details⊇{'kind': 'chart'}] excel:delete:chart → ...（GhostChart 不存在）
  ✅ [3010 details⊇{'kind': 'pivotTable'}] excel:delete:pivotTable → ...（GhostPivot 不存在）
```

- [ ] **B.1** 3010 ELEMENT_NOT_FOUND（`kind:"worksheet"`，旧 5001 退役）
- [ ] **B.2** 3009 RANGE_INVALID（旧 5002 退役）
- [ ] **B.3** 3010 ELEMENT_NOT_FOUND（`kind:"table"`，旧 5006 退役）
- [ ] **B.4** 3010 ELEMENT_NOT_FOUND（`kind:"chart"`，旧 5007 退役）
- [ ] **B.5** 3010 ELEMENT_NOT_FOUND（`kind:"pivotTable"`，旧 5008 退役）

> 其余权威码（3014 ALREADY_MERGED / 3003 DOCUMENT_READ_ONLY / 3017 FORMULA_ERROR /
> 3018 DATA_TYPE_MISMATCH / 3016 API_NOT_SUPPORTED，对应退役的 5003/5004/5005/5009/5010）
> 中，3017 已在 `range_e2e/test_set_formula.py` 建用例；真机触发需特定文档状态
> （受保护表/平台不支持等）的按需手工补验；3018 挂账待 office-editor4ai#80 定义触发条件。

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
`test_worksheet_info.py` 6 例含 3010(kind:worksheet) / `test_selected_range.py` 4 例需手动选区）。

```bash
uv run python manual_tests/excel/read_state_e2e/test_workbook_info.py --test all
uv run python manual_tests/excel/read_state_e2e/test_worksheet_info.py --test all
uv run python manual_tests/excel/read_state_e2e/test_selected_range.py --test all --no-auto-open
```

- [ ] **D.1.1 workbookInfo**：多 sheet 列表完整；隐藏表 `isHidden=true`（openpyxl `sheet_state=hidden` 核对）；`activeSheet` 与 `isActive` 一致；`fileName` 为 `.xlsx`
- [ ] **D.1.2 worksheetInfo**：默认活动表 / 指定 `worksheetName`；`usedRange` 4×3 与 openpyxl `max_row/max_column` 一致；空表 usedRange 极小；`tableCount/chartCount` 字段就绪
- [x] **D.1.3 selectedRange**：单格 1×1；A1:C2 的 2D values（2 行 3 列，AppleScript 自动选区）；空选区 F10；A1:C1 混合类型——**真机实测 falsy `0`/`False`/`''` 已落 wire**：`[['Hello', 42, True], [0, False, '']]`
- [x] **D.1.4 错误码 3010**：`get:worksheetInfo` 传 ghost 表名 → **3010 ELEMENT_NOT_FOUND**（`details.kind:"worksheet"`，oasp#17 定案；当时真机实收 3000，Add-In 接线（office-editor4ai#80）前用例为 XFAIL）
- [ ] **D.1.5 视觉**：底部三个标签页（Sheet1/Data/Report）；隐藏表夹具仅见 Visible 标签

> ✅ **D.1 真机实测（2026-06-18）**：14/14 全过（workbook_info 4/4 · worksheet_info 6/6 · selectedRange 4/4），openpyxl 双重验证通过。
>
> ✅ **错误码已定案（oasp#17，issue #82 校准，影响全 8 子问题 + B 节）**：早期 `5001–5010` 草案块整体退役，权威映射见 `excel_e2e_dev_plan.md`「错误码」节（not-found→3010+`details.kind`、非法 address→3009、公式错→3017 等，规范层 MUST 不得降级 3000）。全部用例已收紧为权威码断言 + `xfail_reason`；Add-In 接线（office-editor4ai#80）前真机实收 3000 兜底 → XFAIL 计通过，接线后摘标转正。

### D.2 Range CRUD + 公式 — `range_e2e/`（#30）

目录：`manual_tests/excel/range_e2e/`（`test_set_get_range.py` 6 例含 3009 /
`test_clear_range.py` 4 例 / `test_copy_range.py` 3 例 /
`test_shift_range.py` 4 例 / `test_set_formula.py` 4 例含 3009）。覆盖 #19 Range 切片
7 事件：`set/get/clear/copy/insert/delete:range` + `set:formula`。夹具 `grid.xlsx`
（Data 数值网格 / Blank 空表 / Shift 自描述 3×3 网格）。

```bash
uv run python manual_tests/excel/range_e2e/test_set_get_range.py --test all
uv run python manual_tests/excel/range_e2e/test_clear_range.py --test all
uv run python manual_tests/excel/range_e2e/test_copy_range.py --test all
uv run python manual_tests/excel/range_e2e/test_shift_range.py --test all
uv run python manual_tests/excel/range_e2e/test_set_formula.py --test all
```

- [x] **D.2.1 set/get:range**：2D 数组往返；标量 `'Z'` 填满；**falsy `0`/`False`/`''` 落 wire 存活**（真机 `[[0, False, '']]`）；读已有 Data!A1:C4（4×3）；`includeFormat=true` 返回 RangeFormatInfo（font=宋体）
- [x] **D.2.2 clear:range**：`contents`/`all` 清空值；`formats` 保留内容；清子区域不影响相邻单元格（openpyxl 核对）
- [x] **D.2.3 copy:range**：单行/数据块/单元格复制到空白区，目标=源且源不变
- [x] **D.2.4 insert/delete:range**：`down`/`right` 让位、`up`/`left` 补位（Shift 自描述网格断言「谁落到哪」）
- [x] **D.2.5 set:formula**：`=SUM(B2:C2)` / `=B3+C3` / `=B4*2` 公式串落盘（openpyxl `data_only=False` 读公式）
- [x] **D.2.6 错误码 3009/3017**：`get:range` / `set:formula` 传非法 address → **3009 RANGE_INVALID**；`set:formula` 畸形公式 → **3017 FORMULA_ERROR**（oasp#17 定案；当时真机实收 3000，接线前 XFAIL）
- [ ] **D.2.7 视觉**：复制后目标区域可见数据；插入/删除后单元格位移正确

> ✅ **D.2 真机实测（2026-06-18）**：21/21 全过（set_get 6/6 · clear 4/4 · copy 3/3 · shift 4/4 · formula 4/4），openpyxl 双重验证通过。
>
> ✅ **「非法 address → 3009」已由 oasp#17 定案为规范层 MUST**。历史观测（`excelErrorCode()` 二值分类、3009 dead code、真机 3000）正是协议裁决要消解的实现缺口，接线归 office-editor4ai#80；接线前本清单相关用例以 XFAIL 运行。仅 schema 违规才得 `4000`。

### D.3 Format / 条件格式 / 合并 — `format_e2e/`（#31）

目录：`manual_tests/excel/format_e2e/`（`test_set_range_format.py` 6 例 /
`test_get_range_format.py` 4 例含 3009 / `test_conditional_format.py` 4 例含 3009 /
`test_merge_cells.py` 4 例含 3009）。覆盖 #20 Format 切片 6 事件：`get/set:rangeFormat` ·
`add/clear:conditionalFormat` · `merge/unmerge:cells`。夹具 `fmt.xlsx`（Data 数值网格 /
Merge 标签行）。

```bash
uv run python manual_tests/excel/format_e2e/test_set_range_format.py --test all
uv run python manual_tests/excel/format_e2e/test_get_range_format.py --test all
uv run python manual_tests/excel/format_e2e/test_conditional_format.py --test all
uv run python manual_tests/excel/format_e2e/test_merge_cells.py --test all
```

- [x] **D.3.1 set:rangeFormat**：font（bold/italic/color）· fill · numberFormat · alignment（center/top/wrap）· borders（top=thick/bottom=thin）· **偏更新**（只改传入属性，其余保留）——openpyxl 读盘核对
- [x] **D.3.2 get:rangeFormat**：RangeFormatInfo 字段完整（numberFormat 为 2D）；set→get 往返（bold + '0.00' 反映）；多格 2×3 维度
- [x] **D.3.3 add/clear:conditionalFormat**：cellValue（>150 红底）/ colorScale 添加；clear 清除（协议成功，clear 后 openpyxl CF 计数=0）
- [x] **D.3.4 merge/unmerge:cells**：合并保留左上 'M1' + B1/C1 清空；取消合并后 merged_ranges=[]（openpyxl 核对）
- [x] **D.3.5 错误码 3009**：四类写/读传非法 address → **3009 RANGE_INVALID**（oasp#17 定案；当时真机实收 3000，接线前 XFAIL）
- [ ] **D.3.6 视觉**：A1:C1 红粗斜体表头；B 列条件格式高亮；Merge 表 A1:C1 合并居中

> ✅ **D.3 真机实测（2026-06-18）**：18/18 全过（set 6/6 · get 4/4 · conditionalFormat 4/4 · merge 4/4），openpyxl 双重验证通过。
>
> ⚠️ **DTO-vs-wire 不一致（get:rangeFormat）**：office4ai `GetRangeFormatData` 声明 `{address, format: RangeFormatInfo}`，但 AddIn `rangeFormat.ts::getRangeFormat()` 实际**直接返回扁平 RangeFormatInfo**（顶层 `font/fill/horizontalAlignment/verticalAlignment/wrapText/numberFormat`，无 `address`、无 `format` 包裹）。E2E 已按真机形态断言；建议后续对齐 DTO 或 AddIn（择一）。注：本问题已由 AddIn 侧 `0be2056 fix(excel): get:rangeFormat 响应按契约包裹 (#52)` 修复并并入 main，后续可回归验证扁平→包裹的形态变化。

---

### D.4 Worksheet 管理 — `worksheet_e2e/`（#32）

目录：`manual_tests/excel/worksheet_e2e/`（`test_add_worksheet.py` 5 例含 3004 /
`test_rename_activate.py` 4 例含 2×3010 / `test_delete_worksheet.py` 3 例含 3010）。覆盖
#21 Worksheet 切片 5 事件：`get:worksheets` · `add:worksheet` · `delete:worksheet` ·
`rename:worksheet` · `activate:worksheet`。夹具 `book.xlsx`（4 张自描述表 Alpha[active]/Beta/Gamma/Delta）。

```bash
uv run python manual_tests/excel/worksheet_e2e/test_add_worksheet.py --test all
uv run python manual_tests/excel/worksheet_e2e/test_rename_activate.py --test all
uv run python manual_tests/excel/worksheet_e2e/test_delete_worksheet.py --test all
```

- [x] **D.4.1 add:worksheet**：命名（'Summary' → `{name,index}`）· 自动命名（省略 → Excel 'Sheet1'）——openpyxl `sheet_names` 核对落盘
- [x] **D.4.2 get:worksheets**：4 表字段完整（name/index/isActive/isHidden），Alpha 唯一 active；实时反映新增 'Extra'
- [x] **D.4.3 rename:worksheet**：Beta → BetaRenamed（`{name}`）；旧名消失、隔离表 Delta 保留
- [x] **D.4.4 activate:worksheet**：activate Gamma → `get:worksheets` 断言 Gamma isActive（openpyxl `wb.active` 佐证）
- [x] **D.4.5 delete:worksheet**：删 Beta → `sheet_names` 不含 Beta、其余保留；`get:worksheets` 实时反映
- [x] **D.4.6 错误码 3004/3010**：add 重名 → **3004 OPERATION_FAILED**；rename·activate·delete 不存在的表 → **3010 ELEMENT_NOT_FOUND**（`kind:"worksheet"`，oasp#17 定案；当时真机实收 3000，接线前 XFAIL）

> ✅ **D.4 真机实测（2026-06-18）**：12/12 全过（add 5/5 · rename_activate 4/4 · delete 3/3），openpyxl `sheet_names`/`wb.active` 双重验证通过。
>
> ⚠️ **DTO-vs-wire（delete/activate 返回 void）**：office4ai `DeleteWorksheetData {deleted}` / `ActivateWorksheetData {activated}` 是声明式的，AddIn handler 实际返回 **void**，真机响应 `data={}` 无该字段。E2E 一律靠 openpyxl 读盘验证，不断言响应体。

---

### D.5 Table 操作 — `table_e2e/`（#33）

目录：`manual_tests/excel/table_e2e/`（`test_insert_table.py` 4 例含 3009 /
`test_get_table.py` 4 例含 3010 / `test_table_rows.py` 6 例含 3010+4004+4000 /
`test_sort_table.py` 4 例含 3010）。覆盖 #22 Table 切片 6 事件：`insert:table` ·
`get:table` · `get:tables` · `add:tableRow` · `delete:tableRow` · `sort:table`。夹具
`tbl.xlsx`（Raw 纯数据 / Blank 空 / Sales[SalesTable A1:C5] / Roster[RosterTable A1:B4]，
后两者由 openpyxl 预置 Excel 表，Office.js 识别带 GUID id）。

```bash
uv run python manual_tests/excel/table_e2e/test_insert_table.py --test all
uv run python manual_tests/excel/table_e2e/test_get_table.py --test all
uv run python manual_tests/excel/table_e2e/test_table_rows.py --test all
uv run python manual_tests/excel/table_e2e/test_sort_table.py --test all
```

- [x] **D.5.1 insert:table**：已有数据建表（Raw!A1:C4）· 带 data 写空白区（hasHeaders=true 表头+正文）· styleName——openpyxl `table_names` 核对落盘
- [x] **D.5.2 get:table**：富 `TableInfo`（name/id/address/rowCount/columnCount/columns[{name,index}]/styleName/showHeaders）；跨表按 `worksheet_name` 取
- [x] **D.5.3 get:tables**：精简列表（每条仅 name/id/address）
- [x] **D.5.4 add:tableRow**：末尾追加 `['R3row',40]`（→`{tableId}`）；openpyxl 读正文核对
- [x] **D.5.5 delete:tableRow**：`rowIndex=0`（falsy 首行）· 中间行——返回 void，openpyxl 读「谁上移」核对
- [x] **D.5.6 sort:table**：单列升/降（Age 唯一）· 多级 Score↑→Age↑（断 tie）——openpyxl 读 Name 列核对行序
- [x] **D.5.7 错误码 3010/4004/4000**：表不存在 → **3010**（`kind:"table"`）；rowIndex 越界 → **4004**；rowIndex 负数（Zod nonnegative）→ **4000**（oasp#17 定案；当时真机前两类实收 3000，接线前 XFAIL）
- [ ] **D.5.8 视觉**：Raw/Blank 新表带表格样式；Sales 排序后行序；Roster 增删行后表范围

> ✅ **D.5 真机实测（2026-06-18）**：18/18 全过（insert 4/4 · get 4/4 · rows 6/6 · sort 4/4），openpyxl `table_names`/单元格读盘双重验证通过。
>
> ✅ **错误码已定案（oasp#17）**：表不存在→`3010`（`kind:"table"`）；rowIndex 越界→`4004`；Zod `nonnegative()` 失败→`4000`「Too small: expected >=0」。旧 `5006/5009` 退役为 `3010/3018`；接线（office-editor4ai#80）前真机实收 `3000` → XFAIL。
>
> ⚠️ **DTO-vs-wire（delete:tableRow 返回 void）**：真机响应 `data={}`，靠 openpyxl 读盘验证，不断言响应体。

---

### D.6 Chart 操作 — `chart_e2e/`（#34）

目录：`manual_tests/excel/chart_e2e/`（`test_insert_chart.py` 5 例含 4002 /
`test_get_charts.py` 4 例含 3010 / `test_update_chart.py` 4 例含 3010 /
`test_delete_chart.py` 3 例含 3010）。覆盖 #23 Chart 切片 4 事件：`insert:chart` ·
`get:charts` · `update:chart` · `delete:chart`。夹具 `chart.xlsx`（Data A1:C4 数值网格 /
Blank 空表）。图表为视觉对象，以**协议层 get:charts 回读**为主验证（openpyxl chart_count best-effort）。

```bash
uv run python manual_tests/excel/chart_e2e/test_insert_chart.py --test all
uv run python manual_tests/excel/chart_e2e/test_get_charts.py --test all
uv run python manual_tests/excel/chart_e2e/test_update_chart.py --test all
uv run python manual_tests/excel/chart_e2e/test_delete_chart.py --test all
```

- [x] **D.6.1 insert:chart**：ColumnClustered+title · Line · Pie（单列）· XYScatter+position → `{name}`（Excel 自动命名 "Chart N"）
- [x] **D.6.2 get:charts**：回读 `[{name,chartType,title,top,left,width,height}]`；单图/多图/空表空列表
- [x] **D.6.3 update:chart**：title · chartType（Column→Line）· position（回读核对生效）
- [x] **D.6.4 delete:chart**：删其一保留其它 · 删唯一→空（返回 void）
- [x] **D.6.5 错误码 4002/3010**：非法 chartType（非空非法枚举）→ **4002 INVALID_PARAM**；图表不存在 → **3010**（`kind:"chart"`）；不存在 worksheet → **3010**（`kind:"worksheet"`）（oasp#17 定案；当时真机实收 3000，接线前 XFAIL）
- [ ] **D.6.6 视觉**：Data 表各类型图表就位；update 后类型/标题/位置变化；双击为原生可编辑图表

> ✅ **D.6 真机实测（2026-06-18）**：16/16 全过（insert 5/5 · get 4/4 · update 4/4 · delete 3/3），协议 get:charts 回读双重验证通过（openpyxl chart_count 旁证读到图表）。
>
> 🔧 **harness 引入 `ExcelCase.flow`**：图表名由 Excel 自动生成（"Chart 1"），update/delete 无法用单 action 表达，故用 `flow`（`async (workspace, uri, reader) -> bool`）做「insert 拿 name → 操作 → get:charts 回读核对」多步流。复用于 #35 PivotTable。
>
> ⚠️ **DTO-vs-wire（delete:chart 返回 void）**：真机响应 `data={}`，靠 get:charts 回读验证，不断言响应体。

---

### D.7 PivotTable 操作 — `pivot_table_e2e/`（#35）

目录：`manual_tests/excel/pivot_table_e2e/`（`test_insert_pivot_table.py` 4 例含 3009+4000 /
`test_get_pivot_tables.py` 4 例含 3010 / `test_delete_pivot_table.py` 3 例含 3010）。覆盖
#24 PivotTable 切片 3 事件：`insert:pivotTable` · `get:pivotTables` · `delete:pivotTable`。
夹具 `pivot.xlsx`（Data A1:C5 数据源 Region/Product/Amount / Blank 空表）。仅用
sourceAddress + targetAddress 创建**空透视表骨架**，不构造 rows/columns/values/filters。

```bash
uv run python manual_tests/excel/pivot_table_e2e/test_insert_pivot_table.py --test all
uv run python manual_tests/excel/pivot_table_e2e/test_get_pivot_tables.py --test all
uv run python manual_tests/excel/pivot_table_e2e/test_delete_pivot_table.py --test all
```

- [x] **D.7.1 insert:pivotTable**：指定 name 'SalesPivot' / 默认 name 'PivotTable' → `{name}`（source 与 target 须同一 worksheet，target 用源表空白落点 E1/E20）
- [x] **D.7.2 get:pivotTables**：回读 `[{name, id}]`（每条仅 2 字段）；单/多透视表/空表空列表
- [x] **D.7.3 delete:pivotTable**：flow 删其一保留其它 / 删唯一→空（返回 void，get:pivotTables 回读核对）
- [x] **D.7.4 错误码 3009/3010/4000**：非法 source → **3009**；透视表不存在 → **3010**（`kind:"pivotTable"`）；不存在 worksheet → **3010**（`kind:"worksheet"`）；source 空串（Zod min(1)）→ **4000**（oasp#17 定案；当时真机前三类实收 3000，接线前 XFAIL）
- [ ] **D.7.5 视觉**：Data!E1 透视表骨架就位（空 PivotTable 占位框）

> ✅ **D.7 真机实测（2026-06-18）**：11/11 全过（insert 4/4 · get 4/4 · delete 3/3），协议 get:pivotTables 回读双重验证通过。
>
> ⚠️ **DTO-vs-wire（delete:pivotTable 返回 void）**：真机响应 `data={}`，靠 get:pivotTables 回读验证，不断言响应体。复用 #34 的 `ExcelCase.flow` 做 insert→delete→get 多步流。

---

### D.8 Find & Filter — `find_filter_e2e/`（#36）

目录：`manual_tests/excel/find_filter_e2e/`（`test_find_values.py` 5 例含 4000 /
`test_auto_filter.py` 5 例含 4000+3000）。覆盖 #25 Find&Filter 切片 3 事件：`find:values` ·
`set:autoFilter` · `clear:autoFilter`。夹具 `filt.xlsx`（Data A1:C5，Product 列大小写混合
Apple/apple 以区分 matchCase）。

```bash
uv run python manual_tests/excel/find_filter_e2e/test_find_values.py --test all
uv run python manual_tests/excel/find_filter_e2e/test_auto_filter.py --test all
```

- [x] **D.8.1 find:values**：默认（不区分大小写·子串）'apple'→B2/B3/B5 共 3；matchCase=true→仅 B3/B5；限定 address=A1:A5 'East'→A2/A4
- [x] **D.8.2 find:values 无命中**：matchEntireCell=true 'App'→`matches=[]`（整单元格匹配排除子串）
- [x] **D.8.3 set:autoFilter**：单列 Region=East / 多列 Region=East & Product∈{Apple,Banana} → `{address}` + openpyxl `auto_filter_ref='A1:C5'`
- [x] **D.8.4 clear:autoFilter**：先 set → clear → `auto_filter_ref=None`（返回 void）
- [x] **D.8.5 错误码 4000/3009**：searchText/address 空串（Zod min(1)）→ **4000**；非法 address → **3009 RANGE_INVALID**（oasp#17 定案；当时真机实收 3000，接线前 XFAIL）

> ✅ **D.8 真机实测（2026-06-18）**：10/10 全过（find 5/5 · filter 5/5），openpyxl `auto_filter_ref` 双重验证通过。
>
> ⚠️ **DTO-vs-wire（clear:autoFilter 返回 void）**：真机响应 `data={}`，靠 openpyxl `auto_filter_ref` 读盘验证，不断言响应体。

---

## 🏁 Excel per-feature E2E 全量收口（#28 子问题 #29–#36 全完成）

| 子问题 | 切片 | 用例 | 真机结果 |
|--------|------|------|---------|
| #29 | Read/Foundation | 14 | ✅ 14/14 |
| #30 | Range CRUD + 公式 | 21 | ✅ 21/21 |
| #31 | Format/条件格式/合并 | 18 | ✅ 18/18 |
| #32 | Worksheet 管理 | 12 | ✅ 12/12 |
| #33 | Table 操作 | 18 | ✅ 18/18 |
| #34 | Chart 操作 | 16 | ✅ 16/16 |
| #35 | PivotTable 操作 | 11 | ✅ 11/11 |
| #36 | Find & Filter | 10 | ✅ 10/10 |
| **合计** | **#18–#25 全 37 事件** | **120** | **✅ 120/120** |

> 错误码（oasp#17 已定案，issue #82 校准）：**not-found → `3010`+`details.kind`、非法 address →
> `3009`、非法枚举 → `4002`、越界 → `4004`、重名 → `3004`、公式错 → `3017`**（规范层 MUST，
> 不得降级 `3000`）；**Zod 校验失败（空串 min(1) / 负数 nonnegative）→ `4000`** 不变。
> 历史全片真机确认时 Add-In 只发 4000/3000 二值——该缺口由 office-editor4ai#80 接线消解，
> 接线前错误码用例以 XFAIL 运行。详见 `excel_e2e_dev_plan.md`「错误码」节权威映射表。

---

## 验收完成后

1. 全部 A/B 项打勾 → 把勾选状态复制到 Issue #26 评论
2. 逐片确认 #18–#25 写工具真机 round-trip（建议从栈底 #18 往上逐片验，#18 返工会牵动其上 rebase）
3. 开 PR 链把 9 深栈（#18→…→#26）并入 main
4. PR 合并后关闭 Issue #26 + 触发 OASP `events-excel.md` 37 事件 **📋 Draft → ✅ Stable**（#27 协议侧 checklist）

---

**最后更新**：2026-06-18
**维护者**：JQQ &lt;jqq1716@gmail.com&gt;
