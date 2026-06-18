# Excel 手动 E2E 开发规划（Tracking #28 → 子问题 #29–#36）

> 目标：把 Excel 手动 E2E 测试套件追平 word/ppt 完整度。
> 分支：`feat/excel-e2e-oasp-0.3.0`（已基于 upstream 同名分支检出）。
> 工作方式：**一次只做一个子问题** → 测试 → 提交 → push，循环推进。

---

## 一、底座（已就绪，8 个子问题共用）

| 文件 | 作用 |
|------|------|
| `manual_tests/excel/e2e_base.py` | `ExcelTestRunner`（AppleScript 激活/保存/关闭）+ `WorkbookReader`（openpyxl 双重验证）+ `create_empty_xlsx` |
| `manual_tests/excel/test_helpers.py` | `ready_workspace` + 通用 `excel_op(workspace, doc_uri, action_name, **params)` |
| `manual_tests/excel/test_excel_e2e.py` | foundation smoke（37 事件 health/full） |
| `manual_tests/excel/fixtures/empty.xlsx` | 空工作簿夹具 |
| `docs/manual_tests/excel_v0.3.0.md` | 验收清单（每个子问题需补该类视觉验收项） |

参考实现（结构镜像目标）：
- 脚本范式：`manual_tests/ppt/insert_table_e2e/test_table_insert.py`（`--test N` / `--test all` + `TEST_CASES` + validator）
- README 范式：`manual_tests/word/insert_text_e2e/README.md`（目录结构 + 编号用例表 + 运行方式 + 前置 + 预期 + 常见问题）
- `WorkbookReader` API：`cell_value` / `cell_number_format` / `cell_fill_hex` / `merged_ranges` / `table_names` / `auto_filter_ref` / `chart_count` / `sheet_names`

---

## 二、推进顺序（依赖友好，逐个认领）

| # | 子问题 | 目录 | 事件数 | 场景脚本数 | 关键错误码 |
|---|--------|------|--------|-----------|-----------|
| 1 | #29 Read/Foundation | `read_state_e2e/` | 3 | 3 | 5001 |
| 2 | #30 Range CRUD+公式 | `range_e2e/` | 7 | 5 | 5002 / 5005 / 5004 |
| 3 | #31 Format/条件格式/合并 | `format_e2e/` | 6 | 4 | 5003 / 5002 |
| 4 | #32 Worksheet 管理 | `worksheet_e2e/` | 5 | 3 | 5001 |
| 5 | #33 Table 操作 | `table_e2e/` | 6 | 4 | 5006 / 5009 / 4004 |
| 6 | #34 Chart 操作 | `chart_e2e/` | 4 | 4 | 4002 / 5007 / 5002 |
| 7 | #35 PivotTable 操作 | `pivot_table_e2e/` | 3 | 3 | 5008 / 5010 / 5002 |
| 8 | #36 Find&Filter | `find_filter_e2e/` | 3 | 2 | 4004 / 5001 / 5002 |

> 顺序理由：Read 先行（其他测试用读类 data 做断言）；Range 次之（多数功能需要先铺数据）；最后 Chart/Pivot/Filter 依赖前面铺的数据夹具。

---

## ⚠️ 错误码现实（真机实测，#29 发现，影响全部子问题）

Issue #28–#36 的 DoD 与旧 DTO 注释写的 **`5001–5010` Excel 错误码在实现里不存在**。Add-In
（`office-editor4ai/packages/shared/src/error-codes.ts`）按 **OASP 0.3.0 error-handling 表用 `3xxx`/`4xxx`**
（#26 已删除历史 5xxx 漂移）。所有子问题的错误码用例**按真实码断言**：

| 旧 DoD（过时） | 真实码 | 含义 |
|---------------|-------|------|
| 5001 WORKSHEET_NOT_FOUND | **3000** DOCUMENT_ERROR ✅#32真机确认 | worksheet/资源不存在（"请求的资源不存在"） |
| 5002 RANGE_INVALID | **3000** DOCUMENT_ERROR ⚠️ | 非法/畸形 address（**非** 3009，见下） |
| 5003 MERGE_CONFLICT | 3014 ALREADY_MERGED（待真机核实） | 合并冲突 |
| 5006 TABLE_NOT_FOUND / 5009 | **3000** DOCUMENT_ERROR ✅#33真机确认 | 表不存在/索引越界（"请求的表格不存在"；3010/3013 未发射） |
| 5007 CHART_NOT_FOUND / 无效 chartType | **3000** DOCUMENT_ERROR ✅#34真机确认 | 图表不存在/非法类型枚举（4002/3015 未发射） |
| 5008 PIVOT_NOT_FOUND / 5010 / 5002 | **3000** DOCUMENT_ERROR ✅#35真机确认 | 透视表/资源不存在（空串→4000） |
| 4002 INVALID_PARAM / 4004 PARAM_OUT_OF_RANGE | **4000** VALIDATION_ERROR ✅#33真机确认 | Zod 校验失败（含 nonnegative）统一 4000，非 4002/4004 |

> 标「待核实」的在做对应子问题时真机确认实际码。后续应回头修订 Issue DoD、
> `docs/manual_tests/excel_v0.3.0.md` B 节、`manual_tests/excel/test_excel_e2e.py` foundation 冒烟里残留的 5xxx。
>
> ⚠️ **#30 真机修正：非法 address → 3000，不是 3009**。`3009 RANGE_INVALID` 在
> `error-codes.ts` 有定义但**全仓 0 个 handler 发射**（dead code，仅 error-codes.ts + 对齐测试引用）。
> `excel-handlers.ts` 的 `excelErrorCode()` 只把 **Zod 失败 → `4000` VALIDATION_ERROR**，其余一切
> Office.js 运行期异常（含 `getRange`/`setFormula` 拒绝畸形地址）→ **`3000` OFFICE_API_ERROR**。
> 故所有「非法范围地址」类错误码用例按 **3000** 断言；仅 **schema 违规**（类型错/缺必填）才得 `4000`。
> 上表 3009/3010/3013/3015 等「待核实」码同样可能实为 3000——做到对应子问题时按此真机复核。

---

## 三、每个子问题的 DoD（统一模板，对齐 #28）

- [ ] `manual_tests/excel/<category>_e2e/` + `__init__.py`
- [ ] **README.md**：目录结构 + 每脚本逐条编号用例表 + 运行方式（`--test N` / `--test all`）+ 前置条件 + 预期结果 + 常见问题
- [ ] **多个场景脚本**（按入参维度拆分），每脚本支持 `--test N` 单跑 + `--test all`
- [ ] **per-feature 夹具**：`manual_tests/excel/fixtures/<category>_e2e/*.xlsx`（必要时预填数据）
- [ ] **openpyxl 双重验证**：写操作读盘核对（值/格式/合并/表/筛选）
- [ ] **错误码场景**：该类相关错误码
- [ ] 更新 `docs/manual_tests/excel_v0.3.0.md` 对应类视觉验收项

---

## 四、单个子问题的标准工作流（每轮重复）

每解决一个子问题，按以下 6 步走，然后才进入下一个：

1. **建目录与夹具**
   - `mkdir manual_tests/excel/<category>_e2e/` + `__init__.py`
   - 生成 per-feature 夹具到 `manual_tests/excel/fixtures/<category>_e2e/`（用 openpyxl 脚本或 `create_empty_xlsx` 派生）

2. **写场景脚本**（镜像 ppt `test_table_insert.py` 范式）
   - 每脚本：`TEST_CASES` 列表 + 每个 case 的 validator（用 `WorkbookReader` 双重验证）+ `--test N/all` argparse
   - 用 `excel_op(workspace, doc_uri, "set:range", address=..., values=...)` 驱动事件
   - **falsy 存活专项**：`0` / `False` / `rowIndex=0` 必须落格/落 wire（#30 #33 #36 重点）
   - 错误码场景：构造非法入参，断言返回 `error.code`

3. **写 README**（镜像 word `insert_text_e2e/README.md`）：编号用例表 + 运行方式 + 前置 + 预期 + 常见问题

4. **静态校验 + 冒烟**（无真机也能跑的部分）
   ```bash
   poe lint && poe format && poe typecheck    # 或 poe check
   uv run python manual_tests/excel/<category>_e2e/<script>.py --list   # 确认用例可枚举、import 无误
   ```
   > 真机用例需打开 Excel + 激活 `excel-editor4ai` Add-In；可用 `/run-office-e2e` skill 半自动跑。无真机时至少保证 `--list` 与静态检查通过。

5. **更新验收清单** `docs/manual_tests/excel_v0.3.0.md` 增加该类视觉验收项。

6. **提交 + push**
   ```bash
   git add manual_tests/excel/<category>_e2e/ manual_tests/excel/fixtures/<category>_e2e/ docs/manual_tests/excel_v0.3.0.md
   git commit -m "test(excel): <category> 手动 E2E — <N> 事件 per-feature 套件 (#28 #<sub>)"
   git push origin feat/excel-e2e-oasp-0.3.0   # 推到自己的 fork（origin=llg0363）
   ```
   > 提交信息沿用仓库前缀风格 `test(excel): ...`，并带上 `#28` 与子问题号。

---

## 五、各子问题脚本拆分速查（来自 issue 建议）

**#29 read_state_e2e/**（验证以返回 data 为主）
- `test_workbook_info.py`：多 sheet / 隐藏 sheet(isHidden) / activeSheet / fileName
- `test_worksheet_info.py`：默认 active / 指定 worksheetName / usedRange 行列 / tableCount,chartCount / 空表
- `test_selected_range.py`：单元格 / 2D values / 空选区 / 混合类型
- 错误码：5001（worksheetInfo 传 ghost 表名）

**#30 range_e2e/**
- `test_set_get_range.py` / `test_clear_range.py` / `test_copy_range.py` / `test_shift_range.py` / `test_set_formula.py`
- 双验证：cell value / 公式字符串 / 清除后为空；falsy `0`/`False` 落格
- 错误码：5002 / 5005 / 5004

**#31 format_e2e/**
- `test_set_range_format.py`（font/fill/alignment/borders/numberFormat/偏更新）/ `test_get_range_format.py` / `test_conditional_format.py` / `test_merge_cells.py`
- 双验证：fill hex / number_format / merged_ranges / 字体属性
- 错误码：5003 / 5002；注意读写不对称（写偏更新模型 vs 读 RangeFormatInfo）

**#32 worksheet_e2e/**
- `test_add_worksheet.py` / `test_rename_activate.py` / `test_delete_worksheet.py`
- 双验证：`sheet_names`（增/删/改名）；错误码：5001

**#33 table_e2e/**
- `test_insert_table.py` / `test_get_table.py` / `test_table_rows.py`（含 `rowIndex=0` falsy）/ `test_sort_table.py`
- 双验证：`table_names` / 排序后行序 / 追加行内容；错误码：5006 / 5009 / 4004

**#34 chart_e2e/**（视觉为主，openpyxl `chart_count` best-effort）
- `test_insert_chart.py`（ColumnClustered/Line/Pie/XYScatter…）/ `test_get_charts.py` / `test_update_chart.py` / `test_delete_chart.py`
- 错误码：4002（无效 chartType）/ 5007 / 5002

**#35 pivot_table_e2e/**（视觉为主，openpyxl 透视支持有限）
- `test_insert_pivot_table.py` / `test_get_pivot_tables.py` / `test_delete_pivot_table.py`
- 只用 sourceAddress + targetAddress 创建，**不构造** rows/columns/values/filters（spec 未定义）
- 错误码：5008 / 5010 / 5002

**#36 find_filter_e2e/**
- `test_find_values.py`（matchCase/matchEntireCell 默认 false 须落 wire，含无命中）/ `test_auto_filter.py`
- 双验证：`auto_filter_ref` / 返回 matches；错误码：4004 / 5001 / 5002

---

## 六、收口

- 8 个子问题全部完成后，逐一在对应 GitHub issue 勾选 DoD 并关闭（`gh issue close <n>`），在 #28 跟踪表回填进度。
- 最终在 fork 上发起到 upstream `feat/excel-e2e-oasp-0.3.0` 的 PR（或按团队约定）。
- 真机门见 #26 / `docs/manual_tests/excel_v0.3.0.md`。

---

**生成日期**：2026-06-18
