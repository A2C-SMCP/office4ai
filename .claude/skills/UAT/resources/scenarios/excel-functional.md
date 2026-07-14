# UAT 场景：Excel 功能验收 (excel-functional) — Phase 2

## 测试目标

在真实 Excel + Add-In 环境下，验证 Excel 工具**不仅注册完整（Phase 1），且功能可正常跑通**——由 `manual_tests/excel/` 的 E2E 用例驱动，openpyxl 双重验证写盘结果。触发人协助连接 Add-In，逐个走查。

> Phase 1（`excel-tools` 注册验收）确认「工具在」；本场景确认「工具能用」。两者互补。

## 前置

- macOS + Excel 桌面版；office-editor4ai `excel-editor4ai` dev server 运行（`pnpm dev:excel`）。
- 触发人：打开对应工作簿后**手动激活 taskpane**（加载项面板点 excel-editor），使其与 workspace 握手（`https://127.0.0.1:4443` 或脚本 workspace :3000/:4443）。
- 运行 E2E 脚本需 `dangerouslyDisableSandbox`（网络 bind + AppleScript + Excel）。各 category runner 连接超时 30s——触发人需及时激活 taskpane。

## Health 档（接线桥接，**无需 Add-In**）

先跑一次接线健康检查，确认 37 个 `excel:*` 事件全部注册到 request_registry（PR/CI 级回归守护）：

```bash
uv run python manual_tests/excel/test_excel_e2e.py --mode health
```

**通过标准**：输出全部 37 事件已注册、无缺失。

## Smoke 档（默认，**一键全 37**）

Excel 有全量编排驱动，一条命令把 37 个事件编成「构建销售报表」工作流逐个跑通（continue-on-error，末尾汇总每事件 ✅/❌）+ openpyxl 读盘双验 + 触发 5001/5002/5006/5007/5008 错误码：

```bash
uv run python manual_tests/excel/test_excel_e2e.py --mode full
```

**通过标准**：37 事件汇总全 ✅；openpyxl 双验通过；错误码用例如期触发。**这一条即覆盖 Excel 功能面的绝大部分。**

## Full 档（逐类别深走，8 类）

在 smoke 基础上，逐个类别 `--test all` 走查（每类 `--list` 可看用例）：

| 类别 | 代表命令（`--test all`） | 覆盖 |
|---|---|---|
| read_state | `manual_tests/excel/read_state_e2e/test_workbook_info.py` · `test_worksheet_info.py` · `test_selected_range.py` | 工作簿/工作表/选区读取 |
| range | `range_e2e/test_set_get_range.py` · `test_copy_range.py` · `test_clear_range.py` · `test_set_formula.py` · `test_shift_range.py` | 读写/复制/清除/公式/插删移位 |
| format | `format_e2e/test_set_range_format.py` · `test_get_range_format.py` · `test_conditional_format.py` · `test_merge_cells.py` | 区域格式/条件格式/合并 |
| worksheet | `worksheet_e2e/test_add_worksheet.py` · `test_delete_worksheet.py` · `test_rename_activate.py` | 增删/重命名/激活 |
| table | `table_e2e/test_insert_table.py` · `test_get_table.py` · `test_table_rows.py` · `test_sort_table.py` | 插入/读取/行增删/排序 |
| chart | `chart_e2e/test_insert_chart.py` · `test_get_charts.py` · `test_update_chart.py` · `test_delete_chart.py` | 插入/读取/更新/删除图表 |
| pivot_table | `pivot_table_e2e/test_insert_pivot_table.py` · `test_get_pivot_tables.py` · `test_delete_pivot_table.py` | 数据透视表 CRUD |
| find_filter | `find_filter_e2e/test_find_values.py` · `test_auto_filter.py` | 查找/自动筛选 |

（命令前缀均为 `uv run python manual_tests/excel/…  --test all`）

**通过标准**：每类别用例全 ✅（协议成功 + openpyxl 内容双验）；失败项记录实际现象。

## 报告片段

```
### Excel 功能验收
- Health（37 事件注册）: ✅
- Smoke（test_excel_e2e --mode full）: 37/37 ✅
- Full（8 类别）: read_state ✅ / range ✅ / format ✅ / worksheet ✅ / table ✅ / chart ✅ / pivot ✅ / find_filter ✅
```
