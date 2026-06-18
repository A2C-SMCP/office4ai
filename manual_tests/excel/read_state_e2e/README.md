# Excel Read/Foundation 状态感知读 E2E 测试

覆盖 **状态感知读** 3 个事件（OASP /excel Draft 0.3.0，对齐功能切片 #18）：

- `excel:get:workbookInfo` — 工作簿信息（sheets / activeSheet / fileName）
- `excel:get:worksheetInfo` — 工作表详情（usedRange / tableCount / chartCount）
- `excel:get:selectedRange` — 当前选中范围及其 2D values

> 读类工具以**返回 data**（状态感知）验证为主，openpyxl 用于核对 sheet 存在 / usedRange 行列。

## 目录结构

```
read_state_e2e/
├── __init__.py               # 包初始化
├── _fixtures.py              # 夹具构建器（multi_sheet / hidden_sheet / prefilled）
├── test_workbook_info.py     # get:workbookInfo（4 个用例）
├── test_worksheet_info.py    # get:worksheetInfo（6 个用例，含错误码 5001）
├── test_selected_range.py    # get:selectedRange（4 个用例，需手动选区）
└── README.md                 # 本文档
```

夹具位于 `manual_tests/excel/fixtures/read_state_e2e/`：

| 夹具 | 内容 | 用途 |
|------|------|------|
| `multi_sheet.xlsx` | Sheet1(空) / Data(A1:C4 预填) / Report(空)，activeSheet=Sheet1 | workbook 多表、active、fileName、worksheet usedRange |
| `hidden_sheet.xlsx` | Visible(可见) / Hidden(隐藏) | workbookInfo isHidden 用例 |
| `prefilled.xlsx` | 单 Data 表，A1:C2 混合类型（含 falsy `0`/`False`/空串） | selectedRange 2D / 混合类型 |

> 夹具已入库；如需重建：`uv run python manual_tests/excel/read_state_e2e/_fixtures.py`

## 测试概览

### 1. 工作簿信息 (`test_workbook_info.py`)

| 编号 | 名称 | 夹具 | 描述 / 验证 |
|------|------|------|------------|
| 1 | 多 sheet 工作簿 | multi_sheet | sheets 列表 ≥3，且每个名都能在 openpyxl 磁盘表名中核对 |
| 2 | 含隐藏 sheet | hidden_sheet | 返回某 sheet `isHidden=true`，openpyxl `sheet_state=hidden` |
| 3 | activeSheet 正确 | multi_sheet | `activeSheet` 合法且与 `sheets[].isActive` 一致 |
| 4 | fileName 字段 | multi_sheet | `fileName` 为 `.xlsx` 文件名 |

```bash
uv run python manual_tests/excel/read_state_e2e/test_workbook_info.py --test 1
uv run python manual_tests/excel/read_state_e2e/test_workbook_info.py --test all
```

### 2. 工作表详细信息 (`test_worksheet_info.py`)

| 编号 | 名称 | worksheetName | 描述 / 验证 |
|------|------|--------------|------------|
| 1 | 默认 active 表 | （省略） | 返回活动表 Sheet1，含 usedRange |
| 2 | 指定 worksheetName | `Data` | 命中指定表 name=Data |
| 3 | usedRange 行列数 | `Data` | usedRange 4 行 × 3 列，与 openpyxl `max_row/max_column` 核对 |
| 4 | tableCount/chartCount | `Data` | 两字段为整数（预填夹具应为 0/0） |
| 5 | 空表 usedRange | `Report` | 空表 usedRange 极小（≤1×1） |
| 6 | 错误码 5001 | `GhostSheet_xyz` | **预期失败**：WORKSHEET_NOT_FOUND（5001） |

```bash
uv run python manual_tests/excel/read_state_e2e/test_worksheet_info.py --test 3
uv run python manual_tests/excel/read_state_e2e/test_worksheet_info.py --test 6   # 错误码
uv run python manual_tests/excel/read_state_e2e/test_worksheet_info.py --test all
```

### 3. 当前选中范围 (`test_selected_range.py`)

读取 Excel **实时选区**，每个用例运行前需**手动在 Excel 中选好区域**。建议加
`--no-auto-open`，在打开夹具后手动选区，再回车继续。

| 编号 | 名称 | 运行前请选中 | 验证 |
|------|------|------------|------|
| 1 | 单元格选区 | `A1` | 1×1，values=`[[x]]` |
| 2 | 多单元格 2D values | `A1:C2` | 2 行 3 列，values 为 2D |
| 3 | 空选区 | `F10`（空白） | 1×1，首格 None/空串 |
| 4 | 混合类型值 | `A1:C1`（字符串/数字/布尔） | 同一选区含多种 Python 类型 |

```bash
uv run python manual_tests/excel/read_state_e2e/test_selected_range.py --test 2 --no-auto-open
uv run python manual_tests/excel/read_state_e2e/test_selected_range.py --test all --no-auto-open
```

> `rowCount`/`columnCount` 与实际选区不符时打印 ⚠️（非失败），提醒你确认选区；结构不变量
> （2D / 行列与 values 维度一致）不满足才判失败。

## 通用参数

| 参数 | 作用 |
|------|------|
| `--test N` / `--test all` | 跑单个用例 / 全部 |
| `--list` | 仅列出用例（不连 Excel，可离线自检脚本） |
| `--no-auto-open` | 不自动打开/激活，逐用例人工配合（selectedRange 必用） |
| `--keep` | 成功后保留工作副本供目测（不自动关闭工作簿） |

## 前置条件

1. ✅ **Excel 已打开夹具工作簿** —— 由 runner 自动复制并打开（`--no-auto-open` 则手动打开 `.test_working/` 下的副本）。
2. ✅ **excel-editor4ai Add-In 已激活** —— 在 Excel「加载项」中点选 `excel-editor4ai`，连接 `http://127.0.0.1:3000`。
3. ✅ **Workspace 自动启动** —— 测试在 `127.0.0.1:3000` 自动启停，无需手动起服务。
4. ✅ **selectedRange 用例**：运行前在 Excel 中按上表选好区域。

## 预期结果

```
✅ Workspace 启动成功
✅ Excel Add-In 已连接
📝 执行: excel:get:workbookInfo (params={})...
   ✅ data={...}
📊 双重验证:
   ✅ 3 个 sheet，与 openpyxl 磁盘表名一致: ['Sheet1', 'Data', 'Report']
✅ 测试 1 通过
```

错误码用例（test 6）预期**失败但命中码**：

```
🎯 预期错误码: 5001
   ✅ 预期失败 [5001] → ok=False err=...WORKSHEET_NOT_FOUND...
✅ 测试 6 通过
```

## 常见问题

### Q1: Add-In 连接超时

确认 Excel 已加载并点选 `excel-editor4ai`；检查能否访问 `http://127.0.0.1:3000`；看浏览器/任务窗格控制台报错。

### Q2: selectedRange 行列数与预期不符（⚠️ 而非 ❌）

读的是 Excel 当前实时选区。请在运行该用例前先在 Excel 中选好 README 指定的区域；`--no-auto-open` 模式会等待你回车，便于先选区。

### Q3: fileName 不以 `multi_sheet` 开头

runner 会把夹具复制成带时间戳的工作副本（`multi_sheet_<时间戳>.xlsx`）。Excel 返回的 `fileName` 为该工作副本名，正常即以 `multi_sheet` 开头；若你手动改名或另存，会打印 ⚠️ 但不判失败。

### Q4: 想离线自检脚本是否可运行

`--list` 不连接 Excel，仅枚举用例，可用于 CI / 改完脚本后的快速自检。

## 相关文档

- [DTO 定义](../../../office4ai/environment/workspace/dtos/excel.py)（`WorkbookInfo` / `WorksheetInfo` / `SelectedRangeInfo`）
- [验收清单](../../../docs/manual_tests/excel_v0.3.0.md)
- [开发规划](../../../docs/manual_tests/excel_e2e_dev_plan.md)
- 上游 Issue：#29（本目录）· #28（Tracking）

## 最后更新

2026-06-18
