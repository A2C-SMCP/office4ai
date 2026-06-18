# Excel Format / 条件格式 / 合并 E2E 测试

覆盖 **#20 Format 切片** 6 个事件（OASP /excel Draft 0.3.0，子问题 #31）：

- `excel:get:rangeFormat` — 读区域格式（RangeFormatInfo：font/fill/对齐/numberFormat 2D）
- `excel:set:rangeFormat` — 设区域格式（font/fill/alignment/borders/numberFormat，**偏更新**）
- `excel:add:conditionalFormat` — 添加条件格式规则（透传，cellValue 需 operator/formula1/format）
- `excel:clear:conditionalFormat` — 清除区域条件格式
- `excel:merge:cells` — 合并（单格，保留左上值）
- `excel:unmerge:cells` — 取消合并

> 读写不对称：**写侧** `set:rangeFormat` 用全可选偏更新模型（含 borders、嵌套 alignment、
> numberFormat 标量串）；**读侧** `get:rangeFormat` 返回扁平 RangeFormatInfo（numberFormat 为
> 2D 数组）。写操作响应仅 `{address}`，故用 openpyxl 读盘核对字体/填充/数字格式/对齐/边框/合并。

## 目录结构

```
format_e2e/
├── __init__.py                 # 包初始化
├── _fixtures.py                # 夹具构建器（fmt.xlsx：Data 数值网格 / Merge 标签行）
├── test_set_range_format.py    # set:rangeFormat（6 例：font/fill/numberFormat/alignment/borders/偏更新）
├── test_get_range_format.py    # get:rangeFormat（4 例，含错误码 3000）
├── test_conditional_format.py  # add/clear:conditionalFormat（4 例，含错误码 3000）
├── test_merge_cells.py         # merge/unmerge:cells（4 例，含错误码 3000）
└── README.md                   # 本文档
```

夹具 `manual_tests/excel/fixtures/format_e2e/fmt.xlsx`（activeSheet=Data）：

| 工作表 | 内容 | 用途 |
|--------|------|------|
| `Data` | A1:C4 数值网格（Region/Q1/Q2 + 数字） | set/get:rangeFormat、conditionalFormat（B 列数值供 cellValue） |
| `Merge` | A1:C1 = M1/M2/M3 | merge/unmerge:cells（合并保留左上、其余清空） |

> 夹具已入库；如需重建：`uv run python manual_tests/excel/format_e2e/_fixtures.py`

## 测试概览

### 1. 设置区域格式 (`test_set_range_format.py`)

| 编号 | 名称 | 描述 / 验证（openpyxl 读盘） |
|------|------|------------|
| 1 | font 加粗+斜体+颜色 | `font={bold,italic,color:#FF0000}` → `font.bold/.italic/.color` |
| 2 | fill 填充色 | `fill={color:#FFFF00}` → `cell_fill_hex='#FFFF00'` |
| 3 | numberFormat | `numberFormat='0.00'` → `cell_number_format='0.00'` |
| 4 | alignment 对齐 | `{horizontal:Center,vertical:Top,wrapText}` → openpyxl 小写 center/top + wrap |
| 5 | borders 边框 | `{top/bottom:Continuous}` → `border.top.style` 非空 |
| 6 | 偏更新保留既有属性 | 先填绿再只设 bold → 填充保留 + bold 生效 |

### 2. 读取区域格式 (`test_get_range_format.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | 读默认格式结构 | RangeFormatInfo 字段完整，numberFormat 为 2D |
| 2 | set 后 get 往返 | 先 set {bold,fill,numberFormat} → get 反映 |
| 3 | 多格 numberFormat 2D 维度 | A1:C2 → numberFormat 为 2×3 |
| 4 | 错误码 3000 | 非法 address → DOCUMENT_ERROR |

### 3. 条件格式 (`test_conditional_format.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | add cellValue 规则 | B2:B4 `>150` 红底 → 协议成功 + best-effort CF 计数 |
| 2 | add colorScale 规则 | B2:B4 colorScale → 协议成功 |
| 3 | clear 条件格式 | 先 add → clear → 协议成功 |
| 4 | 错误码 3000 | 非法 address → DOCUMENT_ERROR |

> 条件格式以**协议成功**为主验证；openpyxl 对 CF 的反序列化差异较大，规则计数为
> best-effort（读不到打印 ⚠️ 不判失败）。视觉验收见 `docs/manual_tests/excel_v0.3.0.md`。

### 4. 合并/取消合并 (`test_merge_cells.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | merge A1:C1 | `merged_ranges` 含 A1:C1，左上 'M1' 保留 |
| 2 | unmerge 还原 | 先 merge → unmerge → merged_ranges 不含该区域 |
| 3 | 合并清空非左上 | merge 后 B1/C1 为 None |
| 4 | 错误码 3000 | 非法 address → DOCUMENT_ERROR |

## 运行方式

```bash
uv run python manual_tests/excel/format_e2e/test_set_range_format.py --test all
uv run python manual_tests/excel/format_e2e/test_get_range_format.py --test 2
uv run python manual_tests/excel/format_e2e/test_merge_cells.py --list
```

| 参数 | 作用 |
|------|------|
| `--test N` / `--test all` | 跑单个用例 / 全部 |
| `--list` | 仅列出用例（不连 Excel） |
| `--no-auto-open` | 不自动打开/激活，逐用例人工配合 |
| `--keep` | 成功后保留工作副本供目测 |

> 真机激活：首个用例弹出工作簿后，在 Excel「加载项」点选本地 TFEditor4Office 激活**一次**，
> 后续用例自动重连。给手动激活留时间可设 `EXCEL_E2E_TIMEOUT=120`。

## 错误码现实（非法 address → **3000**，3009 是 dead code）

> ⚠️ 与 #30 同源：`error-codes.ts` 定义了 `3009 RANGE_INVALID` / `3014`(合并类) 等，但
> **excel handler 0 引用 = dead code**。`excel-handlers.ts` 的 `excelErrorCode()` 仅把 Zod 失败
> → `4000` VALIDATION_ERROR，其余 Office.js 运行期异常（含非法地址、合并冲突）→ `3000`
> OFFICE_API_ERROR。故本套件错误码用例一律按 **3000** 断言。详见
> `docs/manual_tests/excel_e2e_dev_plan.md`「错误码现实」与 `range_e2e/README.md`。

## 常见问题

### Q1: 字体颜色 rgb 对不上

openpyxl 读 ARGB（如 `FFFF0000`）；断言只检查是否含 `FF0000`（后 6 位）。不同 Excel 主题色
存储可能差异，颜色用 ⚠️ 软告警，bold/italic 等布尔属性为硬断言。

### Q2: 条件格式 openpyxl 读不到规则

Excel 对 CF 的序列化与 openpyxl 解析存在差异，规则计数为 best-effort。协议成功即视为通过，
视觉以 Excel 中条件格式管理器为准。

### Q3: 边框 style 名称不一致

Office.js 用 `Continuous` + `weight`，openpyxl 读到的是映射后的 `thin/medium/thick` 等；
用例只断言 `border.top.style` 非空（确有边框）。

## 相关文档

- [DTO 定义](../../../office4ai/environment/workspace/dtos/excel.py)（`SetRangeFormatOptions` / `RangeFormatInfo` / `ConditionalFormatRule` 等 #20 段）
- [开发规划](../../../docs/manual_tests/excel_e2e_dev_plan.md)
- [#30 Range 套件](../range_e2e/README.md)
- 上游 Issue：#31（本目录）· #28（Tracking）

## 最后更新

2026-06-18
