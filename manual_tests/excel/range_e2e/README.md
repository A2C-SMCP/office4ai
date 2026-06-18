# Excel Range CRUD + 公式 E2E 测试

覆盖 **#19 Range 切片** 7 个事件（OASP /excel Draft 0.3.0，子问题 #30）：

- `excel:set:range` — 写入范围（标量填充 / 二维数组）
- `excel:get:range` — 读取范围（2D values / rowCount / includeFormat）
- `excel:clear:range` — 清除内容/格式/全部（clearType）
- `excel:copy:range` — 复制范围到目标位置（sourceAddress → targetAddress）
- `excel:insert:range` — 插入空白单元格并位移现有单元格（down/right）
- `excel:delete:range` — 删除单元格并补位（up/left）
- `excel:set:formula` — 设置单元格公式（透传字符串）

> 写操作响应仅 `{address}`（最小锚点）。验证策略：协议返回 + **openpyxl 读盘双重核对**
> 写入后的真实单元格内容。`reader.reload()` 先经 AppleScript 强制 Excel 存盘再 openpyxl 读。

## 目录结构

```
range_e2e/
├── __init__.py            # 包初始化
├── _fixtures.py           # 夹具构建器（grid.xlsx：Data / Blank / Shift 三表）
├── test_set_get_range.py  # set:range + get:range（6 个用例，含错误码 3009）
├── test_clear_range.py    # clear:range（4 个用例）
├── test_copy_range.py     # copy:range（3 个用例）
├── test_shift_range.py    # insert:range + delete:range（4 个用例）
├── test_set_formula.py    # set:formula（4 个用例，含错误码 3009）
└── README.md              # 本文档
```

夹具位于 `manual_tests/excel/fixtures/range_e2e/grid.xlsx`（activeSheet=Data）：

| 工作表 | 内容 | 用途 |
|--------|------|------|
| `Data` | A1:C4 数值网格（Region/Q1/Q2 + North/South/East，Q 列为数字） | get 已有数据、copy 源、set:formula 数值引用 |
| `Blank` | 空表 | set:range 写入、clear:range（先填后清）不污染 Data |
| `Shift` | 3×3 自描述网格（每格值=自身地址 "A1".."C3"） | insert/delete 后断言「谁落到哪」 |

> 夹具已入库；如需重建：`uv run python manual_tests/excel/range_e2e/_fixtures.py`

## 测试概览

### 1. 范围读写往返 (`test_set_get_range.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | set 2D + get 往返 | 写 A1:C2 二维数组 → get 读回，wire + 磁盘一致 |
| 2 | 标量填充 | `values='Z'` 标量铺满 A1:B2 → 全为 'Z' |
| 3 | **falsy 存活** | `[[0, False, '']]` → get 读回 0/False 必须存活为本身（不被吞为 null） |
| 4 | 读已有数据 | get Data!A1:C4 → 4×3，表头 [Region,Q1,Q2]，C4=350 |
| 5 | includeFormat | get includeFormat=true → 返回 RangeFormatInfo（font/fill） |
| 6 | 错误码 3000 | 非法 address → DOCUMENT_ERROR（真机实测；3009 为 dead code，见下） |

### 2. 清除范围 (`test_clear_range.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | clear contents | 填后清内容 → 值清空（None） |
| 2 | clear all | 填后清全部 → 值清空 |
| 3 | clear formats 保留内容 | 清格式 → 内容 'stay' **保留** |
| 4 | 清子区域不影响相邻 | clear Data!B2:C2 → 仅该区域空，A2/B3 完好 |

### 3. 复制范围 (`test_copy_range.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | 复制单行 | Data!A1:C1 → A6，目标=源且源不变 |
| 2 | 复制数据块 | Data!A2:C4 → E2，E2='North' / G4=350 |
| 3 | 复制单元格 | Data!A1 → E1，E1='Region' |

### 4. 插入/删除位移 (`test_shift_range.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | insert down | insert A2 down → A2 空出，旧 A2 落 A3 |
| 2 | insert right | insert B1 right → B1 空出，旧 B1 落 C1 |
| 3 | delete up | delete A2 up → 旧 A3 上移到 A2 |
| 4 | delete left | delete B1 left → 旧 C1 左移到 B1 |

### 5. 设置公式 (`test_set_formula.py`)

| 编号 | 名称 | 描述 / 验证 |
|------|------|------------|
| 1 | SUM 公式 | E2 = '=SUM(B2:C2)' → 公式落盘 |
| 2 | 算术公式 | E3 = '=B3+C3' → 公式落盘 |
| 3 | 引用/乘法 | E4 = '=B4*2' → 公式落盘 |
| 4 | 错误码 3000 | 非法 address → DOCUMENT_ERROR（同上，3009 为 dead code） |

> openpyxl 不计算公式，读到的是公式串本身；断言对「去空格 + 大写」后做子串包含匹配，
> 容忍 Excel 轻度规范化。

## 运行方式

```bash
# 单脚本全部用例
uv run python manual_tests/excel/range_e2e/test_set_get_range.py --test all
# 单个用例
uv run python manual_tests/excel/range_e2e/test_shift_range.py --test 2
# 仅列出用例（离线自检，不连 Excel）
uv run python manual_tests/excel/range_e2e/test_clear_range.py --list
```

| 参数 | 作用 |
|------|------|
| `--test N` / `--test all` | 跑单个用例 / 全部 |
| `--list` | 仅列出用例（不连 Excel） |
| `--no-auto-open` | 不自动打开/激活，逐用例人工配合 |
| `--keep` | 成功后保留工作副本供目测 |

> 真机激活：首个用例弹出工作簿后，在 Excel「加载项」点选本地 TFEditor4Office 激活**一次**，
> 后续用例自动重连。给手动激活留时间可设 `EXCEL_E2E_TIMEOUT=120`。

## 错误码现实（真机实测：非法 address → **3000**，3009 是 dead code）

> ⚠️ 两层过时：① Issue/旧 DTO 注释写的 `5002` 不存在；② 连「真实码应为 3009 RANGE_INVALID」
> 这个推断也**不成立**。真机实测 `get:range` / `set:formula` 传非法地址 `ZZZZ99999999` 返回
> **`3000 DOCUMENT_ERROR`**。
>
> 源码确认（`office-editor4ai`）：
> - `error-codes.ts` 定义了 `RANGE_INVALID="3009"`，但**全仓 0 个 handler 发射它**（仅
>   error-codes.ts 与对齐测试引用）—— 3009 是 **dead code**。
> - `excel-handlers.ts` 的 `excelErrorCode()` 只把 **Zod 校验失败 → `4000` VALIDATION_ERROR**，
>   其余一切 Office.js 运行期异常（含 `getRange` 拒绝畸形地址）→ **`3000` OFFICE_API_ERROR**。
> - 畸形串 `ZZZZ99999999` 能过 Zod（是合法 string），故在 Office.js 层被拒 → 落到 3000。
>
> **结论**：非法范围地址类用例一律按 **3000** 断言（与 #29 的 5001→3000 同源）。仅当构造
> **schema 违规**（类型错/缺必填）时才会得到 `4000`。详见 `docs/manual_tests/excel_e2e_dev_plan.md`
> 「错误码现实」与 `read_state_e2e/README.md`。

## 前置条件

1. ✅ Excel 已由 runner 自动复制并打开夹具副本。
2. ✅ 本地 TFEditor4Office Add-In 已激活（首个用例点一次）。
3. ✅ Workspace 在 `127.0.0.1:3000` 自动启停。

## 常见问题

### Q1: 公式用例读到 None

确认 set:formula 协议返回成功；openpyxl 读的是 AppleScript 存盘后的磁盘文件。若 Excel
未及时存盘，`reader.reload()` 会重试存盘。

### Q2: copy/shift 目标位置不对

Office.js 的 `copyFrom` 以 targetAddress 为左上角粘贴；insert/delete 的位移方向由
`shiftDirection` 决定。Shift 表的自描述标签可直接看出「谁落到哪」。

### Q3: 错误码不是 3009

不同 Add-In 版本对非法地址的判定可能落到 3000（DOCUMENT_ERROR）等相邻码。真机实测以
Add-In 实际返回为准，必要时按 #29 的做法回填真实码。

## 相关文档

- [DTO 定义](../../../office4ai/environment/workspace/dtos/excel.py)（`ExcelSetRangeRequest` 等 #19 Range 段）
- [开发规划](../../../docs/manual_tests/excel_e2e_dev_plan.md)
- [#29 Read 套件](../read_state_e2e/README.md)
- 上游 Issue：#30（本目录）· #28（Tracking）

## 最后更新

2026-06-18
