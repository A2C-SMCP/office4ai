# Excel Worksheet 管理 E2E（#32）

覆盖 OASP 0.3.0 `/excel` 命名空间 **#21 Worksheet** 切片的 5 个事件，真机在环 + openpyxl 读盘双验证。

## 事件与 wire 形态（以 AddIn `worksheets.ts` 实测为准）

| 事件 | 请求参数 | 返回 data |
|------|---------|-----------|
| `excel:get:worksheets` | （无） | `{worksheets: [{name, index, isActive, isHidden}]}` |
| `excel:add:worksheet` | `name?`（省略→Excel 自动命名 "SheetN"） | `{name, index}` |
| `excel:delete:worksheet` | `worksheetName` | **void**（无 data） |
| `excel:rename:worksheet` | `currentName` + `newName` | `{name}` |
| `excel:activate:worksheet` | `worksheetName` | **void**（无 data） |

> ⚠️ **DTO-vs-wire**：office4ai `DeleteWorksheetData {deleted}` / `ActivateWorksheetData {activated}`
> 是声明式的，AddIn handler 实际返回 **void**，真机响应里**没有** `deleted`/`activated` 字段。
> 故 delete/activate 一律靠 openpyxl `sheet_names` / `wb.active` 读盘验证，不断言响应体。
> 详见 memory `excel-response-dto-not-validated-against-wire`。

## 测试文件（3 个）

| 文件 | 用例数 | 覆盖 |
|------|-------|------|
| `test_add_worksheet.py` | 5 | add（命名/自动命名）、get:worksheets（列举/反映新增）、错误码 |
| `test_rename_activate.py` | 4 | rename、activate（经 get:worksheets 验 isActive）、2 错误码 |
| `test_delete_worksheet.py` | 3 | delete（读盘核对）、delete 后 get:worksheets 反映、错误码 |

**合计 12 case。**

## 夹具 `fixtures/worksheet_e2e/book.xlsx`

4 张自描述表（A1 = `<表名>-A1`）：`Alpha`（activeSheet）/ `Beta`（改名·删除目标）/
`Gamma`（激活目标）/ `Delta`（隔离表，增删改其它表时须原样保留）。每个 case 打开独立工作副本，互不污染。

## 错误码现实（#29/#30 已确立，本套件沿用）

非法/不存在的表名 → **`3000` DOCUMENT_ERROR**，**不是** 旧 DoD 写的 `5001 WORKSHEET_NOT_FOUND`。
`excelErrorCode()` 仅把 Zod 校验失败映射为 `4000`，其余一切 Office.js 运行期异常（含
`worksheets.getItem()` 取不到表、`worksheets.add()` 重名）→ `3000`。`5001`/`3009` 等细分码均为
dead code。

## 运行

```bash
# 单文件全用例
uv run python manual_tests/excel/worksheet_e2e/test_add_worksheet.py --test all
# 单用例
uv run python manual_tests/excel/worksheet_e2e/test_rename_activate.py --test 2
# 仅列用例（不开 Excel）
uv run python manual_tests/excel/worksheet_e2e/test_delete_worksheet.py --list
```

真机激活：首个 case 手动点「加载项 → TFEditor4Office」，后续 case 自动重连。
