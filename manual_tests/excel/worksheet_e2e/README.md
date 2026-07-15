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

## 错误码（oasp#17 已定案，issue #82 校准）

- 工作表不存在（getItem 失败）→ **`3010` ELEMENT_NOT_FOUND**（`details.kind:"worksheet"`）
- add 重名（名称已存在）→ **`3004` OPERATION_FAILED**（events-excel.md add:worksheet）
- Zod 校验失败 → `4000` VALIDATION_ERROR（不变）

以上为规范层 MUST（不得降级 `3000`；旧 `5001` 与「细分码 dead code、真机 3000」口径
均已过时）。Add-In 接线（office-editor4ai#80）前真机仍返 `3000` 兜底，e2e 用例以
**XFAIL** 运行，接线后摘 `xfail_reason` 转正。

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
