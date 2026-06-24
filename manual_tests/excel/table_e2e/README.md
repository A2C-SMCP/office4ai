# Excel Table 操作 E2E（#33）

覆盖 OASP 0.3.0 `/excel` 命名空间 **#22 Table** 切片的 6 个事件，真机在环 + openpyxl 读盘双验证。

## 事件与 wire 形态（以 AddIn `table.ts` 实测为准）

| 事件 | 请求参数 | 返回 data |
|------|---------|-----------|
| `excel:insert:table` | `address, hasHeaders, data?, styleName?, worksheetName?` | `{name, address}` |
| `excel:get:table` | `tableId, worksheetName?` | `TableInfo {name,id,address,rowCount,columnCount,columns:[{name,index}],styleName,showHeaders}` |
| `excel:get:tables` | `worksheetName?` | `{tables:[{name, id, address}]}`（每条仅 3 字段） |
| `excel:add:tableRow` | `tableId, values, worksheetName?` | `{tableId}` |
| `excel:delete:tableRow` | `tableId, rowIndex, worksheetName?` | **void** |
| `excel:sort:table` | `tableId, sortFields:[{columnIndex, ascending?}], worksheetName?` | **void** |

> `insert:table` 约定：`hasHeaders=true` 时 `data` 第一行为表头、覆盖整张表范围；`false` 时 `data` 全为正文。
> `delete:tableRow` 返回 void（无 data）；`rowIndex` 为 **0-based、不含表头**，`rowIndex=0` 是首个正文行（falsy 必须落 wire）。

## 测试文件（4 个）

| 文件 | 用例数 | 覆盖 |
|------|-------|------|
| `test_insert_table.py` | 4 | 在已有数据上建表 / 带 data 写空白区 / styleName / 错误码 3000 |
| `test_get_table.py` | 4 | get:table 富信息 / get:tables 精简列表 / 跨表按名取 / 错误码 3000 |
| `test_table_rows.py` | 6 | add 追加 / delete rowIndex=0(falsy) / delete 中间 / 越界 3000 / 负数 4000 / 表不存在 3000 |
| `test_sort_table.py` | 4 | 单列升 / 单列降 / 多级(断 tie) / 错误码 3000 |

**合计 18 case。**

## 夹具 `fixtures/table_e2e/tbl.xlsx`

- `Raw`（activeSheet）—— A1:C4 纯数据网格（无表），供 insert:table 在已有数据上建表。
- `Blank` —— 空表，供 insert:table 带 data 写入。
- `Sales` —— 预置表 `SalesTable`（A1:C5，Name/Score/Age + 4 行；Score 含并列 85，Age 唯一）。
- `Roster` —— 预置表 `RosterTable`（A1:B4，Item/Qty + 3 行自描述 R0row/R1row/R2row）。

openpyxl 直接写 Excel 表（ListObject），Excel 打开后 Office.js `worksheet.tables` 即可识别。每个 case 打开独立工作副本，互不污染。

## 错误码现实（Zod 校验决定两条路径）

| 触发 | 错误码 |
|------|-------|
| 表不存在（合法非空 id）/ rowIndex·columnIndex **越界**（合法非负但 Office.js 拒绝） | **`3000`** DOCUMENT_ERROR |
| tableId **空串** / rowIndex·columnIndex **负数**（Zod `nonnegative()` 失败） | **`4000`** VALIDATION_ERROR |

> ⚠️ 实测为 `4000`（非旧 DoD 写的 `4004`）：`excelErrorCode()` 仅把 Zod 失败映射为 `4000`，其余一切 Office.js 运行期异常 → `3000`。旧 DoD 的 `5006`/`5009` 均为 dead code。

## 运行

```bash
uv run python manual_tests/excel/table_e2e/test_insert_table.py --test all
uv run python manual_tests/excel/table_e2e/test_get_table.py --test all
uv run python manual_tests/excel/table_e2e/test_table_rows.py --test all
uv run python manual_tests/excel/table_e2e/test_sort_table.py --test all
```

真机激活：首个 case 手动点「加载项 → TFEditor4Office」，后续 case 自动重连。
