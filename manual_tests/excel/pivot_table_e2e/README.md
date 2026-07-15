# Excel PivotTable 操作 E2E（#35）

覆盖 OASP 0.3.0 `/excel` 命名空间 **#24 PivotTable** 切片的 3 个事件。透视表为视觉对象，以
**协议层 get:pivotTables 回读**为唯一验证依据（openpyxl 透视支持有限）。

## 事件与 wire 形态（以 AddIn `pivotTable.ts` 实测为准）

| 事件 | 请求参数 | 返回 data |
|------|---------|-----------|
| `excel:insert:pivotTable` | `sourceAddress, targetAddress, name?, worksheetName?` | `{name}`（name 省略默认 "PivotTable"） |
| `excel:get:pivotTables` | `worksheetName?` | `{pivotTables:[{name, id}]}`（每条仅 2 字段） |
| `excel:delete:pivotTable` | `pivotTableName, worksheetName?` | **void** |

> ⚠️ **source 与 target 须同一 worksheet**：AddIn 两个 `getRange` 都基于同一 worksheet，
> 故 targetAddress 是源表内的空白落点（如 source `A1:C5` → target `E1`）。
> 仅用 sourceAddress + targetAddress 创建**空透视表骨架**，不构造 rows/columns/values/filters（spec 未定义）。

## 测试文件（3 个）

| 文件 | 用例数 | 覆盖 |
|------|-------|------|
| `test_insert_pivot_table.py` | 4 | 指定 name / 默认 name / 非法 source 3009 / 空 source 4000 |
| `test_get_pivot_tables.py` | 4 | 单透视表回读(name/id) / 多透视表 / 空表空列表 / 错误码 3010(kind:worksheet) |
| `test_delete_pivot_table.py` | 3 | flow: 删其一保留其它 / 删唯一→空 / 错误码 3010(kind:pivotTable) |

**合计 11 case。**

## 夹具 `fixtures/pivot_table_e2e/pivot.xlsx`

- `Data`（activeSheet）—— A1:C5 数据源（Region/Product/Amount + 4 行），E1/E20 为透视表落点。
- `Blank` —— 空表，供 get:pivotTables 验证空列表。

每个 case 打开独立工作副本，互不污染。

## 错误码现实

| 触发 | 错误码（oasp#17 权威） |
|------|-------|
| 非法 sourceAddress | **`3009`** RANGE_INVALID |
| 透视表不存在 | **`3010`** ELEMENT_NOT_FOUND（`details.kind:"pivotTable"`） |
| 不存在的 worksheet | **`3010`** ELEMENT_NOT_FOUND（`details.kind:"worksheet"`） |
| 平台不支持（如 Excel for web 限制） | **`3016`** API_NOT_SUPPORTED（`details.requiredApiSet`） |
| sourceAddress·targetAddress·pivotTableName 空串（Zod `min(1)` 失败） | **`4000`** VALIDATION_ERROR |

> oasp#17 已定案上述权威码（规范层 MUST，不得降级 `3000`；旧 `5008`/`5010`/`5002` 退役为
> `3010(kind:pivotTable)`/`3016`/`3009`）。Add-In 接线（office-editor4ai#80）前真机仍返
> `3000` 兜底，e2e 用例以 **XFAIL** 运行，接线后摘 `xfail_reason` 转正。

## 运行

```bash
uv run python manual_tests/excel/pivot_table_e2e/test_insert_pivot_table.py --test all
uv run python manual_tests/excel/pivot_table_e2e/test_get_pivot_tables.py --test all
uv run python manual_tests/excel/pivot_table_e2e/test_delete_pivot_table.py --test all
```

真机激活：首个 case 手动点「加载项 → TFEditor4Office」，后续 case 自动重连。delete 用
`ExcelCase.flow`（见 `e2e_case.py`，#34 引入）。
