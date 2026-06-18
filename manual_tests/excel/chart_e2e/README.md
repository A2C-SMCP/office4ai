# Excel Chart 操作 E2E（#34）

覆盖 OASP 0.3.0 `/excel` 命名空间 **#23 Chart** 切片的 4 个事件。图表为视觉对象，以**协议层
get:charts 回读**为主验证（openpyxl `chart_count` 仅 best-effort，原生图表回读受限）。

## 事件与 wire 形态（以 AddIn `chart.ts` 实测为准）

| 事件 | 请求参数 | 返回 data |
|------|---------|-----------|
| `excel:insert:chart` | `sourceAddress, chartType, title?, position?{top,left,width,height}, worksheetName?` | `{name}` |
| `excel:get:charts` | `worksheetName?` | `{charts:[{name, chartType, title, top, left, width, height}]}` |
| `excel:update:chart` | `chartName, properties{title?,chartType?,sourceAddress?,position?}, worksheetName?` | `{name}` |
| `excel:delete:chart` | `chartName, worksheetName?` | **void** |

> 图表名由 Excel 自动生成（中文环境如「图表 1」），**无法在 insert 时指定**。故 update/delete
> 用 `ExcelCase.flow` 多步流：先 insert 拿 name，再操作，最后 get:charts 回读核对。

## 测试文件（4 个）

| 文件 | 用例数 | 覆盖 |
|------|-------|------|
| `test_insert_chart.py` | 5 | ColumnClustered+title / Line / Pie / XYScatter+position / 错误码 3000 |
| `test_get_charts.py` | 4 | 单图表回读(chartType/title/位置) / 多图表 / 空表空列表 / 错误码 3000 |
| `test_update_chart.py` | 4 | flow: update title / chartType / position（回读核对）/ 错误码 3000 |
| `test_delete_chart.py` | 3 | flow: 删其一保留其它 / 删唯一→空 / 错误码 3000 |

**合计 16 case。**

## 夹具 `fixtures/chart_e2e/chart.xlsx`

- `Data`（activeSheet）—— A1:C4 数值网格（Month/Sales/Cost + 3 行），作所有图表数据源。
- `Blank` —— 空表，供 get:charts 验证空列表。

每个 case 打开独立工作副本，互不污染。

## 错误码现实

| 触发 | 错误码 |
|------|-------|
| 非法 chartType（非空但 Office.js 拒绝枚举）/ 图表不存在 / 不存在的 worksheet | **`3000`** DOCUMENT_ERROR |
| chartType 空串（Zod `min(1)` 失败） | **`4000`** VALIDATION_ERROR（未单独建用例，与 #33 同理） |

> 旧 DoD 的 `4002`（无效 chartType）/ `5007`（图表不存在）均为 dead code：`excelErrorCode()` 仅把
> Zod 失败映射 `4000`，其余 Office.js 运行期异常（含非法图表类型枚举）→ `3000`。

## flow 机制（本片引入，复用于 #35）

`ExcelCase.flow`（`async (workspace, document_uri, reader) -> bool`）取代标准「单 action +
validator」路径，用于「先建对象拿自动生成的名字，再操作，最后回读核对」这类无法用单 action
表达的场景。见 `e2e_case.py`。

## 运行

```bash
uv run python manual_tests/excel/chart_e2e/test_insert_chart.py --test all
uv run python manual_tests/excel/chart_e2e/test_get_charts.py --test all
uv run python manual_tests/excel/chart_e2e/test_update_chart.py --test all
uv run python manual_tests/excel/chart_e2e/test_delete_chart.py --test all
```

真机激活：首个 case 手动点「加载项 → TFEditor4Office」，后续 case 自动重连。视觉验收（图表
就位/类型/标题）见 `docs/manual_tests/excel_v0.3.0.md` D.6。
