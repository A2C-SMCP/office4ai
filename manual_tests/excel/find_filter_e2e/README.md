# Excel Find & Filter E2E（#36）

覆盖 OASP 0.3.0 `/excel` 命名空间 **#25 Find&Filter** 切片的 3 个事件，真机在环 + openpyxl 读盘双验证。

## 事件与 wire 形态（以 AddIn `findValues.ts` / `autoFilter.ts` 实测为准）

| 事件 | 请求参数 | 返回 data |
|------|---------|-----------|
| `excel:find:values` | `searchText, address?, worksheetName?, matchCase?=false, matchEntireCell?=false` | `{matches:[{address, value}]}` |
| `excel:set:autoFilter` | `address, criteria:[{columnIndex, filterOn, values?}], worksheetName?` | `{address}` |
| `excel:clear:autoFilter` | `worksheetName?` | **void** |

> `find:values` 的 `address` 省略时搜整个 usedRange；`matches[].address` 为含表名全地址（如 `Data!B2`）。
> `matchCase`/`matchEntireCell` 默认 false（AddIn 内部默认），不传即不区分大小写的子串匹配。

## 测试文件（2 个，覆盖 3 事件）

| 文件 | 用例数 | 覆盖 |
|------|-------|------|
| `test_find_values.py` | 5 | 默认子串 / matchCase / matchEntireCell 无命中 / 限定 address / 错误码 4000 |
| `test_auto_filter.py` | 5 | set 单列 / set 多列 / clear（ref 清空）/ 错误码 4000 / 错误码 3009 |

**合计 10 case。**

## 夹具 `fixtures/find_filter_e2e/filt.xlsx`

`Data`（activeSheet）—— A1:C5 网格，**Product 列大小写混合**（Apple/apple）以区分 matchCase：

```
Region   Product  Amount
East     Apple    100
West     apple    200
East     Banana   300
South    apple    400
```

- find "apple" 默认 → B2/B3/B5 共 3；matchCase=true → 仅小写 B3/B5 共 2
- find "App" matchEntireCell=true → 无命中（整单元格匹配排除子串）
- set:autoFilter A1:C5 → `auto_filter_ref='A1:C5'`；clear → `None`

每个 case 打开独立工作副本，互不污染。

## 错误码现实

| 触发 | 错误码（oasp#17 权威） |
|------|-------|
| 非法 address | **`3009`** RANGE_INVALID |
| 不存在的 worksheet | **`3010`** ELEMENT_NOT_FOUND（`details.kind:"worksheet"`） |
| searchText / address 空串（Zod `min(1)` 失败） | **`4000`** VALIDATION_ERROR |

> oasp#17 已定案上述权威码（规范层 MUST，不得降级 `3000`；旧 5xxx 块整体退役）。
> Add-In 接线（office-editor4ai#80）前真机仍返 `3000` 兜底，e2e 用例以 **XFAIL** 运行，
> 接线后摘 `xfail_reason` 转正。

## 运行

```bash
uv run python manual_tests/excel/find_filter_e2e/test_find_values.py --test all
uv run python manual_tests/excel/find_filter_e2e/test_auto_filter.py --test all
```

真机激活：首个 case 手动点「加载项 → TFEditor4Office」，后续 case 自动重连。
