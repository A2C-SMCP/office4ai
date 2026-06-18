# Excel 写操作「最小返回」取舍决策

> **状态**: ✅ 已决策（Plan A）
> **创建日期**: 2026-06-18
> **决策范围**: office4ai `/excel` 命名空间全部写操作工具（OASP 0.3.0 events-excel.md）
> **关联**: milestone #3 / issue #26 `[E2E]` / Tracking #27
> **参与方**: office4ai 维护者、OASP 协议维护者、office-editor4ai（Add-In）

---

## 背景

OASP 0.3.0 `events-excel.md` 定义的 Excel 写操作刻意采用**最小返回**——只回操作锚点，
不回写入后的数据快照。典型如：

| 事件 | 返回 data | 不返回 |
|------|-----------|--------|
| `excel:set:range` | `{address}` | 写入后的 values 快照 |
| `excel:set:formula` | `{address}` | 计算后的单元格值 |
| `excel:set:rangeFormat` | `{address}` | 应用后的完整 RangeFormatInfo |
| `excel:merge:cells` | `{address}` | 合并后的区域形状 |
| `excel:add:tableRow` | `{tableId}` | 追加后的表格行数 / 新行内容 |
| `excel:sort:table` | `{sorted}` | 排序后的行序 |
| `excel:insert:chart` | `{name}` | 图表的完整属性 |

这与本项目的一条工具设计取向存在**张力**：返回值应携带足够上下文，方便上层 Agent
免去额外读取、提升单步可用性。issue #26 要求对此给出书面结论。

---

## 张力的两端

- **可组合性（最小动作单元）**：`docs/office4ai_dev_plan.md` 明确把工具定位为
  「最小但可组合的动作单元」——写归写、读归读，组合由 Agent 编排。
- **单步上下文充分性**：让写操作直接回填写入后状态，Agent 一跳拿到结果，无需再读。

---

## 方案对比

### 方案 A：保持最小返回 + 显式文档化（**已采纳**）

写工具透传 Add-In 的最小返回；在 `format_result()`（基类单一源）与本决策文档中
显式说明「写后如需确认结果，再调一次对应读工具」。可组合性由独立读工具承载。

- ✅ 与 OASP 协议、office-editor4ai（Add-In）现状**完全一致**，零协议改动。
- ✅ 契合 dev_plan「最小但可组合」定位；写/读职责清晰，不耦合。
- ✅ 无额外延迟、无额外失败面；写操作语义单一、幂等边界清楚。
- ✅ Add-In 侧无需在每个写路径后做一次代价不定的全量读（大区域/大表格读回可能很贵）。
- ⚠️ Agent 若要写后状态，需自行多调一次读工具（由文档与工具描述引导）。

### 方案 B：Server 侧补读回填上下文（未采纳）

office4ai Server 在写操作成功后自动多发一跳读 RPC，把写入后快照塞进返回值。

- ✅ Agent 一跳即得写后状态。
- ❌ 每次写多一次 round-trip，**延迟翻倍**；大范围/大表格读回开销不可控。
- ❌ 与 Add-In 的最小返回设计**分叉**：office4ai 单方面加语义，三端不再同构。
- ❌ 放大失败面（写成功但回读失败时返回语义含糊：算成功还是失败？）。
- ❌ 读回时机与写之间存在并发窗口（其他客户端可能已改动），快照未必反映本次写。

---

## 决策

**采纳方案 A**：Excel 写操作保持协议定义的最小返回，不在 Server 侧补读。

**理由**：可组合性在本架构里由「独立读工具 + Agent 编排」承载，而非由「胖返回值」承载；
这正是 dev_plan「最小但可组合」的本意。方案 B 的一跳便利换来延迟翻倍、三端分叉、失败面
放大与快照时序问题，得不偿失。最小返回也保证 office4ai / Add-In / OASP **三端同构**。

### 写后确认的推荐用法（write-then-read）

| 写工具 | 确认用读工具 |
|--------|-------------|
| `excel_set_range` / `excel_set_formula` / `excel_clear_range` | `excel_get_range` |
| `excel_set_range_format` / `excel_add_conditional_format` | `excel_get_range_format` |
| `excel_merge_cells` / `excel_unmerge_cells` | `excel_get_range`（读回 address/values 确认合并后形状；合并状态不在 RangeFormatInfo 内） |
| `excel_add_table_row` / `excel_delete_table_row` / `excel_sort_table` | `excel_get_table` |
| `excel_insert_chart` / `excel_update_chart` | `excel_get_charts` |
| `excel_insert_pivot_table` | `excel_get_pivot_tables` |
| `excel_add_worksheet` / `excel_rename_worksheet` | `excel_get_worksheets` |

---

## 留痕

- **office4ai**：
  - 本决策文档（authoritative）。
  - `office4ai/a2c_smcp/tools/base.py` 基类 `format_result()` docstring —— 在写工具
    继承的**单一源**处说明最小返回约定（避免逐工具复制），honoring Plan A 的
    「format_result() 显式说明」。
  - `docs/mcp_server_encapsulation_spec.md` —— 架构决策条目登记。
- **OASP（A2C-SMCP/oasp-protocol）**：建议在 `events-excel.md` 写事件段落补一句
  「写操作按设计返回最小锚点，写后状态请用对应 get 事件查询」，作为 #27 协议侧
  Draft→Stable 转正时的同步项（需在协议仓库走 PR，本仓库不直接改）。

---

## 备注

本决策仅约束 Excel 写操作的返回**取舍**，不改变任何协议字段或 wire 形态——纯文档/约定
层面，零向后兼容风险。读工具的返回保持原样（携带充分上下文，本就是读工具的职责）。
