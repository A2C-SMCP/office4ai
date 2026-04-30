# PPT Chart MCP Tools — 手工验收清单 (OASP /ppt Draft, v0.2.0)

> 配套 Issue [#9](https://github.com/JIAQIA/office4ai/issues/9)。
>
> **背景**：PowerPoint Office.js 不暴露图表创建/更新 API（[office-js#5463](https://github.com/OfficeDev/office-js/issues/5463)），
> 因此 OASP `/ppt` 的 3 个 chart 事件由 office4ai Server 端使用 `python-pptx` 直接改写
> .pptx OOXML — **不下发 Socket.IO 事件给 Add-In**。完成后 MCP 触发 `resource_updated`
> 通知 `window://office4ai/ppt` 订阅者，AI / 用户需在 PowerPoint 中重新打开文档以查看效果。

---

## 自动场景（已由 `manual_tests/ppt/test_chart_e2e.py` 覆盖）

```bash
uv run python manual_tests/ppt/test_chart_e2e.py
```

| 编号 | 场景 | 通过判据 |
|------|------|----------|
| A.1 | `ppt_insert_chart` 柱形图 | `success=True`、返回 `elementId=chart-N`、`requiresReload=True` |
| A.2 | `ppt_get_chart` 回读 | `chartType`/`categories`/`series`/`title` 与插入参数一致 |
| A.3 | `ppt_update_chart` 仅修改标题（同 variant） | `updatedFields` 含 `title`；回读标题为新值 |
| A.4 | `ppt_insert_chart` 散点图 | discriminated union 路由到 `ScatterChartData`；返回 `elementId` |
| A.5 | `ppt_update_chart` 跨 variant（Scatter → Line） | 必须显式提供 `categories`+`series`；返回新 `elementId`（python-pptx 删旧加新） |
| B.1 | `3015 INVALID_CHART_DATA` — categorical 维度不匹配 | `success=False`、`error` 含 `"3015"`；不触发 `resource_updated` |
| B.2 | `3015 INVALID_CHART_DATA` — scatter 非有限值 | `success=False`、`error` 含 `"3015"` |
| B.3 | `3010 ELEMENT_NOT_FOUND` — 不存在的 `chart-N` | `success=False`、`error` 含 `"3010"` |
| C   | OOXML 双重验证 | `python-pptx` 读取 .pptx 后 `chart_count == 2`、标题含「修订版」 |
| D   | `notify_resource_updated` 调用记账 | 写操作精确触发 4 次；错误路径不触发；URIs = `[/ppt, /]` |

---

## 视觉验收清单（必须在 PowerPoint 中肉眼检查）

> 自动脚本完成后，工作副本保留在 `manual_tests/.test_working/ppt_chart_v0_2_0/`。
> 用 PowerPoint 打开最新一份 `charts_demo_*.pptx`，逐项打勾：

- [ ] **A.1 视觉 — 柱形图**：第 1 张幻灯片显示 `ColumnClustered` 图表
  - 标题为「2026 年度业绩（修订版）」
  - 4 个分组（Q1/Q2/Q3/Q4），每组两根柱子（营收 / 成本）
  - 图例可见（默认 `showLegend=True`）；数值标签隐藏（`showDataLabels=False`）
- [ ] **A.4 视觉 — 散点图初版**（A.5 之前的状态）：在 A.5 跨 variant 切换前，第 2 张曾是散点图（5 个数据点，标题「广告投入 vs 销售额」）。运行脚本后该图已被替换为折线图。
- [ ] **A.5 视觉 — 跨 variant 切换为折线图**：第 2 张幻灯片显示 `Line` 图表
  - 4 个 X 轴标签（Jan/Feb/Mar/Apr）
  - 1 条折线（趋势 10 → 25 → 30 → 28）
- [ ] **几何位置正确**：两个图表均靠左上对齐（left=60pt, top=60pt），尺寸 540×360pt
- [ ] **保存兼容**：在 PowerPoint 中保存（`Cmd+S`）后再次打开，所有图表渲染正常（验证 OOXML 写入未破坏 ZIP 包结构）
- [ ] **AI 业务闭环**：通过 MCP 客户端列出工具，确认 `ppt_insert_chart` / `ppt_get_chart` / `ppt_update_chart` 出现在 `tools/list` 响应中；description 含 `(OASP /ppt Draft)`
- [ ] **save() 警告生效**：在 Add-In 已连接、文档有未保存改动时调用 `ppt_insert_chart`，确认 LLM 阅读到 description 中的 `Add-In MUST call save()` 提示并优先调用现有的保存路径（手工模拟 — 没有真实保存 → AI 应主动告警或先 save）

---

## 已知限制

1. **Add-In 持有文档时拒绝写入（fail-loud 防御）**：经实验验证，PowerPoint 打开
   .pptx 后任何 Add-In 触发的 save 都会以内存模型整体覆盖磁盘——Server-OOXML
   写入的 chart 被静默销毁。本 PR 在 `ppt_insert_chart` / `ppt_update_chart`
   入口加了硬防御：当 `workspace.get_document_status(uri) == CONNECTED` 时返回
   `3003 DOCUMENT_READ_ONLY` 并附明确文案，让 LLM 提示用户「先关闭文档」。
   `ppt_get_chart` 不受影响（只读，最坏情况是返回 stale 数据）。
2. **Add-In 自动重载未实装**：当前实现仅返回 `requiresReload: true` 标志 +
   触发 MCP 资源更新通知。Add-In 端的"自动重新打开 .pptx"机制需要新的
   OASP 事件（如 `ppt:notify:reload`），属下一阶段工作。
3. **延迟 >1s**：每次 chart 操作都会完整解析 + 序列化整个 .pptx OOXML。中等大小文档（10–20 页 + 嵌入资源）单次操作通常耗时 1–3 秒。LLM 应在 description 引导下避免短时间内大量串行调用。
4. **并发串行化**：同一 `documentUri` 的并发 chart 调用通过 `DocumentLockManager` 串行化；不同文档的并发调用并行不阻塞。
5. **`update:chart` 必须提供 `chartType`**：作为 Pydantic `discriminator`。AI 应先 `ppt_get_chart` 读出当前 `chartType` 再回填到 `ppt_update_chart` 的入参。
6. **跨 variant 切换会刷新 elementId**：`Scatter → Line` 等跨 variant 操作内部走"删旧加新"路径，返回的新 `elementId` 与原 `elementId` 不同；AI 必须采纳响应中的新 ID。

## 冲突实验数据（`--mode conflict`，2026-04-30 实测）

```
Phase A — Server 写 chart → Add-In 加 image → save
  [起点]                    s0=0c/0p
  [A.1 Server chart]        s0=1c/0p     ← 磁盘有 chart
  [A.2 Add-In image (内存)]  s0=1c/0p     ← PowerPoint 内存独立，磁盘未变
  [A.3 AFTER save]          s0=0c/1p ❗  ← chart 被静默覆盖

Phase B — Add-In image+save → Server chart → Add-In image → save
  [B.1 add-in img + save]   s1=0c/1p
  [B.2 Server chart]        s1=1c/1p     ← image+chart 共存
  [B.3 add-in img (内存)]    s1=1c/1p     ← 磁盘未变
  [B.4 AFTER save]          s1=0c/2p ❗  ← chart 再次被静默覆盖

Phase C — 关闭后再插，验证恢复路径
  [C.0 工具防御]             3003 拒绝   ← Add-In 仍连着，不让插
  [user closes file]
  [C.2 Add-In disconnect]   poll 监测到 → 工具放行
  [C.3 Server chart slide 2] s2=1c/0p   ← 关闭后写盘成功
  [C.4 AppleScript reopen]
  [C.6 disk truth]          s2=1c/0p ✅  ← chart 在 PowerPoint reload 后仍在
```

`c` = chart, `p` = picture. 完整时间线见 `manual_tests/.test_working/` 下脚本输出。

**结论**：PowerPoint for Mac 的设计就是「打开后磁盘是陈旧快照、save 是
dump 而非 merge」。没有外部文件感知、没有冲突提示、没有 merge。唯一安全的
Server-OOXML 写入窗口是「Add-In 未持有此文档」——这正是我们在工具入口
做的检查。

---

## 参考

- 协议：[OASP v0.2.0 events-ppt — `/ppt` Chart 类事件](https://doc.turingfocus.cn/oasp/0.2.0/specification/events-ppt/)
- 数据结构：[OASP v0.2.0 data-structures — ChartType / ChartData](https://doc.turingfocus.cn/oasp/0.2.0/specification/data-structures/)
- 错误码：[OASP v0.2.0 error-handling — `3015 INVALID_CHART_DATA`](https://doc.turingfocus.cn/oasp/0.2.0/specification/error-handling/)
- 实现：`office4ai/environment/workspace/services/chart_engine.py`
- 工具：`office4ai/a2c_smcp/tools/ppt/{insert,get,update}_chart.py`
