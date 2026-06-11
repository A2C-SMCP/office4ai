# PPT Chart MCP Tools — 手工验收清单 (OASP /ppt, v0.3.0 双路径)

> 配套 Issue [#16](https://github.com/JIAQIA/office4ai/issues/16)（Phase 3 E2E），跟踪 [#17](https://github.com/JIAQIA/office4ai/issues/17)；
> 源 [OF4AI-21](https://turingfocus.atlassian.net/browse/OF4AI-21)。承接 v0.2.0 验收（见 [ppt_chart_v0.2.0.md](ppt_chart_v0.2.0.md)）。
>
> **0.3.0 架构转变**：Python MCP Server 从「纯中转」升级为「**具备生产能力**」。
> 图表三工具（`ppt_insert_chart` / `ppt_get_chart` / `ppt_update_chart`）按文档连接状态
> **双路径路由**：
>
> - **关闭态（DISCONNECTED）→ 路径 A**：Server 用 `python-pptx` 离线读写盘（0.2.0 行为，不变）。
> - **打开态（CONNECTED）→ 路径 B**：Server 驱动 Add-In 的两条**通用 OOXML 搬运事件**
>   （`ppt:get:slideOoxml` / `ppt:insert:slidesOoxml`，整页 round-trip）。所有图表 OOXML
>   仍在 Server 端 `python-pptx` 内存完成；Add-In 只导出 / 应用 / 替换 / 移动整页，不碰图表语义。
>
> **守卫语义翻转（[#15](https://github.com/JIAQIA/office4ai/issues/15)）**：0.2.0 在 CONNECTED 时
> 「上抛 3003 拒绝写入」；0.3.0 改为「路由到 path B」。path B 不可用时（Add-In 搬运事件未发布 /
> 超时 / `3016 API_NOT_SUPPORTED` 不支持）才**反应式降级**：写回退到带 `3003` 前缀的「先关闭
> 文档」文案，读回退到磁盘。

---

## 自动场景

### 1. 离线回归（路径 A，CI 保护）— 不变

```bash
uv run python manual_tests/ppt/test_chart_e2e.py            # 或 --mode offline
```

覆盖矩阵（A.1–D）沿用 [ppt_chart_v0.2.0.md](ppt_chart_v0.2.0.md)，无变化。

### 2. 打开态路径 B（模拟 Add-In）— #16 新增

```bash
uv run python manual_tests/ppt/test_chart_e2e.py --mode pathb
```

真实 Add-In（office-editor4ai Task 2 — [#38](https://github.com/JIAQIA/office-editor4ai/issues/38) /
[#39](https://github.com/JIAQIA/office-editor4ai/issues/39)）尚未发布，本模式用进程内 `FakeAddIn`
忠实模拟两条 0.3.0 搬运事件，把 #15 路由器的整条 path-B 编排跑通：

| 编号 | 场景 | 通过判据 |
|------|------|----------|
| P.1 | insert 路径 B（整页 round-trip） | `success`；事件序列 `[ppt:get:slideOoxml, ppt:insert:slidesOoxml]`；`requiresReload=False`；live 单页含 1 个图表；**磁盘副本图表数=0**（path B 不碰盘） |
| P.2 | get 路径 B（读 live 单页） | 回读 `chartType`/`categories`/`title` 与 P.1 插入一致（读自 live 包，非磁盘） |
| P.3 | update 路径 B（就地改 live） | `updatedFields` 含 `title`；live 包标题更新为新值；`requiresReload=False`；写通知 `/ppt` 订阅者 |
| P.4 | 反应式写降级（搬运事件超时） | `success=False`、`error` 含 `3003`；**磁盘字节零变化** |
| P.4b | 反应式写降级（Add-In 回 `3016` 不支持） | 同 P.4：回退 `3003`、磁盘字节零变化（补全 timeout + 3016 降级矩阵） |
| P.5 | 反应式读降级（搬运事件超时） | get 回退读盘，返回磁盘上的图表 |

> `FakeAddIn` 的 `ppt:insert:slidesOoxml` 处理器断言路由器回传的 `replaceSlideId` 与导出时的
> 不透明 `slideId` 一致——验证 #15 的就地替换契约（防重复页）。

---

### 3. 真机联调验收（path B live）— office-editor4ai 0.3.0 已发布

office-editor4ai 0.3.0 搬运事件落地后，用真实 PowerPoint + 真实 Add-In 跑真机验收：

```bash
# macOS + PowerPoint 运行中 + office-editor4ai Add-In 可加载
uv run python manual_tests/ppt/test_chart_e2e.py --mode pathb-live
```

与 `--mode pathb` 同一份 Server 代码，区别仅在搬运事件打到真机 Add-In 而非进程内
`FakeAddIn`。脚本分段断言：

| 段 | 场景 | 通过判据 |
|----|------|----------|
| L1 | 裸搬运 round-trip（直接验新原语） | `ppt:get:slideOoxml` 返回单页 `{slideId, base64}`；服务端内存加图后 `ppt:insert:slidesOoxml` 就地替换 → slide 0 含 1 图、总页数不变（替换非追加） |
| L2 | 三工具 CONNECTED → path B | insert/get/update 均 `success`、`requiresReload=False`；回读/改写一致 |
| L3 | masterLeak 抽测 | 连续 3 次 round-trip 后 `slideMaster` 数 ≤ 基线（spike [office-editor4ai#34](https://github.com/JIAQIA/office-editor4ai/issues/34) 已验单次 `masterLeak:0`，此处验累积） |
| L4 | 视觉验收（人工） | 双击图表为**原生可编辑对象**（可改数据/类型），非图片 |

### 4. 仍需人工逐平台抽测的边界

- [ ] **Web 边界**：在 PowerPoint Web 端重跑上面命令 —— 平台不支持所需 requirement set 时
      应回 `3016` 并反应式降级（写回 3003 / 读回退盘），而非崩溃。
- [ ] **Windows 边界**：在 PowerPoint for Windows 重跑，同上。
- [ ] **双路径切换**：同一文档先打开（path B）再关闭（path A），两次写入均落地、无静默覆盖。
- [ ] **占位符内图表**：`containedType` 过滤属 Office.js（office-editor4ai）侧，本仓不涉及
      （spike#34 已记录占位符内图表 type 报 `Placeholder`）。

---

## 真实环境冲突实验（`--mode conflict`）

`--mode conflict` 仍保留（macOS + 真实 PowerPoint + 真实 Add-In），但 Phase C 的语义已随 #15
翻转更新：CONNECTED 不再「上抛 3003 拒绝」，而是**路由到 path B → 反应式降级回退 3003**（因真实
Add-In 的搬运事件未发布）。一旦 office-editor4ai 发布搬运事件，Phase C.0 将经 path B 直接成功，
届时需把脚本中「应被拒」预期翻转为「应成功」。历史冲突实验数据见
[ppt_chart_v0.2.0.md](ppt_chart_v0.2.0.md) §「冲突实验数据」。

---

## 参考

- 协议：OASP 0.3.0 — `ppt:get:slideOoxml` / `ppt:insert:slidesOoxml` / `3016 API_NOT_SUPPORTED`
- 设计稿：[docs/discussions/ppt-chart-dual-path-design.md](../discussions/ppt-chart-dual-path-design.md)
- 路由器：`office4ai/environment/workspace/services/chart_router.py`（#15）
- 引擎：`office4ai/environment/workspace/services/chart_engine.py`（base64 单页内存读写）
- 工具：`office4ai/a2c_smcp/tools/ppt/{insert,get,update}_chart.py`
- 跨仓：office-editor4ai Task 2 — [#38](https://github.com/JIAQIA/office-editor4ai/issues/38) / [#39](https://github.com/JIAQIA/office-editor4ai/issues/39)
