# UAT 场景：动态工具收敛验收 (tool-convergence)

## 测试目标

验证 **W4a 动态工具收敛（#63）** 的设计预期：平台工具（word/ppt/excel）**随对应 Add-In 连接动态出现/收敛**，常驻工具恒在。通过「断连态」+「连接态」两轮组合验收——**单轮不足以证明，必须两轮对照**。

## 背景（W4a / #63）

服务端 `_is_tool_available` 按连接状态过滤 `list_tools`：

- `requires_connection=False`（`category="authoring"` 常驻工具，如 `office_run_script`）→ **恒暴露**
- `requires_connection=True`（word/ppt/excel 平台工具）→ **仅当对应 namespace 有 Add-In 连接时才暴露**（无连接 → 移除；仅连 Excel → 仍隐藏 word/ppt）
- 连接状态变化时广播 `notifications/tools/list_changed`

> 因此「无 Add-In 时看不到平台工具」是**正确收敛**，不是缺陷。旧 `prerequisites.md`「不需要 Add-In 连接」的说法对平台工具已失效。

## Round 1 — 断连态（不连接任何 Office Add-In）

**前置**：MCP Inspector 已连到 office4ai server；**未**连接任何 Office Add-In（无 word/ppt/excel 握手）。

| # | 验证项 | 预期 | 判定 |
|---|--------|------|------|
| TC1-1 | Tools 列表中的 `word_*` 工具（抽样 `word_insert_text` 等） | **不出现** | 出现即 FAIL |
| TC1-2 | Tools 列表中的 `ppt_*` 工具（抽样 `ppt_insert_shape` 等） | **不出现** | 出现即 FAIL |
| TC1-3 | Tools 列表中的 `excel_*` 工具（抽样 `excel_set_range` 等） | **不出现** | 出现即 FAIL |
| TC1-4 | 常驻工具 `office_run_script`（authoring，`requires_connection=False`） | **出现** | 缺失即 FAIL |

> Round 1 通过标准：**平台工具全不可见 + 常驻工具可见**。

## Round 2 — 连接态（逐平台连接 Add-In）

**前置**：在 Round 1 基础上，**引导用户连接对应 Add-In**。每个平台按下述引导完成（需 office-editor4ai 配合，引导用户逐步做）：

1. 启动对应 Add-In dev server（office-editor4ai）：Word `pnpm dev:word` / PPT `pnpm dev:ppt` / Excel `pnpm dev:excel`
2. 打开对应 Office 应用（Word/PowerPoint/Excel）+ 一个文档，sideload 加载项，taskpane 连上（握手成功）
3. 回到 MCP Inspector **刷新 / 重新拉取** Tools 列表（连接会触发 `tools/list_changed`，通知区应出现）

| # | 验证项 | 预期 | 判定 |
|---|--------|------|------|
| TC2-1 | 连接 Word Add-In 后 → 25 个 `word_*` 工具**出现**（逐项见 `word-tools` 场景） | 出现 | 缺失即 FAIL |
| TC2-2 | 连接 PPT Add-In 后 → 24 个 `ppt_*` 工具**出现**（见 `ppt-tools` 场景） | 出现 | 缺失即 FAIL |
| TC2-3 | 连接 Excel Add-In 后 → 37 个 `excel_*` 工具**出现**（见 `excel-tools` 场景） | 出现 | 缺失即 FAIL |
| TC2-4 | **只连部分平台时**，未连平台的工具仍**不出现**（如仅连 Excel → word/ppt 仍隐） | 未连平台隐藏 | 出现即 FAIL |
| TC2-5 | 连接瞬间收到 `notifications/tools/list_changed`（Inspector 通知区） | 收到 | 未收到记 ⚠️ |
| TC2-6 | 断开某 Add-In 后 → 对应平台工具**再次消失**（收敛回退） | 消失 | 仍在即 FAIL |

## 组合判定（设计预期达成）

- **Round 1「平台工具不暴露」 + Round 2「连上即暴露、断开即收敛」** 同时成立 → W4a 动态收敛符合设计预期 ✅
- 任一轮不符 → FAIL，记录实际现象与截图。

> 关联：资源侧的对偶收敛（per-file 窗口随连接出现/内容实时刷新）见 `resource-convergence` 场景。
