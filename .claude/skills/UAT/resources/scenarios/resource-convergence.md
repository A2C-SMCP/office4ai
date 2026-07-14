# UAT 场景：Desktop 资源收敛与实时内容验收 (resource-convergence)

## 测试目标

验证 **W4b per-file 窗口（#64 / #66）** 的设计预期：Desktop 窗口资源**随对应 Add-In 连接动态出现/收敛**，且资源**内容实时反映所连文件的真实状态**——需在连接前后**重新拉取（re-read）资源最新值**对照。与 `tool-convergence`（工具侧）互为对偶，是独立用例。

## 背景（W4b-1 #64 / W4b-3 #66 / 订阅 #3）

- **静态常驻**：根 `window://office4ai` + `skill://` SKILL 根（恒注册）。
- **动态 per-file**：`window://office4ai/{word|ppt|excel}/{doc_id}` —— 对应 Add-In 连接后**动态注册**，断开后注销。
- 连接/断开广播 `notifications/resources/list_changed`；活动文件切换/写操作后对受影响窗口发 `notifications/resources/updated`。
- **内容是"活的"**：读取 per-file 窗口会向 Add-In 拉取该文件当前摘要（word 字数 / ppt 幻灯片 / excel 工作簿）；根窗口内容随连接列出已连文件。

## Round 1 — 断连态（不连接任何 Office Add-In）

| # | 验证项 | 预期 | 判定 |
|---|--------|------|------|
| RC1-1 | Resources 列表存在根 `window://office4ai` | 出现 | 缺失即 FAIL |
| RC1-2 | **读取**根 `window://office4ai` 内容 | 含"暂无文档连接，等待 Office Add-In 接入" | 否则 FAIL |
| RC1-3 | 存在 3 个 `skill://com.a2c-smcp.office4ai/*` SKILL 根（带 `_meta.source=resources`） | 出现 | 见 `resources` 场景 |
| RC1-4 | `window://office4ai/{word\|ppt\|excel}/{doc_id}` per-file 窗口 | **不出现** | 出现即 FAIL |

## Round 2 — 连接态（连接 Add-In 后动态出现 + 实时内容）

**前置**：引导用户连接对应 Add-In（需 office-editor4ai 配合，逐步引导）：
1. 启动对应 Add-In dev server（`pnpm dev:word|dev:ppt|dev:excel`）
2. 打开对应 Office 应用 + 一个**内容可辨识**的文档（如 Excel 里 A1 填 `UAT-CHECK`），sideload，taskpane 握手成功
3. 回到 Inspector **刷新 Resources 列表**（连接触发 `resources/list_changed`）

### 列表收敛

| # | 验证项 | 预期 | 判定 |
|---|--------|------|------|
| RC2-1 | 连接后对应 `window://office4ai/{type}/{doc_id}` per-file 窗口**出现**，name 形如 `TYPE · 文件名` | 出现 | 缺失即 FAIL |
| RC2-2 | 收到 `notifications/resources/list_changed` | 收到 | 未收到记 ⚠️ |
| RC2-3 | 未连接类型的 per-file 窗口仍不出现 | 隐藏 | 出现即 FAIL |

### 实时内容验收（核心：重新拉取最新值）

| # | 验证项 | 预期 | 判定 |
|---|--------|------|------|
| RC2-4 | **重新读取**根 `window://office4ai` | 内容从"暂无连接"变为**逐条列出**已连文件（`- window://…/{type}/{doc} — TYPE · 文件名`） | 未更新即 FAIL |
| RC2-5 | **读取** per-file 窗口内容 | 渲染该文件**真实摘要**（word=字数/可见文本；ppt=幻灯片信息；excel=工作簿/工作表摘要），与文档实际相符（如能读到 `UAT-CHECK` 所在表） | 不符即 FAIL |
| RC2-6 | 在文档中**改动**（切换活动幻灯片/表、或写入），稍候**再次 re-read** per-file 窗口 | 内容**随之更新**（配合 `resources/updated` 通知）；体现"实时拉取最新值" | 陈旧不变即 FAIL |
| RC2-7 | 断开 Add-In 后 → per-file 窗口**注销消失**，根窗口回到"暂无连接" | 收敛回退 | 仍在即 FAIL |

## 组合判定（设计预期达成）

- **Round 1「无 per-file 窗口 + 根显示暂无连接」 + Round 2「连上即出现、内容实时反映真实文档、改动后 re-read 更新、断开即收敛」** 同时成立 → W4b Desktop 资源收敛符合设计预期 ✅
- 特别强调 RC2-4/5/6：**必须重新拉取资源最新值**（非缓存首值）对照文档真实状态，才算通过。

> 关联：工具侧对偶收敛见 `tool-convergence`；静态资源注册（根 + skill://）逐项见 `resources`。
