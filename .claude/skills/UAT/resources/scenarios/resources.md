# UAT 场景：Resource 注册验收 (resources)

## 测试目标

验证 MCP Resource 在 MCP Inspector 中按**当前资源模型**正确注册（URI / name / mimeType / _meta）。

> ⚠️ 模型变更（W4b-1 / #64 起）：**旧的 per-type 聚合窗口 `window://office4ai/word`、`window://office4ai/ppt` 已被移除**，由**每文件一个的 per-file 动态窗口**取代（`window://office4ai/{word|ppt|excel}/{doc_id}`；excel 由 W4b-3 / #66 纳入）。请勿再期望 per-type 聚合作为静态资源存在。

## 资源模型

- **静态常驻**（无需连接文档即注册）：
  - 根索引 `window://office4ai`
  - `skill://` 能力包（S3 / #59，`resources` source 模式暴露本地 SKILL 文件夹）
- **动态 per-file**（**连接对应文档后**才出现，随连接收敛）：
  - `window://office4ai/{word|ppt|excel}/{doc_id}` —— 每个已连接文件一个独立子窗口

## 验证清单

### A. 静态常驻资源（无需连接文档）

| # | 资源 URI | 预期 name | 预期 mimeType | 关键点 |
|---|----------|-----------|---------------|--------|
| R-01 | `window://office4ai`（读取 URI 带 `?priority=&fullscreen=`） | Office 工作区 | `text/plain` | 根索引；无文档时渲染"暂无文档连接"，有文档时逐条列出 per-file 子窗口 |
| R-02 | `skill://com.a2c-smcp.office4ai/create-office-file` | create-office-file（leaf 名） | **`inode/directory`** | W1 建文件 SKILL **根（目录型）**；**须带 `_meta.source="resources"`** |
| R-03 | `skill://com.a2c-smcp.office4ai/edit-office-file` | edit-office-file | **`inode/directory`** | W2 编辑 SKILL 根；带 `_meta.source="resources"` |
| R-04 | `skill://com.a2c-smcp.office4ai/extract-template` | extract-template | **`inode/directory`** | W3 抽模板 SKILL 根；带 `_meta.source="resources"` |
| R-05 | `skill://com.a2c-smcp.office4ai/<leaf>/<rel>`（子资源） | 包内文件相对路径 | 按扩展名（`SKILL.md`/`references/*.md`→`text/markdown`、`scripts/*.py`→`text/x-python`、`.txt`→`text/plain`） | 每个 SKILL 包的 `SKILL.md` / `references/` / `scripts/` 作为兄弟子资源，**不带** `_meta` |

> skill host 默认 `com.a2c-smcp.office4ai`；可被 `OFFICE4AI_SKILLS_ROOT` 环境变量整体覆盖（则 leaf 集合随之变化）。R-02～R-04 对应包内 `office/skills/` 的 3 个生产 SKILL；每包 ≈ 8 个资源（1 目录根 + `SKILL.md` + 2~3 `references/` + 4 `scripts/`）。
>
> **断连态静态资源基线（实测）**：共 **25 个** = 1 根 window + 3 SKILL 根 + 21 SKILL 子资源；**无任何 per-file 窗口**。

### B. 动态 per-file 窗口资源（**需先连接对应文档**）

> 本节只核对 per-file 窗口的**注册**（URI / name / mimeType）；其**动态收敛行为 + 实时内容**（连接前后 re-read 对照、`resources/list_changed`、`resources/updated`）见 `resource-convergence` 场景。

先分别接入一个 Word / PPT / Excel 文档（Add-In 握手成功），再在 Resources 列表中核对：

| # | 资源 URI 模式 | 预期 name | 预期 mimeType | 关键点 |
|---|---------------|-----------|---------------|--------|
| R-W | `window://office4ai/word/{doc_id}` | `WORD · {文件名}` | `text/plain` | 单个 Word 文件窗口 |
| R-P | `window://office4ai/ppt/{doc_id}` | `PPT · {文件名}` | `text/plain` | 单个 PPT 文件窗口 |
| R-E | `window://office4ai/excel/{doc_id}` | `EXCEL · {文件名}` | `text/plain` | 单个 Excel 文件窗口（#66 纳入） |

> `doc_id` = 可读文件名 + sha1 短摘要（同一文件跨刷新 URI 稳定）。**未连接对应类型文档时，该类型 per-file 窗口不应出现**；根 `window://office4ai` 也不再列它。

## 逐项验证要点

### 静态资源（A）
1. **存在性**：`window://office4ai` 与 3 个 `skill://…` SKILL 根均出现在 Resources 列表
2. **skill 根 `_meta`**：3 个 SKILL 根**必须**带 `_meta.source="resources"`（缺失则会被当作子资源、不注册为根）
3. **skill 子资源**：SKILL 根下的 `SKILL.md`/`scripts/` 等以 `skill://<host>/<leaf>/<rel>` 形式出现，**不带** `_meta`
4. **URI / name / mimeType**：与上表一致

### 动态资源（B）
5. **连接前**：三类 per-file 窗口都不应出现（仅根 + skill）
6. **连接后**：连上哪类文档，就出现哪类 `window://office4ai/{type}/{doc_id}`，name 形如 `TYPE · 文件名`，mimeType `text/plain`
7. **反向（陈旧回归）**：确认**不存在**旧的 `window://office4ai/word`、`window://office4ai/ppt` per-type 聚合静态资源
