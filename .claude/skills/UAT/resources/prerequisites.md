# UAT 前置条件

> 分两阶段：**Phase 1 注册验收**（下方「环境要求」）+ **Phase 2 功能验收**（末尾「Phase 2」节）。按执行的阶段满足对应前置。

## Phase 1 — 注册验收环境要求

### 1. Office4AI MCP Server

MCP Inspector 以 stdio 模式自动管理 MCP Server 进程，无需手动启动。

### 2. MCP Inspector

```bash
npx @anthropic-ai/mcp-inspector
```

- 默认地址：`http://localhost:6274`
- 确保已连接到 Office4AI MCP Server（左侧显示连接状态）

### 3. Office Add-In 连接（**分场景要求**，W4a/W4b 动态收敛）

⚠️ **重要更正**（旧文档「本 UAT 不需要 Add-In 连接」已失效）：自 W4a(#63)/W4b(#64,#66) 起，平台能力**随 Add-In 连接动态收敛**，因此是否需要连接 Add-In **取决于场景**：

| 类别 | 是否需要 Add-In 连接 | 涉及场景 |
|------|---------------------|---------|
| **静态项**：常驻工具（`office_run_script`）+ 静态资源（根 `window://office4ai`、`skill://`） | **不需要** | `resources` 断连态部分 |
| **平台工具**：word/ppt/excel（`requires_connection=True`） | **需要**对应 Add-In 连接才可见 | `word-tools` / `ppt-tools` / `excel-tools`（连接态）；`tool-convergence` Round 2 |
| **per-file 窗口资源**：`window://office4ai/{type}/{doc_id}` | **需要**对应 Add-In 连接才出现 | `resource-convergence` Round 2 |
| **收敛行为本身**（断连隐藏 ↔ 连接暴露） | 两轮都要（先断后连） | `tool-convergence` / `resource-convergence` |

> 连接 Add-In 需 **office-editor4ai** 配合：启动对应 dev server（`pnpm dev:word|dev:ppt|dev:excel`）→ 打开对应 Office 应用 + 文档 → sideload 加载项 → taskpane 握手成功。UAT 执行到需要连接的场景时，**引导用户逐步完成**。

## Phase 2 — 功能验收环境要求（manual_test E2E）

Phase 2 在真实 Office + Add-In 环境跑 `manual_tests/` E2E，验证功能可用。需：

1. **macOS + 对应 Office 桌面版**（Word / PowerPoint / Excel）。
2. **office-editor4ai dev server**：`cd office-editor4ai && pnpm dev:word|dev:ppt|dev:excel`（对应平台）。
3. **打开对应文档 + 激活 taskpane**：触发人打开文档 → 加载项面板点 word/ppt/excel-editor → taskpane 与 workspace 握手（`https://127.0.0.1:4443`；脚本自起的 workspace 双绑 `:3000`+`:4443`）。
4. **依赖就绪**：`poe install-dev`；部分 authoring/word 路径需 LibreOffice `soffice`（脚本会探测）。
5. **运行方式**：功能 E2E 脚本需 `dangerouslyDisableSandbox`（网络 bind + AppleScript + Office）。category runner 连接超时 30s，触发人需**及时激活 taskpane**；部分用例（get_selection / replace_selection / select_text）会**提示手动选中文本**，照做即可。

> **health 档例外**：`test_{excel|word}_e2e.py --mode health` / `test_word_table_e2e.py --mode health` 只校验事件接线，**无需 Add-In**，可先跑做接线回归。

## MCP Inspector 操作指引（Phase 1）

### 查看 Tool 列表

1. 点击左侧 **Tools** 标签页
2. 工具列表中显示所有已注册工具
3. 点击某个工具可查看其参数 schema（inputSchema）

### 查看 Resource 列表

1. 点击左侧 **Resources** 标签页
2. 资源列表中显示所有已注册资源及其 URI、name、mimeType
