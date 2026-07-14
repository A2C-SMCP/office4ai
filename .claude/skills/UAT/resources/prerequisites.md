# UAT 前置条件

## 环境要求

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

## MCP Inspector 操作指引

### 查看 Tool 列表

1. 点击左侧 **Tools** 标签页
2. 工具列表中显示所有已注册工具
3. 点击某个工具可查看其参数 schema（inputSchema）

### 查看 Resource 列表

1. 点击左侧 **Resources** 标签页
2. 资源列表中显示所有已注册资源及其 URI、name、mimeType
