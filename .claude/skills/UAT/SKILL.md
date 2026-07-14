---
name: uat
description: 执行 Office4AI MCP Server 注册级验收测试（tools, resources）
argument-hint: <可选：场景文件名 word-tools | ppt-tools | excel-tools | resources | tool-convergence | resource-convergence，留空则验收所有场景>
---

# UAT - User Acceptance Testing Skill

## Instructions

### 角色定位

你是一名 QA 测试工程师，负责对 Office4AI MCP Server 执行注册级验收。通过 Playwright MCP 操控 MCP Inspector 界面，验证工具和资源的注册、参数定义是否符合预期。

**不做真实调用**——工具的功能正确性由 `manual_test` 体系保证，本 UAT 仅验证 MCP 层面的注册完整性。

### 验收手段

- **Playwright MCP**：Claude Code 驱动浏览器自动化
- **MCP Inspector**：可视化界面，用户实时观察验收过程
- **验收范围**：注册存在性 + 参数 schema 正确性，不执行 Run Tool

### 前置条件检查

1. **MCP Inspector 已启动**：用户已启动 MCP Inspector 并连接到 Office4AI MCP Server
2. **连接成功**：MCP Inspector 左侧显示已连接状态

如未满足，**停止测试**并提示用户准备。详见 `resources/prerequisites.md`。

### 执行协议

1. **加载场景**：读取 `resources/scenarios/<scenario>.md`
2. **通过 Playwright 打开 MCP Inspector**（默认 `http://localhost:6274`）
3. **逐项验证**：
   - 在 MCP Inspector 的 Tools / Resources 列表中逐一确认注册项
   - 点击工具/资源查看参数 schema
   - 截图记录，对比预期，标记 PASS / FAIL
4. **输出报告**

### 验收维度

#### 工具验收（Tool UAT）
- **注册存在性**：工具是否出现在 Tools 列表中
- **名称与描述**：name 和 description 是否正确、可理解
- **参数 Schema**：必填/选填参数是否齐全，类型是否正确，description 是否清晰
- **document_uri**：所有工具是否都包含 `document_uri` 必填参数

#### Resource 验收（Resource UAT）
- **注册存在性**：资源是否出现在 Resources 列表中
- **URI 格式**：URI 是否符合 `window://office4ai[/{type}/{doc_id}]` 或 `skill://<host>/<leaf>` 格式
- **元数据**：name、description、mimeType 是否正确；skill 根须带 `_meta.source=resources`

#### 动态收敛验收（Convergence UAT，W4a/W4b）
- **工具收敛（W4a/#63）**：平台工具（word/ppt/excel）**断连时隐藏、连接对应 Add-In 后暴露**；常驻工具（`office_run_script`）恒在。见 `tool-convergence` 场景（两轮对照）。
- **资源收敛（W4b/#64,#66）**：per-file 窗口 `window://office4ai/{type}/{doc_id}` **随连接出现/断开注销**；内容**实时反映所连文件真实状态**（须重新拉取最新值对照）。见 `resource-convergence` 场景。
- **通知**：连接变化应广播 `tools/list_changed` / `resources/list_changed`；写操作/活动切换发 `resources/updated`。
- ⚠️ 「无 Add-In 时看不到平台工具/per-file 窗口」是**正确收敛，不是缺陷**——两轮组合（先断后连）才能证明设计预期。

### 报告格式

```
## UAT 报告 - [场景名称]
日期：YYYY-MM-DD
环境：MCP Inspector → Office4AI MCP Server

### 测试结果摘要
- 总验证项：N
- 通过：N ✅
- 失败：N ❌

### 验证详情
| # | 验证项 | 结果 | 备注 |
|---|--------|------|------|
| 1 | ...    | ✅   |      |

### 失败项详情
#### [项目名称]
- **预期**：...
- **实际**：...
- **截图**：[MCP Inspector 截图]
```

## Prompt

请执行 UAT 验收。

首先确认：MCP Inspector 是否已启动并连接到 Office4AI MCP Server？

确认后，按以下逻辑加载场景：

- 如果指定了场景参数：加载 `resources/scenarios/<场景名>.md` 执行验收
- 如果未指定（参数为空）：依次加载 `resources/scenarios/` 下所有场景文件，逐一验收

$ARGUMENTS

通过 Playwright 打开 MCP Inspector 界面，逐项验证注册正确性。每个关键步骤截图，供用户实时确认。
