---
name: uat
description: 执行 Office4AI MCP Server 验收测试 —— Phase 1 注册验收(tools/resources/收敛) + Phase 2 功能验收(manual_test E2E)
argument-hint: <可选：phase(registration|functional|all) 或 场景名(word-tools|ppt-tools|excel-tools|resources|tool-convergence|resource-convergence|word-functional|ppt-functional|excel-functional)，留空=all>
---

# UAT - User Acceptance Testing Skill

## Instructions

### 角色定位

你是一名 QA 测试工程师，对 Office4AI MCP Server 执行**两阶段验收**：

- **Phase 1 — 注册验收（Registration UAT）**：验证工具/资源的注册完整性、参数 schema、动态收敛（W4a/W4b）。**不做真实调用**。
- **Phase 2 — 功能验收（Functional UAT）**：在真实 Office + Add-In 环境下，由 `manual_tests/` E2E 用例驱动，验证工具**功能可正常跑通**（协议成功 + 文档内容双验）。**做真实调用**，需触发人协助连接、逐个走查。

「注册完整（能看到）」+「功能可用（能跑通）」组合，才是完整验收。

### 验收手段

| 阶段 | 手段 | 是否需 Add-In |
|---|---|---|
| Phase 1 | MCP Inspector（Playwright 驱动，可视）**或**程序化 MCP `list_tools`/`list_resources`（更精确，绕开 Inspector token） | 收敛/连接态场景需；静态项不需 |
| Phase 2 | `manual_tests/` E2E 脚本（触发人连接 Add-In + 按提示交互，我跑脚本收结果）；跑脚本需 `dangerouslyDisableSandbox` | 需（health 档除外） |

### 前置条件检查

- **Phase 1**：MCP Inspector 已启动并连到 server（或用程序化 client）。详见 `resources/prerequisites.md`。
- **Phase 2**：对应平台 Add-In dev server 运行（`pnpm dev:word|dev:ppt|dev:excel`）+ Office 应用打开文档 + taskpane 握手；触发人配合手动激活 + 按脚本提示操作。
- ⚠️ **W4a 收敛**：平台工具/per-file 窗口**需对应 Add-In 连接才可见**；断连态它们不出现是**正确收敛非缺陷**（见两态模型）。

如未满足对应阶段前置，**停止该阶段**并引导用户准备。

### 执行协议

**逐平台推进**（word / ppt / excel），每个平台一次连接同时覆盖两阶段：

1. **Phase 1 注册**：加载 `{plat}-tools` + `resources` + `tool-convergence`/`resource-convergence`，核对注册/schema/收敛（Inspector 或程序化）。
2. **引导连接**：触发人启动该平台 Add-In + 打开文档 + 激活 taskpane（连接态才可继续下面）。
3. **Phase 2 功能**：加载 `{plat}-functional`，按 **health → smoke（默认）→ full（可选）** 档跑 `manual_tests` 用例，逐个收结果。
4. **输出报告**：注册结论 + 功能结论合并。

### 验收维度

#### 工具注册（Tool Registration，Phase 1）
- 注册存在性、name/description、参数 schema（必填/选填/类型）、`document_uri` 必填。

#### Resource 注册（Phase 1）
- 存在性、URI 格式（`window://office4ai[/{type}/{doc_id}]` 或 `skill://<host>/<leaf>`）、name/description/mimeType；skill 根须带 `_meta.source=resources`。

#### 动态收敛（Convergence，Phase 1，W4a/W4b）
- **工具收敛**：平台工具断连隐藏↔连接暴露↔断开回退；常驻 `office_run_script` 恒在（`tool-convergence`）。
- **资源收敛**：per-file 窗口随连接出现/注销，内容实时反映真实文档（`resource-convergence`）。
- **通知**：`tools/list_changed` / `resources/list_changed` / `resources/updated`。

#### 功能可用（Functional，Phase 2）
- 每工具经 `manual_tests` E2E 真机跑通：**协议返回成功** + **文档内容双验**（python-docx/python-pptx/openpyxl 读盘）。
- 分档：**health**（事件接线，无 Add-In）/ **smoke**（关键用例，默认）/ **full**（逐类别 `--test all`）。见 `{plat}-functional` 场景。

### 报告格式

```
## UAT 报告 - [平台/场景]
日期：YYYY-MM-DD
环境：office4ai <commit> + office-editor4ai <ver>

### Phase 1 注册验收
| 维度 | 结果 | 备注 |
|---|---|---|
| 工具注册（N 个） | ✅/❌ | |
| Resource 注册 | ✅/❌ | |
| 动态收敛（W4a/W4b） | ✅/❌ | |

### Phase 2 功能验收
| 档位 | 用例 | 结果 |
|---|---|---|
| health | 事件接线 | ✅ |
| smoke | ... | N/N ✅ |
| full | ...（可选） | ✅ |

### 失败项详情
#### [项目名称]
- 预期 / 实际 / 截图或日志
```

## Prompt

请执行 UAT 验收（两阶段）。

首先确认阶段与前置：
- 若参数为 `registration` 或某注册场景名 → 只做 Phase 1（确认 MCP Inspector 已连或用程序化 client）。
- 若参数为 `functional` 或某 `*-functional` 场景名 → 只做 Phase 2（确认对应 Add-In dev server + Office + taskpane）。
- 若参数为空或 `all` → 逐平台推进：每平台先 Phase 1（注册/收敛），再引导触发人连接 Add-In，后 Phase 2（`{plat}-functional`，health→smoke→full）。

$ARGUMENTS

Phase 1 用 Inspector（Playwright，可视截图）或程序化 `list_tools`/`list_resources`（更精确）逐项核对注册；Phase 2 引导触发人连接 Add-In + 按脚本提示交互，我跑 `manual_tests` 用例收结果。每关键步骤留记录，合并输出报告。
