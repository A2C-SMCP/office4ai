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

### Phase 2 执行模型：auto_open 语义（关键，勿误判「不可执行」）

**多数 `manual_tests` 用例 agent 可非交互直接跑——PPT 也是。** 常见误判：见脚本含 `input()` 就断定「需人工、跑不了」。**错。** 机制：

```python
if auto_open:                 # auto_open 默认 True
    await asyncio.sleep(2.0)   # 自动续跑，非交互
else:
    input("按回车继续...")       # 仅 --no-auto-open 才走这里
```

- 「按回车继续…」类 `input()` 是**节奏暂停**，受 `auto_open`（默认 True）门控 → 默认 **sleep 2s、不 input**。故 `--test all` 默认可由 agent 非交互执行（Word/PPT/Excel 同理）。
- agent 的 Bash **无 TTY**：只有**真正 GUI 动作型** `input()`（非 pacing）才会 `EOF when reading a line`——那是「需真人操作」的信号，不是脚本 bug。
- **判定可跑与否，看 `input()` 是否 pacing（auto_open 门控），不是看有没有 `input()`。** 反例教训：曾据「PPT 23 文件含 input()」错误结论「PPT 不可自动跑」，实际 PPT full 绝大多数自动通过（update_text_box 4/4、update_table 3/3…）。

### Manual-Only 用例协议（真人 GUI 动作，agent 不可代跑）

少数用例的 `input()` 是**真·GUI 动作屏障**——需真人在 Office 里手动操作，且 **agent 即便有 TTY 也做不了**（无法操作 Word/PPT GUI）。已知清单：
- **word `get_selection`**（手动选中文本）、**cursor-placement**（如 insert_text location「光标位插入」）、**visual color-confirm**（颜色需肉眼确认）。
- 判据：提示语是「完成操作后按 Enter」（要求先做 GUI 动作）且**不受 auto_open 门控**。

**处理协议（强制）**——对 Manual-Only 用例，agent **必须**：
1. **单列**：报告设「🙋 Manual-Only 待人工验收」区块，不混入自动结果；
2. **记录确切命令 + 需人工做的动作**（每条给可复制的 `uv run python …` + 一句「你要做什么」）；
3. **交给用户在自己 TTY 终端跑**（Claude Code 里用 `!<cmd>` 前缀，或本地 shell），跑完贴回结果由 agent 汇总；
4. **不得**把未验的 Manual-Only 标 pass/fail——一律标「⏳ 未验（待人工）」；**更不得**因 agent 跑不了就当通过。

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

### 🙋 Manual-Only 待人工验收（agent 不可代跑，请在你的终端跑后贴回）
| 用例 | 需你做的动作 | 执行命令 |
|---|---|---|
| word get_selection | 提示时在 Word 中选中一段文本 | `uv run python manual_tests/word/get_selection_e2e/test_selection.py --test all` |
| word insert_text 光标位插入 | 提示时把光标放到目标位置 | `uv run python manual_tests/word/insert_text_e2e/test_location_insert.py --test 3` |
| （其余 color-confirm 等按需补） | 肉眼确认颜色/样式 | … |

> 这些用例标「⏳ 未验（待人工）」，用户跑完贴回结果后才更新为 ✅/❌。

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
