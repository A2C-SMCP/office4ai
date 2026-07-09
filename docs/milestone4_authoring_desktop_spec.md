# Milestone #4 拆分规格 — 脚本化文档流水线 + AddIn 感知桌面

> 本文是 milestone #4（`A2C-SMCP/office4ai`）的权威拆分规格，各子 issue 引用本文对应小节。
> 经两路勘察（协议归属 + 跨仓地形）+ 模板库实测 + 三份真实参考脚本 + 多轮访谈定稿。

---

## 1. 目标

让 Office4AI 从「仅在有 Add-In 连接时可用」进化为：**无连接也能用脚本流水线创建/编辑/抽取模板；有连接则动态收敛到精准的 Add-In 工具集，并把每个已连接文件投射成一个可感知的 Window。**

闭环 = **操作提示（SKILL）+ 流水线控制（我们运行脚本）+ 稳定运行环境（内置依赖 + LibreOffice）**。区别于业界只能「改内容」的同类 SKILL：本 milestone 把**企业/个性化模板复用**作为一等公民。

## 2. 关键设计立场

「为什么由我们跑脚本而不是让 LLM 直接用 Bash」——运行时环境（库版本/是否安装）是环境维护问题；我们统一提供稳定运行时 + 通过 `skill://` 暴露操作指导，形成可控闭环。LLM 提交脚本 → 我们在沙箱运行时执行 → 结构化回传。

## 3. 决策记录（全部 locked）

| # | 决策 | 结论 |
|---|---|---|
| D1 工具形态 | 通用 `office_run_script`（提交脚本→沙箱执行→结构化回传）+ SKILL 指导；不做每意图工具 |
| D2 fullscreen 路线 | **路线甲**：office4ai 自保「任一时刻至多一个 window fullscreen」→ **零协议改动、python-sdk 零改动、关闭 #4** |
| D3 模板栈（分层组合） | ② python-docx/pptx/openpyxl + ① **lxml 主力**（SDT 内容控件 / 新建 slideLayout / 保图表改单元格）+ ③ docxtpl（docx `{{}}`）+ **LibreOffice/soffice（运行时必备，dotx/xltx 实例化）** + content-type swap（potx） |
| D4 沙箱 | 子进程 + 超时 + FS 白名单（目标目录 + template_uri）+ import 白名单（office 库 + authoring helper）+ 默认禁网 |
| D5 window 粒度 | per-file 取代 per-type 聚合 + 保留根索引 `window://office4ai` |
| 运行时封装（主决策） | helper 原语封装为**可 import 的库/脚本**，随 SKILL 的 `scripts/` 提供，**渐进式披露**，**不做 MCP 工具**（避免工具膨胀） |
| 实例化策略 | `instantiate_from_template()` 运行时按平台自动选：dotx/xltx→soffice，potx→content-type swap |
| W3 抽取 | LLM 脚本 + 锚点 helper（`locate_by_style/named_range/master` + `wrap_in_sdt`），同构于参考脚本 |
| 返回契约 | `{ok, path, summary, logs, stderr}`。脚本全自主（操作哪个文件/是否 save 都在脚本里）：`ok`=脚本是否运行成功；`path`=脚本运行目录(work_dir=FS 写白名单)；`summary`=运行时事后扫描 work_dir 得到的**产物清单**(新增/改动文件 name/size)，**不解析 office 文件内部**；`logs`=stdout；`stderr`=stderr。富视觉呈现(pages/slides/sheets)交给 W4b per-file window，运行时不做以免过度设计约束脚本 |
| 依赖打包 | `docxtpl` 进核心 dependencies；LibreOffice 系统级、运行时必备（对齐 UNO Bridge） |
| Chart Path A | 重分类为 `requires_connection=True`（W4a 二元模型）；删 Path A 代码 = 独立跟进项 F1 |
| Excel window | word/ppt per-file 先落；excel 渲染 blocked-by milestone #3 |

## 4. 勘察硬结论（背书决策）

- **零协议改动基线**：`skill://`（skill.md#11）、`window://` path 粒度 server 自由（desktop.md）、工具/资源变更链（events.md + computer.md#4.4）均为已发布稳定表面，office4ai 只需遵循。
- **python-sdk 零改动（路线甲）**：`tools/list_changed` 消费链已通；per-file window 对 Computer 开箱即用；单 fullscreen 使 organize.py 的 min-idx 平局成 no-op。
- **模板实测**：python-docx/pptx **拒绝** `.dotx/.potx`（content-type 门），openpyxl **原生**吃 `.xltx`；企业模板高价值特性（SDT / 新建版式 / 保图表改数据）**必须 lxml**；dotx/xltx 忠实实例化用 **LibreOffice**，potx 用 **content-type swap**（LO 会打乱占位符 idx）。

## 5. 子任务清单

### 地基（可并行起步）
- **S1 · authoring 运行时 + 沙箱 + `office_run_script` 工具**：`office/authoring/runtime.py` 子进程沙箱（超时/FS 白名单/import 白名单/禁网）+ 返回契约 + document_lock + soffice 探测/依赖声明 + docxtpl 进核心依赖。`office_run_script` 工具（standalone 常驻）。
- **S2 · authoring helper 原语库**（可 import，随 SKILL 分发）：`fill_sdt_controls` / `fill_cells_lxml` / `create_custom_layout` / `instantiate_from_template`（自动选 LO/swap）/ `locate_by_style/named_range/master`。提炼自参考脚本，含单测。
- **S3 · SkillResource 基础设施 + `skill://` 暴露**：`a2c_smcp/resources/skill.py` + 注册进 `_register_resources` + subscribe + SKILL `scripts/` 打包分发机制。

### SKILL 工作线（执行手段 = 集成测试跑通参考脚本）
- **S4 · W1 `create-office-file` SKILL** + 参考脚本 + 模板资产（含 W4c 文案）
- **S5 · W2 `edit-office-file` SKILL** + 模板操作（占位符替换/母版复用/保图表改数据）（含 W4c 文案）
- **S6 · W3 `extract-template` SKILL** + 抽模板锚点 helper

### W4 AddIn 感知桌面
- **W4a · 动态工具收敛**：BaseTool 建模 `requires_connection`（二元）+ chart 工具重分类 + list_tools 按 `get_clients_by_namespace` 过滤 + 连接变化发 `tools/list_changed` + 声明 `NotificationOptions(tools_changed=True)`。
- **W4b-1 · per-file window（word/ppt）+ 根索引**：`PerFileWindowResource` + 动态注册 + `resources/list_changed` + per-type→per-file（移除聚合，保根索引）。
- **W4b-2 · 单 fullscreen 归属（= #4 服务端解法）**：`update_last_activity` 钩子翻转 fullscreen + 清理上次 + 通知。← W4b-1
- **W4b-3 · excel per-file 渲染**：← W4b-1 + milestone #3。

### 末端 + 跟进
- **INT · 末端跨边界集成回归**：W1/W2/W3 端到端 + W4a 工具收敛（office4ai↔Computer）+ W4b 单 fullscreen 组装（#4 回归）+ skill:// staging。
- **F1 · 删 chart Path A 代码**（独立跟进，← W4a）。

## 6. 依赖图

```
S1(运行时+沙箱+run_script) ─┐
S2(helper 原语库) ──────────┼─→ S4(W1 create) ─┐
S3(SkillResource+skill://) ─┘   S5(W2 edit) ────┼─→ INT(末端集成回归)
                                 S6(W3 extract) ─┤
W4a(动态工具收敛) ───────────────────────────────┤
W4b-1(per-file word/ppt) ─┬─→ W4b-2(单 fullscreen #4) ─┘
                          └─→ W4b-3(excel, +blocked-by #3)
W4a ─→ F1(删 Path A, 跟进)
```

**并行起步**：S1 / S2 / S3 / W4a / W4b-1。关键路径 ≈ S1 → S4/S5/S6 → INT。

## 7. 核验

- **main 可消费（Phase 3）**：每子 issue 单独 merge 后编译/测试/关键行为不破；无协议根；无跨 SDK 平行；Path A 删除独立成 F1，不引入临时兼容层。
- **覆盖完整性（Phase 6）**：父级验收逐条映射（W1→S4 / W2→S5 / W3→S6 / W4a→W4a / W4b→W4b-1+2+3 / W4c→S4,S5），无悬空；跨 sub-task invariant（工具收敛端到端、单 fullscreen 组装#4、skill staging、脚本产物保真）→ INT 唯一汇聚。

## 8. 参考素材

- 参考脚本（作者 llg）：`gen_word.py` / `gen_ppt.py` / `gen_excel.py` —— W1/W2/W3 参考实现，覆盖「造模板 / 自造模板生成 / 真实 .dotx·.potx·.xltx 实例化」三链路，示范 lxml 结合对象模型库 + docxtpl + LibreOffice/content-type swap 的最佳实践。
- 真实素材：两个业务 `.docx` + 真实 `.dotx/.potx/.xltx` 模板 —— 集成测试打磨素材。

---

**维护者**：JQQ <jqq1716@gmail.com>
