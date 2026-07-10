---
name: create-office-file
description: Create Word / PowerPoint / Excel files from scratch or from a reusable template by submitting a short Python script to the office4ai `office_run_script` tool. Works with no Office Add-In connection. Use when the user asks to generate a .docx / .pptx / .xlsx, produce a report / deck / workbook, or instantiate a corporate template.
version: 0.1.0
license: MIT
---

# create-office-file — 脚本化创建 Office 文件

中文：本 SKILL 指导你用 office4ai 的 **`office_run_script`** 工具，提交一小段 Python 脚本，在
稳定沙箱运行环境里**从零或从模板**创建 Word / PowerPoint / Excel 成品——**无需 Office Add-In
连接**。运行环境已内置锁定版本的 `python-docx` / `python-pptx` / `openpyxl` / `lxml` / `docxtpl`，
你不必操心装库/版本漂移。

English: This SKILL guides you to create Word / PowerPoint / Excel files by submitting a short Python
script to office4ai's **`office_run_script`** tool. It runs in a stable sandbox with pinned document
libraries — **no Office Add-In connection required.**

## 何时用 | When

- 用户要「生成/创建」一个 `.docx` / `.pptx` / `.xlsx`（报告、发文、幻灯片、工作簿、仪表盘）。
- 用户有一个**模板**（红头文件模板、品牌 PPT 母版、报表模板）想**复用**着生成新文件。
- 文件当前**未在 Office 里打开**（无 Add-In 连接）。若文件已被 Add-In 连接，见下方「W4c」。

## 核心工作流 | Workflow

1. **想清产物**：哪个平台、从零还是从模板、要哪些结构（表格/图表/KPI/占位符）。
2. **写一小段脚本**：`import` 对应库（和需要时的 authoring helper），把文件**存到当前工作目录**
   （脚本的 cwd）。**你决定文件名、是否 save、一切动作**——运行时不替你决定。
3. **调用 `office_run_script`**：把脚本作为 `script` 传入；可选 `work_dir`（产物落盘目录，也是返回的
   `path`）、`template_uri`（本地模板路径）、`timeout`。
4. **读回结构化结果** `{ok, path, summary, logs, stderr}`：`ok` 是否成功；`path` 是产物目录；
   `summary` 列出新增/改动文件；`logs`/`stderr` 是脚本输出与报错。

> **学模式、别抄源码**：`scripts/` 下是**参考脚本**——读它学「怎么组织几行 glue」，然后写你自己的
> 几行提交给 `office_run_script`。企业模板复用的重活由 helper 承担，脚本里 **`import` helper**
> 而非复制其实现：`from office4ai.office.authoring.helpers import ...`。

## 参考脚本 | Reference scripts（读之学模式）

| 脚本 | 平台 | 演示 |
|------|------|------|
| `scripts/gen_word.py` | Word | 从零建红头文件（红色居中标题 + 红线 + 文号 + 正文 + 落款）|
| `scripts/gen_ppt.py` | PowerPoint | 从零建销售页（数据表 + 柱状图 + KPI 卡片）|
| `scripts/gen_excel.py` | Excel | 从零建工作簿（表头/数据 + 公式 + 柱状图 + 命名区域）|
| `scripts/from_template.py` | 三平台 | **从模板实例化 + 填充**（`.dotx`/`.potx`/`.xltx` → 成品，helper 填占位/单元格/版式）|

## 模板复用 | Template reuse（W1 头牌能力）

从模板生成新文件是本 SKILL 与普通「只改内容」工具的关键差异。用 `instantiate_from_template()`
把 `.dotx`/`.potx`/`.xltx` 实例化成 `.docx`/`.pptx`/`.xlsx`，再用 helper 填占位符/单元格/版式。

⚠️ **沙箱禁网——模板必须是本地路径**：运行环境**默认禁网**，`template_uri` 与脚本读的模板路径
**必须是 Computer 本地文件系统上的路径**，不能是 `http(s)://` 等在线地址。若用户给的是在线模板 URL，
**先用 Computer 提供的网络工具把它下载到本地文件系统，再把本地路径**传给 `office_run_script`。

详见 `references/template-reuse.md`（三平台模板类型、helper 用法、soffice 依赖）。

## W4c · 已连接文件优先用 Add-In 工具

若目标文件**已在 Office 里被 Add-In 连接**（可用 `window://office4ai/*` 资源或连接态判断），
**优先调用该平台的 Add-In 工具**（`word_*` / `ppt_*` / `excel_*`）而非脚本流水线：Add-In 直接操作
打开态文档，更精准、所见即所得。脚本流水线是**无连接**场景的能力；对已连接文件用脚本改盘需重开渲染、
且会与用户的实时编辑冲突。

## 适配坑 | Pitfalls

沙箱 import 白名单、work_dir 写约束、python-docx/pptx/openpyxl 的常见坑，见 `references/pitfalls.md`。
