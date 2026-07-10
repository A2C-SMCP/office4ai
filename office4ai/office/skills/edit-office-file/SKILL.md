---
name: edit-office-file
description: Edit an existing Word / PowerPoint / Excel file — especially template operations (fill {{placeholders}} and SDT content controls, reuse slide masters, change spreadsheet data while preserving charts) — by submitting a short Python script to the office4ai `office_run_script` tool. Works with no Office Add-In connection. Use when the user asks to fill a template, update a report/deck/workbook, replace placeholders, or edit a .docx / .pptx / .xlsx while keeping its styling intact.
version: 0.1.0
license: MIT
---

# edit-office-file — 脚本化编辑 Office 文件（模板操作为核心）

中文：本 SKILL 指导你用 **`office_run_script`** 提交一小段 Python 脚本，**编辑既有** Word / PPT /
Excel 文件——重点是**模板操作**：占位符替换（`{{}}` 与 SDT 内容控件）、母版样式复用、**保图表改数据**。
与「从零创建」（见 `create-office-file` SKILL）互补；运行环境已内置锁定版本的
`python-docx` / `python-pptx` / `openpyxl` / `lxml` / `docxtpl`，**无需 Add-In 连接**。

English: Edit an existing Office file (template operations first-class) by submitting a script to
`office_run_script`. Complements the `create-office-file` SKILL. No Add-In connection required.

## 何时用 | When

- 用户要**改一个已有文件**：填模板占位、更新报表数据、换幻灯片内容、替换红头/落款等。
- 用户有**模板 + 数据**，要把模板**填成**成品（`{{name}}` / SDT 占位 / 命名区）。
- 要改 Excel 数据但**保留图表/条件格式**；要复用 PPT **母版版式**加页而不破坏品牌。
- 文件当前**未在 Office 里打开**。已连接见「W4c」。

## 核心工作流 | Workflow

1. **想清编辑**：改哪个文件、用哪种模板操作（`{{}}` 渲染 / SDT 填充 / 保图表改单元格 / 母版复用）。
2. **写脚本**：`import` 对应库/helper，**读**目标文件（本地路径），把结果**写到工作目录**。
   ⚠️ 沙箱只能写 `work_dir`：**原地编辑类**操作（如 `fill_sdt_controls`）须先把目标文件
   `shutil.copy` 进工作目录再改；**读入-写出类**（`docxtpl` / `fill_cells_lxml`）直接指定工作目录内的输出名。
3. **调 `office_run_script`**：`script` 传脚本；`template_uri` 传要编辑的**本地**文件路径（禁网，见下）。
4. **读回** `{ok, path, summary, logs, stderr}`。

> **学模式、别抄源码**：`scripts/` 是参考脚本，读它学 glue，然后 `import` helper 写你自己的几行。

## 参考脚本 | Reference scripts

| 脚本 | 技术 | 保真点 |
|------|------|--------|
| `scripts/edit_docx_docxtpl.py` | docxtpl `{{}}` 渲染 | 保模板全部样式，仅换占位 |
| `scripts/edit_docx_sdt.py` | `fill_sdt_controls` 填 SDT 控件 | 按 alias 精准替换占位，不动其余 |
| `scripts/edit_xlsx_preserve_chart.py` | `fill_cells_lxml` 保图表改数据 | **保留图表/条件格式** + 删公式缓存逼重算 |
| `scripts/edit_ppt_master.py` | `locate_by_master` 复用母版版式 | 新页沿用模板母版样式，品牌一致 |

## 模板操作要点 | Template ops（与「只改文本」的差异）

- **占位符替换**：`docxtpl` 用 `{{jinja}}` 语法（适合整段/循环/条件）；SDT 内容控件用
  `fill_sdt_controls(path, {alias: value})`（适合企业模板的命名占位框）。
- **保图表改数据**：改含图表/复杂格式的工作簿，用 `fill_cells_lxml(template, output, {sheet: {cell: value}})`
  而**非 `openpyxl` 整体载入-重存**。openpyxl 会按自己**不完整**的模型整体重写工作簿——对它建模不完整的
  部分（复杂图表/样式、条件格式、表单控件、宏等）有丢失或改变的风险，且不保留公式缓存。`fill_cells_lxml`
  在 OOXML 层**只改目标单元格、保全原文件其余一切**，并删公式缓存逼 Office 打开时重算，最稳妥。
- **母版样式复用**：`locate_by_master(prs, 版式名)` 取模板母版里的版式，`add_slide(layout)` 沿用其样式加页。

⚠️ **沙箱禁网——被编辑文件/模板必须是本地路径**：`template_uri` 与脚本读的路径**必须是 Computer 本地
   文件系统路径**，不能是在线地址。若用户给在线 URL，**先用 Computer 网络工具下载到本地再传**。

详见 `references/template-ops.md`；坑见 `references/pitfalls.md`。

## W4c · 已连接文件优先用 Add-In 工具

若目标文件**已在 Office 里被 Add-In 连接**（可查 `window://office4ai/*` 或连接态），**优先调用**该平台
Add-In 工具（`word_*` / `ppt_*` / `excel_*`）——直接编辑打开态文档，更精准、所见即所得。脚本流水线改盘后
需重开渲染、且会与用户实时编辑冲突，仅用于**无连接**场景。
