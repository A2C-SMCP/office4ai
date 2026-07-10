# 抽取锚点方法学 | Anchor-based extraction methodology

「抽取」= 把参考文件里的**可变部分**（实值）升级为**占位**，其余结构/样式原样保留。方向由你**显式指定**
（非自动探测）：先勘锚点，再升级。四条方向如下。

## 1. Word 结构占位 → 命名 SDT（`wrap_in_sdt`）

- **勘锚点**：`locate_by_style(docx, style_id)` 返回该段落样式的所有 `StyleMatch(index, text)`。
  `index` 与 `wrap_in_sdt` 的 `para_index` **同一坐标系**（正文全部 `w:p` 的文档顺序序号），可直接喂过去。
- **升级**：`wrap_in_sdt(docx, para_index, alias, *, placeholder=…)` 把该段包成带 `alias` 的 SDT 内容控件。
- **闭环**：产出的 SDT 被 `fill_sdt_controls(docx, {alias: 值})` 按 alias 回填。**抽取时取的 alias = 回填的键**。
- 适合：标题、字段框、企业模板里「灰显占位」那类命名控件。

## 2. Word 文本占位 → `{{token}}`（`replace_runs_with_token`）

- **升级**：`replace_runs_with_token(docx, {实值: "{{token}}"})` 把具体文本换成 docxtpl 占位。
- **闭环**：产出的 `{{token}}` 被 `docxtpl` 的 `render({"token": 值})` 回填。
- 适合：正文里的实名/日期/金额、需要整段/循环/条件的地方（docxtpl 支持 `{% for %}` / `{% if %}`）。
- SDT vs token 怎么选：**命名控件框、要灰显占位、结构级** → SDT；**正文内联、要 Jinja 逻辑** → token。

## 3. PPT 版式复用 → 母版版式（`promote_slide_to_layout`）

- **勘锚点**：遍历 `prs.slides` 找那张「样板页」；或 `locate_by_master(prs, name)` 确认目标版式名未占用。
- **升级**：`promote_slide_to_layout(prs, slide_index, name, remove_source=True)` 把样板页脱胎为母版版式并删源页。
- **闭环**：产出的版式被 `locate_by_master(prs, name)` + `add_slide(layout)` 复用出图。
- 适合：把一张设计好的页做成「以后照着加页」的品牌版式。样板页**别带外部媒体**（图片/超链接 r:id 会悬空；
  含媒体会 `ExternalMediaError` 快速失败）。

## 4. Excel 区域复用 → 命名区域模板（`locate_by_named_range` + `fill_cells_lxml`）

- **勘锚点**：`locate_by_named_range(xlsx, name)` 返回 `NamedRange(refers_to, sheet, ref, …)`，确认命名区域仍在。
- **升级**：`fill_cells_lxml(xlsx, out, {sheet: {cell: None}})` 把示例值单元格清空，**保图表/条件格式/命名结构**。
- **闭环**：产出的模板被 `fill_cells_lxml(template, out, {sheet: {cell: 值}})` 回填。
- 适合：仪表盘/报表——留公式、图表、命名区域，只抹掉示例数据。

## run-splitting：为什么抽取必须走 helper

Word/PPT 里一个词经常被拆进多个 run（`20|26`、`{{ti|tle}}`）。若你按 run 逐个改，占位会被拆断——SDT 文本
错位、`{{token}}` 被 docxtpl 认不出。`wrap_in_sdt`（把段落 run 归一为单个占位 run）与
`replace_runs_with_token`（按段内拼接全文定位、`{{token}}` 落单个 run）**已经消解了这件事**，直接用。
