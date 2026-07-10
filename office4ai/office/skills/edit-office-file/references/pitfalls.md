# 适配坑 | Pitfalls & gotchas（编辑既有文件）

## 沙箱约束（同 create-office-file，编辑场景尤其注意）

- **原地编辑须先拷进 work_dir**：沙箱只能写工作目录。`fill_sdt_controls` 等原地改的操作，先
  `shutil.copy(源文件, "本地名")` 进 cwd 再改；直接改工作目录外的源文件会被拦截。
- **默认禁网**：被编辑文件/模板/图片素材都要先落到 Computer 本地，再用本地路径。
- **import 白名单**：`docx` / `pptx` / `openpyxl` / `lxml` / `docxtpl` / `PIL` + 安全 stdlib（含
  `shutil`）+ `office4ai.office.authoring.helpers`。

## Excel：保图表是头等坑

- ⚠️ `openpyxl.load_workbook(x)` → 改 → `save`：按 openpyxl **不完整**的模型整体重写。简单图表可能幸存，
  但它建模不完整的部分（复杂图表/样式、条件格式、表单控件、宏）有丢失/改变风险，且不留公式缓存。
- ✅ `fill_cells_lxml(template, output, {sheet: {cell: value}})`：OOXML 层只改目标单元格、**保全原文件其余**。
- openpyxl **不算公式值**；`fill_cells_lxml` 已删缓存 + `fullCalcOnLoad` 逼重算。若你的下游要读值，
  用被 Office 重算并保存过的文件、以 `data_only=True` 读。

## docxtpl（Word 占位）

- 占位是 **Jinja 语法** `{{var}}` / `{% for %}` / `{% if %}`；`render(context)` 的键要与占位对应。
- 渲染在内存进行，`save(out)` 写出；模板本身不被修改。
- 占位若被 Word 拆进多个 run，docxtpl 通常能处理；但手动敲进模板的占位要确保不被自动更正拆断。

## SDT 填充（Word 内容控件）

- `fill_sdt_controls(path, {alias: value})` 按 **alias** 匹配。⚠️ **默认 `strict=True`**：字典里任一
  alias 在文档中找不到就抛 `AnchorNotFoundError`（fail-fast）。要容忍缺失（跳过未命中）请**显式传
  `strict=False`**。注意与 `fill_cells_lxml`（默认 `strict=False`，静默跳过缺失单元格）不对称。
- 是**原地编辑**：先 copy 进 work_dir。

## PPT 母版复用

- `locate_by_master(prs, name)` 按 `cSld/@name` 找版式；名字要与 deck 里真实版式名一致，否则
  `AnchorNotFoundError`。可先遍历 `prs.slide_masters[*].slide_layouts[*].name` 确认可用版式名。
- `add_slide(layout)` 会按版式生成占位；`slide.shapes.title` 可能为 None（版式无标题占位），用前判空。

## 通用

- 编辑前后校验：脚本末尾 `print(...)` 便于从 `logs` 确认；`summary.produced` 列出落盘产物。
- 保真优先：能用 helper（保图表/母版/SDT）就别用「整体重存」，后者易丢结构。
