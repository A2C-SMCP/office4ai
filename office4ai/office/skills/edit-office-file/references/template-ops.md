# 模板操作指南 | Template operations guide

编辑既有文件时，**模板操作**（保样式/母版/图表，只换数据）是与「只改文本」的关键差异。

## 0. 沙箱禁网 + 写约束（先读）

- **禁网**：被编辑的文件/模板**必须是 Computer 本地路径**，不能是在线 URL。在线资源先由 Computer
  网络工具下载到本地，再把本地路径传 `template_uri`。
- **只能写 work_dir**：
  - **读入-写出类**（`docxtpl.render().save(out)`、`fill_cells_lxml(template, out, ...)`）：`out` 用
    工作目录内的名字即可，源文件在别处只读没问题。
  - **原地编辑类**（`fill_sdt_controls(path, ...)`）：先 `shutil.copy(源, "work.docx")` 进工作目录，
    再对副本原地编辑。

## 1. 占位符替换：两种占位

| 方式 | 语法 | helper / 库 | 适合 |
|------|------|-------------|------|
| Jinja 占位 | `{{name}}` / `{% for %}` | `docxtpl.DocxTemplate(t).render(ctx)` | 整段/循环/条件、正文级填充 |
| SDT 内容控件 | 命名占位框（alias） | `fill_sdt_controls(path, {alias: value})` | 企业模板的命名占位框，精准替换 |

```python
# docxtpl：{{}} 渲染
from docxtpl import DocxTemplate
DocxTemplate("/local/letter.docx").render({"name": "张三"}).save("out.docx")

# SDT：按 alias 填（原地，先 copy 进 work_dir）
import shutil
from office4ai.office.authoring.helpers import fill_sdt_controls
shutil.copy("/local/report.docx", "report.docx")
fill_sdt_controls("report.docx", {"文档标题": "2026 年度报告"})
```

## 2. 保图表改数据（Excel 关键坑）

**openpyxl `load_workbook` → 改 → `save` 会按其不完整模型整体重写**（复杂图表/样式、条件格式、控件、宏
有丢失/改变风险，且不留公式缓存）。改含图表/复杂格式的工作簿，用 `fill_cells_lxml` 外科式编辑更稳妥：

```python
from office4ai.office.authoring.helpers import fill_cells_lxml
fill_cells_lxml("/local/dashboard.xlsx", "out.xlsx", {"数据": {"B2": 260, "B3": 175}})
```

- 它在 OOXML 层只改指定单元格，**保图表/条件格式**；
- 删相关公式缓存 + 设 `fullCalcOnLoad`，逼 Office 打开即重算（openpyxl 不算公式值）；
- `sheet_cells = {工作表: {单元格: 值}}`；`strict=True` 时引用不存在会抛 `AnchorNotFoundError`。

## 3. 母版样式复用（PPT）

```python
from pptx import Presentation
from office4ai.office.authoring.helpers import locate_by_master
prs = Presentation("/local/branded_deck.pptx")
layout = locate_by_master(prs, "标题和内容")   # 按名取母版版式（cSld/@name）
prs.slides.add_slide(layout)                    # 沿用其样式加新页
prs.save("out.pptx")
```

新页沿用模板母版的配色/占位/字体，保持全 deck 品牌一致。

## 4. 何时从零、何时编辑

- **有现成文件/模板要改** → 用本 SKILL。
- **从零生成** → 见 `create-office-file` SKILL（W1）。
- **把参考文件脱胎成模板** → 见 `extract-template` SKILL（W3）。
