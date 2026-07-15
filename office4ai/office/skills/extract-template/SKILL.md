---
name: extract-template
description: Extract a reusable template from a reference Word / PowerPoint / Excel file — turn concrete values into placeholders — by submitting a short Python script to the office4ai `office_run_script` tool. Word headings become named SDT content controls, concrete text becomes {{tokens}} for docxtpl, a designed slide becomes a reusable master layout, an Excel named range becomes a blanked template region. Works with no Office Add-In connection. Use when the user has a finished/reference document but no template, and wants to reuse its structure/branding to generate more files.
version: 0.1.0
license: MIT
---

# extract-template — 从参考文件脱胎可复用模板（W3）

中文：用户常有一份**参考文件**（做好的报告/发文/品牌 deck/仪表盘）却没有**模板文件**。本 SKILL 指导你用
office4ai 的 **`office_run_script`** 提交一小段脚本，把参考文件**脱胎**为可复用模板——把具体实值升级成
**占位**：Word 标题段 → 命名 **SDT** 内容控件；Word 实值 → docxtpl **`{{token}}`**；PPT 样板页 → 可复用
**母版版式**；Excel 命名区域 → **留结构、清实值**的模板区。产出的模板可直接被 `create-office-file` /
`edit-office-file` 两个 SKILL 消费。**无需 Office Add-In 连接**。

English: Extract a reusable template from a reference file by submitting a script to `office_run_script`.
Concrete values become placeholders (SDT / `{{token}}` / master layout / blanked named range). The output
is consumable by the `create-office-file` and `edit-office-file` SKILLs. No Add-In connection required.

## 何时用 | When

- 用户有**成品/参考文件**，想**复用它的结构或品牌**批量出新文件，但手上没有模板。
- 用户说「照着这个文件做个模板」「把里面的名字/日期/金额换成占位」「把这页做成可复用版式」。
- 参考文件当前**未在 Office 里打开**（无 Add-In 连接）。已连接见下方「W4c」。

## 核心工作流 | Workflow

1. **想清抽取方向**（由你显式指定，非自动探测）：
   - Word 结构占位 → **SDT**（命名内容控件，适合标题/字段框，回填精准）——`wrap_in_sdt`
   - Word 文本占位 → **`{{token}}`**（适合整段/循环，docxtpl 渲染）——`replace_runs_with_token`
   - PPT 版式复用 → **母版版式**（样板页脱胎为 layout）——`promote_slide_to_layout`
   - Excel 区域复用 → **命名区域**（留公式/图表/命名结构，清示例值）——`locate_by_named_range` + `fill_cells_lxml`
2. **写脚本**：`import` helper，**读**参考文件（本地路径），把模板**写到工作目录**。
   ⚠️ `wrap_in_sdt` 是**原地编辑**：先 `shutil.copy` 参考文件进工作目录再抽；`replace_runs_with_token` /
   PPT / Excel 是**读入-写出**，直接指定工作目录内的输出名。
3. **调 `office_run_script`**：`script` 传脚本；`template_uri` 传**参考文件的本地路径**（进 FS 读白名单；禁网）。
4. **读回** `{ok, path, summary, logs, stderr}`，把产出的模板交给 `create-office-file` / `edit-office-file` 复用。

> **学模式、别抄源码**：`scripts/` 是参考脚本，读它学 glue，然后 `import` helper 写你自己的几行：
> `from office4ai.office.authoring.helpers import wrap_in_sdt, replace_runs_with_token, promote_slide_to_layout`。

## 先勘锚点 · inspect-first（抽取前先看清参考文件结构）

抽取由你**显式指定**可变部分，所以下手前必须先看清参考文件长什么样。无 Add-In 连接时，
把 `office_run_script` 当**只读探针**：先提交一段**只 `print`、不写盘**的脚本，把标题样式 /
命名区域 / 版式名 / 现有 SDT 读出来，内容经返回契约的 **`logs`** 回传；据此决定抽哪些、取什么
alias/token 名。纯读**不产文件**（`summary.produced` 为空），`work_dir` 可省。下列片段均经沙箱实测 `ok=True`。

**勘 Word 标题锚点**（决定哪些标题升级为 SDT / token）
```python
from docx import Document
doc = Document(r"/本地路径/reference.docx")
for i, p in enumerate(doc.paragraphs):
    if p.style.name.startswith(("Heading", "Title")):
        print(f"[{i}] {p.style.name}: {p.text}")   # i 即 wrap_in_sdt 的 para_index
```

**勘 Excel 命名区域**（决定清哪块、`locate_by_named_range` 用哪名）
```python
from openpyxl import load_workbook
wb = load_workbook(r"/本地路径/reference.xlsx")
print("sheets:", wb.sheetnames)
print("named:", [n for n in wb.defined_names])
```

**勘 PPT 版式名**（`promote_slide_to_layout` 命名 / 日后 `locate_by_master` 用）
```python
from pptx import Presentation
prs = Presentation(r"/本地路径/reference.pptx")
for mi, master in enumerate(prs.slide_masters):
    for layout in master.slide_layouts:
        print(f"master[{mi}] layout: {layout.name}")
```

> 已含 SDT 的模板想核对 alias（SDT 文本**不在** `doc.paragraphs`），见 `edit-office-file` SKILL 的
> 「列 SDT 占位 alias」片段（`qn()` + `iter/findall`）；**别用** `sdt.xpath('.//w:alias/@w:val')`——
> 裸 lxml 元素不认 `w:` 前缀，会抛 `XPathEvalError`。

## 参考脚本 | Reference scripts（读之学模式）

| 脚本 | 平台 | 抽取方向 | 产出可被谁消费 |
|------|------|---------|---------------|
| `scripts/extract_docx_sdt.py` | Word | `locate_by_style` → `wrap_in_sdt` 升级标题为命名 SDT | `edit_docx_sdt`（`fill_sdt_controls`）|
| `scripts/extract_docx_token.py` | Word | `replace_runs_with_token` 实值 → `{{token}}` | `edit_docx_docxtpl`（docxtpl 渲染）|
| `scripts/extract_ppt_layout.py` | PowerPoint | `promote_slide_to_layout` 样板页 → 可复用母版版式 | `edit_ppt_master` / `gen_ppt`（`add_slide(layout)`）|
| `scripts/extract_xlsx_named.py` | Excel | `locate_by_named_range` + 清实值，保图表/命名结构 | `edit_xlsx_preserve_chart`（`fill_cells_lxml`）|

## 抽取要点 | Extraction essentials

- **抽取由你显式指定，不自动探测**：你决定哪些段落/实值/页/区域是「可变部分」，helper 只负责把它们
  可靠地升级为占位。先用 `locate_by_style` / `locate_by_named_range` / 遍历 slides **勘一遍锚点**，再抽。
- **run-splitting 已由 helper 消解**：真实文档里一个词常被 Word 拆进多个 run。`wrap_in_sdt`（归一为单 run
  占位）与 `replace_runs_with_token`（跨 run 合并、`{{token}}` 落单个 run）都已处理——**别自己按 run 逐个改**，
  那样占位会被拆断、docxtpl 认不出。
- **抽取↔填充闭环**：`wrap_in_sdt(alias=…)` 产出的 SDT 恰好被 `fill_sdt_controls({alias: 值})` 回填；
  `replace_runs_with_token({实值:"{{k}}"})` 产出的 `{{k}}` 恰好被 docxtpl `render({"k": 值})` 回填。抽取时
  取的 alias/token 名，就是日后回填要用的键——**取有意义的名字**。
- **PPT 抽版式**：`promote_slide_to_layout(prs, i, name, remove_source=True)` 把第 i 页脱胎为母版版式并删源页，
  产出干净模板；日后 `locate_by_master(prs, name)` + `add_slide` 即按该版式出图。样板页**别带图片/超链接等
  外部媒体**（其 `r:id` 不随形状搬运，会在新版式里悬空）。
- **Excel 抽区域**：真实模板的图表/条件格式经 openpyxl round-trip 会损坏，故用 `fill_cells_lxml` 在 OOXML 层
  只清示例单元格（传 `None` 清空）、保全其余；`locate_by_named_range` 先确认命名区域锚点仍在。

⚠️ **沙箱禁网——参考文件必须是本地路径**：`template_uri` 与脚本读的路径**必须是 Computer 本地文件系统
   路径**，不能是在线地址。若用户给在线 URL，**先用 Computer 网络工具下载到本地再传**。

详见 `references/anchors.md`（四种抽取方向的锚点方法学）；坑见 `references/pitfalls.md`。

## W4c · 已连接文件优先用 Add-In 工具

若参考文件**已在 Office 里被 Add-In 连接**（可查 `window://office4ai/*` 或连接态），读取其结构可优先用该平台
Add-In 工具（`word_*` / `ppt_*` / `excel_*`）拿到更准的结构信息；但**抽取改盘**（写模板文件）仍走脚本流水线
落到工作目录，避免与用户实时编辑冲突。
