# 适配坑 | Pitfalls & gotchas

在 `office_run_script` 沙箱里写创建脚本时的常见坑。

## 沙箱约束

- **写只能写工作目录**：脚本的文件写入被限制在 `work_dir`（== cwd，也是返回的 `path`）。
  写到别处会被拦截。把产物存成 **相对路径**（`doc.save("out.docx")`）或 `work_dir` 下的路径。
- **默认禁网**：不能下载字体/图片/模板/任何在线资源。需要的素材先由 Computer 落到本地，
  再作为本地路径给脚本用（模板见 `template-reuse.md`）。
- **import 白名单**：脚本可直接 import 的顶层包 = `docx` / `pptx` / `openpyxl` / `lxml` /
  `docxtpl` / `PIL` + 安全 stdlib 子集 + `office4ai.office.authoring.helpers`。**不要** import
  `subprocess` / `socket` / `requests` 等（会被拒）。helper 内部对 soffice 的调用由运行时放行。
- **超时**：默认 60s wall-clock，超时进程组被杀。大批量生成可传更大的 `timeout`。

## python-docx（Word）

- **边框/底纹无直接 API**：段落下划红线、单元格底纹等需下探 OOXML（`OxmlElement` + `qn`），
  见 `scripts/gen_word.py` 的 `_add_red_bottom_border`。
- **样式名**：`Heading 1` / `Title` 等是内置样式名；用错名字会 KeyError。
- 图片插入用本地路径（`add_picture`），不能用 URL。

## python-pptx（PowerPoint）

- **图表**：`slide.shapes.add_chart(XL_CHART_TYPE.*, x, y, cx, cy, chart_data)` 返回 `GraphicFrame`，
  取图表对象用 `.chart`。分类图用 `CategoryChartData`；散点图用 `XyChartData`。
- **空白版式** = `prs.slide_layouts[6]`（默认模板）。
- **不支持插入视频/音频**（Office.js 也不支持）。
- 颜色用 `from pptx.dml.color import RGBColor`（注意与 python-docx 的 `docx.shared.RGBColor` 不同包）。

## openpyxl（Excel）

- **不计算公式值**：openpyxl 只写公式字符串（`ws["E2"] = "=SUM(...)"`），值待 Office/LibreOffice
  打开时重算。若下游要读值，用 `data_only=True` 打开一个已被 Office 重算过的文件。
- **`wb.active` 类型为 Optional**：运行时新建工作簿总有活动表，直接用即可。
- **图表数据引用**用 `Reference(ws, ...)`；类目用 `set_categories`。

## 通用

- 中文字体/字号：设 `font.name` + `font.size`；跨平台字体可能回退，正文默认字体一般够用。
- 产物校验：脚本末尾 `print(...)` 便于从 `logs` 确认；返回的 `summary.produced` 会列出落盘文件。
