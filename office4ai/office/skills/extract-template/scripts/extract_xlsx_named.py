"""参考脚本：把参考 .xlsx 的命名区域抽成「留结构、清实值」的模板 | Reference: named range -> template.

抽取方向「区域复用」：真实仪表盘/报表模板靠**命名区域**定位数据块。先用 ``locate_by_named_range`` 确认锚点
仍在，再用 ``fill_cells_lxml`` 把命名区域里的**示例值单元格清空**（传 ``None``），保全图表/条件格式/命名区
结构。产出的模板可被 ``fill_cells_lxml(template, out, {sheet: {cell: 值}})`` 回填（见 edit-office-file SKILL
的 ``edit_xlsx_preserve_chart.py``）。

为什么不用 openpyxl 整体重存：它按不完整模型重写工作簿，真实模板的图表/条件格式有丢失风险。
``fill_cells_lxml`` 在 OOXML 层只改目标单元格、保全其余，是抽取保真的关键。这是读入-写出（非原地）。

⚠️ 填路径示意：``REF`` 换成你**本地**的真实参考工作簿路径（禁网）；``RANGE_NAME`` 换成该工作簿里真实存在的
   命名区域名；``BLANK_CELLS`` 由你指定要清空的示例单元格（``{工作表名: [单元格, ...]}``）。
"""

from office4ai.office.authoring.helpers import fill_cells_lxml, locate_by_named_range

REF = "/abs/local/path/to/reference.xlsx"
RANGE_NAME = "金额区"
BLANK_CELLS = {"数据": ["B2", "B3"]}

ranges = locate_by_named_range(REF, RANGE_NAME)  # 勘锚点：命名区域仍在？
if not ranges:
    raise SystemExit(f"命名区域未找到 | named range not found: {RANGE_NAME}")
blanks = {sheet: dict.fromkeys(cells) for sheet, cells in BLANK_CELLS.items()}
written = fill_cells_lxml(REF, "template.xlsx", blanks, strict=False)  # 清示例值，保图表/命名区
print(f"located {RANGE_NAME!r} ({len(ranges)} match), blanked {written} sample cell(s) -> template.xlsx")
