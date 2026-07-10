"""参考脚本：改 Excel 数据但保图表 | Reference: edit Excel cells while preserving charts.

改含图表/复杂格式的工作簿用 ``fill_cells_lxml``，而**非 openpyxl 整体载入-重存**（openpyxl 按其不完整
模型整体重写，复杂图表/条件格式/控件/宏有丢失风险，且不留公式缓存）。``fill_cells_lxml(template,
output, sheet_cells)`` 在 OOXML 层只改指定单元格、**保全原文件其余**，并删相关公式缓存 + 设 fullCalcOnLoad
逼 Office 打开时重算。``sheet_cells = {工作表名: {单元格: 值}}``；值支持 str/int/float/bool，None 清空。

⚠️ 填路径示意：``TEMPLATE`` 换成你**本地**的真实工作簿路径（禁网）；单元格引用要在模板中已存在。
"""

from office4ai.office.authoring.helpers import fill_cells_lxml

TEMPLATE = "/abs/local/path/to/dashboard.xlsx"

n = fill_cells_lxml(
    TEMPLATE,
    "dashboard_updated.xlsx",
    {
        "数据": {"B2": 260, "B3": 175},  # 更新收入/支出，图表随之刷新（打开时重算）
    },
)
print(f"updated {n} cells (charts preserved) -> dashboard_updated.xlsx")
