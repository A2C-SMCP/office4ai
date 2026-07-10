"""参考脚本：从零创建销售汇报幻灯片 | Reference: create a sales-report slide from scratch.

演示一页里放：标题 + 数据表 + 柱状图 + 三张 KPI 卡片。产物存到 cwd。
python-pptx 的图表用 ``add_chart`` + ``CategoryChartData``；KPI 卡片用圆角矩形 + 文本框。
"""

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

BRAND = RGBColor(0x1F, 0x4E, 0x79)
MONTHS = ["1月", "2月", "3月", "4月"]
SALES = [120, 150, 135, 180]

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # 6 = 空白版式

# 标题
title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12.3), Inches(0.9))
tf = title.text_frame
tf.text = "2026 上半年销售汇报"
tf.paragraphs[0].font.size = Pt(32)
tf.paragraphs[0].font.bold = True
tf.paragraphs[0].font.color.rgb = BRAND

# KPI 卡片（三张圆角矩形 + 大数字）
kpis = [("总销售额", "585 万"), ("环比增长", "+18%"), ("目标达成", "112%")]
for i, (label, value) in enumerate(kpis):
    left = Inches(0.5 + i * 4.2)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(1.4), Inches(3.9), Inches(1.5))
    card.fill.solid()
    card.fill.fore_color.rgb = BRAND
    card.line.fill.background()
    cf = card.text_frame
    cf.text = value
    cf.paragraphs[0].font.size = Pt(28)
    cf.paragraphs[0].font.bold = True
    cf.paragraphs[0].alignment = PP_ALIGN.CENTER
    sub = cf.add_paragraph()
    sub.text = label
    sub.font.size = Pt(14)
    sub.alignment = PP_ALIGN.CENTER

# 数据表
rows, cols = len(MONTHS) + 1, 2
table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(3.3), Inches(4.5), Inches(3.2)).table
table.cell(0, 0).text = "月份"
table.cell(0, 1).text = "销售额（万）"
for r, (month, amount) in enumerate(zip(MONTHS, SALES, strict=True), start=1):
    table.cell(r, 0).text = month
    table.cell(r, 1).text = str(amount)

# 柱状图
chart_data = CategoryChartData()
chart_data.categories = MONTHS
chart_data.add_series("销售额", SALES)
chart = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Inches(5.4),
    Inches(3.3),
    Inches(7.4),
    Inches(3.6),
    chart_data,
).chart
chart.has_legend = True
chart.legend.position = XL_LEGEND_POSITION.BOTTOM

prs.save("sales_report.pptx")
print("created sales_report.pptx")
