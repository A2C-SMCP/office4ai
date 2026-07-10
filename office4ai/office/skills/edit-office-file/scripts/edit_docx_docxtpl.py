"""参考脚本：docxtpl {{}} 占位符渲染 | Reference: render a docxtpl template's {{placeholders}}.

适合「模板 + 数据 → 成品」的整段/循环/条件填充。模板是含 Jinja 占位（如 ``{{name}}``、
``{% for row in rows %}``）的本地 .docx。保留模板全部样式，只换占位内容。

⚠️ 填路径示意，不能原样运行：``TEMPLATE`` 换成你**本地**的真实模板路径（禁网：不可为在线 URL）；
   ``context`` 的键要与模板里的 ``{{占位}}`` 对应。
"""

from docxtpl import DocxTemplate

TEMPLATE = "/abs/local/path/to/letter_template.docx"  # 含 {{name}} {{date}} {{amount}} 等占位

tpl = DocxTemplate(TEMPLATE)
tpl.render(
    {
        "name": "张三",
        "date": "2026 年 7 月 10 日",
        "amount": "12,800",
    },
)
tpl.save("letter.docx")
print("rendered letter.docx")
