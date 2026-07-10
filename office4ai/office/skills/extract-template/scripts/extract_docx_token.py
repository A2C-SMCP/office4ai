"""参考脚本：把参考 .docx 的具体实值替换成 docxtpl {{token}} 占位 | Reference: values -> {{tokens}}.

抽取方向「文本占位」：把参考文件里的实名/日期/金额等具体值换成 ``{{jinja}}`` 占位，产出的模板可被
``docxtpl`` 渲染回填（见 edit-office-file SKILL 的 ``edit_docx_docxtpl.py``）。

``replace_runs_with_token`` 已消解 run-splitting——实值/占位常被 Word 拆进多个 run，helper 按段内拼接
全文定位、把 ``{{token}}`` 落进**单个 run**，避免 docxtpl 认不出。**别自己按 run 逐字改**。

这是**读入-写出**（非原地）：直接指定工作目录内的输出名，无需先 copy。

⚠️ 填路径示意：``REF`` 换成你**本地**的真实参考文件路径（禁网）；``REPLACEMENTS`` 的键换成该文件里真实
   出现的实值，值是你选的占位名（= 日后 docxtpl ``render`` 的 context 键）。``strict=False`` 容忍个别实值
   不在文档中；要求「每个实值都必须命中」时改 ``strict=True``。
"""

from office4ai.office.authoring.helpers import replace_runs_with_token

REF = "/abs/local/path/to/reference.docx"

REPLACEMENTS = {
    "张三": "{{name}}",
    "2026-01-01": "{{date}}",
    "12,800": "{{amount}}",
}

n = replace_runs_with_token(REF, REPLACEMENTS, output="template.docx", strict=False)
print(f"replaced {n} concrete value(s) with docxtpl tokens -> template.docx")
