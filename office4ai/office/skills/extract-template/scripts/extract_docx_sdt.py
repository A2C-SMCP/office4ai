"""参考脚本：把参考 .docx 的标题段落抽成命名 SDT 占位模板 | Reference: headings -> named SDT.

抽取方向「结构占位」：先用 ``locate_by_style`` 按段落样式（如标题）勘出锚点，再用 ``wrap_in_sdt`` 把每个
标题段**原地升级**为带 alias 的 SDT 内容控件。产出的模板可被 ``fill_sdt_controls({alias: 值})`` 精准回填
（见 edit-office-file SKILL 的 ``edit_docx_sdt.py``），形成「抽取↔填充」闭环。

wrap_in_sdt 是**原地编辑**——沙箱只能写工作目录，故先 ``shutil.copy`` 把参考文件拷进 cwd 再抽。
wrap_in_sdt 已消解 run-splitting（把段落 run 归一为单个占位 run），别自己按 run 改。

⚠️ 填路径示意：``REF`` 换成你**本地**的真实参考文件路径（禁网）；``STYLE_ID`` 换成该文件里真实存在的
   段落样式 id（如 ``"Heading1"`` / ``"Title"``；注意是样式 **id** 而非显示名，区分大小写）。
"""

import shutil

from office4ai.office.authoring.helpers import locate_by_style, wrap_in_sdt

REF = "/abs/local/path/to/reference.docx"
STYLE_ID = "Heading1"

shutil.copy(REF, "template.docx")  # 拷进工作目录后原地抽取（沙箱写约束）
matches = locate_by_style("template.docx", STYLE_ID)  # 勘锚点：所有该样式的段落
for i, m in enumerate(matches):
    # alias/占位名就是日后回填要用的键——取有意义的名字
    wrap_in_sdt("template.docx", m.index, alias=f"标题{i + 1}", placeholder=f"【{m.text or '标题'}】")
print(f"wrapped {len(matches)} '{STYLE_ID}' heading(s) into named SDT placeholders -> template.docx")
