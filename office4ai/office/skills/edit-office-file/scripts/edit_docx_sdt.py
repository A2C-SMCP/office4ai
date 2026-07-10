"""参考脚本：填充既有 .docx 的 SDT 内容控件 | Reference: fill SDT content controls in a .docx.

企业模板常用 SDT（内容控件）作命名占位框。``fill_sdt_controls(path, {alias: value})`` 按 alias
精准替换，不动其余结构/样式。这是**原地编辑**——沙箱只能写工作目录，故先 ``shutil.copy`` 把目标
文件拷进 cwd，再原地填。

⚠️ 填路径示意：``TARGET`` 换成你**本地**的真实文件路径（禁网）；``values`` 的键要与模板里 SDT 的
   alias 对应。``fill_sdt_controls`` **默认 strict=True**——任一 alias 在文档中不存在即抛
   ``AnchorNotFoundError``；要容忍缺失请显式传 ``strict=False``。
"""

import shutil

from office4ai.office.authoring.helpers import fill_sdt_controls

TARGET = "/abs/local/path/to/report_with_sdt.docx"

shutil.copy(TARGET, "report.docx")  # 拷进工作目录后原地编辑（沙箱写约束）
n = fill_sdt_controls(
    "report.docx",
    {
        "文档标题": "2026 年度工作报告",
        "副标题": "XX 部门",
        "章节一标题": "总体情况",
    },
)
print(f"filled {n} SDT controls -> report.docx")
