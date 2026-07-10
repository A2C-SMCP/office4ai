"""参考脚本：从模板实例化 + 填充 | Reference: instantiate a template and fill it.

这是本 SKILL 的头牌能力——**复用企业/个性化模板**生成新文件，而非只改内容。

⚠️ 沙箱禁网：``TEMPLATE`` 必须是 **Computer 本地文件系统上的路径**，不能是 http(s):// 在线地址。
   若用户给的是在线模板 URL，先用 Computer 的网络工具下载到本地，再把本地路径写进这里。

两步：
1. ``instantiate_from_template(template, output)`` —— 按扩展名分派：``.potx`` 走 content-type 互换
   （纯 zip，无需 LibreOffice）；``.dotx``/``.xltx`` 走 LibreOffice 无头实例化（运行环境需装 soffice）。
2. 用 helper 往成品里**填充**：Word 填 SDT 占位控件 / Excel 保图表改单元格 / PPT 按母版新建版式。

下例以 Word ``.dotx`` 为例；PPT/Excel 见 references/template-reuse.md。

⚠️ **这是「填路径再跑」的示意脚本，不能原样运行**：``TEMPLATE`` 是占位路径，须先换成你**本地**的
   真实模板路径（且模板要含与 ``fill_sdt_controls`` 键匹配的 SDT 占位控件）；``.dotx`` 实例化还需
   运行环境装有 soffice。无 soffice 的 CI 友好链路见 references/template-reuse.md 的 ``.potx`` 示例。
"""

from office4ai.office.authoring.helpers import fill_sdt_controls, instantiate_from_template

TEMPLATE = "/abs/local/path/to/brand.dotx"  # ← 换成你的本地模板路径（禁网：不可为在线 URL）
OUTPUT = "report.docx"

# 1) 实例化：模板 → 成品（等价 Office「基于模板新建」）
instantiate_from_template(TEMPLATE, OUTPUT)

# 2) 填充：把成品里的 SDT 占位控件按 alias 替换为真实内容
filled = fill_sdt_controls(
    OUTPUT,
    {
        "文档标题": "2026 年度工作报告",
        "副标题": "XX 部门",
        "章节一标题": "总体情况",
    },
)
print(f"instantiated + filled: {OUTPUT} ({filled} controls)")
