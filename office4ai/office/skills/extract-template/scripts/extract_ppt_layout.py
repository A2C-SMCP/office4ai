"""参考脚本：把参考 deck 的样板页抽成可复用母版版式模板 | Reference: slide -> reusable layout.

抽取方向「版式复用」：参考 deck 里有一张精心设计的样板页，用 ``promote_slide_to_layout`` 把它**脱胎**为
母版里的命名版式。产出的模板里，日后 ``locate_by_master(prs, name)`` + ``add_slide(layout)`` 即按该版式
批量出图、保持品牌一致（见 edit-office-file SKILL 的 ``edit_ppt_master.py``）。

``remove_source=True`` 抽完删源页，产出干净模板；样板页**别带图片/超链接等外部媒体**（其 r:id 不随形状
搬运，会在新版式里悬空）。这是读入-写出：直接 ``save`` 到工作目录。

⚠️ 填路径示意：``REF`` 换成你**本地**的真实 deck 路径（禁网）；``SLIDE_INDEX`` 换成样板页在 deck 里的序号；
   ``LAYOUT_NAME`` 换成你给新版式取的名字（= 日后 ``locate_by_master`` 要用的名）。
"""

from pptx import Presentation

from office4ai.office.authoring.helpers import promote_slide_to_layout

REF = "/abs/local/path/to/reference.pptx"
SLIDE_INDEX = 0
LAYOUT_NAME = "品牌版式"

prs = Presentation(REF)  # 读参考 deck（禁网：本地路径）
layout = promote_slide_to_layout(prs, SLIDE_INDEX, LAYOUT_NAME, remove_source=True)
prs.save("template.pptx")
print(f"promoted slide {SLIDE_INDEX} -> reusable layout {layout.name!r} -> template.pptx")
