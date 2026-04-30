"""PPT MCP 工具集合 | PPT MCP tools."""

from office4ai.a2c_smcp.tools.ppt.add_slide import PptAddSlideTool
from office4ai.a2c_smcp.tools.ppt.delete_element import PptDeleteElementTool
from office4ai.a2c_smcp.tools.ppt.delete_slide import PptDeleteSlideTool
from office4ai.a2c_smcp.tools.ppt.get_chart import PptGetChartTool
from office4ai.a2c_smcp.tools.ppt.get_current_slide_elements import PptGetCurrentSlideElementsTool
from office4ai.a2c_smcp.tools.ppt.get_slide_elements import PptGetSlideElementsTool
from office4ai.a2c_smcp.tools.ppt.get_slide_info import PptGetSlideInfoTool
from office4ai.a2c_smcp.tools.ppt.get_slide_layouts import PptGetSlideLayoutsTool
from office4ai.a2c_smcp.tools.ppt.get_slide_screenshot import PptGetSlideScreenshotTool
from office4ai.a2c_smcp.tools.ppt.goto_slide import PptGotoSlideTool
from office4ai.a2c_smcp.tools.ppt.insert_chart import PptInsertChartTool
from office4ai.a2c_smcp.tools.ppt.insert_image import PptInsertImageTool
from office4ai.a2c_smcp.tools.ppt.insert_shape import PptInsertShapeTool
from office4ai.a2c_smcp.tools.ppt.insert_table import PptInsertTableTool
from office4ai.a2c_smcp.tools.ppt.insert_text import PptInsertTextTool
from office4ai.a2c_smcp.tools.ppt.move_slide import PptMoveSlideTool
from office4ai.a2c_smcp.tools.ppt.reorder_element import PptReorderElementTool
from office4ai.a2c_smcp.tools.ppt.update_chart import PptUpdateChartTool
from office4ai.a2c_smcp.tools.ppt.update_element import PptUpdateElementTool
from office4ai.a2c_smcp.tools.ppt.update_image import PptUpdateImageTool
from office4ai.a2c_smcp.tools.ppt.update_table_cell import PptUpdateTableCellTool
from office4ai.a2c_smcp.tools.ppt.update_table_format import PptUpdateTableFormatTool
from office4ai.a2c_smcp.tools.ppt.update_table_row_column import PptUpdateTableRowColumnTool
from office4ai.a2c_smcp.tools.ppt.update_text_box import PptUpdateTextBoxTool

__all__ = [
    # Content retrieval tools
    "PptGetCurrentSlideElementsTool",
    "PptGetSlideElementsTool",
    "PptGetSlideScreenshotTool",
    "PptGetSlideInfoTool",
    "PptGetSlideLayoutsTool",
    # Content insertion tools
    "PptInsertTextTool",
    "PptInsertImageTool",
    "PptInsertTableTool",
    "PptInsertShapeTool",
    # Update operation tools
    "PptUpdateTextBoxTool",
    "PptUpdateImageTool",
    "PptUpdateTableCellTool",
    "PptUpdateTableRowColumnTool",
    "PptUpdateTableFormatTool",
    "PptUpdateElementTool",
    # Chart tools (OASP /ppt Draft, v0.2.0 — Server OOXML)
    "PptInsertChartTool",
    "PptGetChartTool",
    "PptUpdateChartTool",
    # Delete & layout tools
    "PptDeleteElementTool",
    "PptReorderElementTool",
    # Slide management tools
    "PptAddSlideTool",
    "PptDeleteSlideTool",
    "PptMoveSlideTool",
    "PptGotoSlideTool",
]
