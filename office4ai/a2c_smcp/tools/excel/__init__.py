"""Excel MCP 工具集合 | Excel MCP tools."""

from office4ai.a2c_smcp.tools.excel.get_selected_range import ExcelGetSelectedRangeTool
from office4ai.a2c_smcp.tools.excel.get_workbook_info import ExcelGetWorkbookInfoTool
from office4ai.a2c_smcp.tools.excel.get_worksheet_info import ExcelGetWorksheetInfoTool

__all__ = [
    # State-awareness read tools (OASP /excel Draft, issue #18 Foundation)
    "ExcelGetWorkbookInfoTool",
    "ExcelGetWorksheetInfoTool",
    "ExcelGetSelectedRangeTool",
]
