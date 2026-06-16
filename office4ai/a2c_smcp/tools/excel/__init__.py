"""Excel MCP 工具集合 | Excel MCP tools."""

from office4ai.a2c_smcp.tools.excel.clear_range import ExcelClearRangeTool
from office4ai.a2c_smcp.tools.excel.copy_range import ExcelCopyRangeTool
from office4ai.a2c_smcp.tools.excel.delete_range import ExcelDeleteRangeTool
from office4ai.a2c_smcp.tools.excel.get_range import ExcelGetRangeTool
from office4ai.a2c_smcp.tools.excel.get_selected_range import ExcelGetSelectedRangeTool
from office4ai.a2c_smcp.tools.excel.get_workbook_info import ExcelGetWorkbookInfoTool
from office4ai.a2c_smcp.tools.excel.get_worksheet_info import ExcelGetWorksheetInfoTool
from office4ai.a2c_smcp.tools.excel.insert_range import ExcelInsertRangeTool
from office4ai.a2c_smcp.tools.excel.set_formula import ExcelSetFormulaTool
from office4ai.a2c_smcp.tools.excel.set_range import ExcelSetRangeTool

__all__ = [
    # State-awareness read tools (OASP /excel Draft, issue #18 Foundation)
    "ExcelGetWorkbookInfoTool",
    "ExcelGetWorksheetInfoTool",
    "ExcelGetSelectedRangeTool",
    # Range CRUD + 公式 (issue #19)
    "ExcelGetRangeTool",
    "ExcelSetRangeTool",
    "ExcelClearRangeTool",
    "ExcelCopyRangeTool",
    "ExcelDeleteRangeTool",
    "ExcelInsertRangeTool",
    "ExcelSetFormulaTool",
]
