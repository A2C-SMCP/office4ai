"""Word MCP 工具集合 | Word MCP tools."""

from office4ai.a2c_smcp.tools.word.append_text import WordAppendTextTool
from office4ai.a2c_smcp.tools.word.delete_comment import WordDeleteCommentTool
from office4ai.a2c_smcp.tools.word.export_content import WordExportContentTool
from office4ai.a2c_smcp.tools.word.get_comments import WordGetCommentsTool
from office4ai.a2c_smcp.tools.word.get_document_stats import WordGetDocumentStatsTool
from office4ai.a2c_smcp.tools.word.get_document_structure import WordGetDocumentStructureTool
from office4ai.a2c_smcp.tools.word.get_selected_content import WordGetSelectedContentTool
from office4ai.a2c_smcp.tools.word.get_selection import WordGetSelectionTool
from office4ai.a2c_smcp.tools.word.get_styles import WordGetStylesTool
from office4ai.a2c_smcp.tools.word.get_visible_content import WordGetVisibleContentTool
from office4ai.a2c_smcp.tools.word.insert_comment import WordInsertCommentTool
from office4ai.a2c_smcp.tools.word.insert_equation import WordInsertEquationTool
from office4ai.a2c_smcp.tools.word.insert_image import WordInsertImageTool
from office4ai.a2c_smcp.tools.word.insert_table import WordInsertTableTool
from office4ai.a2c_smcp.tools.word.insert_text import WordInsertTextTool
from office4ai.a2c_smcp.tools.word.insert_toc import WordInsertTOCTool
from office4ai.a2c_smcp.tools.word.merge_cells import WordMergeCellsTool
from office4ai.a2c_smcp.tools.word.replace_selection import WordReplaceSelectionTool
from office4ai.a2c_smcp.tools.word.replace_text import WordReplaceTextTool
from office4ai.a2c_smcp.tools.word.reply_comment import WordReplyCommentTool
from office4ai.a2c_smcp.tools.word.resolve_comment import WordResolveCommentTool
from office4ai.a2c_smcp.tools.word.select_text import WordSelectTextTool
from office4ai.a2c_smcp.tools.word.update_table_cell import WordUpdateTableCellTool
from office4ai.a2c_smcp.tools.word.update_table_format import WordUpdateTableFormatTool
from office4ai.a2c_smcp.tools.word.update_table_row_column import WordUpdateTableRowColumnTool

__all__ = [
    # Get tools
    "WordGetSelectedContentTool",
    "WordGetVisibleContentTool",
    "WordGetSelectionTool",
    "WordGetDocumentStructureTool",
    "WordGetDocumentStatsTool",
    "WordGetStylesTool",
    # Text operation tools
    "WordInsertTextTool",
    "WordAppendTextTool",
    "WordReplaceTextTool",
    "WordReplaceSelectionTool",
    "WordSelectTextTool",
    # Multimedia tools
    "WordInsertImageTool",
    "WordInsertTableTool",
    "WordInsertEquationTool",
    "WordInsertTOCTool",
    # Table operation tools (OASP /word Draft, v0.2.0)
    "WordMergeCellsTool",
    "WordUpdateTableCellTool",
    "WordUpdateTableRowColumnTool",
    "WordUpdateTableFormatTool",
    # Export tool
    "WordExportContentTool",
    # Comment tools
    "WordGetCommentsTool",
    "WordInsertCommentTool",
    "WordDeleteCommentTool",
    "WordReplyCommentTool",
    "WordResolveCommentTool",
]
