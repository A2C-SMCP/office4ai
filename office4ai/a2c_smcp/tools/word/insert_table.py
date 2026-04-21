"""word_insert_table MCP Tool"""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.dtos.word import TableInsertOptions


class WordInsertTableInput(BaseModel):
    """MCP 输入模型: 插入表格"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/doc.docx)")
    options: TableInsertOptions = Field(..., description="Table insertion options (rows, columns, data, style)")


class WordInsertTableTool(BaseTool):
    """在 Word 文档中插入表格"""

    @property
    def name(self) -> str:
        return "word_insert_table"

    @property
    def description(self) -> str:
        return (
            "Insert a table into a Word document. "
            "Specify the number of rows and columns, optionally provide cell data and table style. "
            "Cell data is a 2D array of strings matching the rows x columns dimensions. "
            "Use options.insertLocation to control placement: "
            "'Start'/'End' (document boundaries), 'Before'/'After' (relative to current selection/cursor), "
            "or 'Replace' (replace current selection). Defaults to 'End' when omitted."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return WordInsertTableInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "word"

    @property
    def event_name(self) -> str:
        return "insert:table"

    @property
    def input_model(self) -> type[BaseModel]:
        return WordInsertTableInput
