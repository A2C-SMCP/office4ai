"""
Word MCP Tools 单元测试 | Word MCP Tools unit tests

测试策略:
- Mock OfficeWorkspace.execute() 的返回值
- 验证 OfficeAction 构建 (category, event_name, params)
- 验证输入校验 (缺少参数、参数类型错误)
- 验证 format_result hook (获取类 vs 操作类)
"""

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.tools.word import (
    WordAppendTextTool,
    WordDeleteCommentTool,
    WordExportContentTool,
    WordGetCommentsTool,
    WordGetDocumentStatsTool,
    WordGetDocumentStructureTool,
    WordGetSelectedContentTool,
    WordGetSelectionTool,
    WordGetStylesTool,
    WordGetVisibleContentTool,
    WordInsertCommentTool,
    WordInsertEquationTool,
    WordInsertImageTool,
    WordInsertTableTool,
    WordInsertTextTool,
    WordInsertTOCTool,
    WordMergeCellsTool,
    WordReplaceSelectionTool,
    WordReplaceTextTool,
    WordReplyCommentTool,
    WordResolveCommentTool,
    WordSelectTextTool,
    WordUpdateTableCellTool,
    WordUpdateTableFormatTool,
    WordUpdateTableRowColumnTool,
)
from office4ai.environment.workspace.base import OfficeObs


@pytest.fixture
def mock_workspace():
    """创建 mock OfficeWorkspace"""
    workspace = MagicMock()
    workspace.execute = AsyncMock()
    return workspace


# ============================================================================
# Tool Metadata Tests
# ============================================================================


class TestToolMetadata:
    """测试所有工具的元数据声明"""

    TOOL_SPECS = [
        # Original 9 MVP tools
        (WordGetSelectedContentTool, "word_get_selected_content", "word", "get:selectedContent"),
        (WordGetVisibleContentTool, "word_get_visible_content", "word", "get:visibleContent"),
        (WordInsertTextTool, "word_insert_text", "word", "insert:text"),
        (WordAppendTextTool, "word_append_text", "word", "append:text"),
        (WordReplaceTextTool, "word_replace_text", "word", "replace:text"),
        (WordInsertImageTool, "word_insert_image", "word", "insert:image"),
        (WordInsertTableTool, "word_insert_table", "word", "insert:table"),
        (WordInsertEquationTool, "word_insert_equation", "word", "insert:equation"),
        (WordInsertTOCTool, "word_insert_toc", "word", "insert:toc"),
        # Phase 1: 7 new tools (DTO already exists)
        (WordGetSelectionTool, "word_get_selection", "word", "get:selection"),
        (WordGetDocumentStructureTool, "word_get_document_structure", "word", "get:documentStructure"),
        (WordGetDocumentStatsTool, "word_get_document_stats", "word", "get:documentStats"),
        (WordGetStylesTool, "word_get_styles", "word", "get:styles"),
        (WordReplaceSelectionTool, "word_replace_selection", "word", "replace:selection"),
        (WordSelectTextTool, "word_select_text", "word", "select:text"),
        (WordExportContentTool, "word_export_content", "word", "export:content"),
        # Phase 2: 5 comment tools
        (WordGetCommentsTool, "word_get_comments", "word", "get:comments"),
        (WordInsertCommentTool, "word_insert_comment", "word", "insert:comment"),
        (WordDeleteCommentTool, "word_delete_comment", "word", "delete:comment"),
        (WordReplyCommentTool, "word_reply_comment", "word", "reply:comment"),
        (WordResolveCommentTool, "word_resolve_comment", "word", "resolve:comment"),
        # Phase 3: 4 table operation tools (OASP /word Draft, v0.2.0)
        (WordMergeCellsTool, "word_merge_cells", "word", "merge:cells"),
        (WordUpdateTableCellTool, "word_update_table_cell", "word", "update:tableCell"),
        (WordUpdateTableRowColumnTool, "word_update_table_row_column", "word", "update:tableRowColumn"),
        (WordUpdateTableFormatTool, "word_update_table_format", "word", "update:tableFormat"),
    ]

    @pytest.mark.parametrize("tool_cls,expected_name,expected_category,expected_event", TOOL_SPECS)
    def test_tool_metadata(self, mock_workspace, tool_cls, expected_name, expected_category, expected_event):
        """验证工具元数据 | Verify tool metadata"""
        tool = tool_cls(mock_workspace)
        assert tool.name == expected_name
        assert tool.category == expected_category
        assert tool.event_name == expected_event
        assert isinstance(tool.description, str)
        assert len(tool.description) > 0
        assert isinstance(tool.input_schema, dict)
        assert "properties" in tool.input_schema

    @pytest.mark.parametrize("tool_cls,expected_name,expected_category,expected_event", TOOL_SPECS)
    def test_input_schema_has_document_uri(
        self, mock_workspace, tool_cls, expected_name, expected_category, expected_event
    ):
        """验证所有工具的 input_schema 都包含 document_uri"""
        tool = tool_cls(mock_workspace)
        schema = tool.input_schema
        assert "document_uri" in schema["properties"]


# ============================================================================
# Execute Flow Tests
# ============================================================================


class TestExecuteFlow:
    """测试通用执行流程"""

    @pytest.mark.asyncio
    async def test_insert_text_builds_correct_action(self, mock_workspace):
        """验证 insert_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello World",
            }
        )

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:text"
        assert action.params["document_uri"] == "file:///test.docx"
        assert action.params["text"] == "Hello World"
        assert action.params["location"] == "Cursor"  # default value

    @pytest.mark.asyncio
    async def test_insert_text_with_format(self, mock_workspace):
        """验证带格式的文本插入"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Bold Text",
                "format": {"bold": True, "font_size": 14},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["text"] == "Bold Text"
        assert action.params["format"]["bold"] is True
        assert action.params["format"]["font_size"] == 14

    @pytest.mark.asyncio
    async def test_append_text_builds_correct_action(self, mock_workspace):
        """验证 append_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordAppendTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Appended",
                "location": "Start",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "append:text"
        assert action.params["location"] == "Start"

    @pytest.mark.asyncio
    async def test_replace_text_builds_correct_action(self, mock_workspace):
        """验证 replace_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"replaceCount": 3})

        tool = WordReplaceTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "foo",
                "replace_text": "bar",
                "options": {"match_case": True, "replace_all": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "replace:text"
        assert action.params["search_text"] == "foo"
        assert action.params["replace_text"] == "bar"
        assert action.params["options"]["match_case"] is True
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_replace_text_with_format(self, mock_workspace):
        """验证 replace_text 带格式参数构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"replaceCount": 2})

        tool = WordReplaceTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "important",
                "replace_text": "important",
                "format": {"bold": True, "color": "#FF0000", "font_size": 16},
                "options": {"replace_all": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "replace:text"
        assert action.params["search_text"] == "important"
        assert action.params["replace_text"] == "important"
        assert action.params["format"]["bold"] is True
        assert action.params["format"]["color"] == "#FF0000"
        assert action.params["format"]["font_size"] == 16
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_replace_text_without_format(self, mock_workspace):
        """验证 replace_text 不带格式时 format 为 None"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"replaceCount": 1})

        tool = WordReplaceTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "old",
                "replace_text": "new",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params.get("format") is None

    @pytest.mark.asyncio
    async def test_replace_text_with_style_name_format(self, mock_workspace):
        """验证 replace_text 使用 styleName 格式化"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"replaceCount": 1})

        tool = WordReplaceTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "Chapter Title",
                "replace_text": "Chapter Title",
                "format": {"style_name": "Heading 1"},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["format"]["style_name"] == "Heading 1"

    @pytest.mark.asyncio
    async def test_insert_image_builds_correct_action(self, mock_workspace):
        """验证 insert_image 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertImageTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "image": {"base64": "iVBORw0KGgo=", "width": 100, "height": 100},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:image"
        assert action.params["image"]["base64"] == "iVBORw0KGgo="

    @pytest.mark.asyncio
    async def test_insert_table_builds_correct_action(self, mock_workspace):
        """验证 insert_table 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {"rows": 3, "columns": 2, "data": [["A", "B"], ["C", "D"], ["E", "F"]]},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:table"
        assert action.params["options"]["rows"] == 3
        assert action.params["options"]["columns"] == 2

    @pytest.mark.asyncio
    async def test_insert_table_omitted_insert_location_is_excluded(self, mock_workspace):
        """未传 insertLocation 时不应出现在 params 中（exclude_none），Add-In 侧回退默认 End"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {"rows": 2, "columns": 2},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert "insert_location" not in action.params["options"]

    @pytest.mark.asyncio
    @pytest.mark.parametrize("insert_location", ["Start", "End", "Before", "After", "Replace"])
    async def test_insert_table_passes_through_insert_location(self, mock_workspace, insert_location: str):
        """五种 insertLocation 枚举值均应透传到 OfficeAction.params"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {
                    "rows": 2,
                    "columns": 2,
                    "insertLocation": insert_location,
                },
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["options"]["insert_location"] == insert_location

    @pytest.mark.asyncio
    async def test_insert_equation_builds_correct_action(self, mock_workspace):
        """验证 insert_equation 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertEquationTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "latex": "E = mc^2",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:equation"
        assert action.params["latex"] == "E = mc^2"

    @pytest.mark.asyncio
    async def test_insert_toc_builds_correct_action(self, mock_workspace):
        """验证 insert_toc 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertTOCTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {"max_level": 2},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:toc"
        assert action.params["options"]["max_level"] == 2

    # ── Phase 1 execute flow tests ──

    @pytest.mark.asyncio
    async def test_get_selection_builds_correct_action(self, mock_workspace):
        """验证 get_selection 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"type": "Normal", "start": 10, "end": 20, "text": "hello"},
        )

        tool = WordGetSelectionTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "get:selection"
        assert action.params["document_uri"] == "file:///test.docx"
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_get_document_structure_builds_correct_action(self, mock_workspace):
        """验证 get_document_structure 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"sectionCount": 4, "paragraphCount": 25, "tableCount": 3, "imageCount": 5},
        )

        tool = WordGetDocumentStructureTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "get:documentStructure"
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_get_document_stats_builds_correct_action(self, mock_workspace):
        """验证 get_document_stats 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"wordCount": 1500, "characterCount": 8500, "paragraphCount": 25},
        )

        tool = WordGetDocumentStatsTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "get:documentStats"
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_get_styles_builds_correct_action(self, mock_workspace):
        """验证 get_styles 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"styles": [{"name": "Normal"}, {"name": "Heading 1"}]},
        )

        tool = WordGetStylesTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {"include_built_in": True, "include_unused": False},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "get:styles"
        assert action.params["options"]["include_built_in"] is True
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_replace_selection_builds_correct_action(self, mock_workspace):
        """验证 replace_selection 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"replaced": True, "characterCount": 10},
        )

        tool = WordReplaceSelectionTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "content": {"text": "Replacement text", "format": {"bold": True}},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "replace:selection"
        assert action.params["content"]["text"] == "Replacement text"
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_select_text_builds_correct_action(self, mock_workspace):
        """验证 select_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"matchCount": 3, "selectedIndex": 1, "selectedText": "Hello"},
        )

        tool = WordSelectTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "Hello",
                "search_options": {"match_case": True},
                "selection_mode": "select",
                "select_index": 2,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "select:text"
        assert action.params["search_text"] == "Hello"
        assert action.params["search_options"]["match_case"] is True
        assert action.params["selection_mode"] == "select"
        assert action.params["select_index"] == 2
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_export_content_builds_correct_action(self, mock_workspace):
        """验证 export_content 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"content": "# Document Title\n\nParagraph text..."},
        )

        tool = WordExportContentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "format": "markdown",
                "options": {"include_images": False, "include_tables": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "export:content"
        assert action.params["format"] == "markdown"
        assert action.params["options"]["include_images"] is False
        assert result["success"] is True

    # ── Phase 2 comment execute flow tests ──

    @pytest.mark.asyncio
    async def test_get_comments_builds_correct_action(self, mock_workspace):
        """验证 get_comments 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={
                "comments": [
                    {"id": "c1", "content": "Fix this", "authorName": "Alice", "resolved": False},
                ]
            },
        )

        tool = WordGetCommentsTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "options": {"include_resolved": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "get:comments"
        assert action.params["options"]["include_resolved"] is True
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_insert_comment_builds_correct_action(self, mock_workspace):
        """验证 insert_comment 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordInsertCommentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Please review this section",
                "target": {"type": "searchText", "search_text": "important"},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "insert:comment"
        assert action.params["text"] == "Please review this section"
        assert action.params["target"]["type"] == "searchText"
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_delete_comment_builds_correct_action(self, mock_workspace):
        """验证 delete_comment 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordDeleteCommentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "comment_id": "comment_123",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "delete:comment"
        assert action.params["comment_id"] == "comment_123"
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_reply_comment_builds_correct_action(self, mock_workspace):
        """验证 reply_comment 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordReplyCommentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "comment_id": "comment_123",
                "text": "Done, fixed it",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "reply:comment"
        assert action.params["comment_id"] == "comment_123"
        assert action.params["text"] == "Done, fixed it"
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_resolve_comment_builds_correct_action(self, mock_workspace):
        """验证 resolve_comment 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordResolveCommentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "comment_id": "comment_123",
                "resolved": True,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "resolve:comment"
        assert action.params["comment_id"] == "comment_123"
        assert action.params["resolved"] is True
        assert result["success"] is True


# ============================================================================
# Input Validation Tests
# ============================================================================


class TestInputValidation:
    """测试输入验证"""

    @pytest.mark.asyncio
    async def test_missing_document_uri(self, mock_workspace):
        """测试缺少 document_uri"""
        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute({"text": "Hello"})

        assert result["success"] is False
        assert "error" in result
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_required_field(self, mock_workspace):
        """测试缺少必要字段"""
        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is False
        assert "error" in result
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_replace_text_search_too_long(self, mock_workspace):
        """测试搜索文本超过 255 字符"""
        tool = WordReplaceTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "a" * 256,
                "replace_text": "b",
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_select_text_search_too_long(self, mock_workspace):
        """测试 select_text 搜索文本超过 255 字符"""
        tool = WordSelectTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "a" * 256,
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_export_content_invalid_format(self, mock_workspace):
        """测试 export_content 无效的导出格式"""
        tool = WordExportContentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "format": "pdf",  # not in Literal["text", "html", "markdown"]
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_delete_comment_missing_comment_id(self, mock_workspace):
        """测试 delete_comment 缺少 comment_id"""
        tool = WordDeleteCommentTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_replace_selection_missing_content(self, mock_workspace):
        """测试 replace_selection 缺少 content"""
        tool = WordReplaceSelectionTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()


# ============================================================================
# Format Result Tests
# ============================================================================


class TestFormatResult:
    """测试 format_result hook"""

    @pytest.mark.asyncio
    async def test_get_selected_content_format(self, mock_workspace):
        """获取类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"text": "Selected text content", "elements": []},
        )

        tool = WordGetSelectedContentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
            }
        )

        assert result["success"] is True
        assert result["content"] == "Selected text content"
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_visible_content_format(self, mock_workspace):
        """获取类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"text": "Visible content here", "elements": []},
        )

        tool = WordGetVisibleContentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
            }
        )

        assert result["success"] is True
        assert result["content"] == "Visible content here"

    @pytest.mark.asyncio
    async def test_operation_tool_format(self, mock_workspace):
        """操作类工具返回标准 JSON (data 字段)"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"inserted": True},
        )

        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello",
            }
        )

        assert result["success"] is True
        assert result["data"] == {"inserted": True}

    @pytest.mark.asyncio
    async def test_error_format(self, mock_workspace):
        """错误返回统一格式"""
        mock_workspace.execute.return_value = OfficeObs(
            success=False,
            data={},
            error="Document not connected: file:///test.docx",
        )

        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello",
            }
        )

        assert result["success"] is False
        assert "Document not connected" in result["error"]

    @pytest.mark.asyncio
    async def test_workspace_exception(self, mock_workspace):
        """workspace 抛出异常时的处理"""
        mock_workspace.execute.side_effect = TimeoutError("Operation timed out")

        tool = WordInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Hello",
            }
        )

        assert result["success"] is False
        assert "timed out" in result["error"]

    # ── Phase 1 format_result tests ──

    @pytest.mark.asyncio
    async def test_get_selection_format(self, mock_workspace):
        """get_selection 返回选区摘要"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"type": "Normal", "start": 10, "end": 20, "text": "hello"},
        )

        tool = WordGetSelectionTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is True
        assert "Normal" in result["content"]
        assert "10" in result["content"]
        assert "20" in result["content"]
        assert "hello" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_document_structure_format(self, mock_workspace):
        """get_document_structure 返回结构摘要"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"sectionCount": 4, "paragraphCount": 25, "tableCount": 3, "imageCount": 5},
        )

        tool = WordGetDocumentStructureTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is True
        assert "4 sections" in result["content"]
        assert "25 paragraphs" in result["content"]
        assert "3 tables" in result["content"]
        assert "5 images" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_document_stats_format(self, mock_workspace):
        """get_document_stats 返回统计摘要"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"wordCount": 1500, "characterCount": 8500, "paragraphCount": 25},
        )

        tool = WordGetDocumentStatsTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is True
        assert "1500 words" in result["content"]
        assert "8500 characters" in result["content"]
        assert "25 paragraphs" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_styles_format(self, mock_workspace):
        """get_styles 返回样式名列表"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"styles": [{"name": "Normal"}, {"name": "Heading 1"}, {"name": "Title"}]},
        )

        tool = WordGetStylesTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is True
        assert "Normal" in result["content"]
        assert "Heading 1" in result["content"]
        assert "Title" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_styles_empty_format(self, mock_workspace):
        """get_styles 空样式列表"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"styles": []},
        )

        tool = WordGetStylesTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is True
        assert "No styles found" in result["content"]

    @pytest.mark.asyncio
    async def test_replace_selection_default_format(self, mock_workspace):
        """replace_selection 使用默认 format_result"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"replaced": True, "characterCount": 15},
        )

        tool = WordReplaceSelectionTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "content": {"text": "new text"},
            }
        )

        assert result["success"] is True
        assert result["data"]["replaced"] is True

    @pytest.mark.asyncio
    async def test_select_text_default_format(self, mock_workspace):
        """select_text 使用默认 format_result"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"matchCount": 3, "selectedIndex": 1, "selectedText": "Hello"},
        )

        tool = WordSelectTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "search_text": "Hello",
            }
        )

        assert result["success"] is True
        assert result["data"]["matchCount"] == 3

    @pytest.mark.asyncio
    async def test_export_content_format(self, mock_workspace):
        """export_content 返回导出内容"""
        exported = "# Title\n\nThis is the document content."
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"content": exported},
        )

        tool = WordExportContentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "format": "markdown",
            }
        )

        assert result["success"] is True
        assert result["content"] == exported
        assert "data" in result

    # ── Phase 2 comment format_result tests ──

    @pytest.mark.asyncio
    async def test_get_comments_format_with_comments(self, mock_workspace):
        """get_comments 返回批注摘要"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={
                "comments": [
                    {"id": "c1", "content": "Fix typo", "authorName": "Alice", "resolved": False},
                    {"id": "c2", "content": "Good point", "authorName": "Bob", "resolved": True},
                ]
            },
        )

        tool = WordGetCommentsTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is True
        assert "2 comment(s)" in result["content"]
        assert "Alice" in result["content"]
        assert "Fix typo" in result["content"]
        assert "[resolved]" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_comments_format_empty(self, mock_workspace):
        """get_comments 空批注列表"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"comments": []},
        )

        tool = WordGetCommentsTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.docx"})

        assert result["success"] is True
        assert "No comments found" in result["content"]

    @pytest.mark.asyncio
    async def test_insert_comment_default_format(self, mock_workspace):
        """insert_comment 使用默认 format_result"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"commentId": "new_comment_123"},
        )

        tool = WordInsertCommentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "text": "Review needed",
            }
        )

        assert result["success"] is True
        assert result["data"]["commentId"] == "new_comment_123"

    @pytest.mark.asyncio
    async def test_resolve_comment_default_format(self, mock_workspace):
        """resolve_comment 使用默认 format_result 且 resolved 默认为 True"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = WordResolveCommentTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "comment_id": "c1",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["resolved"] is True  # default value
        assert result["success"] is True


# ============================================================================
# OF4AI-8: comment_id int Coercion Tests
# ============================================================================


class TestCommentIdIntCoercion:
    """OF4AI-8: LLM 传入 int comment_id 时应自动强转为 str"""

    @pytest.mark.asyncio
    @pytest.mark.parametrize(
        "tool_cls,extra_params",
        [
            (WordDeleteCommentTool, {}),
            (WordReplyCommentTool, {"text": "Reply text"}),
            (WordResolveCommentTool, {}),
        ],
    )
    async def test_int_comment_id_coercion(self, mock_workspace, tool_cls, extra_params):
        """传入 comment_id: 42 (int) 应被强转为 "42" (str)"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = tool_cls(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "comment_id": 42,  # int, not str
                **extra_params,
            }
        )

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.params["comment_id"] == "42"
        assert isinstance(action.params["comment_id"], str)


# ============================================================================
# Word Table Operation Tools (OASP /word Draft, v0.2.0)
# Covers: word_merge_cells, word_update_table_cell,
#         word_update_table_row_column, word_update_table_format
# ============================================================================


class TestWordMergeCellsTool:
    """word_merge_cells normal path + tableId resolution + error code propagation."""

    @pytest.mark.asyncio
    async def test_merge_cells_with_explicit_table_id(self, mock_workspace):
        """显式传 table_id → action.params 透传 tableId 与四个边界索引"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"tableId": "table-0", "requestedRange": {"rowCount": 1, "columnCount": 5}},
        )

        tool = WordMergeCellsTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "table_id": "table-0",
                "start_row_index": 0,
                "start_column_index": 0,
                "end_row_index": 0,
                "end_column_index": 4,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "merge:cells"
        assert action.params["table_id"] == "table-0"
        assert action.params["start_row_index"] == 0
        assert action.params["end_column_index"] == 4
        assert result["success"] is True
        assert result["data"]["requestedRange"] == {"rowCount": 1, "columnCount": 5}

    @pytest.mark.asyncio
    async def test_merge_cells_omitted_table_id_excluded_from_params(self, mock_workspace):
        """缺省 table_id 时应通过 exclude_none 从 params 中省略，由 Add-In 解析当前光标表格"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"tableId": "table-0", "requestedRange": {"rowCount": 1, "columnCount": 3}},
        )

        tool = WordMergeCellsTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "start_row_index": 0,
                "start_column_index": 0,
                "end_row_index": 0,
                "end_column_index": 2,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert "table_id" not in action.params

    @pytest.mark.asyncio
    async def test_merge_cells_negative_index_rejected(self, mock_workspace):
        """负数行/列索引应被 Pydantic 拦截 (ge=0)，不会触达 workspace.execute"""
        tool = WordMergeCellsTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "start_row_index": -1,
                "start_column_index": 0,
                "end_row_index": 0,
                "end_column_index": 2,
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    @pytest.mark.parametrize(
        "code,scenario",
        [
            ("3010", "tableId 在文档中未找到"),
            ("3013", "缺省 tableId + 光标不在表格内"),
            ("3014", "目标区域已存在合并冲突"),
        ],
    )
    async def test_merge_cells_propagates_error_codes(self, mock_workspace, code, scenario):
        """3010 / 3013 / 3014 等错误码应通过 OfficeObs.error 透传到工具返回值"""
        mock_workspace.execute.return_value = OfficeObs(
            success=False,
            data={},
            error=f"{code} {scenario}",
        )

        tool = WordMergeCellsTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "start_row_index": 0,
                "start_column_index": 0,
                "end_row_index": 0,
                "end_column_index": 2,
            }
        )

        assert result["success"] is False
        assert code in result["error"]


class TestWordUpdateTableCellTool:
    """word_update_table_cell normal path + CellFormat enum validation + missing format/text rejection."""

    @pytest.mark.asyncio
    async def test_update_table_cell_with_format(self, mock_workspace):
        """正常路径：单元格文本 + 完整 CellFormat（蓝底居中白字加粗表头场景）"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"tableId": "table-0", "cellsUpdated": 1, "rowCount": 5, "columnCount": 4},
        )

        tool = WordUpdateTableCellTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "table_id": "table-0",
                "cells": [
                    {
                        "rowIndex": 0,
                        "columnIndex": 0,
                        "text": "甲方信息",
                        "format": {
                            "horizontalAlignment": "Centered",
                            "verticalAlignment": "Center",
                            "backgroundColor": "#1F4E79",
                            "fontColor": "#FFFFFF",
                            "bold": True,
                        },
                    }
                ],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "update:tableCell"
        assert action.params["table_id"] == "table-0"
        cell = action.params["cells"][0]
        assert cell["row_index"] == 0
        assert cell["text"] == "甲方信息"
        # snake_case is expected internally because DTO field names are snake_case
        assert cell["format"]["horizontal_alignment"] == "Centered"
        assert cell["format"]["background_color"] == "#1F4E79"
        assert cell["format"]["bold"] is True
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_update_table_cell_rejects_invalid_alignment_value(self, mock_workspace):
        """horizontalAlignment 必须是 Word.Alignment 枚举值: 'Center' 应被拒绝, 'Centered' 才合法"""
        tool = WordUpdateTableCellTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "cells": [
                    {
                        "rowIndex": 0,
                        "columnIndex": 0,
                        "format": {"horizontalAlignment": "Center"},  # ✗ 应该是 Centered
                    }
                ],
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_table_cell_accepts_justified_alignment(self, mock_workspace):
        """horizontalAlignment='Justified' 是 Word.Alignment 合法值"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"tableId": "table-0", "cellsUpdated": 1, "rowCount": 5, "columnCount": 4},
        )

        tool = WordUpdateTableCellTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "cells": [
                    {
                        "rowIndex": 1,
                        "columnIndex": 1,
                        "text": "long paragraph",
                        "format": {"horizontalAlignment": "Justified"},
                    }
                ],
            }
        )
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_update_table_cell_rejects_invalid_vertical_alignment(self, mock_workspace):
        """verticalAlignment 必须是 'Top' | 'Center' | 'Bottom'，'Middle' 应被拒绝"""
        tool = WordUpdateTableCellTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "cells": [
                    {
                        "rowIndex": 0,
                        "columnIndex": 0,
                        "format": {"verticalAlignment": "Middle"},  # ✗ 应该是 Center
                    }
                ],
            }
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_table_cell_empty_cells_rejected(self, mock_workspace):
        """cells 必须 min_length=1，空数组应被拒绝"""
        tool = WordUpdateTableCellTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "cells": [],
            }
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_table_cell_propagates_3010(self, mock_workspace):
        """tableId 不存在 → 3010 ELEMENT_NOT_FOUND 通过 error 透传"""
        mock_workspace.execute.return_value = OfficeObs(success=False, data={}, error="3010 tableId not found")
        tool = WordUpdateTableCellTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "table_id": "table-99",
                "cells": [{"rowIndex": 0, "columnIndex": 0, "text": "x"}],
            }
        )
        assert result["success"] is False
        assert "3010" in result["error"]


class TestWordUpdateTableRowColumnTool:
    """word_update_table_row_column 正常路径 + rows/columns 必须至少二选一."""

    @pytest.mark.asyncio
    async def test_update_table_row_column_with_rows(self, mock_workspace):
        """正常路径：批量按行写入 4 行数据"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"tableId": "table-0", "cellsUpdated": 8, "rowCount": 5, "columnCount": 4},
        )

        tool = WordUpdateTableRowColumnTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "table_id": "table-0",
                "rows": [
                    {"rowIndex": 1, "values": ["甲方", "ACME Corp"]},
                    {"rowIndex": 2, "values": ["地址", "上海市"]},
                    {"rowIndex": 3, "values": ["联系人", "张三"]},
                    {"rowIndex": 4, "values": ["日期", "2026-04-30"]},
                ],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "update:tableRowColumn"
        assert len(action.params["rows"]) == 4
        assert action.params["rows"][0]["values"] == ["甲方", "ACME Corp"]
        # columns 缺省时应通过 exclude_none 省略
        assert "columns" not in action.params
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_update_table_row_column_with_columns(self, mock_workspace):
        """正常路径：按列写入"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"tableId": "table-0", "cellsUpdated": 3, "rowCount": 5, "columnCount": 4},
        )

        tool = WordUpdateTableRowColumnTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "columns": [{"columnIndex": 0, "values": ["甲方", "地址", "联系人"]}],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert "rows" not in action.params
        assert action.params["columns"][0]["column_index"] == 0
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_update_table_row_column_requires_rows_or_columns(self, mock_workspace):
        """rows 和 columns 都缺省应被 model_validator 拦截"""
        tool = WordUpdateTableRowColumnTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
            }
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()


class TestWordUpdateTableFormatTool:
    """word_update_table_format 正常路径 + Word.Alignment 枚举校验 + 错误码透传."""

    @pytest.mark.asyncio
    async def test_update_table_format_full(self, mock_workspace):
        """正常路径：内边距 + 边框 + 列宽 + 整表对齐"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"tableId": "table-0", "rowCount": 5, "columnCount": 4},
        )

        tool = WordUpdateTableFormatTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "table_id": "table-0",
                "style_options": {
                    "cellPadding": {"top": 4, "bottom": 4, "left": 6, "right": 6},
                },
                "border_options": {
                    "location": "inside",
                    "style": "Single",
                    "width": 0.5,
                },
                "column_widths": [120, 80, 80, 80],
                "alignment": "Centered",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "update:tableFormat"
        assert action.params["alignment"] == "Centered"
        assert action.params["column_widths"] == [120, 80, 80, 80]
        assert action.params["border_options"]["location"] == "inside"
        assert action.params["style_options"]["cell_padding"]["top"] == 4
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_update_table_format_rejects_invalid_alignment(self, mock_workspace):
        """alignment='Center' 应被拒绝（必须用 'Centered'）"""
        tool = WordUpdateTableFormatTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "alignment": "Center",  # ✗
            }
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_table_format_rejects_invalid_border_location(self, mock_workspace):
        """borderOptions.location 必须是 'all' | 'inside' | 'outside'"""
        tool = WordUpdateTableFormatTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "border_options": {"location": "diagonal"},
            }
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_table_format_rejects_zero_border_width(self, mock_workspace):
        """borderOptions.width 必须 > 0"""
        tool = WordUpdateTableFormatTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "border_options": {"width": 0},
            }
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_table_format_propagates_3011_style_not_found(self, mock_workspace):
        """styleType 不存在 → 3011 STYLE_NOT_FOUND 通过 error 透传"""
        mock_workspace.execute.return_value = OfficeObs(success=False, data={}, error="3011 The style does not exist.")

        tool = WordUpdateTableFormatTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "style_options": {"styleType": "NoSuchStyle"},
            }
        )
        assert result["success"] is False
        assert "3011" in result["error"]

    @pytest.mark.asyncio
    async def test_update_table_format_propagates_4002_invalid_column_widths(self, mock_workspace):
        """columnWidths 长度超出列数 → 4002 INVALID_PARAM 通过 error 透传 (Add-In 校验)"""
        mock_workspace.execute.return_value = OfficeObs(
            success=False, data={}, error="4002 columnWidths length exceeds column count"
        )

        tool = WordUpdateTableFormatTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.docx",
                "column_widths": [60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60],
            }
        )
        assert result["success"] is False
        assert "4002" in result["error"]
