"""
PPT MCP Tools 单元测试 | PPT MCP Tools unit tests

测试策略:
- Mock OfficeWorkspace.execute() 的返回值
- 验证 OfficeAction 构建 (category, event_name, params)
- 验证输入校验 (缺少参数、参数类型错误)
- 验证 format_result hook (获取类 vs 操作类)
"""

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.tools.ppt import (
    PptAddSlideTool,
    PptDeleteElementTool,
    PptDeleteSlideTool,
    PptGetCurrentSlideElementsTool,
    PptGetSlideElementsTool,
    PptGetSlideInfoTool,
    PptGetSlideLayoutsTool,
    PptGetSlideScreenshotTool,
    PptGotoSlideTool,
    PptInsertImageTool,
    PptInsertShapeTool,
    PptInsertTableTool,
    PptInsertTextTool,
    PptMoveSlideTool,
    PptReorderElementTool,
    PptUpdateElementTool,
    PptUpdateImageTool,
    PptUpdateTableCellTool,
    PptUpdateTableFormatTool,
    PptUpdateTableRowColumnTool,
    PptUpdateTextBoxTool,
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
        # Content retrieval tools
        (PptGetCurrentSlideElementsTool, "ppt_get_current_slide_elements", "ppt", "get:currentSlideElements"),
        (PptGetSlideElementsTool, "ppt_get_slide_elements", "ppt", "get:slideElements"),
        (PptGetSlideScreenshotTool, "ppt_get_slide_screenshot", "ppt", "get:slideScreenshot"),
        (PptGetSlideInfoTool, "ppt_get_slide_info", "ppt", "get:slideInfo"),
        (PptGetSlideLayoutsTool, "ppt_get_slide_layouts", "ppt", "get:slideLayouts"),
        # Content insertion tools
        (PptInsertTextTool, "ppt_insert_text", "ppt", "insert:text"),
        (PptInsertImageTool, "ppt_insert_image", "ppt", "insert:image"),
        (PptInsertTableTool, "ppt_insert_table", "ppt", "insert:table"),
        (PptInsertShapeTool, "ppt_insert_shape", "ppt", "insert:shape"),
        # Update operation tools
        (PptUpdateTextBoxTool, "ppt_update_text_box", "ppt", "update:textBox"),
        (PptUpdateImageTool, "ppt_update_image", "ppt", "update:image"),
        (PptUpdateTableCellTool, "ppt_update_table_cell", "ppt", "update:tableCell"),
        (PptUpdateTableRowColumnTool, "ppt_update_table_row_column", "ppt", "update:tableRowColumn"),
        (PptUpdateTableFormatTool, "ppt_update_table_format", "ppt", "update:tableFormat"),
        (PptUpdateElementTool, "ppt_update_element", "ppt", "update:element"),
        # Delete & layout tools
        (PptDeleteElementTool, "ppt_delete_element", "ppt", "delete:element"),
        (PptReorderElementTool, "ppt_reorder_element", "ppt", "reorder:element"),
        # Slide management tools
        (PptAddSlideTool, "ppt_add_slide", "ppt", "add:slide"),
        (PptDeleteSlideTool, "ppt_delete_slide", "ppt", "delete:slide"),
        (PptMoveSlideTool, "ppt_move_slide", "ppt", "move:slide"),
        (PptGotoSlideTool, "ppt_goto_slide", "ppt", "goto:slide"),
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
    async def test_get_current_slide_elements_action(self, mock_workspace):
        """验证 get_current_slide_elements 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"slideIndex": 0, "elements": []})

        tool = PptGetCurrentSlideElementsTool(mock_workspace)
        await tool.execute({"document_uri": "file:///test.pptx"})

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "get:currentSlideElements"
        assert action.params["document_uri"] == "file:///test.pptx"

    @pytest.mark.asyncio
    async def test_get_slide_elements_with_options(self, mock_workspace):
        """验证 get_slide_elements 带选项构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True, data={"slideIndex": 2, "elements": [{"id": "s1"}]}
        )

        tool = PptGetSlideElementsTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 2,
                "options": {"includeText": True, "includeImages": False},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "get:slideElements"
        assert action.params["slideIndex"] == 2
        assert action.params["options"]["include_text"] is True
        assert action.params["options"]["include_images"] is False

    @pytest.mark.asyncio
    async def test_get_slide_screenshot_action(self, mock_workspace):
        """验证 get_slide_screenshot 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"base64": "iVBORw0KGgo=", "format": "png"})

        tool = PptGetSlideScreenshotTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 0,
                "options": {"format": "png"},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "get:slideScreenshot"
        assert action.params["slideIndex"] == 0
        assert action.params["options"]["format"] == "png"

    @pytest.mark.asyncio
    async def test_get_slide_info_action(self, mock_workspace):
        """验证 get_slide_info 带 slideIndex 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"slideCount": 10, "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"}},
        )

        tool = PptGetSlideInfoTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 0,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "get:slideInfo"
        assert action.params["slideIndex"] == 0

    @pytest.mark.asyncio
    async def test_get_slide_layouts_action(self, mock_workspace):
        """验证 get_slide_layouts 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"layouts": [{"name": "Title Slide"}, {"name": "Blank"}]},
        )

        tool = PptGetSlideLayoutsTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "options": {"includePlaceholders": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "get:slideLayouts"
        assert action.params["options"]["include_placeholders"] is True

    @pytest.mark.asyncio
    async def test_insert_text_action(self, mock_workspace):
        """验证 insert_text 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"elementId": "shape-015"})

        tool = PptInsertTextTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "text": "Hello PPT",
                "options": {"slideIndex": 0, "left": 100, "top": 200, "fontSize": 18},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "insert:text"
        assert action.params["text"] == "Hello PPT"
        assert action.params["options"]["slide_index"] == 0

    @pytest.mark.asyncio
    async def test_insert_image_action(self, mock_workspace):
        """验证 insert_image 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"imageId": "shape-025"})

        tool = PptInsertImageTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "image": {"base64": "iVBORw0KGgo="},
                "options": {"slideIndex": 0, "width": 400, "height": 300},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "insert:image"
        assert action.params["image"]["base64"] == "iVBORw0KGgo="
        assert action.params["options"]["width"] == 400

    @pytest.mark.asyncio
    async def test_insert_table_action(self, mock_workspace):
        """验证 insert_table 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"elementId": "shape-030"})

        tool = PptInsertTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "options": {
                    "rows": 3,
                    "columns": 4,
                    "data": [["A", "B", "C", "D"], ["1", "2", "3", "4"], ["5", "6", "7", "8"]],
                },
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "insert:table"
        assert action.params["options"]["rows"] == 3
        assert action.params["options"]["columns"] == 4

    @pytest.mark.asyncio
    async def test_insert_shape_action(self, mock_workspace):
        """验证 insert_shape 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"shapeId": "shape-020"})

        tool = PptInsertShapeTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "shapeType": "RoundedRectangle",
                "options": {"fillColor": "#4472C4", "text": "Click here"},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "insert:shape"
        assert action.params["shapeType"] == "RoundedRectangle"
        assert action.params["options"]["fill_color"] == "#4472C4"

    @pytest.mark.asyncio
    async def test_update_text_box_action(self, mock_workspace):
        """验证 update_text_box 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"elementId": "shape-001"})

        tool = PptUpdateTextBoxTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-001",
                "updates": {"text": "Updated title", "fontSize": 28, "bold": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "update:textBox"
        assert action.params["elementId"] == "shape-001"
        assert action.params["updates"]["text"] == "Updated title"
        assert action.params["updates"]["bold"] is True

    @pytest.mark.asyncio
    async def test_update_image_action(self, mock_workspace):
        """验证 update_image 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"elementId": "shape-025"})

        tool = PptUpdateImageTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-025",
                "image": {"base64": "newBase64Data=="},
                "options": {"keepDimensions": True},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "update:image"
        assert action.params["elementId"] == "shape-025"
        assert action.params["image"]["base64"] == "newBase64Data=="
        assert action.params["options"]["keep_dimensions"] is True

    @pytest.mark.asyncio
    async def test_update_table_cell_action(self, mock_workspace):
        """验证 update_table_cell 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"cellsUpdated": 2})

        tool = PptUpdateTableCellTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-030",
                "cells": [
                    {"rowIndex": 0, "columnIndex": 0, "text": "Name"},
                    {"rowIndex": 0, "columnIndex": 1, "text": "Age"},
                ],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "update:tableCell"
        assert action.params["elementId"] == "shape-030"
        assert len(action.params["cells"]) == 2

    @pytest.mark.asyncio
    async def test_update_table_row_column_action(self, mock_workspace):
        """验证 update_table_row_column 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"cellsUpdated": 8})

        tool = PptUpdateTableRowColumnTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-030",
                "rows": [
                    {"rowIndex": 0, "values": ["Name", "Age", "City", "Job"]},
                    {"rowIndex": 1, "values": ["Alice", "28", "Beijing", "Engineer"]},
                ],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "update:tableRowColumn"
        assert action.params["elementId"] == "shape-030"
        assert len(action.params["rows"]) == 2
        assert action.params["rows"][0]["values"][0] == "Name"

    @pytest.mark.asyncio
    async def test_update_table_format_action(self, mock_workspace):
        """验证 update_table_format 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"cellsFormatted": 5})

        tool = PptUpdateTableFormatTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-030",
                "rowFormats": [{"rowIndex": 0, "backgroundColor": "#4472C4", "fontSize": 14}],
                "cellFormats": [{"rowIndex": 1, "columnIndex": 0, "bold": True, "fontColor": "#333333"}],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "update:tableFormat"
        assert action.params["elementId"] == "shape-030"
        assert len(action.params["rowFormats"]) == 1
        assert len(action.params["cellFormats"]) == 1

    @pytest.mark.asyncio
    async def test_update_element_action(self, mock_workspace):
        """验证 update_element 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"elementId": "shape-015"})

        tool = PptUpdateElementTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-015",
                "slideIndex": 0,
                "updates": {"left": 200, "top": 150, "width": 300, "height": 200},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "update:element"
        assert action.params["elementId"] == "shape-015"
        assert action.params["updates"]["left"] == 200
        assert action.params["slideIndex"] == 0

    @pytest.mark.asyncio
    async def test_delete_element_single(self, mock_workspace):
        """验证 delete_element 单个删除构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deletedCount": 1})

        tool = PptDeleteElementTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-015",
                "slideIndex": 0,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "delete:element"
        assert action.params["elementId"] == "shape-015"
        assert action.params["slideIndex"] == 0

    @pytest.mark.asyncio
    async def test_delete_element_batch(self, mock_workspace):
        """验证 delete_element 批量删除构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deletedCount": 3})

        tool = PptDeleteElementTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementIds": ["shape-015", "shape-016", "shape-017"],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "delete:element"
        assert len(action.params["elementIds"]) == 3

    @pytest.mark.asyncio
    async def test_reorder_element_action(self, mock_workspace):
        """验证 reorder_element 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"zOrder": 5})

        tool = PptReorderElementTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-015",
                "action": "bringToFront",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "reorder:element"
        assert action.params["elementId"] == "shape-015"
        assert action.params["action"] == "bringToFront"

    @pytest.mark.asyncio
    async def test_add_slide_action(self, mock_workspace):
        """验证 add_slide 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"slideIndex": 2, "slideId": "slide-003"})

        tool = PptAddSlideTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "options": {"insertIndex": 2, "layout": "Title Slide"},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "add:slide"
        assert action.params["options"]["insert_index"] == 2
        assert action.params["options"]["layout"] == "Title Slide"

    @pytest.mark.asyncio
    async def test_delete_slide_action(self, mock_workspace):
        """验证 delete_slide 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deleted": True, "totalSlides": 9})

        tool = PptDeleteSlideTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 3,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "delete:slide"
        assert action.params["slideIndex"] == 3

    @pytest.mark.asyncio
    async def test_move_slide_action(self, mock_workspace):
        """验证 move_slide 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True, data={"fromIndex": 0, "toIndex": 3, "totalSlides": 10}
        )

        tool = PptMoveSlideTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "fromIndex": 0,
                "toIndex": 3,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "move:slide"
        assert action.params["fromIndex"] == 0
        assert action.params["toIndex"] == 3

    @pytest.mark.asyncio
    async def test_goto_slide_action(self, mock_workspace):
        """验证 goto_slide 构建正确的 OfficeAction"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"slideIndex": 5})

        tool = PptGotoSlideTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 5,
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "ppt"
        assert action.action_name == "goto:slide"
        assert action.params["slideIndex"] == 5


# ============================================================================
# Input Validation Tests
# ============================================================================


class TestInputValidation:
    """测试输入验证"""

    @pytest.mark.asyncio
    async def test_missing_document_uri(self, mock_workspace):
        """测试缺少 document_uri"""
        tool = PptInsertTextTool(mock_workspace)
        result = await tool.execute({"text": "Hello"})

        assert result["success"] is False
        assert "error" in result
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_required_text(self, mock_workspace):
        """测试 insert_text 缺少 text"""
        tool = PptInsertTextTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.pptx"})

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_required_shape_type(self, mock_workspace):
        """测试 insert_shape 缺少 shapeType"""
        tool = PptInsertShapeTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.pptx"})

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_shape_type(self, mock_workspace):
        """测试 insert_shape 无效的 shapeType"""
        tool = PptInsertShapeTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "shapeType": "InvalidShape",
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_negative_slide_index(self, mock_workspace):
        """测试负数 slideIndex"""
        tool = PptGetSlideElementsTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": -1,
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_element_id_for_update(self, mock_workspace):
        """测试 update_text_box 缺少 elementId"""
        tool = PptUpdateTextBoxTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "updates": {"text": "new text"},
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_invalid_reorder_action(self, mock_workspace):
        """测试 reorder_element 无效的 action"""
        tool = PptReorderElementTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-001",
                "action": "invalidAction",
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_slide_index_for_delete_slide(self, mock_workspace):
        """测试 delete_slide 缺少 slideIndex"""
        tool = PptDeleteSlideTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.pptx"})

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_missing_from_to_index_for_move_slide(self, mock_workspace):
        """测试 move_slide 缺少 fromIndex/toIndex"""
        tool = PptMoveSlideTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "fromIndex": 0,
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_empty_cells_for_update_table_cell(self, mock_workspace):
        """测试 update_table_cell 空 cells 列表"""
        tool = PptUpdateTableCellTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-030",
                "cells": [],
            }
        )

        assert result["success"] is False
        mock_workspace.execute.assert_not_called()


# ============================================================================
# Format Result Tests
# ============================================================================


class TestFormatResult:
    """测试 format_result hook"""

    @pytest.mark.asyncio
    async def test_get_current_slide_elements_format(self, mock_workspace):
        """获取类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"slideIndex": 0, "elements": [{"id": "s1"}, {"id": "s2"}]},
        )

        tool = PptGetCurrentSlideElementsTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.pptx"})

        assert result["success"] is True
        assert "2 element(s)" in result["content"]
        assert "Slide 0" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_slide_elements_format(self, mock_workspace):
        """获取类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"slideIndex": 2, "elements": [{"id": "s1"}]},
        )

        tool = PptGetSlideElementsTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 2,
            }
        )

        assert result["success"] is True
        assert "1 element(s)" in result["content"]
        assert "Slide 2" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_slide_screenshot_format(self, mock_workspace):
        """获取截图类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"base64": "iVBORw0KGgo=", "format": "png"},
        )

        tool = PptGetSlideScreenshotTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 0,
            }
        )

        assert result["success"] is True
        assert "png" in result["content"]
        assert "base64" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_slide_info_format(self, mock_workspace):
        """获取演示信息类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={
                "slideCount": 10,
                "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
            },
        )

        tool = PptGetSlideInfoTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.pptx"})

        assert result["success"] is True
        assert "10 slides" in result["content"]
        assert "960x540" in result["content"]
        assert "16:9" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_slide_layouts_format(self, mock_workspace):
        """获取版式类工具返回 content 字段"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"layouts": [{"name": "Title Slide"}, {"name": "Blank"}, {"name": "Title and Content"}]},
        )

        tool = PptGetSlideLayoutsTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.pptx"})

        assert result["success"] is True
        assert "3 layout(s)" in result["content"]
        assert "Title Slide" in result["content"]
        assert "Blank" in result["content"]
        assert "data" in result

    @pytest.mark.asyncio
    async def test_get_slide_layouts_empty_format(self, mock_workspace):
        """获取版式空列表"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"layouts": []},
        )

        tool = PptGetSlideLayoutsTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///test.pptx"})

        assert result["success"] is True
        assert "No layouts found" in result["content"]

    @pytest.mark.asyncio
    async def test_operation_tool_format(self, mock_workspace):
        """操作类工具返回标准 JSON (data 字段)"""
        mock_workspace.execute.return_value = OfficeObs(
            success=True,
            data={"elementId": "shape-015"},
        )

        tool = PptInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "text": "Hello",
            }
        )

        assert result["success"] is True
        assert result["data"] == {"elementId": "shape-015"}

    @pytest.mark.asyncio
    async def test_error_format(self, mock_workspace):
        """错误返回统一格式"""
        mock_workspace.execute.return_value = OfficeObs(
            success=False,
            data={},
            error="Document not connected: file:///test.pptx",
        )

        tool = PptInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "text": "Hello",
            }
        )

        assert result["success"] is False
        assert "Document not connected" in result["error"]

    @pytest.mark.asyncio
    async def test_workspace_exception(self, mock_workspace):
        """workspace 抛出异常时的处理"""
        mock_workspace.execute.side_effect = TimeoutError("Operation timed out")

        tool = PptInsertTextTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "text": "Hello",
            }
        )

        assert result["success"] is False
        assert "timed out" in result["error"]

    @pytest.mark.asyncio
    async def test_get_format_error(self, mock_workspace):
        """获取类工具错误格式"""
        mock_workspace.execute.return_value = OfficeObs(
            success=False,
            data={},
            error="Slide index out of range",
        )

        tool = PptGetSlideElementsTool(mock_workspace)
        result = await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "slideIndex": 99,
            }
        )

        assert result["success"] is False
        assert "Slide index out of range" in result["error"]


# ============================================================================
# OF4AI-8: elementId int Coercion Tests
# ============================================================================


class TestElementIdIntCoercion:
    """OF4AI-8: LLM 传入 int elementId 时应自动强转为 str"""

    @pytest.mark.asyncio
    async def test_int_element_id_coercion(self, mock_workspace):
        """传入 elementId: 5 (int) 应被强转为 "5" (str)"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deletedCount": 1})

        tool = PptDeleteElementTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": 5,  # int, not str
            }
        )

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.params["elementId"] == "5"
        assert isinstance(action.params["elementId"], str)

    @pytest.mark.asyncio
    async def test_int_element_ids_coercion(self, mock_workspace):
        """传入 elementIds: [5, 3] (int list) 应被强转为 ["5", "3"]"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deletedCount": 2})

        tool = PptDeleteElementTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementIds": [5, 3],  # int list, not str list
            }
        )

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.params["elementIds"] == ["5", "3"]
        assert all(isinstance(x, str) for x in action.params["elementIds"])

    @pytest.mark.asyncio
    async def test_str_element_id_unchanged(self, mock_workspace):
        """传入 elementId: "shape-001" (str) 应保持不变"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"elementId": "shape-001"})

        tool = PptUpdateTextBoxTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": "shape-001",
                "updates": {"text": "Hello"},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["elementId"] == "shape-001"
        assert isinstance(action.params["elementId"], str)

    def test_json_schema_accepts_int(self, mock_workspace):
        """验证 input_schema 的 elementId 包含 integer 类型"""
        tool = PptDeleteElementTool(mock_workspace)
        schema = tool.input_schema
        element_id_schema = schema["properties"]["elementId"]
        # Should have anyOf with both string and integer
        type_strs = _extract_types_from_schema(element_id_schema)
        assert "string" in type_strs
        assert "integer" in type_strs

    @pytest.mark.asyncio
    @pytest.mark.parametrize(
        "tool_cls,extra_params",
        [
            (PptUpdateElementTool, {"updates": {"left": 100}}),
            (PptUpdateImageTool, {"image": {"base64": "abc=="}}),
            (PptUpdateTextBoxTool, {"updates": {"text": "hi"}}),
            (PptReorderElementTool, {"action": "bringToFront"}),
            (PptUpdateTableCellTool, {"cells": [{"rowIndex": 0, "columnIndex": 0, "text": "A"}]}),
            (PptUpdateTableRowColumnTool, {"rows": [{"rowIndex": 0, "values": ["A"]}]}),
            (PptUpdateTableFormatTool, {"rowFormats": [{"rowIndex": 0, "bold": True}]}),
        ],
    )
    async def test_all_tools_int_element_id(self, mock_workspace, tool_cls, extra_params):
        """所有 8 个 PPT 工具都应接受 int elementId 并强转为 str"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = tool_cls(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///test.pptx",
                "elementId": 42,  # int
                **extra_params,
            }
        )

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.params["elementId"] == "42"
        assert isinstance(action.params["elementId"], str)


def _extract_types_from_schema(schema: dict) -> set[str]:
    """从 JSON Schema 中提取所有 type 值"""
    types = set()
    if "type" in schema:
        types.add(schema["type"])
    for key in ("anyOf", "oneOf"):
        if key in schema:
            for item in schema[key]:
                types.update(_extract_types_from_schema(item))
    return types
