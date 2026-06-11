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
    PptGetChartTool,
    PptGetCurrentSlideElementsTool,
    PptGetSlideElementsTool,
    PptGetSlideInfoTool,
    PptGetSlideLayoutsTool,
    PptGetSlideScreenshotTool,
    PptGotoSlideTool,
    PptInsertChartTool,
    PptInsertImageTool,
    PptInsertShapeTool,
    PptInsertTableTool,
    PptInsertTextTool,
    PptMoveSlideTool,
    PptReorderElementTool,
    PptUpdateChartTool,
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
        # Chart tools (OASP /ppt Draft, v0.2.0 — Server OOXML)
        (PptInsertChartTool, "ppt_insert_chart", "ppt", "insert:chart"),
        (PptGetChartTool, "ppt_get_chart", "ppt", "get:chart"),
        (PptUpdateChartTool, "ppt_update_chart", "ppt", "update:chart"),
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


# ============================================================================
# Chart Tool Tests (OASP /ppt Draft, v0.2.0 — Server OOXML path)
# ============================================================================
# These tools override BaseTool.execute() to bypass workspace.execute() and
# instead drive the OOXML chart engine directly (because Office.js does not
# expose chart APIs). Tests use a real .pptx fixture so the engine path is
# exercised end-to-end.


class TestChartToolExecute:
    """End-to-end execute() tests for the 3 chart tools (real .pptx fixture)."""

    @pytest.fixture
    def deck_uri(self, tmp_path):
        from pptx import Presentation

        path = tmp_path / "deck.pptx"
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[5])
        prs.slides.add_slide(prs.slide_layouts[5])
        prs.save(str(path))
        return path.as_uri()

    @pytest.fixture
    def workspace_with_notify(self):
        # Default: no Add-In holding the file → chart writes are allowed.
        # Tests that exercise the CONNECTED-reject path override this per-test.
        from office4ai.environment.workspace.base import DocumentStatus

        ws = MagicMock()
        ws.notify_resource_updated = MagicMock()
        ws.update_last_activity = MagicMock()
        ws.get_document_status = MagicMock(return_value=DocumentStatus.DISCONNECTED)
        return ws

    @pytest.mark.asyncio
    async def test_insert_chart_returns_element_id_and_fires_notify(self, workspace_with_notify, deck_uri):
        tool = PptInsertChartTool(workspace_with_notify)
        result = await tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "ColumnClustered",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                    "title": "T",
                },
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is True
        assert result["data"]["elementId"].startswith("oasp-chart-")
        assert result["data"]["requiresReload"] is True
        # Server-OOXML mutations must trigger MCP resource_updated notifications.
        workspace_with_notify.notify_resource_updated.assert_called_once_with(
            ["window://office4ai/ppt", "window://office4ai"]
        )
        workspace_with_notify.update_last_activity.assert_called_once()

    @pytest.mark.asyncio
    async def test_insert_scatter_chart(self, workspace_with_notify, deck_uri):
        tool = PptInsertChartTool(workspace_with_notify)
        result = await tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "Scatter",
                    "series": [{"name": "ads", "points": [{"x": 1, "y": 2}, {"x": 3, "y": 4}]}],
                },
            }
        )
        assert result["success"] is True
        assert result["data"]["chartType"] == "Scatter"

    @pytest.mark.asyncio
    async def test_insert_dimension_mismatch_returns_3015(self, workspace_with_notify, deck_uri):
        tool = PptInsertChartTool(workspace_with_notify)
        result = await tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "Pie",
                    "categories": ["A", "B", "C"],
                    "series": [{"name": "x", "values": [1, 2]}],  # 2 != 3
                },
            }
        )
        assert result["success"] is False
        assert "3015" in result["error"]
        # Failed insert should NOT fire reload notification.
        workspace_with_notify.notify_resource_updated.assert_not_called()

    @pytest.mark.asyncio
    async def test_insert_invalid_chart_type_rejected_at_validation(self, workspace_with_notify, deck_uri):
        tool = PptInsertChartTool(workspace_with_notify)
        result = await tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {"chartType": "Bogus", "categories": [], "series": []},
            }
        )
        assert result["success"] is False
        assert "error" in result

    @pytest.mark.asyncio
    async def test_get_chart_returns_data_without_notify(self, workspace_with_notify, deck_uri):
        # Insert first to obtain an elementId.
        i_tool = PptInsertChartTool(workspace_with_notify)
        ins = await i_tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "Line",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                    "title": "Trend",
                },
            }
        )
        eid = ins["data"]["elementId"]
        workspace_with_notify.notify_resource_updated.reset_mock()

        g_tool = PptGetChartTool(workspace_with_notify)
        result = await g_tool.execute({"document_uri": deck_uri, "elementId": eid})
        assert result["success"] is True
        assert result["data"]["chart"]["chartType"] == "Line"
        assert result["data"]["chart"]["title"] == "Trend"
        # Read-only — must not fire reload notification.
        workspace_with_notify.notify_resource_updated.assert_not_called()

    @pytest.mark.asyncio
    async def test_get_chart_int_element_id_coerced(self, workspace_with_notify, deck_uri):
        # Insert and grab the numeric portion.
        i_tool = PptInsertChartTool(workspace_with_notify)
        ins = await i_tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "Pie",
                    "categories": ["A"],
                    "series": [{"name": "x", "values": [1]}],
                },
            }
        )
        eid = ins["data"]["elementId"]  # "chart-N"
        # Tool input model coerces int elementId — but here we test str works (default path).
        g_tool = PptGetChartTool(workspace_with_notify)
        result = await g_tool.execute({"document_uri": deck_uri, "elementId": eid})
        assert result["success"] is True

    @pytest.mark.asyncio
    async def test_get_chart_unknown_element_returns_3010(self, workspace_with_notify, deck_uri):
        g_tool = PptGetChartTool(workspace_with_notify)
        result = await g_tool.execute({"document_uri": deck_uri, "elementId": "chart-99999"})
        assert result["success"] is False
        assert "3010" in result["error"]

    @pytest.mark.asyncio
    async def test_update_chart_title_only(self, workspace_with_notify, deck_uri):
        i_tool = PptInsertChartTool(workspace_with_notify)
        ins = await i_tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "BarClustered",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                    "title": "Old",
                },
            }
        )
        eid = ins["data"]["elementId"]
        workspace_with_notify.notify_resource_updated.reset_mock()

        u_tool = PptUpdateChartTool(workspace_with_notify)
        result = await u_tool.execute(
            {
                "document_uri": deck_uri,
                "elementId": eid,
                "chart": {"chartType": "BarClustered", "title": "New"},
            }
        )
        assert result["success"] is True
        assert "title" in result["data"]["updatedFields"]
        assert result["data"]["requiresReload"] is True
        workspace_with_notify.notify_resource_updated.assert_called_once()

    @pytest.mark.asyncio
    async def test_update_chart_dimension_mismatch_returns_3015(self, workspace_with_notify, deck_uri):
        i_tool = PptInsertChartTool(workspace_with_notify)
        ins = await i_tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "ColumnClustered",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                },
            }
        )
        eid = ins["data"]["elementId"]

        u_tool = PptUpdateChartTool(workspace_with_notify)
        result = await u_tool.execute(
            {
                "document_uri": deck_uri,
                "elementId": eid,
                "chart": {
                    "chartType": "ColumnClustered",
                    "categories": ["X", "Y", "Z"],
                    "series": [{"name": "t", "values": [10, 20]}],  # 2 != 3
                },
            }
        )
        assert result["success"] is False
        assert "3015" in result["error"]

    @pytest.mark.asyncio
    async def test_update_cross_variant_returns_new_element_id(self, workspace_with_notify, deck_uri):
        i_tool = PptInsertChartTool(workspace_with_notify)
        ins = await i_tool.execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "ColumnClustered",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                },
            }
        )
        old_eid = ins["data"]["elementId"]

        u_tool = PptUpdateChartTool(workspace_with_notify)
        result = await u_tool.execute(
            {
                "document_uri": deck_uri,
                "elementId": old_eid,
                "chart": {
                    "chartType": "Scatter",
                    "series": [{"name": "p", "points": [{"x": 1, "y": 2}, {"x": 3, "y": 4}]}],
                },
            }
        )
        assert result["success"] is True
        assert result["data"]["chartType"] == "Scatter"
        assert "chartType" in result["data"]["updatedFields"]
        # Recreate preserves the opaque id (written to cNvPr/@name); OASP returns the latest elementId.
        new_eid = result["data"]["elementId"]
        assert new_eid.startswith("oasp-chart-")
        assert new_eid == old_eid

    @pytest.mark.asyncio
    async def test_update_unknown_element_returns_3010(self, workspace_with_notify, deck_uri):
        u_tool = PptUpdateChartTool(workspace_with_notify)
        result = await u_tool.execute(
            {
                "document_uri": deck_uri,
                "elementId": "chart-99999",
                "chart": {"chartType": "ColumnClustered", "title": "x"},
            }
        )
        assert result["success"] is False
        assert "3010" in result["error"]

    @pytest.mark.asyncio
    async def test_insert_missing_document_returns_3001(self, workspace_with_notify):
        tool = PptInsertChartTool(workspace_with_notify)
        result = await tool.execute(
            {
                "document_uri": "file:///does/not/exist.pptx",
                "chart": {
                    "chartType": "Line",
                    "categories": ["A"],
                    "series": [{"name": "x", "values": [1]}],
                },
            }
        )
        assert result["success"] is False
        assert "3001" in result["error"]


# ============================================================================
# Dual-path routing (#15, OASP 0.3.0) — route by DocumentStatus
# ============================================================================
# CONNECTED → path B (client round-trip: ppt:get:slideOoxml / ppt:insert:slidesOoxml,
# all chart OOXML built Server-side). DISCONNECTED → path A (on-disk python-pptx).
# Because the Add-In carrier-event handlers ship later (#16), a CONNECTED path-B
# attempt currently fails (timeout / 3016) → reactive degradation: writes surface
# the flipped 3003 guidance, reads fall back to the on-disk read.


def _disconnected_ws():
    from office4ai.environment.workspace.base import DocumentStatus

    ws = MagicMock()
    ws.notify_resource_updated = MagicMock()
    ws.update_last_activity = MagicMock()
    ws.get_document_status = MagicMock(return_value=DocumentStatus.DISCONNECTED)
    return ws


def _connected_ws(emit_side_effect=None):
    """A CONNECTED workspace whose emit_to_document runs ``emit_side_effect``.

    ``emit_side_effect`` may be a plain sync function returning the wire response
    dict (AsyncMock awaits the call and yields that value), or an exception to raise.
    """
    from office4ai.environment.workspace.base import DocumentStatus

    ws = MagicMock()
    ws.notify_resource_updated = MagicMock()
    ws.update_last_activity = MagicMock()
    ws.get_document_status = MagicMock(return_value=DocumentStatus.CONNECTED)
    ws.emit_to_document = AsyncMock(side_effect=emit_side_effect)
    return ws


def _blank_slide_b64():
    """A base64 single-slide .pptx package (blank layout) — the 'exported live slide'."""
    import base64
    import io

    from pptx import Presentation

    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])  # blank
    buf = io.BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _make_path_b_emit(slide_b64, slide_id="sid-export"):
    """Sync side_effect emulating a successful path-B round-trip (AsyncMock awaits it)."""

    def _emit(document_uri, event, data):
        if event == "ppt:get:slideOoxml":
            return {
                "success": True,
                "data": {"slideIndex": data["slideIndex"], "slideId": slide_id, "base64": slide_b64},
            }
        if event == "ppt:insert:slidesOoxml":
            final = data.get("finalSlideIndex", data.get("targetSlideIndex", 0))
            return {"success": True, "data": {"insertedSlideIndices": [final], "insertedSlideIds": ["sid-new"]}}
        raise AssertionError(f"unexpected path-B event: {event!r}")

    return _emit


def _emit_returns_3016(document_uri, event, data):
    return {"success": False, "error": {"code": "3016", "message": "PowerPointApi 1.8 unavailable"}}


async def _insert_on_disk(deck_uri, slide_index=0, chart_type="Line"):
    """Insert a chart via the on-disk (DISCONNECTED) path; return its elementId."""
    ins = await PptInsertChartTool(_disconnected_ws()).execute(
        {
            "document_uri": deck_uri,
            "chart": {"chartType": chart_type, "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
            "options": {"slideIndex": slide_index},
        }
    )
    assert ins["success"], ins
    return ins["data"]["elementId"]


class TestChartToolDualPathRouting:
    """#15: CONNECTED → client round-trip (path B); DISCONNECTED → on-disk (path A);
    path B failure → reactive degradation."""

    @pytest.fixture
    def deck_uri(self, tmp_path):
        from pptx import Presentation

        path = tmp_path / "deck.pptx"
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[5])
        prs.slides.add_slide(prs.slide_layouts[5])
        prs.save(str(path))
        return path.as_uri()

    # -- DISCONNECTED → on-disk, no client round-trip ----------------------------

    @pytest.mark.asyncio
    async def test_disconnected_insert_uses_disk_and_does_not_emit(self, deck_uri):
        ws = _disconnected_ws()
        ws.emit_to_document = AsyncMock()
        result = await PptInsertChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "chart": {"chartType": "Pie", "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is True
        # On-disk write → the deck must be reopened to render.
        assert result["data"]["requiresReload"] is True
        ws.emit_to_document.assert_not_awaited()

    # -- CONNECTED → path B happy path (mocked carrier events) -------------------

    @pytest.mark.asyncio
    async def test_insert_routes_to_path_b_when_connected(self, deck_uri):
        ws = _connected_ws(_make_path_b_emit(_blank_slide_b64()))
        result = await PptInsertChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "chart": {
                    "chartType": "ColumnClustered",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                },
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is True
        assert result["data"]["elementId"].startswith("oasp-chart-")
        assert result["data"]["slideIndex"] == 0
        # Live round-trip already updated the open document → no reopen needed.
        assert result["data"]["requiresReload"] is False
        # Path B drove the two generic carrier events, export then in-place re-insert.
        events = [call.args[1] for call in ws.emit_to_document.call_args_list]
        assert events == ["ppt:get:slideOoxml", "ppt:insert:slidesOoxml"]
        apply_payload = ws.emit_to_document.call_args_list[1].args[2]
        assert apply_payload["formatting"] == "keepSourceFormatting"
        assert apply_payload["replaceSlideId"] == "sid-export"
        assert apply_payload["finalSlideIndex"] == 0
        ws.notify_resource_updated.assert_called_once()

    @pytest.mark.asyncio
    async def test_get_routes_to_path_b_when_connected_with_hint(self, deck_uri):
        from office4ai.environment.workspace.dtos.ppt import CategoricalChartData
        from office4ai.environment.workspace.services import chart_engine

        # Build a live single-slide package carrying a known chart.
        built = await chart_engine.generate_chart_slide_base64(
            CategoricalChartData.model_validate(
                {
                    "chartType": "Line",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                    "title": "Live",
                }
            )
        )
        ws = _connected_ws(_make_path_b_emit(built["slideBase64"]))
        result = await PptGetChartTool(ws).execute(
            {"document_uri": deck_uri, "elementId": built["elementId"], "slideIndex": 1}
        )
        assert result["success"] is True
        assert result["data"]["chart"]["chartType"] == "Line"
        assert result["data"]["chart"]["title"] == "Live"
        # Reports the deck slide index from the hint, not the mini-package's 0.
        assert result["data"]["slideIndex"] == 1
        events = [call.args[1] for call in ws.emit_to_document.call_args_list]
        assert events == ["ppt:get:slideOoxml"]

    @pytest.mark.asyncio
    async def test_update_routes_to_path_b_when_connected_with_hint(self, deck_uri):
        from office4ai.environment.workspace.dtos.ppt import CategoricalChartData
        from office4ai.environment.workspace.services import chart_engine

        built = await chart_engine.generate_chart_slide_base64(
            CategoricalChartData.model_validate(
                {
                    "chartType": "Line",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                    "title": "Old",
                }
            )
        )
        ws = _connected_ws(_make_path_b_emit(built["slideBase64"]))
        result = await PptUpdateChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "elementId": built["elementId"],
                "slideIndex": 0,
                "chart": {"chartType": "Line", "title": "New"},
            }
        )
        assert result["success"] is True
        assert "title" in result["data"]["updatedFields"]
        assert result["data"]["requiresReload"] is False
        events = [call.args[1] for call in ws.emit_to_document.call_args_list]
        assert events == ["ppt:get:slideOoxml", "ppt:insert:slidesOoxml"]
        ws.notify_resource_updated.assert_called_once()

    # -- Reactive degradation: writes → 3003, reads → on-disk fallback -----------

    @pytest.mark.asyncio
    async def test_insert_degrades_when_path_b_times_out(self, deck_uri):
        ws = _connected_ws(TimeoutError("no ack from Add-In"))
        result = await PptInsertChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "chart": {"chartType": "Pie", "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is False
        assert "3003" in result["error"]
        ws.notify_resource_updated.assert_not_called()
        ws.update_last_activity.assert_not_called()

    @pytest.mark.asyncio
    async def test_insert_degrades_when_path_b_returns_3016(self, deck_uri):
        ws = _connected_ws(_emit_returns_3016)
        result = await PptInsertChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "chart": {"chartType": "Pie", "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is False
        assert "3003" in result["error"]

    @pytest.mark.asyncio
    async def test_path_b_business_error_is_surfaced_not_degraded(self, deck_uri):
        # get:slideOoxml succeeds; the apply fails with a genuine 3004 (working Add-In).
        def _emit(document_uri, event, data):
            if event == "ppt:get:slideOoxml":
                return {"success": True, "data": {"slideIndex": 0, "slideId": "sid", "base64": _blank_slide_b64()}}
            return {"success": False, "error": {"code": "3004", "message": "insert failed on client"}}

        ws = _connected_ws(_emit)
        result = await PptInsertChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "chart": {"chartType": "Pie", "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is False
        # A real business error must propagate verbatim, NOT be masked by the 3003 degrade.
        assert "3004" in result["error"]
        assert "3003" not in result["error"]

    @pytest.mark.asyncio
    async def test_update_degrades_when_connected_without_slide_hint(self, deck_uri):
        ws = _connected_ws()  # emit must never be awaited — degrade is up-front
        result = await PptUpdateChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "elementId": "oasp-chart-abc",
                "chart": {"chartType": "Line", "title": "x"},
            }
        )
        assert result["success"] is False
        assert "3003" in result["error"]
        ws.emit_to_document.assert_not_awaited()
        ws.notify_resource_updated.assert_not_called()

    @pytest.mark.asyncio
    async def test_update_degrades_when_path_b_times_out(self, deck_uri):
        ws = _connected_ws(TimeoutError("no ack from Add-In"))
        result = await PptUpdateChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "elementId": "oasp-chart-abc",
                "slideIndex": 0,
                "chart": {"chartType": "Line", "title": "x"},
            }
        )
        assert result["success"] is False
        assert "3003" in result["error"]
        ws.notify_resource_updated.assert_not_called()

    @pytest.mark.asyncio
    async def test_get_degrades_to_disk_when_path_b_times_out(self, deck_uri):
        eid = await _insert_on_disk(deck_uri, slide_index=0, chart_type="Line")
        ws = _connected_ws(TimeoutError("no ack from Add-In"))
        result = await PptGetChartTool(ws).execute({"document_uri": deck_uri, "elementId": eid, "slideIndex": 0})
        # Read-only degradation falls back to the on-disk read — still succeeds.
        assert result["success"] is True
        assert result["data"]["chart"]["chartType"] == "Line"

    @pytest.mark.asyncio
    async def test_get_uses_disk_when_connected_without_hint(self, deck_uri):
        eid = await _insert_on_disk(deck_uri, slide_index=0, chart_type="Pie")
        ws = _connected_ws()  # no slideIndex → on-disk read, no round-trip
        result = await PptGetChartTool(ws).execute({"document_uri": deck_uri, "elementId": eid})
        assert result["success"] is True
        assert result["data"]["chart"]["chartType"] == "Pie"
        ws.emit_to_document.assert_not_awaited()

    @pytest.mark.asyncio
    async def test_insert_degrades_when_export_omits_slide_id(self, deck_uri):
        # A malformed export (base64 present, opaque slideId missing) must NOT fall through
        # to an append (which would leave a duplicate page) — it degrades, and the apply
        # event is never sent.
        def _emit(document_uri, event, data):
            return {"success": True, "data": {"slideIndex": 0, "base64": _blank_slide_b64()}}

        ws = _connected_ws(_emit)
        result = await PptInsertChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "chart": {"chartType": "Pie", "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is False
        assert "3003" in result["error"]
        events = [call.args[1] for call in ws.emit_to_document.call_args_list]
        assert events == ["ppt:get:slideOoxml"]  # stopped before ppt:insert:slidesOoxml
        ws.notify_resource_updated.assert_not_called()

    @pytest.mark.asyncio
    async def test_connected_write_degrade_leaves_disk_untouched(self, deck_uri):
        # The core safety invariant of the flipped guard: a CONNECTED write that degrades
        # must NEVER write the .pptx (else PowerPoint's next save would overwrite it).
        from pathlib import Path
        from urllib.parse import unquote, urlparse

        disk_path = Path(unquote(urlparse(deck_uri).path))
        before = disk_path.read_bytes()
        ws = _connected_ws(TimeoutError("no ack from Add-In"))
        result = await PptInsertChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "chart": {"chartType": "Pie", "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
                "options": {"slideIndex": 0},
            }
        )
        assert result["success"] is False
        assert disk_path.read_bytes() == before

    @pytest.mark.asyncio
    async def test_get_path_b_business_error_surfaced_not_disk_fallback(self, deck_uri):
        # A working Add-In returning a genuine business error (3010) on export must be
        # SURFACED, not silently swallowed into a stale on-disk read. Seed the disk WITH a
        # chart at this elementId so a (wrong) silent fallback would *succeed* — the test
        # then proves the tool fails with 3010 instead.
        eid = await _insert_on_disk(deck_uri, slide_index=0, chart_type="Line")

        def _emit(document_uri, event, data):
            return {"success": False, "error": {"code": "3010", "message": "chart not found on live slide"}}

        ws = _connected_ws(_emit)
        result = await PptGetChartTool(ws).execute({"document_uri": deck_uri, "elementId": eid, "slideIndex": 0})
        assert result["success"] is False
        assert "3010" in result["error"]  # surfaced, NOT a stale disk success

    @pytest.mark.asyncio
    async def test_update_path_b_business_error_is_surfaced_not_degraded(self, deck_uri):
        from office4ai.environment.workspace.dtos.ppt import CategoricalChartData
        from office4ai.environment.workspace.services import chart_engine

        built = await chart_engine.generate_chart_slide_base64(
            CategoricalChartData.model_validate(
                {
                    "chartType": "Line",
                    "categories": ["A", "B"],
                    "series": [{"name": "x", "values": [1, 2]}],
                    "title": "Old",
                }
            )
        )

        def _emit(document_uri, event, data):
            if event == "ppt:get:slideOoxml":
                return {"success": True, "data": {"slideIndex": 0, "slideId": "sid", "base64": built["slideBase64"]}}
            return {"success": False, "error": {"code": "3004", "message": "apply failed on client"}}

        ws = _connected_ws(_emit)
        result = await PptUpdateChartTool(ws).execute(
            {
                "document_uri": deck_uri,
                "elementId": built["elementId"],
                "slideIndex": 0,
                "chart": {"chartType": "Line", "title": "New"},
            }
        )
        assert result["success"] is False
        # Real business error propagates verbatim — never masked by the 3003 degrade.
        assert "3004" in result["error"]
        assert "3003" not in result["error"]
        ws.notify_resource_updated.assert_not_called()

    @pytest.mark.asyncio
    async def test_path_b_holds_document_lock_serializing_concurrent_writes(self, deck_uri):
        # Lost-update invariant: the document lock must wrap the WHOLE path-B round-trip
        # (export → build → re-insert), so two concurrent writes on the same documentUri
        # serialize. The `await asyncio.to_thread(...)` inside chart_engine (between the two
        # carrier emits) is the yield point that WOULD let them interleave absent the lock.
        import asyncio

        order: list[tuple[str, str]] = []

        def _make_emit(tag: str):
            def _emit(document_uri, event, data):
                order.append((tag, event))
                if event == "ppt:get:slideOoxml":
                    return {
                        "success": True,
                        "data": {"slideIndex": 0, "slideId": f"sid-{tag}", "base64": _blank_slide_b64()},
                    }
                return {"success": True, "data": {"insertedSlideIndices": [0], "insertedSlideIds": ["new"]}}

            return _emit

        payload = {
            "document_uri": deck_uri,  # SAME uri for both → same lock
            "chart": {"chartType": "Pie", "categories": ["A", "B"], "series": [{"name": "x", "values": [1, 2]}]},
            "options": {"slideIndex": 0},
        }
        ra, rb = await asyncio.gather(
            PptInsertChartTool(_connected_ws(_make_emit("A"))).execute(payload),
            PptInsertChartTool(_connected_ws(_make_emit("B"))).execute(payload),
        )
        assert ra["success"] is True and rb["success"] is True
        # Each task's [get, insert] pair must complete before the other's begins —
        # a fully grouped order, never interleaved like ["A","B","A","B"].
        tags = [tag for tag, _ in order]
        assert tags in (["A", "A", "B", "B"], ["B", "B", "A", "A"]), order


class TestChartRouterPayloadContract:
    """The router's path-B payloads must validate through the *real* wrap_request → DTO path.

    The dual-path tests mock ``emit_to_document`` (and therefore never exercise wrap_request),
    so this is the only coverage proving the router emits keys the #14 carrier DTOs accept —
    i.e. a server-side payload bug can't masquerade as a 'path B unavailable' degrade.
    """

    def test_apply_payload_validates_through_insert_slides_ooxml_dto(self):
        from office4ai.environment.workspace.services.chart_router import _apply_payload
        from office4ai.environment.workspace.socketio.request_wrapper import wrap_request

        wrapped = wrap_request(
            "ppt:insert:slidesOoxml",
            _apply_payload("UEsDBBQ=", slide_index=2, slide_id="sid-1"),
            "file:///t.pptx",
        )
        assert wrapped["base64"] == "UEsDBBQ="
        assert wrapped["targetSlideIndex"] == 2
        assert wrapped["replaceSlideId"] == "sid-1"
        assert wrapped["finalSlideIndex"] == 2
        assert wrapped["formatting"] == "keepSourceFormatting"

    def test_get_slide_ooxml_payload_validates(self):
        from office4ai.environment.workspace.socketio.request_wrapper import wrap_request

        wrapped = wrap_request("ppt:get:slideOoxml", {"slideIndex": 3}, "file:///t.pptx")
        assert wrapped["slideIndex"] == 3
