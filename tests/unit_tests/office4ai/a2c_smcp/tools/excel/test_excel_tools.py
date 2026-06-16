"""
Excel MCP Tools 单元测试 | Excel MCP Tools unit tests

测试策略 (mirror tests/.../tools/word/test_word_tools.py):
- Mock OfficeWorkspace.execute() 的返回值
- 验证工具元数据 (name, category, event_name, input_schema)
- 验证 OfficeAction 构建 (category, action_name, params)
- 验证 format_result hook (获取类工具返回 content + data)
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.tools.excel import (
    ExcelGetSelectedRangeTool,
    ExcelGetWorkbookInfoTool,
    ExcelGetWorksheetInfoTool,
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
    """测试 Excel 工具的元数据声明"""

    TOOL_SPECS = [
        (ExcelGetWorkbookInfoTool, "excel_get_workbook_info", "excel", "get:workbookInfo"),
        (ExcelGetWorksheetInfoTool, "excel_get_worksheet_info", "excel", "get:worksheetInfo"),
        (ExcelGetSelectedRangeTool, "excel_get_selected_range", "excel", "get:selectedRange"),
    ]

    @pytest.mark.parametrize("tool_cls,expected_name,expected_category,expected_event", TOOL_SPECS)
    def test_tool_metadata(self, mock_workspace, tool_cls, expected_name, expected_category, expected_event):
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
        tool = tool_cls(mock_workspace)
        assert "document_uri" in tool.input_schema["properties"]

    def test_full_event_names_compose_to_oasp(self, mock_workspace):
        """category + event_name 拼出 OASP 全名 (workspace.execute 用 f'{category}:{action_name}')。"""
        expected = {
            "excel_get_workbook_info": "excel:get:workbookInfo",
            "excel_get_worksheet_info": "excel:get:worksheetInfo",
            "excel_get_selected_range": "excel:get:selectedRange",
        }
        for tool_cls, name, _, _ in self.TOOL_SPECS:
            tool = tool_cls(mock_workspace)
            assert f"{tool.category}:{tool.event_name}" == expected[name]


# ============================================================================
# Execute Flow Tests
# ============================================================================


class TestExecuteFlow:
    """测试通用执行流程与 OfficeAction 构建"""

    @pytest.mark.asyncio
    async def test_get_workbook_info_builds_correct_action(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelGetWorkbookInfoTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        mock_workspace.execute.assert_called_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "get:workbookInfo"
        assert action.params["document_uri"] == "file:///data.xlsx"

    @pytest.mark.asyncio
    async def test_get_worksheet_info_passes_worksheet_name(self, mock_workspace):
        """worksheet_name 以 snake_case 流入 params (DTO populate_by_name 接受)。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelGetWorksheetInfoTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "worksheet_name": "Sheet2"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "get:worksheetInfo"
        assert action.params["worksheet_name"] == "Sheet2"

    @pytest.mark.asyncio
    async def test_get_worksheet_info_omits_none_worksheet_name(self, mock_workspace):
        """worksheet_name 省略时, exclude_none 不应把它放进 params。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelGetWorksheetInfoTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        action = mock_workspace.execute.call_args[0][0]
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_get_selected_range_builds_correct_action(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelGetSelectedRangeTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "get:selectedRange"

    @pytest.mark.asyncio
    async def test_missing_document_uri_returns_validation_error(self, mock_workspace):
        tool = ExcelGetWorkbookInfoTool(mock_workspace)
        result = await tool.execute({})
        assert result["success"] is False
        assert "error" in result
        mock_workspace.execute.assert_not_called()


# ============================================================================
# format_result Tests (获取类工具: content + data)
# ============================================================================


class TestFormatResult:
    """测试 format_result hook"""

    def test_workbook_info_summary(self, mock_workspace):
        tool = ExcelGetWorkbookInfoTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "sheets": [
                    {"name": "Sheet1", "index": 0, "isActive": True, "isHidden": False},
                    {"name": "Sheet2", "index": 1, "isActive": False, "isHidden": False},
                ],
                "activeSheet": "Sheet1",
                "fileName": "data.xlsx",
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "2 sheet(s)" in result["content"]
        assert "data.xlsx" in result["content"]
        assert result["data"]["activeSheet"] == "Sheet1"

    def test_worksheet_info_summary(self, mock_workspace):
        tool = ExcelGetWorksheetInfoTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "name": "Sheet1",
                "usedRange": {"address": "Sheet1!A1:D10", "rowCount": 10, "columnCount": 4},
                "tableCount": 1,
                "chartCount": 2,
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "Sheet1!A1:D10" in result["content"]
        assert "tables=1" in result["content"]
        assert "charts=2" in result["content"]

    def test_selected_range_summary(self, mock_workspace):
        tool = ExcelGetSelectedRangeTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={"address": "Sheet1!A1:C3", "values": [[1, 2, 3]], "rowCount": 3, "columnCount": 3},
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "Sheet1!A1:C3" in result["content"]
        assert result["data"]["rowCount"] == 3

    def test_failure_returns_error(self, mock_workspace):
        tool = ExcelGetWorkbookInfoTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="Document not connected")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "Document not connected"
