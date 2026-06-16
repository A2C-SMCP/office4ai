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
    ExcelAddConditionalFormatTool,
    ExcelClearConditionalFormatTool,
    ExcelClearRangeTool,
    ExcelCopyRangeTool,
    ExcelDeleteRangeTool,
    ExcelGetRangeFormatTool,
    ExcelGetRangeTool,
    ExcelGetSelectedRangeTool,
    ExcelGetWorkbookInfoTool,
    ExcelGetWorksheetInfoTool,
    ExcelInsertRangeTool,
    ExcelMergeCellsTool,
    ExcelSetFormulaTool,
    ExcelSetRangeFormatTool,
    ExcelSetRangeTool,
    ExcelUnmergeCellsTool,
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
        # Range CRUD + 公式 (#19)
        (ExcelGetRangeTool, "excel_get_range", "excel", "get:range"),
        (ExcelSetRangeTool, "excel_set_range", "excel", "set:range"),
        (ExcelClearRangeTool, "excel_clear_range", "excel", "clear:range"),
        (ExcelCopyRangeTool, "excel_copy_range", "excel", "copy:range"),
        (ExcelDeleteRangeTool, "excel_delete_range", "excel", "delete:range"),
        (ExcelInsertRangeTool, "excel_insert_range", "excel", "insert:range"),
        (ExcelSetFormulaTool, "excel_set_formula", "excel", "set:formula"),
        # Format / 条件格式 / 合并单元格 (#20)
        (ExcelGetRangeFormatTool, "excel_get_range_format", "excel", "get:rangeFormat"),
        (ExcelSetRangeFormatTool, "excel_set_range_format", "excel", "set:rangeFormat"),
        (ExcelAddConditionalFormatTool, "excel_add_conditional_format", "excel", "add:conditionalFormat"),
        (ExcelClearConditionalFormatTool, "excel_clear_conditional_format", "excel", "clear:conditionalFormat"),
        (ExcelMergeCellsTool, "excel_merge_cells", "excel", "merge:cells"),
        (ExcelUnmergeCellsTool, "excel_unmerge_cells", "excel", "unmerge:cells"),
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
            "excel_get_range": "excel:get:range",
            "excel_set_range": "excel:set:range",
            "excel_clear_range": "excel:clear:range",
            "excel_copy_range": "excel:copy:range",
            "excel_delete_range": "excel:delete:range",
            "excel_insert_range": "excel:insert:range",
            "excel_set_formula": "excel:set:formula",
            "excel_get_range_format": "excel:get:rangeFormat",
            "excel_set_range_format": "excel:set:rangeFormat",
            "excel_add_conditional_format": "excel:add:conditionalFormat",
            "excel_clear_conditional_format": "excel:clear:conditionalFormat",
            "excel_merge_cells": "excel:merge:cells",
            "excel_unmerge_cells": "excel:unmerge:cells",
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

    @pytest.mark.asyncio
    async def test_get_range_builds_action_with_params(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelGetRangeTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:C3", "include_format": True})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "get:range"
        assert action.params["address"] == "A1:C3"
        assert action.params["include_format"] is True

    @pytest.mark.asyncio
    async def test_set_range_passes_scalar_or_array_values(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelSetRangeTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:B2", "values": [[1, 2], [3, 4]]})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "set:range"
        assert action.params["values"] == [[1, 2], [3, 4]]

    @pytest.mark.asyncio
    async def test_clear_range_passes_clear_type(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelClearRangeTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:C3", "clear_type": "contents"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "clear:range"
        assert action.params["clear_type"] == "contents"

    @pytest.mark.asyncio
    async def test_clear_range_rejects_invalid_clear_type(self, mock_workspace):
        tool = ExcelClearRangeTool(mock_workspace)
        result = await tool.execute(
            {"document_uri": "file:///data.xlsx", "address": "A1:C3", "clear_type": "everything"}
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_copy_range_passes_source_target(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelCopyRangeTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "source_address": "A1:C3", "target_address": "E1:G3"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "copy:range"
        assert action.params["source_address"] == "A1:C3"
        assert action.params["target_address"] == "E1:G3"

    @pytest.mark.asyncio
    async def test_delete_range_rejects_insert_direction(self, mock_workspace):
        """delete 仅接受 up/left；'down' 应在输入校验阶段被拒。"""
        tool = ExcelDeleteRangeTool(mock_workspace)
        result = await tool.execute(
            {"document_uri": "file:///data.xlsx", "address": "B2:B5", "shift_direction": "down"}
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_insert_range_passes_shift_direction(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelInsertRangeTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "B2:B5", "shift_direction": "down"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "insert:range"
        assert action.params["shift_direction"] == "down"

    @pytest.mark.asyncio
    async def test_set_formula_passes_formula(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelSetFormulaTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "D1", "formula": "=SUM(A1:C1)"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "set:formula"
        assert action.params["formula"] == "=SUM(A1:C1)"

    @pytest.mark.asyncio
    async def test_get_range_omits_none_worksheet_name(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelGetRangeTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1"})

        action = mock_workspace.execute.call_args[0][0]
        assert "worksheet_name" not in action.params

    # ---- #20 Format / 条件格式 / 合并单元格 -------------------------------

    @pytest.mark.asyncio
    async def test_get_range_format_builds_correct_action(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelGetRangeFormatTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:C3"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "get:rangeFormat"
        assert action.params["address"] == "A1:C3"
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_set_range_format_passes_partial_format(self, mock_workspace):
        """format 偏更新结构以 snake_case 流入 params；未传字段被 exclude_none 剔除。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelSetRangeFormatTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///data.xlsx",
                "address": "A1:C1",
                "format": {"font": {"bold": True}, "alignment": {"horizontal": "Center"}},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "set:rangeFormat"
        assert action.params["format"]["font"]["bold"] is True
        assert action.params["format"]["alignment"]["horizontal"] == "Center"
        # 未传入的可选字段不应出现
        assert "fill" not in action.params["format"]

    @pytest.mark.asyncio
    async def test_add_conditional_format_passes_rule_passthrough(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelAddConditionalFormatTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///data.xlsx",
                "address": "B2:B100",
                "rule": {"type": "cellValue", "operator": "greaterThan", "value": 90},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "add:conditionalFormat"
        assert action.params["rule"]["type"] == "cellValue"
        assert action.params["rule"]["operator"] == "greaterThan"
        assert action.params["rule"]["value"] == 90

    @pytest.mark.asyncio
    async def test_add_conditional_format_rejects_rule_without_type(self, mock_workspace):
        """rule 透传但 type 必填；缺失应在输入校验阶段被拒。"""
        tool = ExcelAddConditionalFormatTool(mock_workspace)
        result = await tool.execute(
            {"document_uri": "file:///data.xlsx", "address": "B2:B100", "rule": {"operator": "greaterThan"}}
        )
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_clear_conditional_format_builds_correct_action(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelClearConditionalFormatTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "B2:B100"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "clear:conditionalFormat"
        assert action.params["address"] == "B2:B100"

    @pytest.mark.asyncio
    async def test_merge_cells_builds_correct_action(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelMergeCellsTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:C1", "worksheet_name": "Sheet1"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "merge:cells"
        assert action.params["address"] == "A1:C1"
        assert action.params["worksheet_name"] == "Sheet1"

    @pytest.mark.asyncio
    async def test_unmerge_cells_builds_correct_action(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={})

        tool = ExcelUnmergeCellsTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:C1"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "unmerge:cells"
        assert action.params["address"] == "A1:C1"


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

    def test_get_range_summary_without_format(self, mock_workspace):
        tool = ExcelGetRangeTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={"address": "Sheet1!A1:C3", "values": [[1, 2, 3]], "rowCount": 1, "columnCount": 3},
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "Sheet1!A1:C3" in result["content"]
        assert "with format" not in result["content"]
        assert result["data"]["columnCount"] == 3

    def test_get_range_summary_with_format(self, mock_workspace):
        tool = ExcelGetRangeTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "address": "Sheet1!A1",
                "values": [[1]],
                "rowCount": 1,
                "columnCount": 1,
                "format": {"fill": {"color": "#FFFFFF"}},
            },
        )
        result = tool.format_result(obs)
        assert "with format" in result["content"]

    def test_write_tool_uses_base_format_result(self, mock_workspace):
        """写工具不覆写 format_result: 成功时返回 {success, data}，无 content。"""
        tool = ExcelSetRangeTool(mock_workspace)
        obs = OfficeObs(success=True, data={"address": "Sheet1!A1:C3"})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"address": "Sheet1!A1:C3"}}
        assert "content" not in result

    def test_write_tool_failure_returns_error(self, mock_workspace):
        tool = ExcelSetFormulaTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5005 FORMULA_ERROR")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5005 FORMULA_ERROR"

    def test_get_range_surfaces_range_invalid_error(self, mock_workspace):
        """非法地址由 AddIn 返回 5002 RANGE_INVALID；consumer 透传 (address 对 consumer 透明)。"""
        tool = ExcelGetRangeTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5002 RANGE_INVALID")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5002 RANGE_INVALID"

    # ---- #20 Format / 条件格式 / 合并单元格 -------------------------------

    def test_get_range_format_summary(self, mock_workspace):
        """get:rangeFormat 覆写 format_result: 返回字体摘要 content + 完整 data。"""
        tool = ExcelGetRangeFormatTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "address": "Sheet1!A1:C3",
                "format": {
                    "font": {"name": "Calibri", "size": 12, "bold": True},
                    "fill": {"color": "#FFFF00"},
                },
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "Sheet1!A1:C3" in result["content"]
        assert "Calibri" in result["content"]
        assert result["data"]["format"]["fill"]["color"] == "#FFFF00"

    def test_get_range_format_failure_returns_error(self, mock_workspace):
        tool = ExcelGetRangeFormatTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5001 WORKSHEET_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5001 WORKSHEET_NOT_FOUND"

    def test_set_range_format_uses_base_format_result(self, mock_workspace):
        """写工具不覆写 format_result: 成功时返回 {success, data}，无 content。"""
        tool = ExcelSetRangeFormatTool(mock_workspace)
        obs = OfficeObs(success=True, data={"address": "Sheet1!A1:C1"})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"address": "Sheet1!A1:C1"}}
        assert "content" not in result

    def test_merge_cells_surfaces_merge_conflict_error(self, mock_workspace):
        """合并冲突由 AddIn 返回 5003 MERGE_CONFLICT；consumer 透传。"""
        tool = ExcelMergeCellsTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5003 MERGE_CONFLICT")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5003 MERGE_CONFLICT"

    def test_set_range_format_surfaces_protected_sheet_error(self, mock_workspace):
        """受保护工作表由 AddIn 返回 5004 PROTECTED_SHEET；consumer 透传。"""
        tool = ExcelSetRangeFormatTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5004 PROTECTED_SHEET")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5004 PROTECTED_SHEET"
