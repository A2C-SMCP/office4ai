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
    ExcelActivateWorksheetTool,
    ExcelAddConditionalFormatTool,
    ExcelAddTableRowTool,
    ExcelAddWorksheetTool,
    ExcelClearConditionalFormatTool,
    ExcelClearRangeTool,
    ExcelCopyRangeTool,
    ExcelDeleteChartTool,
    ExcelDeletePivotTableTool,
    ExcelDeleteRangeTool,
    ExcelDeleteTableRowTool,
    ExcelDeleteWorksheetTool,
    ExcelGetChartsTool,
    ExcelGetPivotTablesTool,
    ExcelGetRangeFormatTool,
    ExcelGetRangeTool,
    ExcelGetSelectedRangeTool,
    ExcelGetTablesTool,
    ExcelGetTableTool,
    ExcelGetWorkbookInfoTool,
    ExcelGetWorksheetInfoTool,
    ExcelGetWorksheetsTool,
    ExcelInsertChartTool,
    ExcelInsertPivotTableTool,
    ExcelInsertRangeTool,
    ExcelInsertTableTool,
    ExcelMergeCellsTool,
    ExcelRenameWorksheetTool,
    ExcelSetFormulaTool,
    ExcelSetRangeFormatTool,
    ExcelSetRangeTool,
    ExcelSortTableTool,
    ExcelUnmergeCellsTool,
    ExcelUpdateChartTool,
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
        # Worksheet 管理 (#21)
        (ExcelGetWorksheetsTool, "excel_get_worksheets", "excel", "get:worksheets"),
        (ExcelAddWorksheetTool, "excel_add_worksheet", "excel", "add:worksheet"),
        (ExcelDeleteWorksheetTool, "excel_delete_worksheet", "excel", "delete:worksheet"),
        (ExcelRenameWorksheetTool, "excel_rename_worksheet", "excel", "rename:worksheet"),
        (ExcelActivateWorksheetTool, "excel_activate_worksheet", "excel", "activate:worksheet"),
        # Table 操作 (#22)
        (ExcelInsertTableTool, "excel_insert_table", "excel", "insert:table"),
        (ExcelGetTableTool, "excel_get_table", "excel", "get:table"),
        (ExcelGetTablesTool, "excel_get_tables", "excel", "get:tables"),
        (ExcelAddTableRowTool, "excel_add_table_row", "excel", "add:tableRow"),
        (ExcelDeleteTableRowTool, "excel_delete_table_row", "excel", "delete:tableRow"),
        (ExcelSortTableTool, "excel_sort_table", "excel", "sort:table"),
        # Chart 操作 (#23)
        (ExcelInsertChartTool, "excel_insert_chart", "excel", "insert:chart"),
        (ExcelGetChartsTool, "excel_get_charts", "excel", "get:charts"),
        (ExcelUpdateChartTool, "excel_update_chart", "excel", "update:chart"),
        (ExcelDeleteChartTool, "excel_delete_chart", "excel", "delete:chart"),
        # PivotTable 操作 (#24)
        (ExcelInsertPivotTableTool, "excel_insert_pivot_table", "excel", "insert:pivotTable"),
        (ExcelGetPivotTablesTool, "excel_get_pivot_tables", "excel", "get:pivotTables"),
        (ExcelDeletePivotTableTool, "excel_delete_pivot_table", "excel", "delete:pivotTable"),
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
            "excel_get_worksheets": "excel:get:worksheets",
            "excel_add_worksheet": "excel:add:worksheet",
            "excel_delete_worksheet": "excel:delete:worksheet",
            "excel_rename_worksheet": "excel:rename:worksheet",
            "excel_activate_worksheet": "excel:activate:worksheet",
            "excel_insert_table": "excel:insert:table",
            "excel_get_table": "excel:get:table",
            "excel_get_tables": "excel:get:tables",
            "excel_add_table_row": "excel:add:tableRow",
            "excel_delete_table_row": "excel:delete:tableRow",
            "excel_sort_table": "excel:sort:table",
            "excel_insert_chart": "excel:insert:chart",
            "excel_get_charts": "excel:get:charts",
            "excel_update_chart": "excel:update:chart",
            "excel_delete_chart": "excel:delete:chart",
            "excel_insert_pivot_table": "excel:insert:pivotTable",
            "excel_get_pivot_tables": "excel:get:pivotTables",
            "excel_delete_pivot_table": "excel:delete:pivotTable",
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

    # ---- #21 Worksheet 管理 -----------------------------------------------

    @pytest.mark.asyncio
    async def test_get_worksheets_builds_correct_action(self, mock_workspace):
        """get:worksheets 无业务参数, params 仅有 document_uri。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"worksheets": []})

        tool = ExcelGetWorksheetsTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "get:worksheets"
        assert "worksheet_name" not in action.params
        assert "name" not in action.params

    @pytest.mark.asyncio
    async def test_add_worksheet_passes_name(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "数据分析", "index": 2})

        tool = ExcelAddWorksheetTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "name": "数据分析"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "add:worksheet"
        assert action.params["name"] == "数据分析"

    @pytest.mark.asyncio
    async def test_add_worksheet_omits_none_name(self, mock_workspace):
        """name 省略时, exclude_none 不应把它放进 params（Excel 自动命名）。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "Sheet3", "index": 2})

        tool = ExcelAddWorksheetTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "add:worksheet"
        assert "name" not in action.params

    @pytest.mark.asyncio
    async def test_delete_worksheet_passes_worksheet_name(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deleted": True})

        tool = ExcelDeleteWorksheetTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "worksheet_name": "Sheet3"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "delete:worksheet"
        assert action.params["worksheet_name"] == "Sheet3"

    @pytest.mark.asyncio
    async def test_delete_worksheet_requires_worksheet_name(self, mock_workspace):
        """worksheet_name 必填: 缺失时 MCP 输入校验失败, 不应触达 workspace。"""
        tool = ExcelDeleteWorksheetTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_rename_worksheet_passes_both_names(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "销售数据"})

        tool = ExcelRenameWorksheetTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "current_name": "Sheet1", "new_name": "销售数据"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "rename:worksheet"
        assert action.params["current_name"] == "Sheet1"
        assert action.params["new_name"] == "销售数据"

    @pytest.mark.asyncio
    async def test_rename_worksheet_requires_both_names(self, mock_workspace):
        """current_name 与 new_name 均必填: 缺其一即校验失败。"""
        tool = ExcelRenameWorksheetTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "current_name": "Sheet1"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_activate_worksheet_passes_worksheet_name(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"activated": True})

        tool = ExcelActivateWorksheetTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "worksheet_name": "Sheet2"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "activate:worksheet"
        assert action.params["worksheet_name"] == "Sheet2"

    # ---- #22 Table 操作 ----------------------------------------------------

    @pytest.mark.asyncio
    async def test_insert_table_passes_required_and_drops_optionals(self, mock_workspace):
        """address + has_headers 流入 params；data/style_name/worksheet_name 省略时被剔除。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "Table1", "address": "A1:C4"})

        tool = ExcelInsertTableTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:C4", "has_headers": True})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "insert:table"
        assert action.params["address"] == "A1:C4"
        assert action.params["has_headers"] is True
        assert "data" not in action.params
        assert "style_name" not in action.params
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_insert_table_passes_data_and_style(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "Table1", "address": "A1:C2"})

        tool = ExcelInsertTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///data.xlsx",
                "address": "A1:C2",
                "has_headers": True,
                "data": [["姓名", "年龄", "城市"], ["张三", 25, "北京"]],
                "style_name": "TableStyleMedium2",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["data"] == [["姓名", "年龄", "城市"], ["张三", 25, "北京"]]
        assert action.params["style_name"] == "TableStyleMedium2"

    @pytest.mark.asyncio
    async def test_insert_table_requires_address_and_has_headers(self, mock_workspace):
        """address 与 has_headers 必填: 缺其一即校验失败, 不触达 workspace。"""
        tool = ExcelInsertTableTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "address": "A1:C4"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_get_table_passes_table_id(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "Table1"})

        tool = ExcelGetTableTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "table_id": "Table1"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "get:table"
        assert action.params["table_id"] == "Table1"
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_get_table_requires_table_id(self, mock_workspace):
        tool = ExcelGetTableTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_get_tables_builds_correct_action(self, mock_workspace):
        """get:tables 仅 worksheet_name 可选, 省略时 params 仅有 document_uri。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"tables": []})

        tool = ExcelGetTablesTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "get:tables"
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_add_table_row_passes_values(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"tableId": "Table1"})

        tool = ExcelAddTableRowTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "table_id": "Table1", "values": ["赵六", 35, "深圳"]})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "add:tableRow"
        assert action.params["table_id"] == "Table1"
        assert action.params["values"] == ["赵六", 35, "深圳"]

    @pytest.mark.asyncio
    async def test_add_table_row_requires_table_id_and_values(self, mock_workspace):
        tool = ExcelAddTableRowTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "table_id": "Table1"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_delete_table_row_passes_row_index_zero(self, mock_workspace):
        """row_index=0 是合法首行: exclude_none 不得剔除 (0 ≠ None)。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deleted": True})

        tool = ExcelDeleteTableRowTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "table_id": "Table1", "row_index": 0})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "delete:tableRow"
        assert action.params["table_id"] == "Table1"
        assert action.params["row_index"] == 0

    @pytest.mark.asyncio
    async def test_delete_table_row_requires_table_id_and_row_index(self, mock_workspace):
        tool = ExcelDeleteTableRowTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "table_id": "Table1"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_sort_table_passes_nested_sort_fields(self, mock_workspace):
        """sort_fields 经 MCP 输入 (snake) → params 以 snake_case 嵌套字典流入。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"sorted": True})

        tool = ExcelSortTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///data.xlsx",
                "table_id": "Table1",
                "sort_fields": [{"column_index": 1, "ascending": False}, {"column_index": 0}],
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "sort:table"
        assert action.params["table_id"] == "Table1"
        # 第一项含 ascending=False；第二项省略 ascending → exclude_none 剔除
        assert action.params["sort_fields"][0] == {"column_index": 1, "ascending": False}
        assert action.params["sort_fields"][1] == {"column_index": 0}

    @pytest.mark.asyncio
    async def test_sort_table_requires_sort_fields(self, mock_workspace):
        tool = ExcelSortTableTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "table_id": "Table1"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    # ---- #23 Chart 操作 ----------------------------------------------------

    @pytest.mark.asyncio
    async def test_insert_chart_passes_required_and_drops_optionals(self, mock_workspace):
        """source_address + chart_type 流入 params；title/position/worksheet_name 省略时被剔除。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "Chart 1"})

        tool = ExcelInsertChartTool(mock_workspace)
        await tool.execute(
            {"document_uri": "file:///data.xlsx", "source_address": "A1:C4", "chart_type": "ColumnClustered"}
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "insert:chart"
        assert action.params["source_address"] == "A1:C4"
        assert action.params["chart_type"] == "ColumnClustered"
        assert "title" not in action.params
        assert "position" not in action.params
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_insert_chart_passes_title_and_nested_position(self, mock_workspace):
        """position 经 MCP 输入 (snake) → params 以嵌套字典流入；top=0 保留, height 省略剔除。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "Chart 1"})

        tool = ExcelInsertChartTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///data.xlsx",
                "source_address": "A1:C4",
                "chart_type": "Pie",
                "title": "占比",
                "position": {"top": 0, "left": 300, "width": 400},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["title"] == "占比"
        assert action.params["position"] == {"top": 0, "left": 300, "width": 400}

    @pytest.mark.asyncio
    async def test_insert_chart_requires_source_address_and_chart_type(self, mock_workspace):
        """source_address 与 chart_type 必填: 缺其一即校验失败, 不触达 workspace。"""
        tool = ExcelInsertChartTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "source_address": "A1:C4"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_get_charts_builds_correct_action(self, mock_workspace):
        """get:charts 仅 worksheet_name 可选, 省略时 params 仅有 document_uri。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"charts": []})

        tool = ExcelGetChartsTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "get:charts"
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_update_chart_passes_nested_properties(self, mock_workspace):
        """properties 经 MCP 输入 (snake) → params 以 snake_case 嵌套字典流入；偏更新省略字段剔除。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "Chart 1"})

        tool = ExcelUpdateChartTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///data.xlsx",
                "chart_name": "Chart 1",
                "properties": {"chart_type": "Line", "source_address": "A1:D9"},
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "update:chart"
        assert action.params["chart_name"] == "Chart 1"
        assert action.params["properties"] == {"chart_type": "Line", "source_address": "A1:D9"}

    @pytest.mark.asyncio
    async def test_update_chart_requires_chart_name_and_properties(self, mock_workspace):
        tool = ExcelUpdateChartTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "chart_name": "Chart 1"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_delete_chart_passes_chart_name(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deleted": True})

        tool = ExcelDeleteChartTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "chart_name": "Chart 1"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "delete:chart"
        assert action.params["chart_name"] == "Chart 1"
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_delete_chart_requires_chart_name(self, mock_workspace):
        tool = ExcelDeleteChartTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    # ---- #24 PivotTable 操作 ------------------------------------------------

    @pytest.mark.asyncio
    async def test_insert_pivot_table_passes_required_and_drops_optionals(self, mock_workspace):
        """source_address + target_address 流入 params；name/worksheet_name 省略时被剔除。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "PivotTable1"})

        tool = ExcelInsertPivotTableTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "source_address": "A1:D100", "target_address": "F1"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "excel"
        assert action.action_name == "insert:pivotTable"
        assert action.params["source_address"] == "A1:D100"
        assert action.params["target_address"] == "F1"
        assert "name" not in action.params
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_insert_pivot_table_passes_name_and_worksheet(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"name": "销售汇总"})

        tool = ExcelInsertPivotTableTool(mock_workspace)
        await tool.execute(
            {
                "document_uri": "file:///data.xlsx",
                "source_address": "A1:D100",
                "target_address": "F1",
                "name": "销售汇总",
                "worksheet_name": "Sheet2",
            }
        )

        action = mock_workspace.execute.call_args[0][0]
        assert action.params["name"] == "销售汇总"
        assert action.params["worksheet_name"] == "Sheet2"

    @pytest.mark.asyncio
    async def test_insert_pivot_table_requires_source_and_target(self, mock_workspace):
        """source_address 与 target_address 必填: 缺其一即校验失败, 不触达 workspace。"""
        tool = ExcelInsertPivotTableTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx", "source_address": "A1:D100"})
        assert result["success"] is False
        mock_workspace.execute.assert_not_called()

    @pytest.mark.asyncio
    async def test_get_pivot_tables_builds_correct_action(self, mock_workspace):
        """get:pivotTables 仅 worksheet_name 可选, 省略时 params 仅有 document_uri。"""
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"pivotTables": []})

        tool = ExcelGetPivotTablesTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "get:pivotTables"
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_delete_pivot_table_passes_pivot_table_name(self, mock_workspace):
        mock_workspace.execute.return_value = OfficeObs(success=True, data={"deleted": True})

        tool = ExcelDeletePivotTableTool(mock_workspace)
        await tool.execute({"document_uri": "file:///data.xlsx", "pivot_table_name": "销售汇总"})

        action = mock_workspace.execute.call_args[0][0]
        assert action.action_name == "delete:pivotTable"
        assert action.params["pivot_table_name"] == "销售汇总"
        assert "worksheet_name" not in action.params

    @pytest.mark.asyncio
    async def test_delete_pivot_table_requires_pivot_table_name(self, mock_workspace):
        tool = ExcelDeletePivotTableTool(mock_workspace)
        result = await tool.execute({"document_uri": "file:///data.xlsx"})
        assert result["success"] is False
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

    # ---- #21 Worksheet 管理 -----------------------------------------------

    def test_get_worksheets_summary(self, mock_workspace):
        """get:worksheets 覆写 format_result: 返回工作表名列表 content + 完整 data。"""
        tool = ExcelGetWorksheetsTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "worksheets": [
                    {"name": "Sheet1", "index": 0, "isActive": True, "isHidden": False},
                    {"name": "销售数据", "index": 1, "isActive": False, "isHidden": False},
                ]
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "2 worksheet(s)" in result["content"]
        assert "Sheet1" in result["content"]
        assert "销售数据" in result["content"]
        assert len(result["data"]["worksheets"]) == 2

    def test_get_worksheets_empty_summary(self, mock_workspace):
        tool = ExcelGetWorksheetsTool(mock_workspace)
        obs = OfficeObs(success=True, data={"worksheets": []})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 worksheet(s)" in result["content"]

    def test_get_worksheets_missing_key_falls_back(self, mock_workspace):
        """data 缺 'worksheets' 键时, .get 兜底为空列表 → '0 worksheet(s)' 不崩溃。"""
        tool = ExcelGetWorksheetsTool(mock_workspace)
        obs = OfficeObs(success=True, data={})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 worksheet(s)" in result["content"]

    def test_get_worksheets_failure_returns_error(self, mock_workspace):
        tool = ExcelGetWorksheetsTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="3001 DOCUMENT_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "3001 DOCUMENT_NOT_FOUND"

    def test_add_worksheet_uses_base_format_result(self, mock_workspace):
        """写工具不覆写 format_result: 成功时返回 {success, data}，无 content。"""
        tool = ExcelAddWorksheetTool(mock_workspace)
        obs = OfficeObs(success=True, data={"name": "数据分析", "index": 2})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"name": "数据分析", "index": 2}}
        assert "content" not in result

    def test_delete_worksheet_surfaces_worksheet_not_found(self, mock_workspace):
        """删除不存在的工作表由 AddIn 返回 5001 WORKSHEET_NOT_FOUND；consumer 透传。"""
        tool = ExcelDeleteWorksheetTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5001 WORKSHEET_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5001 WORKSHEET_NOT_FOUND"

    def test_rename_worksheet_surfaces_operation_failed(self, mock_workspace):
        """重命名为已存在名称由 AddIn 返回 3004 OPERATION_FAILED；consumer 透传。"""
        tool = ExcelRenameWorksheetTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="3004 OPERATION_FAILED")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "3004 OPERATION_FAILED"

    # ---- #22 Table 操作 ----------------------------------------------------

    def test_get_table_summary(self, mock_workspace):
        """get:table 覆写 format_result: 返回 name/address/行列数 content + 完整 data。"""
        tool = ExcelGetTableTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "name": "Table1",
                "id": "{1}",
                "address": "Sheet1!A1:C4",
                "rowCount": 3,
                "columnCount": 3,
                "columns": [{"name": "姓名", "index": 0}],
                "styleName": "TableStyleMedium2",
                "showHeaders": True,
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "Table1" in result["content"]
        assert "Sheet1!A1:C4" in result["content"]
        assert "3 row(s)" in result["content"]
        assert "3 column(s)" in result["content"]
        assert result["data"]["styleName"] == "TableStyleMedium2"

    def test_get_table_surfaces_table_not_found(self, mock_workspace):
        """获取不存在的表格由 AddIn 返回 5006 TABLE_NOT_FOUND；consumer 透传。"""
        tool = ExcelGetTableTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5006 TABLE_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5006 TABLE_NOT_FOUND"

    def test_get_tables_summary(self, mock_workspace):
        """get:tables 覆写 format_result: 返回表格名列表 content + 完整 data。"""
        tool = ExcelGetTablesTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "tables": [
                    {"name": "Table1", "id": "{1}", "address": "Sheet1!A1:C4"},
                    {"name": "销售表", "id": "{2}", "address": "Sheet1!E1:G10"},
                ]
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "2 table(s)" in result["content"]
        assert "Table1" in result["content"]
        assert "销售表" in result["content"]
        assert len(result["data"]["tables"]) == 2

    def test_get_tables_empty_summary(self, mock_workspace):
        tool = ExcelGetTablesTool(mock_workspace)
        obs = OfficeObs(success=True, data={"tables": []})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 table(s)" in result["content"]

    def test_get_tables_missing_key_falls_back(self, mock_workspace):
        """data 缺 'tables' 键时, .get 兜底为空列表 → '0 table(s)' 不崩溃。"""
        tool = ExcelGetTablesTool(mock_workspace)
        obs = OfficeObs(success=True, data={})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 table(s)" in result["content"]

    def test_get_tables_filters_non_dict_entries(self, mock_workspace):
        """data['tables'] 含非 dict 畸形条目时, isinstance 过滤 → 计数与名称列表一致 (不计入畸形项)。"""
        tool = ExcelGetTablesTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={"tables": [{"name": "T1", "id": "{1}", "address": "Sheet1!A1:C4"}, "garbage", 42]},
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        # 仅 1 个合法 dict 条目 → 计数为 1, 与名称列表一致 (不是原始 len 3)
        assert "1 table(s)" in result["content"]
        assert "T1" in result["content"]
        assert "garbage" not in result["content"]

    def test_insert_table_uses_base_format_result(self, mock_workspace):
        """写工具不覆写 format_result: 成功时返回 {success, data}，无 content。"""
        tool = ExcelInsertTableTool(mock_workspace)
        obs = OfficeObs(success=True, data={"name": "Table1", "address": "Sheet1!A1:C4"})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"name": "Table1", "address": "Sheet1!A1:C4"}}
        assert "content" not in result

    def test_add_table_row_surfaces_data_type_mismatch(self, mock_workspace):
        """类型不匹配由 AddIn 返回 5009 DATA_TYPE_MISMATCH；consumer 透传。"""
        tool = ExcelAddTableRowTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5009 DATA_TYPE_MISMATCH")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5009 DATA_TYPE_MISMATCH"

    def test_delete_table_row_surfaces_param_out_of_range(self, mock_workspace):
        """行索引越界由 AddIn 返回 4004 PARAM_OUT_OF_RANGE；consumer 透传。"""
        tool = ExcelDeleteTableRowTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="4004 PARAM_OUT_OF_RANGE")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "4004 PARAM_OUT_OF_RANGE"

    def test_sort_table_uses_base_format_result(self, mock_workspace):
        tool = ExcelSortTableTool(mock_workspace)
        obs = OfficeObs(success=True, data={"sorted": True})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"sorted": True}}
        assert "content" not in result

    # ---- #23 Chart 操作 ----------------------------------------------------

    def test_get_charts_summary(self, mock_workspace):
        """get:charts 覆写 format_result: 返回图表名列表 content + 完整 data。"""
        tool = ExcelGetChartsTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "charts": [
                    {
                        "name": "Chart 1",
                        "chartType": "ColumnClustered",
                        "title": "销售",
                        "top": 200.0,
                        "left": 300.0,
                        "width": 400.0,
                        "height": 300.0,
                    },
                    {
                        "name": "趋势图",
                        "chartType": "Line",
                        "title": "趋势",
                        "top": 10.0,
                        "left": 20.0,
                        "width": 300.0,
                        "height": 200.0,
                    },
                ]
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "2 chart(s)" in result["content"]
        assert "Chart 1" in result["content"]
        assert "趋势图" in result["content"]
        assert len(result["data"]["charts"]) == 2

    def test_get_charts_empty_summary(self, mock_workspace):
        tool = ExcelGetChartsTool(mock_workspace)
        obs = OfficeObs(success=True, data={"charts": []})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 chart(s)" in result["content"]

    def test_get_charts_missing_key_falls_back(self, mock_workspace):
        """data 缺 'charts' 键时, .get 兜底为空列表 → '0 chart(s)' 不崩溃。"""
        tool = ExcelGetChartsTool(mock_workspace)
        obs = OfficeObs(success=True, data={})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 chart(s)" in result["content"]

    def test_get_charts_filters_non_dict_entries(self, mock_workspace):
        """data['charts'] 含非 dict 畸形条目时, isinstance 过滤 → 计数与名称列表一致 (不计入畸形项)。"""
        tool = ExcelGetChartsTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "charts": [
                    {
                        "name": "C1",
                        "chartType": "Pie",
                        "title": "",
                        "top": 0.0,
                        "left": 0.0,
                        "width": 1.0,
                        "height": 1.0,
                    },
                    "garbage",
                    7,
                ]
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        # 仅 1 个合法 dict 条目 → 计数为 1, 与名称列表一致 (不是原始 len 3)
        assert "1 chart(s)" in result["content"]
        assert "C1" in result["content"]
        assert "garbage" not in result["content"]

    def test_get_charts_dict_entry_missing_name_uses_placeholder(self, mock_workspace):
        """合法 dict 但缺 'name' 键时, .get 兜底 '?' 占位并计入 (与 get_tables 同语义)。"""
        tool = ExcelGetChartsTool(mock_workspace)
        obs = OfficeObs(success=True, data={"charts": [{"chartType": "Line", "title": "无名图"}]})
        result = tool.format_result(obs)
        assert result["success"] is True
        # dict 条目被保留并计数 (isinstance 通过)，缺名以 '?' 占位
        assert "1 chart(s): ?" in result["content"]

    def test_get_charts_surfaces_worksheet_not_found(self, mock_workspace):
        """工作表不存在由 AddIn 返回 5001 WORKSHEET_NOT_FOUND；consumer 透传。"""
        tool = ExcelGetChartsTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5001 WORKSHEET_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5001 WORKSHEET_NOT_FOUND"

    def test_insert_chart_uses_base_format_result(self, mock_workspace):
        """写工具不覆写 format_result: 成功时返回 {success, data}，无 content。"""
        tool = ExcelInsertChartTool(mock_workspace)
        obs = OfficeObs(success=True, data={"name": "Chart 1"})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"name": "Chart 1"}}
        assert "content" not in result

    def test_insert_chart_surfaces_invalid_chart_type(self, mock_workspace):
        """无效 chartType 由 AddIn 返回 4002 INVALID_PARAM；consumer 透传 (按 spec 非 3015)。"""
        tool = ExcelInsertChartTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="4002 INVALID_PARAM")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "4002 INVALID_PARAM"

    def test_insert_chart_surfaces_range_invalid(self, mock_workspace):
        """无效数据源范围由 AddIn 返回 5002 RANGE_INVALID；consumer 透传。"""
        tool = ExcelInsertChartTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5002 RANGE_INVALID")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5002 RANGE_INVALID"

    def test_update_chart_uses_base_format_result(self, mock_workspace):
        tool = ExcelUpdateChartTool(mock_workspace)
        obs = OfficeObs(success=True, data={"name": "Chart 1"})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"name": "Chart 1"}}

    def test_update_chart_surfaces_chart_not_found(self, mock_workspace):
        """图表不存在由 AddIn 返回 5007 CHART_NOT_FOUND；consumer 透传。"""
        tool = ExcelUpdateChartTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5007 CHART_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5007 CHART_NOT_FOUND"

    def test_delete_chart_uses_base_format_result(self, mock_workspace):
        tool = ExcelDeleteChartTool(mock_workspace)
        obs = OfficeObs(success=True, data={"deleted": True})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"deleted": True}}

    def test_delete_chart_surfaces_chart_not_found(self, mock_workspace):
        tool = ExcelDeleteChartTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5007 CHART_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5007 CHART_NOT_FOUND"

    # ---- #24 PivotTable 操作 ------------------------------------------------

    def test_get_pivot_tables_summary(self, mock_workspace):
        """get:pivotTables 覆写 format_result: 返回透视表名列表 content + 完整 data。"""
        tool = ExcelGetPivotTablesTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={
                "pivotTables": [
                    {"name": "销售汇总", "id": "{abcd-1234}"},
                    {"name": "PivotTable2", "id": "{efgh-5678}"},
                ]
            },
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "2 pivot table(s)" in result["content"]
        assert "销售汇总" in result["content"]
        assert "PivotTable2" in result["content"]
        assert len(result["data"]["pivotTables"]) == 2

    def test_get_pivot_tables_empty_summary(self, mock_workspace):
        tool = ExcelGetPivotTablesTool(mock_workspace)
        obs = OfficeObs(success=True, data={"pivotTables": []})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 pivot table(s)" in result["content"]

    def test_get_pivot_tables_missing_key_falls_back(self, mock_workspace):
        """data 缺 'pivotTables' 键时, .get 兜底为空列表 → '0 pivot table(s)' 不崩溃。"""
        tool = ExcelGetPivotTablesTool(mock_workspace)
        obs = OfficeObs(success=True, data={})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "0 pivot table(s)" in result["content"]

    def test_get_pivot_tables_filters_non_dict_entries(self, mock_workspace):
        """data['pivotTables'] 含非 dict 畸形条目时, isinstance 过滤 → 计数与名称列表一致。"""
        tool = ExcelGetPivotTablesTool(mock_workspace)
        obs = OfficeObs(
            success=True,
            data={"pivotTables": [{"name": "P1", "id": "{x}"}, "garbage", 7]},
        )
        result = tool.format_result(obs)
        assert result["success"] is True
        # 仅 1 个合法 dict 条目 → 计数为 1, 与名称列表一致 (不是原始 len 3)
        assert "1 pivot table(s)" in result["content"]
        assert "P1" in result["content"]
        assert "garbage" not in result["content"]

    def test_get_pivot_tables_dict_entry_missing_name_uses_placeholder(self, mock_workspace):
        """合法 dict 但缺 'name' 键时, .get 兜底 '?' 占位并计入 (与 get_charts 同语义)。"""
        tool = ExcelGetPivotTablesTool(mock_workspace)
        obs = OfficeObs(success=True, data={"pivotTables": [{"id": "{no-name}"}]})
        result = tool.format_result(obs)
        assert result["success"] is True
        assert "1 pivot table(s): ?" in result["content"]

    def test_get_pivot_tables_surfaces_worksheet_not_found(self, mock_workspace):
        """工作表不存在由 AddIn 返回 5001 WORKSHEET_NOT_FOUND；consumer 透传。"""
        tool = ExcelGetPivotTablesTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5001 WORKSHEET_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5001 WORKSHEET_NOT_FOUND"

    def test_insert_pivot_table_uses_base_format_result(self, mock_workspace):
        """写工具不覆写 format_result: 成功时返回 {success, data}，无 content。"""
        tool = ExcelInsertPivotTableTool(mock_workspace)
        obs = OfficeObs(success=True, data={"name": "PivotTable1"})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"name": "PivotTable1"}}
        assert "content" not in result

    def test_insert_pivot_table_surfaces_range_invalid(self, mock_workspace):
        """无效数据源范围由 AddIn 返回 5002 RANGE_INVALID；consumer 透传。"""
        tool = ExcelInsertPivotTableTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5002 RANGE_INVALID")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5002 RANGE_INVALID"

    def test_insert_pivot_table_surfaces_not_supported(self, mock_workspace):
        """平台不支持透视表时由 AddIn 返回 5010 NOT_SUPPORTED；consumer 透传。"""
        tool = ExcelInsertPivotTableTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5010 NOT_SUPPORTED")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5010 NOT_SUPPORTED"

    def test_delete_pivot_table_uses_base_format_result(self, mock_workspace):
        tool = ExcelDeletePivotTableTool(mock_workspace)
        obs = OfficeObs(success=True, data={"deleted": True})
        result = tool.format_result(obs)
        assert result == {"success": True, "data": {"deleted": True}}

    def test_delete_pivot_table_surfaces_pivot_not_found(self, mock_workspace):
        """透视表不存在由 AddIn 返回 5008 PIVOT_NOT_FOUND；consumer 透传。"""
        tool = ExcelDeletePivotTableTool(mock_workspace)
        obs = OfficeObs(success=False, data={}, error="5008 PIVOT_NOT_FOUND")
        result = tool.format_result(obs)
        assert result["success"] is False
        assert result["error"] == "5008 PIVOT_NOT_FOUND"
