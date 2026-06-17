"""
Test Request Wrapper

测试请求包装功能。
"""

import pytest

from office4ai.environment.workspace.socketio.request_wrapper import (
    RequestWrapperError,
    get_registered_events,
    is_wrappable_event,
    wrap_request,
)


class TestWrapRequest:
    """Test wrap_request function"""

    def test_wrap_word_get_selected_content(self) -> None:
        """Test wrapping word:get:selectedContent"""
        business_params = {
            "document_uri": "file:///test.docx",
            "options": {"includeText": True},
        }

        wrapped = wrap_request("word:get:selectedContent", business_params)

        assert "requestId" in wrapped
        assert wrapped["documentUri"] == "file:///test.docx"
        assert wrapped["options"]["includeText"] is True
        assert "timestamp" in wrapped
        assert isinstance(wrapped["timestamp"], int)

    def test_wrap_with_explicit_document_uri(self) -> None:
        """Test wrapping when document_uri passed explicitly"""
        business_params = {"options": {"includeText": False}}

        wrapped = wrap_request(
            "word:get:selectedContent",
            business_params,
            document_uri="file:///explicit.docx",
        )

        assert wrapped["documentUri"] == "file:///explicit.docx"

    def test_wrap_missing_document_uri(self) -> None:
        """Test error when document_uri missing"""
        business_params = {"options": {"includeText": True}}

        with pytest.raises(RequestWrapperError) as exc_info:
            wrap_request("word:get:selectedContent", business_params)

        assert "document_uri" in str(exc_info.value)

    def test_wrap_unknown_event(self) -> None:
        """Test error for unregistered event"""
        with pytest.raises(RequestWrapperError) as exc_info:
            wrap_request(
                "unknown:event",
                {"document_uri": "file:///test.docx"},
            )

        assert "Unknown event" in str(exc_info.value)

    def test_wrap_generates_unique_request_ids(self) -> None:
        """Test that each request gets unique requestId"""
        params = {"document_uri": "file:///test.docx"}

        wrapped1 = wrap_request("word:get:selectedContent", params)
        wrapped2 = wrap_request("word:get:selectedContent", params)

        assert wrapped1["requestId"] != wrapped2["requestId"]

    def test_wrap_word_insert_text(self) -> None:
        """Test wrapping word:insert:text with all parameters"""
        business_params = {
            "document_uri": "file:///test.docx",
            "text": "Hello World",
            "location": "Cursor",
            "format": {"bold": True, "fontSize": 14},
        }

        wrapped = wrap_request("word:insert:text", business_params)

        assert wrapped["text"] == "Hello World"
        assert wrapped["location"] == "Cursor"
        assert wrapped["format"]["bold"] is True
        assert wrapped["format"]["fontSize"] == 14

    def test_wrap_excel_get_worksheet_info(self) -> None:
        """Test wrapping excel:get:worksheetInfo (snake_case worksheet_name → camelCase alias)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "worksheet_name": "Sheet2",
        }

        wrapped = wrap_request("excel:get:worksheetInfo", business_params)

        assert wrapped["worksheetName"] == "Sheet2"
        assert wrapped["documentUri"] == "file:///test.xlsx"

    def test_wrap_excel_set_range(self) -> None:
        """Test wrapping excel:set:range (2D values + snake_case worksheet_name → camelCase)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "address": "A1:B2",
            "values": [[1, 2], [3, 4]],
            "worksheet_name": "Sheet1",
        }

        wrapped = wrap_request("excel:set:range", business_params)

        assert wrapped["address"] == "A1:B2"
        assert wrapped["values"] == [[1, 2], [3, 4]]
        assert wrapped["worksheetName"] == "Sheet1"
        assert "worksheet_name" not in wrapped

    def test_wrap_excel_get_range_includes_format_flag(self) -> None:
        """Test wrapping excel:get:range (includeFormat camelCase alias)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "address": "A1:C3",
            "include_format": True,
        }

        wrapped = wrap_request("excel:get:range", business_params)

        assert wrapped["address"] == "A1:C3"
        assert wrapped["includeFormat"] is True

    def test_wrap_excel_set_range_format(self) -> None:
        """Test wrapping excel:set:rangeFormat (nested partial format → camelCase, None dropped)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "address": "A1:C1",
            "format": {
                "font": {"bold": True, "size": 14},
                "alignment": {"horizontal": "Center", "wrap_text": True},
                "number_format": "0.00",
            },
        }

        wrapped = wrap_request("excel:set:rangeFormat", business_params)

        assert wrapped["address"] == "A1:C1"
        assert wrapped["format"]["font"]["bold"] is True
        # snake_case input fields surface as camelCase wire aliases
        assert wrapped["format"]["alignment"]["wrapText"] is True
        assert wrapped["format"]["numberFormat"] == "0.00"
        # unset optional font fields are dropped (exclude_none)
        assert "italic" not in wrapped["format"]["font"]
        assert "wrap_text" not in wrapped["format"]["alignment"]

    def test_wrap_excel_add_conditional_format_passthrough(self) -> None:
        """Test wrapping excel:add:conditionalFormat (rule passthrough keeps extra keys)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "address": "B2:B100",
            "rule": {
                "type": "cellValue",
                "operator": "greaterThan",
                "value": 90,
                "format": {"fill": {"color": "#C6EFCE"}},
            },
        }

        wrapped = wrap_request("excel:add:conditionalFormat", business_params)

        assert wrapped["address"] == "B2:B100"
        # passthrough: type required + arbitrary extra keys preserved verbatim
        assert wrapped["rule"]["type"] == "cellValue"
        assert wrapped["rule"]["operator"] == "greaterThan"
        assert wrapped["rule"]["value"] == 90
        assert wrapped["rule"]["format"]["fill"]["color"] == "#C6EFCE"

    def test_wrap_excel_merge_cells(self) -> None:
        """Test wrapping excel:merge:cells (minimal {address})"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "address": "A1:C1",
        }

        wrapped = wrap_request("excel:merge:cells", business_params)

        assert wrapped["address"] == "A1:C1"
        assert wrapped["documentUri"] == "file:///test.xlsx"

    def test_wrap_excel_add_worksheet(self) -> None:
        """Test wrapping excel:add:worksheet (optional name passes through verbatim)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "name": "数据分析",
        }

        wrapped = wrap_request("excel:add:worksheet", business_params)

        assert wrapped["name"] == "数据分析"
        assert wrapped["documentUri"] == "file:///test.xlsx"

    def test_wrap_excel_rename_worksheet(self) -> None:
        """Test wrapping excel:rename:worksheet (snake_case current_name/new_name → camelCase)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "current_name": "Sheet1",
            "new_name": "销售数据",
        }

        wrapped = wrap_request("excel:rename:worksheet", business_params)

        assert wrapped["currentName"] == "Sheet1"
        assert wrapped["newName"] == "销售数据"
        # snake_case keys must not leak to the wire payload
        assert "current_name" not in wrapped
        assert "new_name" not in wrapped

    def test_wrap_excel_delete_worksheet(self) -> None:
        """Test wrapping excel:delete:worksheet (required worksheet_name → worksheetName)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "worksheet_name": "Sheet3",
        }

        wrapped = wrap_request("excel:delete:worksheet", business_params)

        assert wrapped["worksheetName"] == "Sheet3"
        assert "worksheet_name" not in wrapped

    def test_wrap_excel_insert_table(self) -> None:
        """Test wrapping excel:insert:table (snake_case has_headers/style_name → camelCase)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "address": "A1:C2",
            "has_headers": True,
            "data": [["姓名", "年龄", "城市"], ["张三", 25, "北京"]],
            "style_name": "TableStyleMedium2",
        }

        wrapped = wrap_request("excel:insert:table", business_params)

        assert wrapped["address"] == "A1:C2"
        assert wrapped["hasHeaders"] is True
        assert wrapped["data"] == [["姓名", "年龄", "城市"], ["张三", 25, "北京"]]
        assert wrapped["styleName"] == "TableStyleMedium2"
        # snake_case keys must not leak to the wire payload
        assert "has_headers" not in wrapped
        assert "style_name" not in wrapped

    def test_wrap_excel_delete_table_row_retains_zero_index(self) -> None:
        """Test wrapping excel:delete:tableRow (row_index=0 retained, snake → camelCase)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "table_id": "Table1",
            "row_index": 0,
        }

        wrapped = wrap_request("excel:delete:tableRow", business_params)

        assert wrapped["tableId"] == "Table1"
        # rowIndex=0 是合法首行: exclude_none 不得剔除 (0 ≠ None)
        assert wrapped["rowIndex"] == 0
        assert "table_id" not in wrapped
        assert "row_index" not in wrapped

    def test_wrap_excel_sort_table_nested_fields(self) -> None:
        """Test wrapping excel:sort:table (nested sortFields → columnIndex camelCase, None dropped)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "table_id": "Table1",
            "sort_fields": [{"column_index": 1, "ascending": False}, {"column_index": 0}],
        }

        wrapped = wrap_request("excel:sort:table", business_params)

        assert wrapped["tableId"] == "Table1"
        assert wrapped["sortFields"][0] == {"columnIndex": 1, "ascending": False}
        # 第二项省略 ascending → exclude_none 剔除, 仅余 columnIndex
        assert wrapped["sortFields"][1] == {"columnIndex": 0}
        # snake_case keys must not leak (top-level or nested)
        assert "sort_fields" not in wrapped
        assert "column_index" not in wrapped["sortFields"][0]

    def test_wrap_excel_insert_chart(self) -> None:
        """Test wrapping excel:insert:chart (sourceAddress/chartType + nested position camelCase)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "source_address": "A1:C4",
            "chart_type": "ColumnClustered",
            "title": "销售对比",
            "position": {"top": 0, "left": 300, "width": 400},
        }

        wrapped = wrap_request("excel:insert:chart", business_params)

        assert wrapped["sourceAddress"] == "A1:C4"
        assert wrapped["chartType"] == "ColumnClustered"
        assert wrapped["title"] == "销售对比"
        # nested position: top=0 是合法位置, exclude_none 不得剔除 (0 ≠ None); height 省略 → 剔除
        assert wrapped["position"] == {"top": 0, "left": 300, "width": 400}
        # snake_case keys must not leak to the wire payload
        assert "source_address" not in wrapped
        assert "chart_type" not in wrapped

    def test_wrap_excel_update_chart_nested_properties(self) -> None:
        """Test wrapping excel:update:chart (nested properties partial → camelCase, None dropped)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "chart_name": "Chart 1",
            "properties": {"chart_type": "Line", "source_address": "A1:D9"},
        }

        wrapped = wrap_request("excel:update:chart", business_params)

        assert wrapped["chartName"] == "Chart 1"
        # 偏更新: 仅传入 chartType/sourceAddress; title/position 省略 → exclude_none 剔除
        assert wrapped["properties"] == {"chartType": "Line", "sourceAddress": "A1:D9"}
        # snake_case keys must not leak (top-level or nested)
        assert "chart_name" not in wrapped
        assert "chart_type" not in wrapped["properties"]
        assert "source_address" not in wrapped["properties"]

    def test_wrap_excel_insert_pivot_table(self) -> None:
        """Test wrapping excel:insert:pivotTable (sourceAddress/targetAddress snake → camelCase)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "source_address": "A1:D100",
            "target_address": "F1",
            "name": "销售汇总",
        }

        wrapped = wrap_request("excel:insert:pivotTable", business_params)

        assert wrapped["sourceAddress"] == "A1:D100"
        assert wrapped["targetAddress"] == "F1"
        assert wrapped["name"] == "销售汇总"
        # worksheetName omitted → exclude_none drops it
        assert "worksheetName" not in wrapped
        # snake_case keys must not leak to the wire payload
        assert "source_address" not in wrapped
        assert "target_address" not in wrapped

    def test_wrap_excel_delete_pivot_table(self) -> None:
        """Test wrapping excel:delete:pivotTable (pivot_table_name → pivotTableName)"""
        business_params = {
            "document_uri": "file:///test.xlsx",
            "pivot_table_name": "销售汇总",
        }

        wrapped = wrap_request("excel:delete:pivotTable", business_params)

        assert wrapped["pivotTableName"] == "销售汇总"
        assert "pivot_table_name" not in wrapped
        assert "worksheetName" not in wrapped

    def test_wrap_ppt_insert_text(self) -> None:
        """Test wrapping ppt:insert:text"""
        business_params = {
            "document_uri": "file:///test.pptx",
            "text": "Slide Title",
            "options": {"fontSize": 32},
        }

        wrapped = wrap_request("ppt:insert:text", business_params)

        assert wrapped["text"] == "Slide Title"
        assert wrapped["options"]["fontSize"] == 32


class TestIsWrappableEvent:
    """Test is_wrappable_event function"""

    def test_known_word_event(self) -> None:
        assert is_wrappable_event("word:get:selectedContent") is True

    def test_known_excel_event(self) -> None:
        assert is_wrappable_event("excel:get:selectedRange") is True

    def test_known_ppt_event(self) -> None:
        assert is_wrappable_event("ppt:insert:text") is True

    def test_unknown_event(self) -> None:
        assert is_wrappable_event("unknown:event") is False


class TestGetRegisteredEvents:
    """Test get_registered_events function"""

    def test_returns_non_empty_list(self) -> None:
        events = get_registered_events()
        assert len(events) > 0
        assert "word:get:selectedContent" in events
        assert "excel:get:selectedRange" in events
        assert "ppt:insert:text" in events

    def test_events_are_sorted(self) -> None:
        events = get_registered_events()
        # Check if sorted
        assert events == sorted(events)

    def test_contains_all_word_events(self) -> None:
        events = get_registered_events()
        word_events = [e for e in events if e.startswith("word:")]
        # Should have at least 13 Word events
        assert len(word_events) >= 13

    def test_contains_all_excel_events(self) -> None:
        events = get_registered_events()
        excel_events = [e for e in events if e.startswith("excel:")]
        # #18 read slice (3) + #19 Range CRUD + 公式 (7) + #20 Format/条件格式/合并 (6)
        # + #21 Worksheet 管理 (5) + #22 Table 操作 (6) + #23 Chart 操作 (4)
        # + #24 PivotTable 操作 (3); remaining /excel events land with #25–#26.
        assert set(excel_events) >= {
            "excel:get:workbookInfo",
            "excel:get:worksheetInfo",
            "excel:get:selectedRange",
            "excel:get:range",
            "excel:set:range",
            "excel:clear:range",
            "excel:copy:range",
            "excel:delete:range",
            "excel:insert:range",
            "excel:set:formula",
            "excel:get:rangeFormat",
            "excel:set:rangeFormat",
            "excel:add:conditionalFormat",
            "excel:clear:conditionalFormat",
            "excel:merge:cells",
            "excel:unmerge:cells",
            "excel:get:worksheets",
            "excel:add:worksheet",
            "excel:delete:worksheet",
            "excel:rename:worksheet",
            "excel:activate:worksheet",
            "excel:insert:table",
            "excel:get:table",
            "excel:get:tables",
            "excel:add:tableRow",
            "excel:delete:tableRow",
            "excel:sort:table",
            "excel:insert:chart",
            "excel:get:charts",
            "excel:update:chart",
            "excel:delete:chart",
            "excel:insert:pivotTable",
            "excel:get:pivotTables",
            "excel:delete:pivotTable",
        }

    def test_contains_all_ppt_events(self) -> None:
        events = get_registered_events()
        ppt_events = [e for e in events if e.startswith("ppt:")]
        # Should have at least 10 PPT events
        assert len(ppt_events) >= 10
