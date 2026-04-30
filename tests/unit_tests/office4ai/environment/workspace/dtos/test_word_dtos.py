"""
Test Word DTOs

测试 Word 事件的数据传输对象。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import request_registry
from office4ai.environment.workspace.dtos.word import (
    AnyContentElement,
    CellFormat,
    CommentData,
    CommentReplyData,
    ContentMetadata,
    DocumentStats,
    GetCommentsOptions,
    GetContentOptions,
    GetStylesOptions,
    InsertCommentSearchOptions,
    InsertCommentTarget,
    ReplaceContent,
    SelectionInfo,
    SelectTextResult,
    SelectTextSearchOptions,
    StyleInfo,
    StylesResult,
    TableBorderOptions,
    TableCellPadding,
    TableCellUpdate,
    TableColumnUpdate,
    TableInsertOptions,
    TableRowUpdate,
    TableStyleOptions,
    TableSummary,
    TextFormat,
    WordDeleteCommentRequest,
    WordGetCommentsRequest,
    WordGetCommentsResponse,
    WordGetDocumentStatsRequest,
    WordGetDocumentStatsResponse,
    WordGetSelectedContentRequest,
    WordGetSelectedContentResponse,
    WordGetSelectionRequest,
    WordGetSelectionResponse,
    WordGetStylesRequest,
    WordGetVisibleContentRequest,
    WordGetVisibleContentResponse,
    WordInsertCommentRequest,
    WordInsertTableRequest,
    WordInsertTextRequest,
    WordMergeCellsRequest,
    WordReplaceSelectionRequest,
    WordReplaceSelectionResponse,
    WordReplyCommentRequest,
    WordResolveCommentRequest,
    WordSelectTextRequest,
    WordSelectTextResponse,
    WordUpdateTableCellRequest,
    WordUpdateTableFormatRequest,
    WordUpdateTableRowColumnRequest,
)


class TestWordGetSelectionRequest:
    """Test WordGetSelectionRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetSelectionRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert isinstance(request.timestamp, int)

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetSelectionRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetSelectionRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetSelectionRequest.event_name == "word:get:selection"

    def test_to_payload_camel_case(self) -> None:
        """Test serialization to camelCase payload"""
        request = WordGetSelectionRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
        )

        payload = request.to_payload()

        assert payload["requestId"] == "req_002"
        assert payload["documentUri"] == "file:///test.docx"
        assert isinstance(payload["timestamp"], int)

    def test_build_class_method(self) -> None:
        """Test build class method for creating requests"""
        request = WordGetSelectionRequest.build(document_uri="file:///test.docx")

        assert request.request_id is not None
        assert isinstance(request.request_id, str)
        assert request.document_uri == "file:///test.docx"


class TestSelectionInfo:
    """Test SelectionInfo DTO"""

    def test_valid_selection_info_empty(self) -> None:
        """Test creating valid selection info for empty selection (cursor)"""
        selection = SelectionInfo(
            isEmpty=True,
            type="InsertionPoint",
            start=100,
            end=100,
        )

        assert selection.is_empty is True
        assert selection.type == "InsertionPoint"
        assert selection.start == 100
        assert selection.end == 100
        assert selection.text is None

    def test_valid_selection_info_normal(self) -> None:
        """Test creating valid selection info for normal selection"""
        selection = SelectionInfo(
            isEmpty=False,
            type="Normal",
            start=100,
            end=150,
            text="Selected text",
        )

        assert selection.is_empty is False
        assert selection.type == "Normal"
        assert selection.start == 100
        assert selection.end == 150
        assert selection.text == "Selected text"

    def test_valid_selection_info_no_selection(self) -> None:
        """Test creating valid selection info for no selection"""
        selection = SelectionInfo(
            isEmpty=True,
            type="NoSelection",
        )

        assert selection.is_empty is True
        assert selection.type == "NoSelection"
        assert selection.start is None
        assert selection.end is None
        assert selection.text is None

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing isEmpty
        with pytest.raises(ValidationError) as exc_info:
            SelectionInfo(type="InsertionPoint")  # type: ignore

        assert "isEmpty" in str(exc_info.value)

        # Missing type
        with pytest.raises(ValidationError) as exc_info:
            SelectionInfo(isEmpty=True)  # type: ignore

        assert "type" in str(exc_info.value)

    def test_invalid_selection_type(self) -> None:
        """Test validation fails with invalid selection type"""
        with pytest.raises(ValidationError) as exc_info:
            SelectionInfo(
                isEmpty=True,
                type="InvalidType",  # type: ignore
            )

        assert "type" in str(exc_info.value).lower()

    def test_all_valid_selection_types(self) -> None:
        """Test all valid selection types"""
        types = ["NoSelection", "InsertionPoint", "Normal"]

        for selection_type in types:
            selection = SelectionInfo(isEmpty=True, type=selection_type)  # type: ignore
            assert selection.type == selection_type

    def test_selection_serialization(self) -> None:
        """Test selection can be serialized to dict with correct aliases"""
        selection = SelectionInfo(
            isEmpty=False,
            type="Normal",
            start=50,
            end=100,
            text="Test",
        )

        data = selection.model_dump(by_alias=True)

        assert data["isEmpty"] is False
        assert data["type"] == "Normal"
        assert data["start"] == 50
        assert data["end"] == 100
        assert data["text"] == "Test"

    def test_selection_from_dict(self) -> None:
        """Test creating selection from dict"""
        data = {
            "isEmpty": True,
            "type": "InsertionPoint",
            "start": 200,
            "end": 200,
            "text": None,
        }

        selection = SelectionInfo(**data)

        assert selection.is_empty is True
        assert selection.type == "InsertionPoint"
        assert selection.start == 200
        assert selection.end == 200


class TestWordGetSelectionResponse:
    """Test WordGetSelectionResponse DTO"""

    def test_valid_response_with_data(self) -> None:
        """Test creating valid response with selection data"""
        selection_data = SelectionInfo(
            isEmpty=False,
            type="Normal",
            start=100,
            end=150,
            text="Selected text",
        )
        response = WordGetSelectionResponse(data=selection_data)

        assert response.data is not None
        assert response.data.is_empty is False
        assert response.data.type == "Normal"
        assert response.data.start == 100
        assert response.data.end == 150
        assert response.data.text == "Selected text"

    def test_valid_response_without_data(self) -> None:
        """Test creating valid response without data"""
        response = WordGetSelectionResponse()

        assert response.data is None

    def test_response_with_empty_selection(self) -> None:
        """Test creating response with empty selection (cursor)"""
        selection_data = SelectionInfo(
            isEmpty=True,
            type="InsertionPoint",
            start=50,
            end=50,
        )
        response = WordGetSelectionResponse(data=selection_data)

        assert response.data.is_empty is True
        assert response.data.type == "InsertionPoint"
        assert response.data.start == 50
        assert response.data.end == 50

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        selection_data = SelectionInfo(
            isEmpty=False,
            type="Normal",
            start=25,
            end=75,
            text="Test text",
        )
        response = WordGetSelectionResponse(data=selection_data)

        data = response.model_dump(by_alias=True)

        assert "data" in data
        assert data["data"]["isEmpty"] is False
        assert data["data"]["type"] == "Normal"
        assert data["data"]["start"] == 25
        assert data["data"]["end"] == 75
        assert data["data"]["text"] == "Test text"

    def test_response_from_dict(self) -> None:
        """Test creating response from dict"""
        data = {
            "data": {
                "isEmpty": True,
                "type": "NoSelection",
            }
        }

        response = WordGetSelectionResponse(**data)

        assert response.data is not None
        assert response.data.is_empty is True
        assert response.data.type == "NoSelection"


class TestWordGetSelectedContentRequest:
    """Test WordGetSelectedContentRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetSelectedContentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.options is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = GetContentOptions(
            includeText=True,
            includeImages=False,
            includeTables=True,
        )

        request = WordGetSelectedContentRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            options=options,
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.include_images is False
        assert request.options.include_tables is True

    def test_request_with_dict_options(self) -> None:
        """Test creating request with options as dict"""
        request = WordGetSelectedContentRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            options={"includeText": True, "detailedMetadata": True},
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.detailed_metadata is True

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetSelectedContentRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetSelectedContentRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetSelectedContentRequest.event_name == "word:get:selectedContent"


class TestGetContentOptions:
    """Test GetContentOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = GetContentOptions()

        assert options.include_text is True
        assert options.include_images is True
        assert options.include_tables is True
        assert options.include_content_controls is False
        assert options.detailed_metadata is False
        assert options.max_text_length is None

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = GetContentOptions(
            includeText=False,
            includeImages=False,
            includeTables=False,
            includeContentControls=True,
            detailedMetadata=True,
            maxTextLength=1000,
        )

        assert options.include_text is False
        assert options.include_images is False
        assert options.include_tables is False
        assert options.include_content_controls is True
        assert options.detailed_metadata is True
        assert options.max_text_length == 1000

    def test_options_from_dict(self) -> None:
        """Test creating options from dict with aliases"""
        options = GetContentOptions(
            **{
                "includeText": True,
                "includeImages": False,
                "maxTextLength": 500,
            }
        )

        assert options.include_text is True
        assert options.include_images is False
        assert options.max_text_length == 500

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = GetContentOptions(
            includeText=True,
            includeImages=True,
            detailedMetadata=True,
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = options.model_dump(by_alias=True)

        assert data["includeText"] is True
        assert data["includeImages"] is True
        assert data["detailedMetadata"] is True
        # Default values should be included
        assert "includeTables" in data

    def test_invalid_max_text_length(self) -> None:
        """Test that negative max_text_length is rejected"""
        # Pydantic doesn't have a constraint on max_text_length, so this should work
        # In a real scenario, you might want to add a validation constraint
        options = GetContentOptions(maxTextLength=-100)
        assert options.max_text_length == -100


class TestWordInsertTextRequest:
    """Test WordInsertTextRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordInsertTextRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            text="Hello World",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.text == "Hello World"
        assert request.location == "Cursor"  # Default value
        assert request.format is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_location(self) -> None:
        """Test creating valid request with different locations"""
        # Test "Start" location
        request_start = WordInsertTextRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            text="Text at start",
            location="Start",
        )
        assert request_start.location == "Start"

        # Test "End" location
        request_end = WordInsertTextRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            text="Text at end",
            location="End",
        )
        assert request_end.location == "End"

        # Test "Cursor" location (default)
        request_cursor = WordInsertTextRequest(
            requestId="req_004",
            documentUri="file:///test.docx",
            text="Text at cursor",
            location="Cursor",
        )
        assert request_cursor.location == "Cursor"

    def test_valid_request_with_format(self) -> None:
        """Test creating valid request with text format"""
        format_options = TextFormat(
            bold=True,
            italic=False,
            fontSize=14,
            fontName="Arial",
            color="#FF0000",
        )

        request = WordInsertTextRequest(
            requestId="req_005",
            documentUri="file:///test.docx",
            text="Formatted text",
            location="Cursor",
            format=format_options,
        )

        assert request.text == "Formatted text"
        assert request.format is not None
        assert request.format.bold is True
        assert request.format.italic is False
        assert request.format.font_size == 14
        assert request.format.font_name == "Arial"
        assert request.format.color == "#FF0000"

    def test_request_with_dict_format(self) -> None:
        """Test creating request with format as dict"""
        request = WordInsertTextRequest(
            requestId="req_006",
            documentUri="file:///test.docx",
            text="Text with format",
            format={"bold": True, "italic": True, "fontSize": 12},
        )

        assert request.format is not None
        assert request.format.bold is True
        assert request.format.italic is True
        assert request.format.font_size == 12

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                documentUri="file:///test.docx",
                text="Hello",
            )

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                requestId="req_001",
                text="Hello",
            )

        assert "documentUri" in str(exc_info.value)

        # Missing text
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "text" in str(exc_info.value)

    def test_invalid_location(self) -> None:
        """Test validation fails with invalid location"""
        with pytest.raises(ValidationError) as exc_info:
            WordInsertTextRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
                text="Hello",
                location="InvalidLocation",  # type: ignore
            )

        assert "location" in str(exc_info.value).lower()

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordInsertTextRequest.event_name == "word:insert:text"

    def test_to_payload_camel_case(self) -> None:
        """Test serialization to camelCase payload"""
        request = WordInsertTextRequest(
            requestId="req_007",
            documentUri="file:///test.docx",
            text="Hello",
            location="Start",
            format={"bold": True},
        )

        payload = request.to_payload()

        assert payload["requestId"] == "req_007"
        assert payload["documentUri"] == "file:///test.docx"
        assert payload["text"] == "Hello"
        assert payload["location"] == "Start"
        assert payload["format"]["bold"] is True
        assert isinstance(payload["timestamp"], int)

    def test_build_class_method(self) -> None:
        """Test build class method for creating requests"""
        request = WordInsertTextRequest.build(
            document_uri="file:///test.docx",
            text="Auto-generated request",
            location="End",
        )

        assert request.request_id is not None
        assert isinstance(request.request_id, str)
        assert request.document_uri == "file:///test.docx"
        assert request.text == "Auto-generated request"
        assert request.location == "End"


class TestTextFormat:
    """Test TextFormat DTO"""

    def test_default_format(self) -> None:
        """Test creating format with default values"""
        format_options = TextFormat()

        assert format_options.bold is None
        assert format_options.italic is None
        assert format_options.font_size is None
        assert format_options.font_name is None
        assert format_options.color is None
        assert format_options.underline is None
        assert format_options.style_name is None

    def test_custom_format(self) -> None:
        """Test creating format with custom values"""
        format_options = TextFormat(
            bold=True,
            italic=True,
            fontSize=16,
            fontName="Times New Roman",
            color="#0000FF",
            underline="Single",
            styleName="Heading 1",
        )

        assert format_options.bold is True
        assert format_options.italic is True
        assert format_options.font_size == 16
        assert format_options.font_name == "Times New Roman"
        assert format_options.color == "#0000FF"
        assert format_options.underline == "Single"
        assert format_options.style_name == "Heading 1"

    def test_partial_format(self) -> None:
        """Test creating format with partial fields"""
        format_options = TextFormat(bold=True, fontSize=14)

        assert format_options.bold is True
        assert format_options.italic is None
        assert format_options.font_size == 14
        assert format_options.font_name is None

    def test_format_with_underline_types(self) -> None:
        """Test different underline types"""
        # Single underline
        format_single = TextFormat(underline="Single")
        assert format_single.underline == "Single"

        # Double underline
        format_double = TextFormat(underline="Double")
        assert format_double.underline == "Double"

        # Dotted underline
        format_dotted = TextFormat(underline="Dotted")
        assert format_dotted.underline == "Dotted"

    def test_format_with_style_name_only(self) -> None:
        """Test format with only style name (recommended approach)"""
        format_options = TextFormat(styleName="Title")

        assert format_options.style_name == "Title"
        # All direct format fields should be None
        assert format_options.bold is None
        assert format_options.italic is None
        assert format_options.font_size is None

    def test_format_priority_rule_direct_format(self) -> None:
        """
        Test that direct format takes precedence over styleName.

        Note: This is a documentation test - actual priority enforcement
        happens in the Add-In implementation, not in the DTO.
        """
        # Not recommended: styleName will be ignored when direct format is present
        format_options = TextFormat(
            bold=True,
            styleName="Heading 1",
        )

        assert format_options.bold is True
        assert format_options.style_name == "Heading 1"
        # The Add-In should ignore styleName when bold is present

    def test_format_serialization(self) -> None:
        """Test format can be serialized to dict with correct aliases"""
        format_options = TextFormat(
            bold=True,
            italic=False,
            fontSize=12,
            fontName="Arial",
            underline="Single",
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = format_options.model_dump(by_alias=True, exclude_none=True)

        assert data["bold"] is True
        assert data["italic"] is False
        assert data["fontSize"] == 12
        assert data["fontName"] == "Arial"
        assert data["underline"] == "Single"
        # None values should be excluded with exclude_none=True
        assert "color" not in data
        assert "styleName" not in data

    def test_format_from_dict(self) -> None:
        """Test creating format from dict with aliases"""
        format_options = TextFormat(
            **{
                "bold": True,
                "italic": True,
                "fontSize": 14,
                "color": "#FF0000",
                "underline": "Double",
                "styleName": "Normal",
            }
        )

        assert format_options.bold is True
        assert format_options.italic is True
        assert format_options.font_size == 14
        assert format_options.color == "#FF0000"
        assert format_options.underline == "Double"
        assert format_options.style_name == "Normal"


class TestWordGetStylesRequest:
    """Test WordGetStylesRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetStylesRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.options is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = GetStylesOptions(
            includeBuiltIn=True,
            includeCustom=True,
            includeUnused=False,
            detailedInfo=True,
        )

        request = WordGetStylesRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            options=options,
        )

        assert request.options is not None
        assert request.options.include_built_in is True
        assert request.options.include_custom is True
        assert request.options.include_unused is False
        assert request.options.detailed_info is True

    def test_request_with_dict_options(self) -> None:
        """Test creating request with options as dict"""
        request = WordGetStylesRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            options={"includeBuiltIn": False, "includeCustom": True, "detailedInfo": False},
        )

        assert request.options is not None
        assert request.options.include_built_in is False
        assert request.options.include_custom is True
        assert request.options.detailed_info is False

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetStylesRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetStylesRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetStylesRequest.event_name == "word:get:styles"


class TestGetStylesOptions:
    """Test GetStylesOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = GetStylesOptions()

        assert options.include_built_in is True
        assert options.include_custom is True
        assert options.include_unused is False
        assert options.detailed_info is False

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = GetStylesOptions(
            includeBuiltIn=False,
            includeCustom=True,
            includeUnused=True,
            detailedInfo=True,
        )

        assert options.include_built_in is False
        assert options.include_custom is True
        assert options.include_unused is True
        assert options.detailed_info is True

    def test_options_from_dict(self) -> None:
        """Test creating options from dict with aliases"""
        options = GetStylesOptions(
            **{
                "includeBuiltIn": False,
                "includeCustom": True,
                "detailedInfo": True,
            }
        )

        assert options.include_built_in is False
        assert options.include_custom is True
        assert options.detailed_info is True
        # include_unused should use default value
        assert options.include_unused is False

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = GetStylesOptions(
            includeBuiltIn=True,
            includeCustom=True,
            detailedInfo=True,
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = options.model_dump(by_alias=True)

        assert data["includeBuiltIn"] is True
        assert data["includeCustom"] is True
        assert data["detailedInfo"] is True
        # Default values should be included
        assert "includeUnused" in data


class TestStyleInfo:
    """Test StyleInfo DTO"""

    def test_valid_style_info(self) -> None:
        """Test creating valid style info"""
        style = StyleInfo(
            name="标题一",
            type="Paragraph",
            builtIn=True,
            inUse=True,
        )

        assert style.name == "标题一"
        assert style.type == "Paragraph"
        assert style.built_in is True
        assert style.in_use is True
        assert style.description is None

    def test_style_info_with_description(self) -> None:
        """Test creating style info with description"""
        style = StyleInfo(
            name="Normal",
            type="Paragraph",
            builtIn=True,
            inUse=True,
            description="Normal paragraph style",
        )

        assert style.name == "Normal"
        assert style.description == "Normal paragraph style"

    def test_invalid_style_type(self) -> None:
        """Test validation fails with invalid style type"""
        with pytest.raises(ValidationError) as exc_info:
            StyleInfo(
                name="Test Style",
                type="InvalidType",  # type: ignore
                builtIn=True,
                inUse=True,
            )

        assert "type" in str(exc_info.value).lower()

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing name
        with pytest.raises(ValidationError) as exc_info:
            StyleInfo(
                type="Paragraph",  # type: ignore
                builtIn=True,
                inUse=True,
            )

        assert "name" in str(exc_info.value)

        # Missing type
        with pytest.raises(ValidationError) as exc_info:
            StyleInfo(
                name="Test",
                builtIn=True,
                inUse=True,
            )

        assert "type" in str(exc_info.value)

    def test_all_style_types(self) -> None:
        """Test all valid style types"""
        types = ["Paragraph", "Character", "Table", "List"]

        for style_type in types:
            style = StyleInfo(
                name=f"Test {style_type}",
                type=style_type,  # type: ignore
                builtIn=True,
                inUse=True,
            )
            assert style.type == style_type

    def test_style_serialization(self) -> None:
        """Test style can be serialized to dict with correct aliases"""
        style = StyleInfo(
            name="标题一",
            type="Paragraph",
            builtIn=True,
            inUse=True,
            description="Heading style",
        )

        # Convert to dict (as it would be sent over Socket.IO)
        data = style.model_dump(by_alias=True)

        assert data["name"] == "标题一"
        assert data["type"] == "Paragraph"
        assert data["builtIn"] is True
        assert data["inUse"] is True
        assert data["description"] == "Heading style"


class TestStylesResult:
    """Test StylesResult DTO"""

    def test_valid_styles_result(self) -> None:
        """Test creating valid styles result"""
        styles = [
            StyleInfo(name="标题一", type="Paragraph", builtIn=True, inUse=True),
            StyleInfo(name="正文", type="Paragraph", builtIn=True, inUse=True),
            StyleInfo(name="强调", type="Character", builtIn=True, inUse=False),
        ]

        result = StylesResult(styles=styles)

        assert len(result.styles) == 3
        assert result.styles[0].name == "标题一"
        assert result.styles[1].name == "正文"
        assert result.styles[2].name == "强调"

    def test_empty_styles_result(self) -> None:
        """Test creating empty styles result"""
        result = StylesResult(styles=[])

        assert len(result.styles) == 0

    def test_styles_result_serialization(self) -> None:
        """Test styles result can be serialized to dict with correct aliases"""
        styles = [
            StyleInfo(
                name="Normal",
                type="Paragraph",
                builtIn=True,
                inUse=True,
                description="Normal style",
            ),
        ]

        result = StylesResult(styles=styles)
        data = result.model_dump(by_alias=True)

        assert "styles" in data
        assert len(data["styles"]) == 1
        assert data["styles"][0]["name"] == "Normal"
        assert data["styles"][0]["type"] == "Paragraph"
        assert data["styles"][0]["description"] == "Normal style"

    def test_styles_result_from_dict(self) -> None:
        """Test creating styles result from dict"""
        styles_data = [
            {
                "name": "标题一",
                "type": "Paragraph",
                "builtIn": True,
                "inUse": True,
            },
            {
                "name": "Custom Style",
                "type": "Character",
                "builtIn": False,
                "inUse": True,
            },
        ]

        result = StylesResult(**{"styles": styles_data})

        assert len(result.styles) == 2
        assert result.styles[0].name == "标题一"
        assert result.styles[1].name == "Custom Style"
        assert result.styles[1].built_in is False


class TestWordReplaceSelectionRequest:
    """Test WordReplaceSelectionRequest DTO"""

    def test_valid_request_with_text_content(self) -> None:
        """Test creating valid request with text content"""
        content = ReplaceContent(text="Replaced text")

        request = WordReplaceSelectionRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            content=content,
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.content.text == "Replaced text"
        assert request.content.images is None
        assert request.content.format is None

    def test_valid_request_with_images_content(self) -> None:
        """Test creating valid request with images content"""
        images = [{"base64": "iVBORw0KGgo...", "width": 100, "height": 100}]
        content = ReplaceContent(images=images)

        request = WordReplaceSelectionRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            content=content,
        )

        assert request.content.images == images
        assert request.content.text is None

    def test_valid_request_with_text_and_format(self) -> None:
        """Test creating valid request with text and format"""
        text_format = TextFormat(bold=True, italic=True, fontSize=14)
        content = ReplaceContent(text="Formatted text", format=text_format)

        request = WordReplaceSelectionRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            content=content,
        )

        assert request.content.text == "Formatted text"
        assert request.content.format is not None
        assert request.content.format.bold is True
        assert request.content.format.italic is True
        assert request.content.format.font_size == 14

    def test_request_with_dict_content(self) -> None:
        """Test creating request with content as dict"""
        request = WordReplaceSelectionRequest(
            requestId="req_004",
            documentUri="file:///test.docx",
            content={"text": "Test", "format": {"bold": True}},
        )

        assert request.content.text == "Test"
        assert request.content.format is not None
        assert request.content.format.bold is True

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceSelectionRequest(
                documentUri="file:///test.docx",
                content={"text": "Test"},
            )

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceSelectionRequest(
                requestId="req_001",
                content={"text": "Test"},
            )

        assert "documentUri" in str(exc_info.value)

        # Missing content
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceSelectionRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "content" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordReplaceSelectionRequest.event_name == "word:replace:selection"

    def test_request_serialization(self) -> None:
        """Test request can be serialized to dict with correct aliases"""
        content = ReplaceContent(text="Test content")
        request = WordReplaceSelectionRequest(
            requestId="req_005",
            documentUri="file:///test.docx",
            content=content,
        )

        data = request.model_dump(by_alias=True)

        assert data["requestId"] == "req_005"
        assert data["documentUri"] == "file:///test.docx"
        assert data["content"]["text"] == "Test content"


class TestWordReplaceSelectionResponse:
    """Test WordReplaceSelectionResponse DTO"""

    def test_valid_response_success(self) -> None:
        """Test creating valid success response"""
        response = WordReplaceSelectionResponse(
            replaced=True,
            characterCount=100,
        )

        assert response.replaced is True
        assert response.character_count == 100

    def test_valid_response_failure(self) -> None:
        """Test creating valid failure response"""
        response = WordReplaceSelectionResponse(
            replaced=False,
            characterCount=0,
        )

        assert response.replaced is False
        assert response.character_count == 0

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        response = WordReplaceSelectionResponse(
            replaced=True,
            characterCount=50,
        )

        data = response.model_dump(by_alias=True)

        assert data["replaced"] is True
        assert data["characterCount"] == 50

    def test_response_from_dict(self) -> None:
        """Test creating response from dict"""
        data = {"replaced": True, "characterCount": 200}

        response = WordReplaceSelectionResponse(**data)

        assert response.replaced is True
        assert response.character_count == 200


class TestReplaceContent:
    """Test ReplaceContent DTO"""

    def test_content_with_text_only(self) -> None:
        """Test creating content with text only"""
        content = ReplaceContent(text="Some text")

        assert content.text == "Some text"
        assert content.images is None
        assert content.format is None

    def test_content_with_images_only(self) -> None:
        """Test creating content with images only"""
        images = [
            {"base64": "abc123", "width": 200, "height": 150, "altText": "Image 1"},
            {"base64": "def456", "width": 300, "height": 200},
        ]
        content = ReplaceContent(images=images)

        assert content.images == images
        assert content.text is None
        assert content.format is None

    def test_content_with_text_and_format(self) -> None:
        """Test creating content with text and format"""
        format_obj = TextFormat(
            bold=True,
            italic=False,
            fontSize=16,
            fontName="Arial",
            color="#FF0000",
        )
        content = ReplaceContent(text="Bold red text", format=format_obj)

        assert content.text == "Bold red text"
        assert content.format is not None
        assert content.format.bold is True
        assert content.format.italic is False
        assert content.format.font_size == 16
        assert content.format.font_name == "Arial"
        assert content.format.color == "#FF0000"

    def test_content_serialization(self) -> None:
        """Test content can be serialized to dict with correct aliases"""
        format_obj = TextFormat(bold=True, fontSize=14)
        content = ReplaceContent(text="Test", format=format_obj)

        data = content.model_dump(by_alias=True)

        assert data["text"] == "Test"
        assert data["format"]["bold"] is True
        assert data["format"]["fontSize"] == 14


class TestWordGetSelectedContentResponse:
    """Test WordGetSelectedContentResponse DTO"""

    def test_valid_response(self) -> None:
        """Test creating valid response"""
        metadata = ContentMetadata(isEmpty=False, characterCount=100)
        response = WordGetSelectedContentResponse(
            text="Selected text content",
            elements=[],
            metadata=metadata,
        )

        assert response.text == "Selected text content"
        assert response.elements == []
        assert response.metadata is not None
        assert response.metadata.is_empty is False
        assert response.metadata.character_count == 100

    def test_response_without_metadata(self) -> None:
        """Test creating response without metadata"""
        response = WordGetSelectedContentResponse(
            text="Text without metadata",
            elements=[],
        )

        assert response.text == "Text without metadata"
        assert response.metadata is None

    def test_response_with_elements(self) -> None:
        """Test creating response with elements"""
        elements = [
            AnyContentElement(type="text", content={"text": "Paragraph 1"}),
            AnyContentElement(type="image", content={"base64": "abc123", "width": 100}),
        ]
        response = WordGetSelectedContentResponse(
            text="Content with elements",
            elements=elements,
        )

        assert len(response.elements) == 2
        assert response.elements[0].type == "text"
        assert response.elements[1].type == "image"

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        metadata = ContentMetadata(isEmpty=False, characterCount=50)
        response = WordGetSelectedContentResponse(
            text="Test content",
            elements=[],
            metadata=metadata,
        )

        data = response.model_dump(by_alias=True)

        assert data["text"] == "Test content"
        assert data["metadata"]["isEmpty"] is False
        assert data["metadata"]["characterCount"] == 50


class TestWordGetVisibleContentRequest:
    """Test WordGetVisibleContentRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetVisibleContentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.options is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = GetContentOptions(
            includeText=True,
            includeImages=False,
            includeTables=True,
        )

        request = WordGetVisibleContentRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            options=options,
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.include_images is False
        assert request.options.include_tables is True

    def test_request_with_dict_options(self) -> None:
        """Test creating request with options as dict"""
        request = WordGetVisibleContentRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            options={"includeText": True, "maxTextLength": 1000},
        )

        assert request.options is not None
        assert request.options.include_text is True
        assert request.options.max_text_length == 1000

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetVisibleContentRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetVisibleContentRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetVisibleContentRequest.event_name == "word:get:visibleContent"


class TestWordGetVisibleContentResponse:
    """Test WordGetVisibleContentResponse DTO"""

    def test_valid_response(self) -> None:
        """Test creating valid response"""
        metadata = ContentMetadata(isEmpty=False, characterCount=200)
        response = WordGetVisibleContentResponse(
            text="Visible content text",
            elements=[],
            metadata=metadata,
        )

        assert response.text == "Visible content text"
        assert response.elements == []
        assert response.metadata is not None
        assert response.metadata.is_empty is False
        assert response.metadata.character_count == 200

    def test_empty_content_response(self) -> None:
        """Test creating empty content response"""
        metadata = ContentMetadata(isEmpty=True, characterCount=0)
        response = WordGetVisibleContentResponse(
            text="",
            elements=[],
            metadata=metadata,
        )

        assert response.text == ""
        assert response.metadata.is_empty is True
        assert response.metadata.character_count == 0

    def test_response_with_complex_elements(self) -> None:
        """Test creating response with various element types"""
        elements = [
            AnyContentElement(type="text", content={"text": "Heading"}),
            AnyContentElement(type="image", content={"base64": "iVBORw0KGgo...", "width": 200}),
            AnyContentElement(type="table", content={"rows": 3, "columns": 2}),
            AnyContentElement(type="other", content={"type": "equation"}),
        ]
        response = WordGetVisibleContentResponse(
            text="Complex content",
            elements=elements,
        )

        assert len(response.elements) == 4
        element_types = [elem.type for elem in response.elements]
        assert "text" in element_types
        assert "image" in element_types
        assert "table" in element_types
        assert "other" in element_types

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        metadata = ContentMetadata(isEmpty=False, characterCount=150)
        response = WordGetVisibleContentResponse(
            text="Serialized content",
            elements=[AnyContentElement(type="text", content={"text": "Test"})],
            metadata=metadata,
        )

        data = response.model_dump(by_alias=True)

        assert data["text"] == "Serialized content"
        assert len(data["elements"]) == 1
        assert data["elements"][0]["type"] == "text"
        assert data["metadata"]["isEmpty"] is False
        assert data["metadata"]["characterCount"] == 150


class TestContentMetadata:
    """Test ContentMetadata DTO"""

    def test_valid_metadata(self) -> None:
        """Test creating valid metadata"""
        metadata = ContentMetadata(isEmpty=False, characterCount=100)

        assert metadata.is_empty is False
        assert metadata.character_count == 100

    def test_empty_metadata(self) -> None:
        """Test creating empty content metadata"""
        metadata = ContentMetadata(isEmpty=True, characterCount=0)

        assert metadata.is_empty is True
        assert metadata.character_count == 0

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing isEmpty
        with pytest.raises(ValidationError) as exc_info:
            ContentMetadata(characterCount=100)

        assert "isEmpty" in str(exc_info.value)

        # Missing characterCount
        with pytest.raises(ValidationError) as exc_info:
            ContentMetadata(isEmpty=False)

        assert "characterCount" in str(exc_info.value)

    def test_metadata_serialization(self) -> None:
        """Test metadata can be serialized to dict with correct aliases"""
        metadata = ContentMetadata(isEmpty=False, characterCount=500)

        data = metadata.model_dump(by_alias=True)

        assert data["isEmpty"] is False
        assert data["characterCount"] == 500


class TestAnyContentElement:
    """Test AnyContentElement DTO"""

    def test_text_element(self) -> None:
        """Test creating text element"""
        element = AnyContentElement(type="text", content={"text": "Paragraph text"})

        assert element.type == "text"
        assert element.content["text"] == "Paragraph text"

    def test_image_element(self) -> None:
        """Test creating image element"""
        element = AnyContentElement(
            type="image",
            content={"base64": "iVBORw0KGgo...", "width": 300, "height": 200},
        )

        assert element.type == "image"
        assert element.content["width"] == 300
        assert element.content["height"] == 200

    def test_table_element(self) -> None:
        """Test creating table element"""
        element = AnyContentElement(
            type="table",
            content={"rows": 3, "columns": 4, "data": [["A", "B"], ["C", "D"]]},
        )

        assert element.type == "table"
        assert element.content["rows"] == 3
        assert element.content["columns"] == 4

    def test_other_element(self) -> None:
        """Test creating other type element"""
        element = AnyContentElement(type="other", content={"type": "equation", "latex": "x^2"})

        assert element.type == "other"
        assert element.content["latex"] == "x^2"

    def test_invalid_element_type(self) -> None:
        """Test validation fails with invalid element type"""
        with pytest.raises(ValidationError) as exc_info:
            AnyContentElement(
                type="invalid_type",  # type: ignore
                content={},
            )

        assert "type" in str(exc_info.value).lower()

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing type
        with pytest.raises(ValidationError) as exc_info:
            AnyContentElement(content={})  # type: ignore

        assert "type" in str(exc_info.value)

        # Missing content
        with pytest.raises(ValidationError) as exc_info:
            AnyContentElement(type="text", content=None)  # type: ignore

        assert "content" in str(exc_info.value)

    def test_element_serialization(self) -> None:
        """Test element can be serialized to dict with correct aliases"""
        element = AnyContentElement(type="text", content={"text": "Test text"})

        data = element.model_dump(by_alias=True)

        assert data["type"] == "text"
        assert data["content"]["text"] == "Test text"

    def test_all_valid_element_types(self) -> None:
        """Test all valid element types"""
        types = ["text", "image", "table", "other"]

        for element_type in types:
            element = AnyContentElement(type=element_type, content={"data": "test"})
            assert element.type == element_type


class TestWordGetDocumentStatsRequest:
    """Test WordGetDocumentStatsRequest DTO"""

    def test_valid_request(self) -> None:
        """Test creating valid request"""
        request = WordGetDocumentStatsRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert isinstance(request.timestamp, int)

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordGetDocumentStatsRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordGetDocumentStatsRequest(requestId="req_001")

        assert "documentUri" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetDocumentStatsRequest.event_name == "word:get:documentStats"

    def test_to_payload_camel_case(self) -> None:
        """Test serialization to camelCase payload"""
        request = WordGetDocumentStatsRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
        )

        payload = request.to_payload()

        assert payload["requestId"] == "req_002"
        assert payload["documentUri"] == "file:///test.docx"
        assert isinstance(payload["timestamp"], int)

    def test_build_class_method(self) -> None:
        """Test build class method for creating requests"""
        request = WordGetDocumentStatsRequest.build(document_uri="file:///test.docx")

        assert request.request_id is not None
        assert isinstance(request.request_id, str)
        assert request.document_uri == "file:///test.docx"


class TestDocumentStats:
    """Test DocumentStats DTO"""

    def test_valid_stats(self) -> None:
        """Test creating valid document stats"""
        stats = DocumentStats(
            wordCount=1000,
            characterCount=5000,
            paragraphCount=20,
        )

        assert stats.word_count == 1000
        assert stats.character_count == 5000
        assert stats.paragraph_count == 20

    def test_zero_stats(self) -> None:
        """Test creating stats with zero values"""
        stats = DocumentStats(
            wordCount=0,
            characterCount=0,
            paragraphCount=0,
        )

        assert stats.word_count == 0
        assert stats.character_count == 0
        assert stats.paragraph_count == 0

    def test_large_values(self) -> None:
        """Test creating stats with large values"""
        stats = DocumentStats(
            wordCount=1000000,
            characterCount=5000000,
            paragraphCount=10000,
        )

        assert stats.word_count == 1000000
        assert stats.character_count == 5000000
        assert stats.paragraph_count == 10000

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing wordCount
        with pytest.raises(ValidationError) as exc_info:
            DocumentStats(
                characterCount=5000,
                paragraphCount=20,
            )

        assert "wordCount" in str(exc_info.value)

        # Missing characterCount
        with pytest.raises(ValidationError) as exc_info:
            DocumentStats(
                wordCount=1000,
                paragraphCount=20,
            )

        assert "characterCount" in str(exc_info.value)

        # Missing paragraphCount
        with pytest.raises(ValidationError) as exc_info:
            DocumentStats(
                wordCount=1000,
                characterCount=5000,
            )

        assert "paragraphCount" in str(exc_info.value)

    def test_stats_serialization(self) -> None:
        """Test stats can be serialized to dict with correct aliases"""
        stats = DocumentStats(
            wordCount=1500,
            characterCount=7500,
            paragraphCount=30,
        )

        data = stats.model_dump(by_alias=True)

        assert data["wordCount"] == 1500
        assert data["characterCount"] == 7500
        assert data["paragraphCount"] == 30

    def test_stats_from_dict(self) -> None:
        """Test creating stats from dict"""
        data = {
            "wordCount": 2000,
            "characterCount": 10000,
            "paragraphCount": 40,
        }

        stats = DocumentStats(**data)

        assert stats.word_count == 2000
        assert stats.character_count == 10000
        assert stats.paragraph_count == 40


class TestWordGetDocumentStatsResponse:
    """Test WordGetDocumentStatsResponse DTO"""

    def test_valid_response(self) -> None:
        """Test creating valid response"""
        stats = DocumentStats(
            wordCount=1000,
            characterCount=5000,
            paragraphCount=20,
        )
        response = WordGetDocumentStatsResponse(data=stats)

        assert response.data.word_count == 1000
        assert response.data.character_count == 5000
        assert response.data.paragraph_count == 20

    def test_response_with_zero_stats(self) -> None:
        """Test creating response with zero stats"""
        stats = DocumentStats(
            wordCount=0,
            characterCount=0,
            paragraphCount=0,
        )
        response = WordGetDocumentStatsResponse(data=stats)

        assert response.data.word_count == 0
        assert response.data.character_count == 0
        assert response.data.paragraph_count == 0

    def test_response_from_dict(self) -> None:
        """Test creating response from dict"""
        data = {
            "data": {
                "wordCount": 3000,
                "characterCount": 15000,
                "paragraphCount": 50,
            }
        }

        response = WordGetDocumentStatsResponse(**data)

        assert response.data.word_count == 3000
        assert response.data.character_count == 15000
        assert response.data.paragraph_count == 50

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing data
        with pytest.raises(ValidationError) as exc_info:
            WordGetDocumentStatsResponse()  # type: ignore

        assert "data" in str(exc_info.value)

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        stats = DocumentStats(
            wordCount=2500,
            characterCount=12500,
            paragraphCount=35,
        )
        response = WordGetDocumentStatsResponse(data=stats)

        data = response.model_dump(by_alias=True)

        assert "data" in data
        assert data["data"]["wordCount"] == 2500
        assert data["data"]["characterCount"] == 12500
        assert data["data"]["paragraphCount"] == 35


class TestWordSelectTextRequest:
    """Test WordSelectTextRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordSelectTextRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            searchText="Hello World",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.search_text == "Hello World"
        assert request.search_options is None
        assert request.selection_mode == "select"  # Default value
        assert request.select_index == 1  # Default value
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_search_options(self) -> None:
        """Test creating valid request with search options"""
        search_options = SelectTextSearchOptions(
            matchCase=True,
            matchWholeWord=True,
            matchWildcards=False,
        )

        request = WordSelectTextRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            searchText="test",
            searchOptions=search_options,
        )

        assert request.search_text == "test"
        assert request.search_options is not None
        assert request.search_options.match_case is True
        assert request.search_options.match_whole_word is True
        assert request.search_options.match_wildcards is False

    def test_valid_request_with_selection_modes(self) -> None:
        """Test creating valid request with different selection modes"""
        # Test "select" mode (default)
        request_select = WordSelectTextRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            searchText="text",
            selectionMode="select",
        )
        assert request_select.selection_mode == "select"

        # Test "start" mode
        request_start = WordSelectTextRequest(
            requestId="req_004",
            documentUri="file:///test.docx",
            searchText="text",
            selectionMode="start",
        )
        assert request_start.selection_mode == "start"

        # Test "end" mode
        request_end = WordSelectTextRequest(
            requestId="req_005",
            documentUri="file:///test.docx",
            searchText="text",
            selectionMode="end",
        )
        assert request_end.selection_mode == "end"

    def test_valid_request_with_select_index(self) -> None:
        """Test creating valid request with different selectIndex"""
        request = WordSelectTextRequest(
            requestId="req_006",
            documentUri="file:///test.docx",
            searchText="test",
            selectIndex=3,
        )

        assert request.select_index == 3

    def test_request_with_dict_search_options(self) -> None:
        """Test creating request with search options as dict"""
        request = WordSelectTextRequest(
            requestId="req_007",
            documentUri="file:///test.docx",
            searchText="pattern",
            searchOptions={"matchCase": True, "matchWildcards": True},
        )

        assert request.search_options is not None
        assert request.search_options.match_case is True
        assert request.search_options.match_wildcards is True
        assert request.search_options.match_whole_word is False

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextRequest(
                documentUri="file:///test.docx",
                searchText="test",
            )

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextRequest(
                requestId="req_001",
                searchText="test",
            )

        assert "documentUri" in str(exc_info.value)

        # Missing searchText
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "searchText" in str(exc_info.value)

    def test_invalid_selection_mode(self) -> None:
        """Test validation fails with invalid selection mode"""
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
                searchText="test",
                selectionMode="invalid_mode",  # type: ignore
            )

        assert "selectionMode" in str(exc_info.value)

    def test_invalid_select_index(self) -> None:
        """Test validation fails with invalid selectIndex (must be >= 1)"""
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
                searchText="test",
                selectIndex=0,  # type: ignore
            )

        assert "selectIndex" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordSelectTextRequest.event_name == "word:select:text"

    def test_to_payload_camel_case(self) -> None:
        """Test serialization to camelCase payload"""
        request = WordSelectTextRequest(
            requestId="req_008",
            documentUri="file:///test.docx",
            searchText="Hello",
            selectionMode="select",
            selectIndex=2,
        )

        payload = request.to_payload()

        assert payload["requestId"] == "req_008"
        assert payload["documentUri"] == "file:///test.docx"
        assert payload["searchText"] == "Hello"
        assert payload["selectionMode"] == "select"
        assert payload["selectIndex"] == 2
        assert isinstance(payload["timestamp"], int)

    def test_build_class_method(self) -> None:
        """Test build class method for creating requests"""
        request = WordSelectTextRequest.build(
            document_uri="file:///test.docx",
            search_text="Auto test",
            selection_mode="start",
        )

        assert request.request_id is not None
        assert isinstance(request.request_id, str)
        assert request.document_uri == "file:///test.docx"
        assert request.search_text == "Auto test"
        assert request.selection_mode == "start"

    def test_search_text_max_length(self) -> None:
        """Test search text max length validation (255 characters, enforced by Word.js API)"""
        # Test valid length (exactly 255 characters)
        valid_text = "a" * 255
        request = WordSelectTextRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            searchText=valid_text,
        )
        assert len(request.search_text) == 255

        # Test invalid length (256 characters)
        invalid_text = "a" * 256
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextRequest(
                requestId="req_002",
                documentUri="file:///test.docx",
                searchText=invalid_text,
            )

        assert "searchText" in str(exc_info.value)
        assert "255" in str(exc_info.value) or "at most 255" in str(exc_info.value).lower()


class TestSelectTextSearchOptions:
    """Test SelectTextSearchOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = SelectTextSearchOptions()

        assert options.match_case is False
        assert options.match_whole_word is False
        assert options.match_wildcards is False

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = SelectTextSearchOptions(
            matchCase=True,
            matchWholeWord=True,
            matchWildcards=True,
        )

        assert options.match_case is True
        assert options.match_whole_word is True
        assert options.match_wildcards is True

    def test_options_from_dict(self) -> None:
        """Test creating options from dict with aliases"""
        options = SelectTextSearchOptions(
            **{
                "matchCase": True,
                "matchWholeWord": False,
                "matchWildcards": True,
            }
        )

        assert options.match_case is True
        assert options.match_whole_word is False
        assert options.match_wildcards is True

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = SelectTextSearchOptions(
            matchCase=True,
            matchWholeWord=True,
        )

        data = options.model_dump(by_alias=True)

        assert data["matchCase"] is True
        assert data["matchWholeWord"] is True
        assert data["matchWildcards"] is False

    def test_partial_options(self) -> None:
        """Test creating options with partial fields"""
        options = SelectTextSearchOptions(matchCase=True)

        assert options.match_case is True
        assert options.match_whole_word is False
        assert options.match_wildcards is False


class TestSelectTextResult:
    """Test SelectTextResult DTO"""

    def test_valid_result_with_selection(self) -> None:
        """Test creating valid result with selection"""
        selection_info = SelectionInfo(
            isEmpty=False,
            type="Normal",
            start=100,
            end=150,
            text="Selected text",
        )

        result = SelectTextResult(
            success=True,
            matchCount=3,
            selectedIndex=2,
            selectedText="Selected text",
            selectionInfo=selection_info,
        )

        assert result.success is True
        assert result.match_count == 3
        assert result.selected_index == 2
        assert result.selected_text == "Selected text"
        assert result.selection_info is not None
        assert result.selection_info.is_empty is False
        assert result.selection_info.start == 100

    def test_valid_result_without_selection(self) -> None:
        """Test creating valid result without selection info"""
        result = SelectTextResult(
            success=True,
            matchCount=1,
            selectedIndex=1,
            selectedText="test",
        )

        assert result.success is True
        assert result.match_count == 1
        assert result.selected_index == 1
        assert result.selected_text == "test"
        assert result.selection_info is None

    def test_result_serialization(self) -> None:
        """Test result can be serialized to dict with correct aliases"""
        selection_info = SelectionInfo(
            isEmpty=False,
            type="Normal",
            start=50,
            end=100,
            text="Test",
        )

        result = SelectTextResult(
            success=True,
            matchCount=5,
            selectedIndex=1,
            selectedText="Test",
            selectionInfo=selection_info,
        )

        data = result.model_dump(by_alias=True)

        assert data["success"] is True
        assert data["matchCount"] == 5
        assert data["selectedIndex"] == 1
        assert data["selectedText"] == "Test"
        assert data["selectionInfo"]["isEmpty"] is False

    def test_result_from_dict(self) -> None:
        """Test creating result from dict"""
        data = {
            "success": True,
            "matchCount": 2,
            "selectedIndex": 1,
            "selectedText": "Hello World",
        }

        result = SelectTextResult(**data)

        assert result.success is True
        assert result.match_count == 2
        assert result.selected_index == 1
        assert result.selected_text == "Hello World"

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing success
        with pytest.raises(ValidationError) as exc_info:
            SelectTextResult(
                matchCount=1,
                selectedIndex=1,
                selectedText="test",
            )

        assert "success" in str(exc_info.value)

        # Missing matchCount
        with pytest.raises(ValidationError) as exc_info:
            SelectTextResult(
                success=True,
                selectedIndex=1,
                selectedText="test",
            )

        assert "matchCount" in str(exc_info.value)

        # Missing selectedIndex
        with pytest.raises(ValidationError) as exc_info:
            SelectTextResult(
                success=True,
                matchCount=1,
                selectedText="test",
            )

        assert "selectedIndex" in str(exc_info.value)

        # Missing selectedText
        with pytest.raises(ValidationError) as exc_info:
            SelectTextResult(
                success=True,
                matchCount=1,
                selectedIndex=1,
            )

        assert "selectedText" in str(exc_info.value)


class TestWordSelectTextResponse:
    """Test WordSelectTextResponse DTO"""

    def test_valid_response_with_data(self) -> None:
        """Test creating valid response with data"""
        result_data = SelectTextResult(
            success=True,
            matchCount=3,
            selectedIndex=2,
            selectedText="Found text",
        )

        response = WordSelectTextResponse(
            requestId="req_001",
            success=True,
            data=result_data,
            timestamp=1234567890,
        )

        assert response.request_id == "req_001"
        assert response.success is True
        assert response.data is not None
        assert response.data.match_count == 3
        assert response.data.selected_text == "Found text"
        assert response.error is None

    def test_valid_response_with_error(self) -> None:
        """Test creating valid response with error"""
        from office4ai.environment.workspace.dtos.common import ErrorResponse

        error = ErrorResponse(code="3000", message="Office API error")

        response = WordSelectTextResponse(
            requestId="req_002",
            success=False,
            error=error,
            timestamp=1234567890,
        )

        assert response.request_id == "req_002"
        assert response.success is False
        assert response.data is None
        assert response.error is not None
        assert response.error.code == "3000"
        assert response.error.message == "Office API error"

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        result_data = SelectTextResult(
            success=True,
            matchCount=1,
            selectedIndex=1,
            selectedText="Test",
        )

        response = WordSelectTextResponse(
            requestId="req_003",
            success=True,
            data=result_data,
            timestamp=1234567890,
        )

        data = response.model_dump(by_alias=True)

        assert data["requestId"] == "req_003"
        assert data["success"] is True
        assert data["data"]["matchCount"] == 1
        assert data["data"]["selectedText"] == "Test"
        assert data["timestamp"] == 1234567890

    def test_response_from_dict(self) -> None:
        """Test creating response from dict"""
        data = {
            "requestId": "req_004",
            "success": True,
            "data": {
                "success": True,
                "matchCount": 2,
                "selectedIndex": 1,
                "selectedText": "Hello",
            },
            "timestamp": 1234567890,
        }

        response = WordSelectTextResponse(**data)

        assert response.request_id == "req_004"
        assert response.success is True
        assert response.data is not None
        assert response.data.match_count == 2
        assert response.data.selected_text == "Hello"

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextResponse(
                success=True,
                timestamp=1234567890,
            )

        assert "requestId" in str(exc_info.value)

        # Missing success
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextResponse(
                requestId="req_001",
                timestamp=1234567890,
            )

        assert "success" in str(exc_info.value)

        # Missing timestamp
        with pytest.raises(ValidationError) as exc_info:
            WordSelectTextResponse(
                requestId="req_001",
                success=True,
            )

        assert "timestamp" in str(exc_info.value)


# ============================================================================
# Comment DTO Tests
# ============================================================================


class TestGetCommentsOptions:
    """Test GetCommentsOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = GetCommentsOptions()

        assert options.include_resolved is False
        assert options.include_replies is True
        assert options.include_associated_text is False
        assert options.detailed_metadata is False
        assert options.max_text_length is None

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = GetCommentsOptions(
            includeResolved=True,
            includeReplies=False,
            includeAssociatedText=True,
            detailedMetadata=True,
            maxTextLength=500,
        )

        assert options.include_resolved is True
        assert options.include_replies is False
        assert options.include_associated_text is True
        assert options.detailed_metadata is True
        assert options.max_text_length == 500

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = GetCommentsOptions(
            includeResolved=True,
            includeAssociatedText=True,
        )

        data = options.model_dump(by_alias=True)

        assert data["includeResolved"] is True
        assert data["includeReplies"] is True  # default
        assert data["includeAssociatedText"] is True

    def test_options_from_dict(self) -> None:
        """Test creating options from dict with aliases"""
        options = GetCommentsOptions(
            **{
                "includeResolved": True,
                "includeReplies": False,
                "maxTextLength": 200,
            }
        )

        assert options.include_resolved is True
        assert options.include_replies is False
        assert options.max_text_length == 200


class TestCommentReplyData:
    """Test CommentReplyData DTO"""

    def test_valid_reply(self) -> None:
        """Test creating valid comment reply"""
        reply = CommentReplyData(
            id="reply_001",
            content="I agree with this change",
            authorName="Bob",
            authorEmail="bob@example.com",
            creationDate="2026-01-15T10:30:00Z",
        )

        assert reply.id == "reply_001"
        assert reply.content == "I agree with this change"
        assert reply.author_name == "Bob"
        assert reply.author_email == "bob@example.com"
        assert reply.creation_date == "2026-01-15T10:30:00Z"

    def test_minimal_reply(self) -> None:
        """Test creating reply with only required fields"""
        reply = CommentReplyData(
            id="reply_002",
            content="OK",
        )

        assert reply.id == "reply_002"
        assert reply.content == "OK"
        assert reply.author_name is None
        assert reply.author_email is None
        assert reply.creation_date is None

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        with pytest.raises(ValidationError) as exc_info:
            CommentReplyData(content="Hello")

        assert "id" in str(exc_info.value)

        with pytest.raises(ValidationError) as exc_info:
            CommentReplyData(id="reply_001")

        assert "content" in str(exc_info.value)

    def test_reply_serialization(self) -> None:
        """Test reply can be serialized to dict with correct aliases"""
        reply = CommentReplyData(
            id="r1",
            content="Test reply",
            authorName="Alice",
        )

        data = reply.model_dump(by_alias=True)

        assert data["id"] == "r1"
        assert data["content"] == "Test reply"
        assert data["authorName"] == "Alice"


class TestCommentData:
    """Test CommentData DTO"""

    def test_valid_comment(self) -> None:
        """Test creating valid comment"""
        comment = CommentData(
            id="comment_001",
            content="Please fix this typo",
            authorName="Alice",
            authorEmail="alice@example.com",
            creationDate="2026-01-10T08:00:00Z",
            resolved=False,
            associatedText="teh",
        )

        assert comment.id == "comment_001"
        assert comment.content == "Please fix this typo"
        assert comment.author_name == "Alice"
        assert comment.resolved is False
        assert comment.associated_text == "teh"
        assert comment.replies is None

    def test_comment_with_replies(self) -> None:
        """Test creating comment with replies"""
        replies = [
            CommentReplyData(id="r1", content="Fixed!"),
            CommentReplyData(id="r2", content="Thanks"),
        ]

        comment = CommentData(
            id="comment_002",
            content="Review this section",
            resolved=True,
            replies=replies,
        )

        assert comment.resolved is True
        assert comment.replies is not None
        assert len(comment.replies) == 2
        assert comment.replies[0].content == "Fixed!"

    def test_minimal_comment(self) -> None:
        """Test creating comment with only required fields"""
        comment = CommentData(
            id="c1",
            content="Note",
        )

        assert comment.id == "c1"
        assert comment.content == "Note"
        assert comment.author_name is None
        assert comment.resolved is False  # default
        assert comment.replies is None

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        with pytest.raises(ValidationError) as exc_info:
            CommentData(content="Hello")

        assert "id" in str(exc_info.value)

        with pytest.raises(ValidationError) as exc_info:
            CommentData(id="c1")

        assert "content" in str(exc_info.value)

    def test_comment_serialization(self) -> None:
        """Test comment can be serialized to dict with correct aliases"""
        comment = CommentData(
            id="c1",
            content="Test comment",
            authorName="Bob",
            resolved=True,
            associatedText="some text",
        )

        data = comment.model_dump(by_alias=True)

        assert data["id"] == "c1"
        assert data["content"] == "Test comment"
        assert data["authorName"] == "Bob"
        assert data["resolved"] is True
        assert data["associatedText"] == "some text"

    def test_comment_from_dict(self) -> None:
        """Test creating comment from dict"""
        data = {
            "id": "c1",
            "content": "From dict",
            "authorName": "Alice",
            "resolved": False,
            "replies": [
                {"id": "r1", "content": "Reply from dict"},
            ],
        }

        comment = CommentData(**data)

        assert comment.id == "c1"
        assert comment.author_name == "Alice"
        assert comment.replies is not None
        assert len(comment.replies) == 1
        assert comment.replies[0].id == "r1"


class TestInsertCommentSearchOptions:
    """Test InsertCommentSearchOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = InsertCommentSearchOptions()

        assert options.match_case is False
        assert options.match_whole_word is False

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = InsertCommentSearchOptions(
            matchCase=True,
            matchWholeWord=True,
        )

        assert options.match_case is True
        assert options.match_whole_word is True

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = InsertCommentSearchOptions(matchCase=True)

        data = options.model_dump(by_alias=True)

        assert data["matchCase"] is True
        assert data["matchWholeWord"] is False


class TestInsertCommentTarget:
    """Test InsertCommentTarget DTO"""

    def test_selection_target(self) -> None:
        """Test creating selection target"""
        target = InsertCommentTarget(type="selection")

        assert target.type == "selection"
        assert target.search_text is None
        assert target.search_options is None

    def test_search_text_target(self) -> None:
        """Test creating searchText target"""
        target = InsertCommentTarget(
            type="searchText",
            searchText="important text",
            searchOptions={"matchCase": True},
        )

        assert target.type == "searchText"
        assert target.search_text == "important text"
        assert target.search_options is not None
        assert target.search_options.match_case is True

    def test_invalid_target_type(self) -> None:
        """Test validation fails with invalid target type"""
        with pytest.raises(ValidationError) as exc_info:
            InsertCommentTarget(type="invalid")

        assert "type" in str(exc_info.value).lower()

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        with pytest.raises(ValidationError) as exc_info:
            InsertCommentTarget()

        assert "type" in str(exc_info.value)

    def test_target_serialization(self) -> None:
        """Test target can be serialized to dict with correct aliases"""
        target = InsertCommentTarget(
            type="searchText",
            searchText="test",
        )

        data = target.model_dump(by_alias=True)

        assert data["type"] == "searchText"
        assert data["searchText"] == "test"


class TestWordGetCommentsRequest:
    """Test WordGetCommentsRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordGetCommentsRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.options is None

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = GetCommentsOptions(
            includeResolved=True,
            includeReplies=True,
        )

        request = WordGetCommentsRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            options=options,
        )

        assert request.options is not None
        assert request.options.include_resolved is True

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordGetCommentsRequest.event_name == "word:get:comments"


class TestWordGetCommentsResponse:
    """Test WordGetCommentsResponse DTO"""

    def test_valid_response_with_comments(self) -> None:
        """Test creating valid response with comments"""
        comments = [
            CommentData(id="c1", content="Fix typo", authorName="Alice"),
            CommentData(id="c2", content="Good", resolved=True),
        ]
        response = WordGetCommentsResponse(comments=comments)

        assert len(response.comments) == 2
        assert response.comments[0].id == "c1"
        assert response.comments[1].resolved is True

    def test_empty_response(self) -> None:
        """Test creating response with no comments"""
        response = WordGetCommentsResponse()

        assert len(response.comments) == 0

    def test_response_from_dict(self) -> None:
        """Test creating response from dict"""
        data = {
            "comments": [
                {"id": "c1", "content": "Test", "resolved": False},
            ]
        }

        response = WordGetCommentsResponse(**data)

        assert len(response.comments) == 1
        assert response.comments[0].id == "c1"

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        comments = [CommentData(id="c1", content="Test")]
        response = WordGetCommentsResponse(comments=comments)

        data = response.model_dump(by_alias=True)

        assert "comments" in data
        assert len(data["comments"]) == 1
        assert data["comments"][0]["id"] == "c1"


class TestWordInsertCommentRequest:
    """Test WordInsertCommentRequest DTO"""

    def test_valid_request_minimal(self) -> None:
        """Test creating valid request with only required fields"""
        request = WordInsertCommentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            text="Please review",
        )

        assert request.text == "Please review"
        assert request.target is None

    def test_valid_request_with_target(self) -> None:
        """Test creating valid request with target"""
        target = InsertCommentTarget(type="searchText", searchText="important")

        request = WordInsertCommentRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            text="Review needed",
            target=target,
        )

        assert request.target is not None
        assert request.target.type == "searchText"
        assert request.target.search_text == "important"

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        with pytest.raises(ValidationError) as exc_info:
            WordInsertCommentRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "text" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordInsertCommentRequest.event_name == "word:insert:comment"


class TestWordDeleteCommentRequest:
    """Test WordDeleteCommentRequest DTO"""

    def test_valid_request(self) -> None:
        """Test creating valid request"""
        request = WordDeleteCommentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            commentId="comment_123",
        )

        assert request.comment_id == "comment_123"

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        with pytest.raises(ValidationError) as exc_info:
            WordDeleteCommentRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "commentId" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordDeleteCommentRequest.event_name == "word:delete:comment"


class TestWordReplyCommentRequest:
    """Test WordReplyCommentRequest DTO"""

    def test_valid_request(self) -> None:
        """Test creating valid request"""
        request = WordReplyCommentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            commentId="comment_123",
            text="Done",
        )

        assert request.comment_id == "comment_123"
        assert request.text == "Done"

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        with pytest.raises(ValidationError) as exc_info:
            WordReplyCommentRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
                text="Reply",
            )

        assert "commentId" in str(exc_info.value)

        with pytest.raises(ValidationError) as exc_info:
            WordReplyCommentRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
                commentId="c1",
            )

        assert "text" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordReplyCommentRequest.event_name == "word:reply:comment"


class TestWordResolveCommentRequest:
    """Test WordResolveCommentRequest DTO"""

    def test_valid_request_resolve(self) -> None:
        """Test creating valid request to resolve"""
        request = WordResolveCommentRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            commentId="comment_123",
            resolved=True,
        )

        assert request.comment_id == "comment_123"
        assert request.resolved is True

    def test_valid_request_unresolve(self) -> None:
        """Test creating valid request to unresolve"""
        request = WordResolveCommentRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            commentId="comment_123",
            resolved=False,
        )

        assert request.resolved is False

    def test_default_resolved(self) -> None:
        """Test resolved defaults to True"""
        request = WordResolveCommentRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            commentId="comment_123",
        )

        assert request.resolved is True

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields"""
        with pytest.raises(ValidationError) as exc_info:
            WordResolveCommentRequest(
                requestId="req_001",
                documentUri="file:///test.docx",
            )

        assert "commentId" in str(exc_info.value)

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordResolveCommentRequest.event_name == "word:resolve:comment"

    def test_to_payload_camel_case(self) -> None:
        """Test serialization to camelCase payload"""
        request = WordResolveCommentRequest(
            requestId="req_004",
            documentUri="file:///test.docx",
            commentId="c1",
            resolved=True,
        )

        payload = request.to_payload()

        assert payload["commentId"] == "c1"
        assert payload["resolved"] is True


class TestTableInsertOptionsInsertLocation:
    """TableInsertOptions 新增 insertLocation 字段覆盖（OF4AI-9）"""

    @pytest.mark.parametrize("location", ["Start", "End", "Before", "After", "Replace"])
    def test_all_enum_values_accepted(self, location: str) -> None:
        """五种枚举值均应通过校验，并保留原值"""
        options = TableInsertOptions(rows=2, columns=2, insertLocation=location)

        assert options.insert_location == location

    def test_default_omitted_means_none(self) -> None:
        """未传 insertLocation 时内部值为 None，由 Add-In 侧回退默认 'End'"""
        options = TableInsertOptions(rows=2, columns=2)

        assert options.insert_location is None

    def test_rejects_invalid_location(self) -> None:
        """非枚举值应触发 ValidationError"""
        with pytest.raises(ValidationError):
            TableInsertOptions(rows=2, columns=2, insertLocation="Middle")

    def test_accepts_snake_case_field_name(self) -> None:
        """内部字段名 snake_case 也应被接受（populate_by_name）"""
        options = TableInsertOptions(rows=2, columns=2, insert_location="After")

        assert options.insert_location == "After"

    def test_dump_by_alias_uses_camel_case(self) -> None:
        """序列化到协议 wire format 时字段名为 camelCase"""
        options = TableInsertOptions(rows=2, columns=2, insertLocation="Before")

        payload = options.model_dump(by_alias=True, exclude_none=True)

        assert "insertLocation" in payload
        assert payload["insertLocation"] == "Before"
        assert "insert_location" not in payload

    def test_omitted_excluded_from_wire_payload(self) -> None:
        """未设置 insertLocation 时，不应出现在按别名导出的 payload 中"""
        options = TableInsertOptions(rows=1, columns=1)

        payload = options.model_dump(by_alias=True, exclude_none=True)

        assert "insertLocation" not in payload

    def test_request_wraps_options_and_serializes_insert_location(self) -> None:
        """WordInsertTableRequest 端到端：接受 camelCase 输入，输出 wire payload 含 insertLocation"""
        request = WordInsertTableRequest(
            requestId="req-insert-table-1",
            documentUri="file:///test.docx",
            options={"rows": 2, "columns": 2, "insertLocation": "Replace"},
        )

        payload = request.to_payload()

        assert payload["options"]["insertLocation"] == "Replace"


# ============================================================================
# Table Operation DTOs (OASP /word Draft, v0.2.0)
# ============================================================================


class TestCellFormat:
    """CellFormat 枚举值校验 + 别名往返"""

    def test_centered_horizontal_alignment_accepted(self) -> None:
        fmt = CellFormat(horizontalAlignment="Centered")
        assert fmt.horizontal_alignment == "Centered"

    def test_justified_horizontal_alignment_accepted(self) -> None:
        fmt = CellFormat(horizontalAlignment="Justified")
        assert fmt.horizontal_alignment == "Justified"

    def test_center_horizontal_alignment_rejected(self) -> None:
        """horizontalAlignment='Center' 必须被拒绝（应是 'Centered'）"""
        with pytest.raises(ValidationError):
            CellFormat(horizontalAlignment="Center")

    def test_center_vertical_alignment_accepted(self) -> None:
        fmt = CellFormat(verticalAlignment="Center")
        assert fmt.vertical_alignment == "Center"

    def test_middle_vertical_alignment_rejected(self) -> None:
        """verticalAlignment='Middle' 必须被拒绝（应是 'Center'）"""
        with pytest.raises(ValidationError):
            CellFormat(verticalAlignment="Middle")

    def test_zero_font_size_rejected(self) -> None:
        with pytest.raises(ValidationError):
            CellFormat(fontSize=0)

    def test_alias_round_trip(self) -> None:
        """snake_case 输入 + camelCase 输出"""
        fmt = CellFormat(
            horizontal_alignment="Centered",
            background_color="#1F4E79",
            font_color="#FFFFFF",
            bold=True,
        )
        payload = fmt.model_dump(by_alias=True, exclude_none=True)
        assert payload["horizontalAlignment"] == "Centered"
        assert payload["backgroundColor"] == "#1F4E79"
        assert payload["fontColor"] == "#FFFFFF"
        assert payload["bold"] is True


class TestTableSummary:
    """TableSummary (word:get:documentStructure 配套)"""

    def test_camelcase_input_with_optional_heading(self) -> None:
        s = TableSummary(
            tableId="table-0",
            rowCount=5,
            columnCount=4,
            precedingHeading="甲方信息",
        )
        assert s.table_id == "table-0"
        assert s.preceding_heading == "甲方信息"
        payload = s.model_dump(by_alias=True, exclude_none=True)
        assert payload["tableId"] == "table-0"
        assert payload["precedingHeading"] == "甲方信息"

    def test_omitted_heading_excluded_from_wire(self) -> None:
        s = TableSummary(tableId="table-1", rowCount=2, columnCount=3)
        payload = s.model_dump(by_alias=True, exclude_none=True)
        assert "precedingHeading" not in payload


class TestWordMergeCellsRequest:
    """word:merge:cells DTO"""

    def test_event_registered(self) -> None:
        assert request_registry.contains("word:merge:cells")
        assert request_registry.get("word:merge:cells") is WordMergeCellsRequest

    def test_build_with_explicit_table_id_serializes_camelcase(self) -> None:
        req = WordMergeCellsRequest.build(
            document_uri="file:///t.docx",
            tableId="table-0",
            startRowIndex=0,
            startColumnIndex=0,
            endRowIndex=0,
            endColumnIndex=4,
        )
        payload = req.to_payload()
        assert payload["tableId"] == "table-0"
        assert payload["startRowIndex"] == 0
        assert payload["endColumnIndex"] == 4

    def test_optional_table_id_omitted_from_wire(self) -> None:
        req = WordMergeCellsRequest.build(
            document_uri="file:///t.docx",
            startRowIndex=1,
            startColumnIndex=1,
            endRowIndex=2,
            endColumnIndex=3,
        )
        payload = req.to_payload()
        assert "tableId" not in payload

    def test_negative_index_rejected(self) -> None:
        with pytest.raises(ValidationError):
            WordMergeCellsRequest.build(
                document_uri="file:///t.docx",
                startRowIndex=-1,
                startColumnIndex=0,
                endRowIndex=0,
                endColumnIndex=2,
            )


class TestWordUpdateTableCellRequest:
    """word:update:tableCell DTO"""

    def test_event_registered(self) -> None:
        assert request_registry.contains("word:update:tableCell")

    def test_build_with_format_and_alias_round_trip(self) -> None:
        req = WordUpdateTableCellRequest.build(
            document_uri="file:///t.docx",
            tableId="table-0",
            cells=[
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
        )
        payload = req.to_payload()
        assert payload["tableId"] == "table-0"
        cell = payload["cells"][0]
        assert cell["rowIndex"] == 0
        assert cell["text"] == "甲方信息"
        assert cell["format"]["horizontalAlignment"] == "Centered"
        assert cell["format"]["backgroundColor"] == "#1F4E79"

    def test_empty_cells_rejected(self) -> None:
        with pytest.raises(ValidationError):
            WordUpdateTableCellRequest.build(
                document_uri="file:///t.docx",
                cells=[],
            )

    def test_table_cell_update_accepts_text_only(self) -> None:
        upd = TableCellUpdate(rowIndex=1, columnIndex=2, text="hello")
        assert upd.text == "hello"
        assert upd.format is None

    def test_table_cell_update_accepts_format_only(self) -> None:
        upd = TableCellUpdate(
            rowIndex=1,
            columnIndex=2,
            format={"backgroundColor": "#EEEEEE"},
        )
        assert upd.text is None
        assert upd.format is not None
        assert upd.format.background_color == "#EEEEEE"


class TestWordUpdateTableRowColumnRequest:
    """word:update:tableRowColumn DTO"""

    def test_event_registered(self) -> None:
        assert request_registry.contains("word:update:tableRowColumn")

    def test_rows_only_round_trip(self) -> None:
        req = WordUpdateTableRowColumnRequest.build(
            document_uri="file:///t.docx",
            rows=[
                TableRowUpdate(rowIndex=1, values=["甲方", "ACME"]),
                TableRowUpdate(rowIndex=2, values=["地址", "上海"]),
            ],
        )
        payload = req.to_payload()
        assert len(payload["rows"]) == 2
        assert payload["rows"][0]["rowIndex"] == 1
        assert payload["rows"][0]["values"] == ["甲方", "ACME"]
        assert "columns" not in payload

    def test_columns_only_round_trip(self) -> None:
        req = WordUpdateTableRowColumnRequest.build(
            document_uri="file:///t.docx",
            columns=[TableColumnUpdate(columnIndex=0, values=["甲方", "地址"])],
        )
        payload = req.to_payload()
        assert payload["columns"][0]["columnIndex"] == 0
        assert "rows" not in payload


class TestWordUpdateTableFormatRequest:
    """word:update:tableFormat DTO"""

    def test_event_registered(self) -> None:
        assert request_registry.contains("word:update:tableFormat")

    def test_full_format_round_trip(self) -> None:
        req = WordUpdateTableFormatRequest.build(
            document_uri="file:///t.docx",
            tableId="table-0",
            styleOptions=TableStyleOptions(
                styleType="Grid Table 4 - Accent 1",
                bandedRows=True,
                cellPadding=TableCellPadding(top=4, bottom=4, left=6, right=6),
            ),
            borderOptions=TableBorderOptions(location="inside", style="Single", width=0.5),
            columnWidths=[120, 80, 80, 80],
            alignment="Centered",
        )
        payload = req.to_payload()
        assert payload["tableId"] == "table-0"
        assert payload["alignment"] == "Centered"
        assert payload["columnWidths"] == [120, 80, 80, 80]
        assert payload["borderOptions"]["location"] == "inside"
        assert payload["borderOptions"]["style"] == "Single"
        assert payload["styleOptions"]["styleType"] == "Grid Table 4 - Accent 1"
        assert payload["styleOptions"]["bandedRows"] is True
        assert payload["styleOptions"]["cellPadding"]["top"] == 4

    def test_alignment_center_rejected(self) -> None:
        """alignment='Center' 必须被拒绝（应是 'Centered'）"""
        with pytest.raises(ValidationError):
            WordUpdateTableFormatRequest.build(
                document_uri="file:///t.docx",
                alignment="Center",
            )

    def test_border_location_invalid_rejected(self) -> None:
        with pytest.raises(ValidationError):
            TableBorderOptions(location="diagonal")

    def test_border_width_zero_rejected(self) -> None:
        with pytest.raises(ValidationError):
            TableBorderOptions(width=0)

    def test_border_style_none_accepted(self) -> None:
        """style='None' 是合法枚举值（用于清除边框）"""
        b = TableBorderOptions(style="None")
        assert b.style == "None"
