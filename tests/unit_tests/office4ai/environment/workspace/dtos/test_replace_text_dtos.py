"""
Test Word Replace Text DTOs

测试 word:replace:text 事件的数据传输对象。
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import ErrorResponse
from office4ai.environment.workspace.dtos.word import (
    ReplaceOptions,
    ReplaceTextResult,
    TextFormat,
    WordFont,
    WordReplaceTextRequest,
    WordReplaceTextResponse,
)


class TestWordReplaceTextRequest:
    """Test WordReplaceTextRequest DTO"""

    def test_valid_request_with_defaults(self) -> None:
        """Test creating valid request with default values"""
        request = WordReplaceTextRequest(
            requestId="req_001",
            documentUri="file:///test.docx",
            searchText="old text",
            replaceText="new text",
        )

        assert request.request_id == "req_001"
        assert request.document_uri == "file:///test.docx"
        assert request.search_text == "old text"
        assert request.replace_text == "new text"
        assert request.options is None
        assert isinstance(request.timestamp, int)

    def test_valid_request_with_options(self) -> None:
        """Test creating valid request with options"""
        options = ReplaceOptions(
            matchCase=True,
            matchWholeWord=True,
            replaceAll=True,
        )

        request = WordReplaceTextRequest(
            requestId="req_002",
            documentUri="file:///test.docx",
            searchText="old",
            replaceText="new",
            options=options,
        )

        assert request.search_text == "old"
        assert request.replace_text == "new"
        assert request.options is not None
        assert request.options.match_case is True
        assert request.options.match_whole_word is True
        assert request.options.replace_all is True

    def test_request_with_dict_options(self) -> None:
        """Test creating request with options as dict"""
        request = WordReplaceTextRequest(
            requestId="req_003",
            documentUri="file:///test.docx",
            searchText="find",
            replaceText="replace",
            options=ReplaceOptions(matchCase=True, matchWholeWord=False),
        )

        assert request.options is not None
        assert request.options.match_case is True
        assert request.options.match_whole_word is False
        assert request.options.replace_all is False  # Default value

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields

        Uses model_validate to test runtime validation without
        triggering static type checker errors.
        """
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceTextRequest.model_validate(
                {
                    "documentUri": "file:///test.docx",
                    "searchText": "old",
                    "replaceText": "new",
                }
            )

        assert "requestId" in str(exc_info.value)

        # Missing documentUri
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceTextRequest.model_validate(
                {
                    "requestId": "req_001",
                    "searchText": "old",
                    "replaceText": "new",
                }
            )

        assert "documentUri" in str(exc_info.value)

        # Missing searchText
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceTextRequest.model_validate(
                {
                    "requestId": "req_001",
                    "documentUri": "file:///test.docx",
                    "replaceText": "new",
                }
            )

        assert "searchText" in str(exc_info.value)

        # Missing replaceText
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceTextRequest.model_validate(
                {
                    "requestId": "req_001",
                    "documentUri": "file:///test.docx",
                    "searchText": "old",
                }
            )

        assert "replaceText" in str(exc_info.value)

    def test_empty_search_and_replace_text(self) -> None:
        """Test that empty strings are valid for searchText and replaceText"""
        # Empty strings should be valid at DTO level (validation happens at handler level)
        request = WordReplaceTextRequest(
            requestId="req_004",
            documentUri="file:///test.docx",
            searchText="",
            replaceText="",
        )

        assert request.search_text == ""
        assert request.replace_text == ""

    def test_event_name_attribute(self) -> None:
        """Test event name class variable"""
        assert WordReplaceTextRequest.event_name == "word:replace:text"

    def test_to_payload_camel_case(self) -> None:
        """Test serialization to camelCase payload"""
        request = WordReplaceTextRequest(
            requestId="req_005",
            documentUri="file:///test.docx",
            searchText="old text",
            replaceText="new text",
            options=ReplaceOptions(matchCase=False, replaceAll=True),
        )

        payload = request.to_payload()

        assert payload["requestId"] == "req_005"
        assert payload["documentUri"] == "file:///test.docx"
        assert payload["searchText"] == "old text"
        assert payload["replaceText"] == "new text"
        assert payload["options"]["matchCase"] is False
        assert payload["options"]["replaceAll"] is True
        assert isinstance(payload["timestamp"], int)

    def test_build_class_method(self) -> None:
        """Test build class method for creating requests"""
        request = WordReplaceTextRequest.build(
            document_uri="file:///test.docx",
            search_text="find me",
            replace_text="replaced",
            options=ReplaceOptions(replaceAll=True),
        )

        assert request.request_id is not None
        assert isinstance(request.request_id, str)
        assert request.document_uri == "file:///test.docx"
        assert request.search_text == "find me"
        assert request.replace_text == "replaced"

    def test_valid_request_with_format(self) -> None:
        """Test creating valid request with nested-font text formatting"""
        fmt = TextFormat(font=WordFont(bold=True, italic=False, size=14, color="#FF0000"))

        request = WordReplaceTextRequest(
            requestId="req_010",
            documentUri="file:///test.docx",
            searchText="hello",
            replaceText="hello",
            format=fmt,
        )

        assert request.format is not None
        assert request.format.font is not None
        assert request.format.font.bold is True
        assert request.format.font.italic is False
        assert request.format.font.size == 14
        assert request.format.font.color == "#FF0000"

    def test_request_with_format_serialization(self) -> None:
        """Test nested font + styleName serialize to camelCase payload"""
        fmt = TextFormat(font=WordFont(bold=True), styleName="Heading 1")

        request = WordReplaceTextRequest(
            requestId="req_011",
            documentUri="file:///test.docx",
            searchText="title",
            replaceText="title",
            format=fmt,
        )

        payload = request.to_payload()

        assert payload["format"]["font"]["bold"] is True
        assert payload["format"]["styleName"] == "Heading 1"

    def test_request_without_format(self) -> None:
        """Test format field defaults to None"""
        request = WordReplaceTextRequest(
            requestId="req_012",
            documentUri="file:///test.docx",
            searchText="old",
            replaceText="new",
        )

        assert request.format is None

    def test_unicode_and_special_characters(self) -> None:
        """Test handling of unicode and special characters"""
        request = WordReplaceTextRequest(
            requestId="req_006",
            documentUri="file:///test.docx",
            searchText="Hello 世界",
            replaceText="Bonjour 🌍",
        )

        assert request.search_text == "Hello 世界"
        assert request.replace_text == "Bonjour 🌍"


class TestReplaceOptions:
    """Test ReplaceOptions DTO"""

    def test_default_options(self) -> None:
        """Test creating options with default values"""
        options = ReplaceOptions()

        assert options.match_case is False
        assert options.match_whole_word is False
        assert options.replace_all is False

    def test_custom_options(self) -> None:
        """Test creating options with custom values"""
        options = ReplaceOptions(
            matchCase=True,
            matchWholeWord=True,
            replaceAll=True,
        )

        assert options.match_case is True
        assert options.match_whole_word is True
        assert options.replace_all is True

    def test_partial_options(self) -> None:
        """Test creating options with partial fields"""
        options = ReplaceOptions(matchCase=True)

        assert options.match_case is True
        assert options.match_whole_word is False  # Default
        assert options.replace_all is False  # Default

    def test_options_from_dict(self) -> None:
        """Test creating options from dict with aliases"""
        options = ReplaceOptions(
            **{
                "matchCase": True,
                "matchWholeWord": False,
                "replaceAll": True,
            }
        )

        assert options.match_case is True
        assert options.match_whole_word is False
        assert options.replace_all is True

    def test_options_serialization(self) -> None:
        """Test options can be serialized to dict with correct aliases"""
        options = ReplaceOptions(
            matchCase=True,
            matchWholeWord=True,
            replaceAll=False,
        )

        data = options.model_dump(by_alias=True)

        assert data["matchCase"] is True
        assert data["matchWholeWord"] is True
        assert data["replaceAll"] is False

    def test_all_false_combinations(self) -> None:
        """Test all possible False/True combinations"""
        # All False
        options1 = ReplaceOptions(matchCase=False, matchWholeWord=False, replaceAll=False)
        assert options1.match_case is False
        assert options1.match_whole_word is False
        assert options1.replace_all is False

        # All True
        options2 = ReplaceOptions(matchCase=True, matchWholeWord=True, replaceAll=True)
        assert options2.match_case is True
        assert options2.match_whole_word is True
        assert options2.replace_all is True


class TestReplaceTextResult:
    """Test ReplaceTextResult DTO"""

    def test_valid_result(self) -> None:
        """Test creating valid result"""
        result = ReplaceTextResult(replaceCount=5)

        assert result.replace_count == 5

    def test_zero_replace_count(self) -> None:
        """Test result with zero replacements"""
        result = ReplaceTextResult(replaceCount=0)

        assert result.replace_count == 0

    def test_large_replace_count(self) -> None:
        """Test result with large replacement count"""
        result = ReplaceTextResult(replaceCount=10000)

        assert result.replace_count == 10000

    def test_missing_required_field(self) -> None:
        """Test validation fails without replaceCount

        Uses model_validate to test runtime validation without
        triggering static type checker errors.
        """
        with pytest.raises(ValidationError) as exc_info:
            ReplaceTextResult.model_validate({})

        assert "replaceCount" in str(exc_info.value)

    def test_negative_replace_count(self) -> None:
        """Test that negative count is accepted (validation should happen at handler level)"""
        # Pydantic doesn't have a constraint on replaceCount, so this should work
        result = ReplaceTextResult(replaceCount=-1)
        assert result.replace_count == -1

    def test_result_serialization(self) -> None:
        """Test result can be serialized to dict with correct aliases"""
        result = ReplaceTextResult(replaceCount=42)

        data = result.model_dump(by_alias=True)

        assert data["replaceCount"] == 42

    def test_result_from_dict(self) -> None:
        """Test creating result from dict"""
        data = {"replaceCount": 100}
        result = ReplaceTextResult(**data)

        assert result.replace_count == 100


class TestWordReplaceTextResponse:
    """Test WordReplaceTextResponse DTO"""

    def test_valid_success_response(self) -> None:
        """Test creating valid success response"""
        result_data = ReplaceTextResult(replaceCount=3)
        response = WordReplaceTextResponse(
            requestId="req_001",
            success=True,
            data=result_data,
            timestamp=1234567890,
        )

        assert response.request_id == "req_001"
        assert response.success is True
        assert response.data is not None
        assert response.data.replace_count == 3
        assert response.error is None
        assert response.timestamp == 1234567890

    def test_valid_error_response(self) -> None:
        """Test creating valid error response"""
        error = ErrorResponse(code="4001", message="Missing required parameters")
        response = WordReplaceTextResponse(
            requestId="req_002",
            success=False,
            data=None,
            error=error,
            timestamp=1234567890,
        )

        assert response.request_id == "req_002"
        assert response.success is False
        assert response.data is None
        assert response.error is not None
        assert response.error.code == "4001"
        assert response.timestamp == 1234567890

    def test_response_with_zero_replacements(self) -> None:
        """Test response with zero replacements"""
        result_data = ReplaceTextResult(replaceCount=0)
        response = WordReplaceTextResponse(
            requestId="req_003",
            success=True,
            data=result_data,
            timestamp=1234567890,
        )

        assert response.data is not None
        assert response.data.replace_count == 0

    def test_response_with_many_replacements(self) -> None:
        """Test response with many replacements"""
        result_data = ReplaceTextResult(replaceCount=5000)
        response = WordReplaceTextResponse(
            requestId="req_004",
            success=True,
            data=result_data,
            timestamp=1234567890,
        )

        assert response.data is not None
        assert response.data.replace_count == 5000

    def test_missing_required_fields(self) -> None:
        """Test validation fails without required fields

        Uses model_validate to test runtime validation without
        triggering static type checker errors.
        """
        # Missing requestId
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceTextResponse.model_validate(
                {
                    "success": True,
                    "timestamp": 1234567890,
                }
            )

        assert "requestId" in str(exc_info.value)

        # Missing success
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceTextResponse.model_validate(
                {
                    "requestId": "req_001",
                    "timestamp": 1234567890,
                }
            )

        assert "success" in str(exc_info.value)

        # Missing timestamp
        with pytest.raises(ValidationError) as exc_info:
            WordReplaceTextResponse.model_validate(
                {
                    "requestId": "req_001",
                    "success": True,
                }
            )

        assert "timestamp" in str(exc_info.value)

    def test_response_serialization(self) -> None:
        """Test response can be serialized to dict with correct aliases"""
        result_data = ReplaceTextResult(replaceCount=10)
        response = WordReplaceTextResponse(
            requestId="req_005",
            success=True,
            data=result_data,
            timestamp=1234567890,
        )

        data = response.model_dump(by_alias=True)

        assert data["requestId"] == "req_005"
        assert data["success"] is True
        assert data["data"]["replaceCount"] == 10
        assert data["timestamp"] == 1234567890

    def test_response_from_dict(self) -> None:
        """Test creating response from dict"""
        data = {
            "requestId": "req_006",
            "success": True,
            "data": {"replaceCount": 25},
            "timestamp": 1234567890,
        }

        response = WordReplaceTextResponse(**data)

        assert response.request_id == "req_006"
        assert response.success is True
        assert response.data is not None
        assert response.data.replace_count == 25
        assert response.timestamp == 1234567890
