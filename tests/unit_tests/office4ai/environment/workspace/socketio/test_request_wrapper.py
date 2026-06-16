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
        # Foundation slice (#18) registers the 3 state-awareness read events;
        # the remaining /excel events land with #19–#26.
        assert set(excel_events) >= {
            "excel:get:workbookInfo",
            "excel:get:worksheetInfo",
            "excel:get:selectedRange",
        }

    def test_contains_all_ppt_events(self) -> None:
        events = get_registered_events()
        ppt_events = [e for e in events if e.startswith("ppt:")]
        # Should have at least 10 PPT events
        assert len(ppt_events) >= 10
