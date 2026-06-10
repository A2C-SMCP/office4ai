"""Unit tests for PPT OOXML carrier DTOs (OASP /ppt Draft, v0.3.0).

Covers ppt:get:slideOoxml / ppt:insert:slidesOoxml request DTOs and the
3016 API_NOT_SUPPORTED error code. Focus: camelCase wire alias round-trips
and snake_case input acceptance (populate_by_name).
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import ErrorCode, request_registry
from office4ai.environment.workspace.dtos.ppt import (
    PptGetSlideOoxmlRequest,
    PptInsertSlidesOoxmlRequest,
)


class TestErrorCodeAndRegistry:
    def test_api_not_supported_code(self) -> None:
        assert ErrorCode.API_NOT_SUPPORTED == "3016"

    def test_registry_has_ooxml_events(self) -> None:
        assert request_registry.contains("ppt:get:slideOoxml")
        assert request_registry.contains("ppt:insert:slidesOoxml")

    def test_registry_resolves_to_dto_classes(self) -> None:
        assert request_registry.get("ppt:get:slideOoxml") is PptGetSlideOoxmlRequest
        assert request_registry.get("ppt:insert:slidesOoxml") is PptInsertSlidesOoxmlRequest


class TestPptGetSlideOoxmlRequest:
    def test_round_trip_emits_camel_case(self) -> None:
        req = PptGetSlideOoxmlRequest.build(document_uri="file:///t.pptx", slide_index=2)
        payload = req.to_payload()
        assert payload["documentUri"] == "file:///t.pptx"
        assert payload["slideIndex"] == 2
        assert "slide_index" not in payload

    def test_accepts_snake_case_input(self) -> None:
        # populate_by_name=True: snake_case field name is accepted alongside the alias
        req = PptGetSlideOoxmlRequest.model_validate(
            {"requestId": "r1", "documentUri": "file:///t.pptx", "slide_index": 0}
        )
        assert req.slide_index == 0

    def test_accepts_camel_case_input(self) -> None:
        req = PptGetSlideOoxmlRequest.model_validate(
            {"requestId": "r1", "documentUri": "file:///t.pptx", "slideIndex": 5}
        )
        assert req.slide_index == 5

    def test_slide_index_required(self) -> None:
        with pytest.raises(ValidationError):
            PptGetSlideOoxmlRequest.build(document_uri="file:///t.pptx")

    def test_slide_index_must_be_non_negative(self) -> None:
        with pytest.raises(ValidationError):
            PptGetSlideOoxmlRequest.build(document_uri="file:///t.pptx", slide_index=-1)


class TestPptInsertSlidesOoxmlRequest:
    def test_minimal_request_only_base64(self) -> None:
        req = PptInsertSlidesOoxmlRequest.build(document_uri="file:///t.pptx", base64="UEsDBBQ=")
        payload = req.to_payload()
        assert payload["base64"] == "UEsDBBQ="
        # BaseRequest still supplies its default timestamp even on a minimal request
        assert "timestamp" in payload
        # all optional fields excluded when None
        assert "formatting" not in payload
        assert "targetSlideIndex" not in payload
        assert "replaceSlideId" not in payload
        assert "finalSlideIndex" not in payload

    def test_base64_required(self) -> None:
        with pytest.raises(ValidationError):
            PptInsertSlidesOoxmlRequest.build(document_uri="file:///t.pptx")

    def test_full_round_trip_fields_camel_case(self) -> None:
        req = PptInsertSlidesOoxmlRequest.build(
            document_uri="file:///t.pptx",
            base64="UEsDBBQ=",
            formatting="keepSourceFormatting",
            target_slide_index=2,
            replace_slide_id="slide-003",
            final_slide_index=2,
        )
        payload = req.to_payload()
        assert payload["base64"] == "UEsDBBQ="
        assert payload["formatting"] == "keepSourceFormatting"
        assert payload["targetSlideIndex"] == 2
        assert payload["replaceSlideId"] == "slide-003"
        assert payload["finalSlideIndex"] == 2
        # no snake_case leakage
        assert "target_slide_index" not in payload
        assert "replace_slide_id" not in payload
        assert "final_slide_index" not in payload

    def test_accepts_camel_case_input(self) -> None:
        req = PptInsertSlidesOoxmlRequest.model_validate(
            {
                "requestId": "r1",
                "documentUri": "file:///t.pptx",
                "base64": "UEsDBBQ=",
                "formatting": "useDestinationTheme",
                "targetSlideIndex": 1,
                "replaceSlideId": "slide-001",
                "finalSlideIndex": 1,
            }
        )
        assert req.formatting == "useDestinationTheme"
        assert req.target_slide_index == 1
        assert req.replace_slide_id == "slide-001"
        assert req.final_slide_index == 1

    def test_accepts_snake_case_input(self) -> None:
        req = PptInsertSlidesOoxmlRequest.model_validate(
            {
                "requestId": "r1",
                "documentUri": "file:///t.pptx",
                "base64": "UEsDBBQ=",
                "target_slide_index": 3,
                "replace_slide_id": "slide-009",
                "final_slide_index": 3,
            }
        )
        assert req.target_slide_index == 3
        assert req.replace_slide_id == "slide-009"
        assert req.final_slide_index == 3

    def test_use_destination_theme_round_trips_to_payload(self) -> None:
        req = PptInsertSlidesOoxmlRequest.build(
            document_uri="file:///t.pptx",
            base64="UEsDBBQ=",
            formatting="useDestinationTheme",
        )
        assert req.to_payload()["formatting"] == "useDestinationTheme"

    def test_invalid_formatting_rejected(self) -> None:
        with pytest.raises(ValidationError):
            PptInsertSlidesOoxmlRequest.build(
                document_uri="file:///t.pptx",
                base64="UEsDBBQ=",
                formatting="bogusFormatting",
            )

    def test_negative_indices_rejected(self) -> None:
        with pytest.raises(ValidationError):
            PptInsertSlidesOoxmlRequest.build(document_uri="file:///t.pptx", base64="UEsDBBQ=", target_slide_index=-1)
        with pytest.raises(ValidationError):
            PptInsertSlidesOoxmlRequest.build(document_uri="file:///t.pptx", base64="UEsDBBQ=", final_slide_index=-1)
