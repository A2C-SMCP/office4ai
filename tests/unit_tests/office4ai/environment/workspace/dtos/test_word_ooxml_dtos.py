"""Unit tests for Word OOXML carrier DTOs (OASP /word Draft, v0.3.0).

Covers word:get:ooxml / word:insert:ooxml request DTOs. Focus: camelCase wire
alias round-trips, snake_case input acceptance (populate_by_name), and
scope / insertLocation enum + required-field validation.

Contract reference: cross-ask with office-editor4ai (office4ai#46) — the carrier
is a Flat OPC XML *string* (not base64), and get echoes the effective scope.
"""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import request_registry
from office4ai.environment.workspace.dtos.word import (
    WordGetOoxmlRequest,
    WordGetOoxmlResponse,
    WordInsertOoxmlRequest,
    WordInsertOoxmlResponse,
    WordInsertOoxmlResult,
)


class TestRegistry:
    def test_registry_has_ooxml_events(self) -> None:
        assert request_registry.contains("word:get:ooxml")
        assert request_registry.contains("word:insert:ooxml")

    def test_registry_resolves_to_dto_classes(self) -> None:
        assert request_registry.get("word:get:ooxml") is WordGetOoxmlRequest
        assert request_registry.get("word:insert:ooxml") is WordInsertOoxmlRequest


class TestWordGetOoxmlRequest:
    def test_scope_defaults_to_selection(self) -> None:
        req = WordGetOoxmlRequest.build(document_uri="file:///t.docx")
        assert req.scope == "selection"
        assert req.to_payload()["scope"] == "selection"

    def test_round_trip_emits_camel_case(self) -> None:
        req = WordGetOoxmlRequest.build(document_uri="file:///t.docx", scope="body")
        payload = req.to_payload()
        assert payload["documentUri"] == "file:///t.docx"
        assert payload["scope"] == "body"

    def test_accepts_camel_case_input(self) -> None:
        req = WordGetOoxmlRequest.model_validate({"requestId": "r1", "documentUri": "file:///t.docx", "scope": "body"})
        assert req.scope == "body"

    def test_invalid_scope_rejected(self) -> None:
        with pytest.raises(ValidationError):
            WordGetOoxmlRequest.build(document_uri="file:///t.docx", scope="paragraph")


class TestWordInsertOoxmlRequest:
    def test_minimal_request(self) -> None:
        req = WordInsertOoxmlRequest.build(
            document_uri="file:///t.docx", ooxml="<pkg:package/>", insert_location="Replace"
        )
        payload = req.to_payload()
        assert payload["ooxml"] == "<pkg:package/>"
        assert payload["insertLocation"] == "Replace"
        # scope falls back to its default and is emitted in camelCase form
        assert payload["scope"] == "selection"
        # BaseRequest still supplies its default timestamp
        assert "timestamp" in payload
        # no snake_case leakage
        assert "insert_location" not in payload

    def test_ooxml_required(self) -> None:
        with pytest.raises(ValidationError):
            WordInsertOoxmlRequest.build(document_uri="file:///t.docx", insert_location="Start")

    def test_ooxml_must_be_non_empty(self) -> None:
        with pytest.raises(ValidationError):
            WordInsertOoxmlRequest.build(document_uri="file:///t.docx", ooxml="", insert_location="End")

    def test_insert_location_required(self) -> None:
        with pytest.raises(ValidationError):
            WordInsertOoxmlRequest.build(document_uri="file:///t.docx", ooxml="<pkg:package/>")

    def test_invalid_insert_location_rejected(self) -> None:
        with pytest.raises(ValidationError):
            WordInsertOoxmlRequest.build(
                document_uri="file:///t.docx", ooxml="<pkg:package/>", insert_location="Before"
            )

    def test_invalid_scope_rejected(self) -> None:
        with pytest.raises(ValidationError):
            WordInsertOoxmlRequest.build(
                document_uri="file:///t.docx",
                ooxml="<pkg:package/>",
                insert_location="Start",
                scope="table",
            )

    def test_full_round_trip_fields_camel_case(self) -> None:
        req = WordInsertOoxmlRequest.build(
            document_uri="file:///t.docx",
            ooxml="<pkg:package/>",
            insert_location="End",
            scope="body",
        )
        payload = req.to_payload()
        assert payload["insertLocation"] == "End"
        assert payload["scope"] == "body"
        assert "insert_location" not in payload

    def test_accepts_camel_case_input(self) -> None:
        req = WordInsertOoxmlRequest.model_validate(
            {
                "requestId": "r1",
                "documentUri": "file:///t.docx",
                "ooxml": "<pkg:package/>",
                "insertLocation": "Start",
                "scope": "body",
            }
        )
        assert req.insert_location == "Start"
        assert req.scope == "body"

    def test_accepts_snake_case_input(self) -> None:
        # populate_by_name=True: snake_case insert_location is accepted alongside the alias
        req = WordInsertOoxmlRequest.model_validate(
            {
                "requestId": "r1",
                "documentUri": "file:///t.docx",
                "ooxml": "<pkg:package/>",
                "insert_location": "End",
            }
        )
        assert req.insert_location == "End"
        assert req.scope == "selection"


class TestOoxmlResponses:
    def test_get_response_deserializes_success(self) -> None:
        resp = WordGetOoxmlResponse.model_validate(
            {
                "requestId": "r1",
                "success": True,
                "data": {"scope": "body", "ooxml": "<pkg:package/>"},
                "timestamp": 123,
            }
        )
        assert resp.success is True
        assert resp.data is not None
        assert resp.data.scope == "body"  # effective scope echoed
        assert resp.data.ooxml == "<pkg:package/>"
        assert resp.error is None

    def test_get_response_error_branch(self) -> None:
        resp = WordGetOoxmlResponse.model_validate(
            {
                "requestId": "r1",
                "success": False,
                "error": {"code": "3016", "message": "not supported"},
                "timestamp": 123,
            }
        )
        assert resp.success is False
        assert resp.data is None
        assert resp.error is not None
        assert resp.error.code == "3016"

    def test_insert_response_deserializes_camel_case(self) -> None:
        resp = WordInsertOoxmlResponse.model_validate(
            {
                "requestId": "r1",
                "success": True,
                "data": {"scope": "selection", "insertLocation": "Replace"},
                "timestamp": 123,
            }
        )
        assert resp.data is not None
        assert resp.data.scope == "selection"
        assert resp.data.insert_location == "Replace"

    def test_insert_result_round_trips_camel_case(self) -> None:
        result = WordInsertOoxmlResult.model_validate({"scope": "body", "insertLocation": "End"})
        payload = result.model_dump(by_alias=True)
        assert payload["insertLocation"] == "End"
        assert "insert_location" not in payload
