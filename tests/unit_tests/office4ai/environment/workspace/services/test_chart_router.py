"""Direct unit tests for the chart dual-path router (#15, OASP 0.3.0).

These are deliberately WHITE-BOX: they exercise ``chart_router._emit`` and the
``_require_*`` guards directly with crafted carrier-event responses, so every
branch of the "path B available vs degrade vs surface-business-error" classifier
is pinned in isolation — no ``.pptx`` fixture, no ``chart_engine`` OOXML work, no
``document_lock``. The tool-level tests in ``test_ppt_tools.py`` cover the
end-to-end wiring; this file localises a regression to the router's pure logic
(e.g. someone narrowing the ``except`` tuple, or swapping degrade-vs-surface).
"""

from __future__ import annotations

from typing import Any
from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.services import chart_router
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.chart_router import PathBUnavailable

_EVENT = "ppt:get:slideOoxml"
_URI = "file:///deck.pptx"


def _ws(side_effect: Any) -> Any:
    """A workspace whose ``emit_to_document`` runs ``side_effect`` (sync fn returning
    the wire response dict, or an exception instance to raise)."""
    ws = MagicMock()
    ws.emit_to_document = AsyncMock(side_effect=side_effect)
    return ws


# ---------------------------------------------------------------------------
# _emit — success branches
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_emit_success_returns_data_payload() -> None:
    ws = _ws(lambda u, e, d: {"success": True, "data": {"slideId": "s", "base64": "b"}})
    assert await chart_router._emit(ws, _URI, _EVENT, {}) == {"slideId": "s", "base64": "b"}


@pytest.mark.asyncio
async def test_emit_success_non_dict_payload_returns_empty_dict() -> None:
    # Defensive: a working ack with a non-dict ``data`` is treated as empty,
    # then the _require_* guards downstream force a degrade.
    ws = _ws(lambda u, e, d: {"success": True, "data": "oops-not-a-dict"})
    assert await chart_router._emit(ws, _URI, _EVENT, {}) == {}


@pytest.mark.asyncio
async def test_emit_success_missing_data_returns_empty_dict() -> None:
    ws = _ws(lambda u, e, d: {"success": True})
    assert await chart_router._emit(ws, _URI, _EVENT, {}) == {}


# ---------------------------------------------------------------------------
# _emit — "path B unavailable" → PathBUnavailable (degrade)
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "exc",
    [
        TimeoutError("no ack from Add-In"),  # event unimplemented / slow client
        ValueError("no socket / client not found"),  # status raced to disconnected
        RuntimeError("Socket.IO server not running"),
    ],
)
async def test_emit_transport_exceptions_degrade(exc: Exception) -> None:
    ws = _ws(exc)
    with pytest.raises(PathBUnavailable):
        await chart_router._emit(ws, _URI, _EVENT, {})


@pytest.mark.asyncio
@pytest.mark.parametrize("response", ["a-bare-string", None, 123, ["a", "list"]])
async def test_emit_non_dict_response_degrades(response: Any) -> None:
    # A buggy/legacy Add-In returning a non-dict ack must degrade, not crash.
    ws = _ws(lambda u, e, d: response)
    with pytest.raises(PathBUnavailable):
        await chart_router._emit(ws, _URI, _EVENT, {})


@pytest.mark.asyncio
async def test_emit_3016_api_not_supported_degrades() -> None:
    ws = _ws(lambda u, e, d: {"success": False, "error": {"code": "3016", "message": "PowerPointApi 1.8 unavailable"}})
    with pytest.raises(PathBUnavailable):
        await chart_router._emit(ws, _URI, _EVENT, {})


# ---------------------------------------------------------------------------
# _emit — genuine business error from a WORKING Add-In → ChartEngineError (surfaced)
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_emit_business_error_surfaced_not_degraded() -> None:
    # A non-degrade error code (e.g. 3010 chart-not-found) means the Add-In worked
    # and reported a real failure — must surface as ChartEngineError, never degrade.
    ws = _ws(lambda u, e, d: {"success": False, "error": {"code": "3010", "message": "chart not found"}})
    with pytest.raises(ChartEngineError) as exc_info:
        await chart_router._emit(ws, _URI, _EVENT, {})
    assert exc_info.value.code == "3010"
    assert exc_info.value.message == "chart not found"


@pytest.mark.asyncio
async def test_emit_falsy_success_without_error_object_surfaces_3004() -> None:
    ws = _ws(lambda u, e, d: {"success": False})
    with pytest.raises(ChartEngineError) as exc_info:
        await chart_router._emit(ws, _URI, _EVENT, {})
    assert exc_info.value.code == ErrorCode.OPERATION_FAILED  # "3004"


@pytest.mark.asyncio
async def test_emit_falsy_success_with_non_dict_error_surfaces_3004() -> None:
    ws = _ws(lambda u, e, d: {"success": False, "error": "boom"})
    with pytest.raises(ChartEngineError) as exc_info:
        await chart_router._emit(ws, _URI, _EVENT, {})
    assert exc_info.value.code == ErrorCode.OPERATION_FAILED


# ---------------------------------------------------------------------------
# _require_base64 / _require_slide_id guards
# ---------------------------------------------------------------------------


def test_require_base64_returns_present_value() -> None:
    assert chart_router._require_base64({"base64": "abc"}, _EVENT) == "abc"


@pytest.mark.parametrize("payload", [{}, {"base64": ""}, {"base64": None}, {"base64": 123}])
def test_require_base64_missing_or_blank_degrades(payload: dict[str, Any]) -> None:
    with pytest.raises(PathBUnavailable):
        chart_router._require_base64(payload, _EVENT)


def test_require_slide_id_returns_present_value() -> None:
    assert chart_router._require_slide_id({"slideId": "sid-1"}, _EVENT) == "sid-1"


@pytest.mark.parametrize("payload", [{}, {"slideId": ""}, {"slideId": None}, {"slideId": 5}])
def test_require_slide_id_missing_or_blank_degrades(payload: dict[str, Any]) -> None:
    # Missing/blank slideId on an in-place replace would orphan the old slide
    # (silent duplicate page) → must degrade rather than risk it.
    with pytest.raises(PathBUnavailable):
        chart_router._require_slide_id(payload, _EVENT)


# ---------------------------------------------------------------------------
# _apply_payload wire contract
# ---------------------------------------------------------------------------


def test_apply_payload_builds_in_place_replace_contract() -> None:
    payload = chart_router._apply_payload("B64DATA", 3, "sid-export")
    assert payload == {
        "base64": "B64DATA",
        "targetSlideIndex": 3,
        "formatting": "keepSourceFormatting",
        "replaceSlideId": "sid-export",
        "finalSlideIndex": 3,
    }
