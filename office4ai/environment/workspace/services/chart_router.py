"""Client round-trip driver for the OASP /ppt chart tools (#15, OASP 0.3.0).

The three chart tools (``ppt_insert_chart`` / ``ppt_get_chart`` /
``ppt_update_chart``) operate on a single execution path: the CONNECTED client
round-trip. The Server drives the Add-In's *generic*, chart-agnostic OOXML
carrier events (``ppt:get:slideOoxml`` / ``ppt:insert:slidesOoxml``, published
normative in OASP 0.3.0) and performs *all* chart OOXML work in-memory via
``chart_engine``'s base64 helpers. The Add-In only exports / inserts / replaces /
repositions slides — it never touches chart semantics.

Connected-only (Path A removed in F1 #68)
=========================================
The former on-disk "Path A" (a Server-side python-pptx read/write of a closed
``.pptx``) was removed once the authoring pipeline (``office_run_script`` +
python-pptx) took over offline chart work. The tools now refuse a DISCONNECTED
target with :data:`OFFLINE_MESSAGE`, steering the caller to that pipeline rather
than writing the file directly. This keeps chart behaviour aligned with the W4a
binary tool-convergence model (#63): a chart tool is only exposed while a /ppt
Add-In connection exists, and only truly operates on the document that is open.

Reactive degradation
=====================
office-editor4ai's Add-In handlers for the carrier events ship in a later
milestone (#16). Until they do, a CONNECTED round-trip attempt fails — the Add-In
has no handler (Socket.IO ``.call()`` times out), acks ``3016 API_NOT_SUPPORTED``,
or the platform lacks the required Office.js requirement set. We *react* to that
failure with :data:`DEGRADE_MESSAGE` (close the document and edit the chart
offline via the authoring pipeline) rather than gating on a capability flag
up-front. When the Add-In ships, the same round-trip code starts succeeding with
no Server change. A genuine business error from a *working* Add-In (e.g. 3010
chart not found) is surfaced as :class:`ChartEngineError`, never degraded.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.dtos.ppt import (
    PptGetSlideOoxmlRequest,
    PptInsertSlidesOoxmlRequest,
)
from office4ai.environment.workspace.services import chart_engine
from office4ai.environment.workspace.services.chart_engine import ChartEngineError

if TYPE_CHECKING:
    from office4ai.environment.workspace.base import BaseWorkspace
    from office4ai.environment.workspace.dtos.ppt import (
        CategoricalChartData,
        CategoricalChartUpdate,
        ChartInsertOptions,
        ScatterChartData,
        ScatterChartUpdate,
    )

_GET_SLIDE_OOXML = PptGetSlideOoxmlRequest.event_name  # "ppt:get:slideOoxml"
_INSERT_SLIDES_OOXML = PptInsertSlidesOoxmlRequest.event_name  # "ppt:insert:slidesOoxml"

#: Insert/replace must preserve the source slide faithfully — never re-theme it.
_KEEP_SOURCE_FORMATTING = "keepSourceFormatting"

#: Error codes that mean "path B is not usable on this client right now" (degrade).
_DEGRADE_CODES = frozenset({ErrorCode.API_NOT_SUPPORTED})  # "3016"

#: Refused when a chart tool targets a DISCONNECTED document: there is no live
#: round-trip to drive and the Server no longer edits closed files on disk (Path A
#: removed in F1 #68). The "3003" prefix keeps the familiar code; the guidance
#: steers offline chart work to the authoring pipeline.
OFFLINE_MESSAGE = (
    "3003: Target document is not open in PowerPoint via the Add-In, so live chart "
    "operations are unavailable. For chart work on a closed .pptx, use the authoring "
    "pipeline instead — call office_run_script with a python-pptx script (see the "
    "create-office-file / edit-office-file SKILL)."
)

#: Surfaced when the document is open but the client round-trip is not usable yet
#: (Add-In handler pending / unsupported requirement set / timeout). Keeps the
#: "3003" prefix; steers offline chart work to the authoring pipeline (there is no
#: longer an on-disk fallback — Path A was removed in F1 #68).
DEGRADE_MESSAGE = (
    "3003: The document is open in PowerPoint via the Add-In, but the client round-trip "
    "path (OASP ppt:get:slideOoxml / ppt:insert:slidesOoxml) is not available yet "
    "(Add-In handler pending or required Office.js requirement set unsupported). "
    "To work on the chart offline, ask the user to close the document, then use the "
    "authoring pipeline (office_run_script + python-pptx). This restriction lifts "
    "automatically once the Add-In ships the OOXML carrier events."
)

#: Surfaced when a chart operation on an OPEN document lacks the slideIndex needed
#: to export that live slide (both get and update require it for the round-trip).
LIVE_NEEDS_SLIDE_INDEX = (
    "3003: slideIndex (0-based) is required to operate on a chart while the document is "
    "open in PowerPoint — the live round-trip must export that specific slide. Supply the "
    "slide's index and retry."
)


class PathBUnavailable(Exception):
    """Path B (client round-trip) could not be used; the caller should degrade.

    Raised on transport failures (timeout / no socket / server down), a ``3016``
    ack, or a malformed success response — i.e. signals that mean "not available
    right now", as opposed to a genuine business error (which propagates as
    :class:`ChartEngineError`).
    """


async def _emit(workspace: BaseWorkspace, document_uri: str, event: str, data: dict[str, Any]) -> dict[str, Any]:
    """Drive one carrier event and return its business ``data`` payload.

    Raises :class:`PathBUnavailable` on any "path B not available" condition, and
    :class:`ChartEngineError` on a genuine business error from a working Add-In.
    """
    try:
        response = await workspace.emit_to_document(document_uri, event, data)
    except (TimeoutError, ValueError, RuntimeError) as exc:
        # Timeout = no Add-In handler (event unimplemented) or slow client.
        # ValueError = no socket / client not found (status raced to disconnected).
        # RuntimeError = Socket.IO server not running. All → degrade.
        raise PathBUnavailable(f"{event}: {type(exc).__name__}: {exc}") from exc

    if not isinstance(response, dict):
        raise PathBUnavailable(f"{event}: malformed response {response!r}")

    if response.get("success"):
        payload = response.get("data")
        return payload if isinstance(payload, dict) else {}

    # success is falsy — distinguish "not available" from a real business failure.
    error = response.get("error")
    code = error.get("code") if isinstance(error, dict) else None
    message = error.get("message") if isinstance(error, dict) else None
    if code in _DEGRADE_CODES:
        raise PathBUnavailable(f"{event}: {code} {message or ''}".strip())
    raise ChartEngineError(
        code=str(code or ErrorCode.OPERATION_FAILED),
        message=message or f"{event} failed on the client",
    )


def _require_base64(payload: dict[str, Any], event: str) -> str:
    """Pull the slide package out of a carrier response; degrade if malformed."""
    base64 = payload.get("base64")
    if not isinstance(base64, str) or not base64:
        raise PathBUnavailable(f"{event}: response missing 'base64' slide package")
    return base64


def _require_slide_id(payload: dict[str, Any], event: str) -> str:
    """Pull the opaque slideId out of an export response.

    An in-place write (replace + reposition) MUST have it: without a slideId the
    composite round-trip cannot delete the old slide, so the rebuilt slide would be
    appended and the original left behind (a silent duplicate-page corruption). A
    missing slideId means a malformed/partial Add-In response → degrade rather than
    risk that, so the caller falls back to the explicit "close the document" guidance.
    """
    slide_id = payload.get("slideId")
    if not isinstance(slide_id, str) or not slide_id:
        raise PathBUnavailable(f"{event}: response missing opaque 'slideId' (required for in-place replace)")
    return slide_id


async def insert_chart_path_b(
    workspace: BaseWorkspace,
    document_uri: str,
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None,
) -> dict[str, Any]:
    """CONNECTED insert: export the live target slide, add the chart in-memory,
    then re-insert it in place (replace the old slide, restore its position).

    The chart lands on the existing target slide (default slide 0), keeping that
    slide's other content, and operates on the live unsaved document.

    Deliberately does NOT use ``chart_engine.generate_chart_slide_base64`` (the
    "standalone new page" mode): insert always targets an existing slide here. The
    generate-a-new-page mode is reserved for a future explicit "append slide"
    capability (it needs an input signal to request it, and on an open document a
    deck-length query to detect the append case), wired alongside the #16 E2E work,
    not inferred here.
    """
    slide_index = options.slide_index if options is not None and options.slide_index is not None else 0

    exported = await _emit(workspace, document_uri, _GET_SLIDE_OOXML, {"slideIndex": slide_index})
    slide_b64 = _require_base64(exported, _GET_SLIDE_OOXML)
    slide_id = _require_slide_id(exported, _GET_SLIDE_OOXML)

    built = await chart_engine.insert_chart_into_slide_base64(slide_b64, chart_data, options)

    await _emit(
        workspace, document_uri, _INSERT_SLIDES_OOXML, _apply_payload(built["slideBase64"], slide_index, slide_id)
    )

    result = {key: value for key, value in built.items() if key != "slideBase64"}
    result["slideIndex"] = slide_index
    return result


async def get_chart_path_b(
    workspace: BaseWorkspace,
    document_uri: str,
    element_id: str,
    slide_index: int,
) -> dict[str, Any]:
    """CONNECTED get: export the (hinted) live slide and parse the chart Server-side."""
    exported = await _emit(workspace, document_uri, _GET_SLIDE_OOXML, {"slideIndex": slide_index})
    slide_b64 = _require_base64(exported, _GET_SLIDE_OOXML)
    result = await chart_engine.get_chart_from_slide_base64(slide_b64, element_id)
    result["slideIndex"] = slide_index  # report the deck index, not the mini-package's 0
    return result


async def update_chart_path_b(
    workspace: BaseWorkspace,
    document_uri: str,
    element_id: str,
    update: CategoricalChartUpdate | ScatterChartUpdate,
    slide_index: int,
) -> dict[str, Any]:
    """CONNECTED update: export the live slide, mutate the chart in-memory, re-insert in place."""
    exported = await _emit(workspace, document_uri, _GET_SLIDE_OOXML, {"slideIndex": slide_index})
    slide_b64 = _require_base64(exported, _GET_SLIDE_OOXML)
    slide_id = _require_slide_id(exported, _GET_SLIDE_OOXML)

    built = await chart_engine.update_chart_in_slide_base64(slide_b64, element_id, update)

    await _emit(
        workspace, document_uri, _INSERT_SLIDES_OOXML, _apply_payload(built["slideBase64"], slide_index, slide_id)
    )

    return {key: value for key, value in built.items() if key != "slideBase64"}


def _apply_payload(slide_base64: str, slide_index: int, slide_id: str) -> dict[str, Any]:
    """Build the ``ppt:insert:slidesOoxml`` payload for an in-place slide replacement.

    Insert the rebuilt slide after the target, delete the old slide (by its opaque
    ``slideId`` from the export, guaranteed non-empty by :func:`_require_slide_id`),
    and move the new slide back to the original index — the composite round-trip from
    OASP 0.3.0 (best-effort, non-atomic; on partial failure the Add-In returns
    ``error.details = {stage, partiallyApplied, createdSlideId}``, surfaced here as a
    genuine business error via :func:`_emit`, never silently degraded).
    """
    return {
        "base64": slide_base64,
        "targetSlideIndex": slide_index,
        "formatting": _KEEP_SOURCE_FORMATTING,
        "replaceSlideId": slide_id,
        "finalSlideIndex": slide_index,
    }
