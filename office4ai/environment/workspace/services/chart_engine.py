"""
OOXML chart engine for OASP /ppt chart events (insert / get / update).

PowerPoint Office.js does not expose chart creation or data-update APIs (see
office-js#5463). The OASP protocol therefore routes ppt:insert:chart /
ppt:get:chart / ppt:update:chart through this Server-side engine, which builds
and mutates the .pptx OOXML directly via ``python-pptx``.

Single execution path: in-memory base64 (open document, OASP 0.3.0)
===================================================================
The Server never touches disk; it exchanges single-slide ``.pptx`` packages as
base64 over the wire with the Add-In (which exports the live slide via
``Slide.exportAsBase64`` and re-inserts via ``insertSlidesFromBase64``). The
engine offers two in-memory shapes:

* *generate standalone single page* — build a fresh 1-slide deck containing the
  chart and return it as base64 (used by "insert → new page"; needs no input
  package);
* *modify a single page* — load a base64 single-slide package, locate/add/edit
  the chart, return the package back as base64 (used by get / update /
  insert-into-existing-page).

All chart OOXML logic stays here in python-pptx; the Add-In only carries generic,
chart-agnostic primitives. The chart tools (#15) drive this path only when the
target document is CONNECTED; a closed document is refused with guidance to the
authoring pipeline (``office_run_script`` + python-pptx). The former on-disk
"Path A" was removed in F1 (#68) — offline chart authoring is served by the
scripting runtime, not by this engine writing ``.pptx`` files directly.

Operational constraints:

- Concurrent chart calls on the same document must be serialized (see
  ``document_lock.DocumentLockManager``); two simultaneous writes corrupt the
  OOXML package.
- Latency is dominated by python-pptx parsing of the exported slide package.

Element identity
================
The wire ``elementId`` is an **opaque, server-assigned string** of the form
``"oasp-chart-<uuid>"`` (per OASP ``data-structures.md#element-id-opacity``).
Consumers must not parse it — they round-trip it verbatim. The engine stores the
id in the chart shape's OOXML ``<p:cNvPr name="...">`` (python-pptx
``graphic_frame.name``), which survives both ``save()``/reload and the Add-In's
whole-slide round-trip (spike office-editor4ai#34: ``@name`` survives, geometry
readable, ``masterLeak: 0``). Charts are therefore relocated by ``shape.name``.

For backward compatibility we also accept the legacy ``"chart-<slide>-<shape>"``
(or ``"chart-<shape>"``) form emitted by 0.2.0, resolving it by native
python-pptx ``shape_id``. A legacy chart that is recreated during an update is
migrated to a fresh opaque id.
"""

from __future__ import annotations

import asyncio
import base64
import binascii
import io
import math
import uuid
from dataclasses import dataclass
from typing import Any

from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Emu, Pt

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.dtos.ppt import (
    CategoricalChartData,
    CategoricalChartUpdate,
    ChartInsertOptions,
    ScatterChartData,
    ScatterChartUpdate,
    ScatterPoint,
    ScatterSeries,
)

# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


@dataclass
class ChartEngineError(Exception):
    """Raised by the engine; carries an OASP error code + human-readable message."""

    code: str
    message: str

    def __str__(self) -> str:  # pragma: no cover - trivial formatting
        return f"{self.code}: {self.message}"


# ---------------------------------------------------------------------------
# ChartType ↔ XL_CHART_TYPE mapping
# ---------------------------------------------------------------------------

_CHART_TYPE_MAP: dict[str, XL_CHART_TYPE] = {
    "ColumnClustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "ColumnStacked": XL_CHART_TYPE.COLUMN_STACKED,
    "BarClustered": XL_CHART_TYPE.BAR_CLUSTERED,
    "Line": XL_CHART_TYPE.LINE,
    "LineMarkers": XL_CHART_TYPE.LINE_MARKERS,
    "Pie": XL_CHART_TYPE.PIE,
    "Doughnut": XL_CHART_TYPE.DOUGHNUT,
    "Area": XL_CHART_TYPE.AREA,
    "Radar": XL_CHART_TYPE.RADAR,
    "Scatter": XL_CHART_TYPE.XY_SCATTER,
}

_CATEGORICAL_TYPES = {
    "ColumnClustered",
    "ColumnStacked",
    "BarClustered",
    "Line",
    "LineMarkers",
    "Pie",
    "Doughnut",
    "Area",
    "Radar",
}


def _xl_chart_type(chart_type: str) -> XL_CHART_TYPE:
    try:
        return _CHART_TYPE_MAP[chart_type]
    except KeyError as exc:
        raise ChartEngineError(
            code=ErrorCode.INVALID_PARAM,
            message=f"Unsupported chartType: {chart_type!r}",
        ) from exc


# ---------------------------------------------------------------------------
# Validation (raises 3015 INVALID_CHART_DATA on dimension issues)
# ---------------------------------------------------------------------------


def _validate_categorical(data: CategoricalChartData) -> None:
    expected = len(data.categories)
    for idx, series in enumerate(data.series):
        if len(series.values) != expected:
            raise ChartEngineError(
                code=ErrorCode.INVALID_CHART_DATA,
                message=(
                    f"series[{idx}].values length ({len(series.values)}) does not match categories length ({expected})"
                ),
            )


def _validate_scatter(data: ScatterChartData) -> None:
    for idx, series in enumerate(data.series):
        if not series.points:
            raise ChartEngineError(
                code=ErrorCode.INVALID_CHART_DATA,
                message=f"series[{idx}].points must not be empty",
            )
        for pidx, p in enumerate(series.points):
            if not (math.isfinite(p.x) and math.isfinite(p.y)):
                raise ChartEngineError(
                    code=ErrorCode.INVALID_CHART_DATA,
                    message=f"series[{idx}].points[{pidx}] contains non-finite x/y",
                )


def _validate_chart_data(chart_data: CategoricalChartData | ScatterChartData) -> None:
    if isinstance(chart_data, CategoricalChartData):
        _validate_categorical(chart_data)
    else:
        _validate_scatter(chart_data)


# ---------------------------------------------------------------------------
# python-pptx helpers — chart construction / extraction / mutation
# ---------------------------------------------------------------------------

_DEFAULT_WIDTH_PT = 480.0
_DEFAULT_HEIGHT_PT = 320.0

#: python-pptx default template "Blank" layout index (no placeholders).
_BLANK_LAYOUT_INDEX = 6


def _resolve_geometry(prs: Any, options: ChartInsertOptions | None) -> tuple[Any, Any, Any, Any]:
    """Resolve (x, y, cx, cy) as python-pptx ``Length`` values. Default 480x320 pt, centered."""
    width_pt = options.width if options and options.width else _DEFAULT_WIDTH_PT
    height_pt = options.height if options and options.height else _DEFAULT_HEIGHT_PT
    cx = Pt(width_pt)
    cy = Pt(height_pt)

    if options and options.left is not None:
        x: Any = Pt(options.left)
    else:
        x = Emu(max(0, (int(prs.slide_width) - int(cx)) // 2))
    if options and options.top is not None:
        y: Any = Pt(options.top)
    else:
        y = Emu(max(0, (int(prs.slide_height) - int(cy)) // 2))
    return x, y, cx, cy


def _build_categorical_data(data: CategoricalChartData) -> CategoryChartData:
    cd = CategoryChartData()  # type: ignore[no-untyped-call]
    cd.categories = list(data.categories)
    for series in data.series:
        cd.add_series(series.name, list(series.values))  # type: ignore[no-untyped-call]
    return cd


def _build_xy_data(data: ScatterChartData) -> XyChartData:
    cd = XyChartData()  # type: ignore[no-untyped-call]
    for series in data.series:
        s = cd.add_series(series.name)  # type: ignore[no-untyped-call]
        for p in series.points:
            s.add_data_point(p.x, p.y)
    return cd


def _apply_display_options(
    chart: Any,
    title: str | None,
    show_legend: bool | None,
    show_data_labels: bool | None,
    *,
    title_explicit_none: bool = False,
) -> None:
    """Apply title / legend / data-label toggles. ``title_explicit_none`` deletes the title."""
    if title is not None:
        chart.has_title = True
        chart.chart_title.text_frame.text = title
    elif title_explicit_none:
        chart.has_title = False

    if show_legend is not None:
        chart.has_legend = bool(show_legend)

    if show_data_labels is not None and chart.plots:
        try:
            chart.plots[0].has_data_labels = bool(show_data_labels)
        except (AttributeError, ValueError):
            # Scatter plots don't support data labels via this code path — silently skip.
            pass


def _slide_index(prs: Any, slide: Any) -> int:
    for idx, s in enumerate(prs.slides):
        if s is slide:
            return idx
    return -1


# ---------------------------------------------------------------------------
# Element identity — opaque ``oasp-chart-<uuid>`` + legacy ``chart-<slide>-<shape>``
# ---------------------------------------------------------------------------

_OPAQUE_PREFIX = "oasp-chart-"


def _new_element_id() -> str:
    """Mint a fresh opaque, server-assigned chart id (``oasp-chart-<uuid4>``)."""
    return f"{_OPAQUE_PREFIX}{uuid.uuid4()}"


def _is_opaque_id(element_id: str) -> bool:
    return element_id.startswith(_OPAQUE_PREFIX)


def _parse_legacy_element_id(element_id: str) -> tuple[int | None, int]:
    """Parse legacy ``chart-<slide>-<shape>`` (preferred) or ``chart-<shape>``.

    Only called for ids that start with ``"chart-"``. Returns
    ``(slide_index_or_None, shape_id)``. Raises 3010 on malformed input.
    """
    rest = element_id[len("chart-") :]
    parts = rest.split("-")
    try:
        if len(parts) == 2:
            return int(parts[0]), int(parts[1])
        if len(parts) == 1:
            return None, int(parts[0])
    except ValueError:
        pass
    raise ChartEngineError(
        code=ErrorCode.ELEMENT_NOT_FOUND,
        message=f"Invalid chart elementId: {element_id!r}",
    )


def _ordered_slides(prs: Any, hint_index: int | None) -> list[Any]:
    """Slides with ``hint_index`` (if valid) searched first, then the rest in order."""
    slides = list(prs.slides)
    if hint_index is not None and 0 <= hint_index < len(slides):
        return [slides[hint_index]] + [s for i, s in enumerate(slides) if i != hint_index]
    return slides


def _find_chart(prs: Any, element_id: str, hint_index: int | None = None) -> tuple[Any, Any, int]:
    """Return (slide, chart_shape, slide_index) for the chart with this elementId.

    Primary lookup is by opaque ``shape.name`` (== OOXML ``cNvPr/@name``). If the
    id is in the legacy ``chart-<slide>-<shape>`` form and no shape carries that
    name, fall back to matching the native python-pptx ``shape_id``.

    Raises 3010 ELEMENT_NOT_FOUND if no chart shape matches.
    """
    # 1. Primary: opaque id written to cNvPr/@name (survives save/reload + round-trip).
    #    First match wins. opaque UUID ids make same-name collisions improbable in normal
    #    operation; a stronger guard (customXmlParts registry, design §5) is deferred to the
    #    routing work (#15) where the open-document registry becomes the source of truth.
    for slide in _ordered_slides(prs, hint_index):
        for shape in slide.shapes:
            if getattr(shape, "has_chart", False) and shape.name == element_id:
                return slide, shape, _slide_index(prs, slide)

    # 2. Legacy fallback: chart-<slide>-<shape> / chart-<shape> by native shape id.
    if element_id.startswith("chart-"):
        parsed_slide, shape_id = _parse_legacy_element_id(element_id)
        target = parsed_slide if parsed_slide is not None else hint_index
        for slide in _ordered_slides(prs, target):
            for shape in slide.shapes:
                if getattr(shape, "has_chart", False) and int(shape.shape_id) == shape_id:
                    return slide, shape, _slide_index(prs, slide)

    raise ChartEngineError(
        code=ErrorCode.ELEMENT_NOT_FOUND,
        message=f"Chart not found: {element_id}",
    )


_C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _extract_scatter_points(series: Any) -> list[dict[str, float]]:
    """Read XY series points by parsing c:xVal / c:yVal numeric refs from OOXML."""

    def _read_axis(parent_tag: str) -> list[float]:
        node = series._element.find(f"{{{_C_NS}}}{parent_tag}")
        if node is None:
            return []
        num_ref = node.find(f"{{{_C_NS}}}numRef")
        num_lit = node.find(f"{{{_C_NS}}}numLit")
        cache = num_ref.find(f"{{{_C_NS}}}numCache") if num_ref is not None else None
        source = cache if cache is not None else num_lit
        if source is None:
            return []
        out: list[float] = []
        for pt in source.findall(f"{{{_C_NS}}}pt"):
            v = pt.find(f"{{{_C_NS}}}v")
            if v is None or v.text is None:
                continue
            try:
                out.append(float(v.text))
            except ValueError:
                continue
        return out

    xs = _read_axis("xVal")
    ys = _read_axis("yVal")
    return [{"x": x, "y": y} for x, y in zip(xs, ys, strict=False)]


def _extract_chart_data(chart: Any) -> dict[str, Any]:
    """Read python-pptx chart back into the OASP ChartData wire shape (camelCase)."""
    xl_type = chart.chart_type
    # Reverse-map XL_CHART_TYPE → OASP chartType string.
    reverse = {v: k for k, v in _CHART_TYPE_MAP.items()}
    chart_type_str = reverse.get(xl_type)
    if chart_type_str is None:
        # Unknown / unmapped concrete enum (e.g. variant subtype) — fall back to closest base.
        # We pick the first key whose XL_CHART_TYPE shares the same name root.
        for key, val in _CHART_TYPE_MAP.items():
            if val == xl_type:
                chart_type_str = key
                break
        chart_type_str = chart_type_str or "ColumnClustered"

    title = None
    if chart.has_title:
        title = chart.chart_title.text_frame.text or None
    show_legend = bool(chart.has_legend)
    # Scatter plots (CT_ScatterChart) don't expose dLbls — fall back to False.
    show_data_labels = False
    if chart.plots:
        try:
            show_data_labels = bool(chart.plots[0].has_data_labels)
        except (AttributeError, ValueError):
            show_data_labels = False

    if chart_type_str == "Scatter":
        series_out: list[dict[str, Any]] = []
        for series in chart.series:
            points = _extract_scatter_points(series)
            series_out.append({"name": series.name, "points": points})
        return {
            "chartType": "Scatter",
            "series": series_out,
            "title": title,
            "showLegend": show_legend,
            "showDataLabels": show_data_labels,
        }

    # Categorical
    plot = chart.plots[0] if chart.plots else None
    categories_raw = list(plot.categories) if plot is not None else []
    categories = [str(c) for c in categories_raw]
    series_out_cat: list[dict[str, Any]] = []
    for series in chart.series:
        values = [float(v) if v is not None else 0.0 for v in series.values]
        series_out_cat.append({"name": series.name, "values": values})
    return {
        "chartType": chart_type_str,
        "categories": categories,
        "series": series_out_cat,
        "title": title,
        "showLegend": show_legend,
        "showDataLabels": show_data_labels,
    }


def _is_categorical(chart_type: str) -> bool:
    return chart_type in _CATEGORICAL_TYPES


# ---------------------------------------------------------------------------
# OOXML core — operate on an already-loaded ``Presentation`` (shared by every entry point)
# ---------------------------------------------------------------------------


def _geometry_dict(shape: Any) -> dict[str, float]:
    return {
        "left": Emu(int(shape.left or 0)).pt,
        "top": Emu(int(shape.top or 0)).pt,
        "width": Emu(int(shape.width or 0)).pt,
        "height": Emu(int(shape.height or 0)).pt,
    }


def _add_chart_to_slide(
    prs: Any,
    slide: Any,
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None,
) -> tuple[Any, str]:
    """Add a chart shape to ``slide``, tag it with a fresh opaque id, return (shape, id)."""
    x, y, cx, cy = _resolve_geometry(prs, options)

    # Core invariant: never build a chart from invalid data. The public async wrappers
    # also validate up-front (fast-fail before the expensive Presentation parse / disk read);
    # this re-validation guards direct core callers and is idempotent.
    if isinstance(chart_data, CategoricalChartData):
        _validate_categorical(chart_data)
        xl_type = _xl_chart_type(chart_data.chart_type)
        cd: Any = _build_categorical_data(chart_data)
    else:
        _validate_scatter(chart_data)
        xl_type = _xl_chart_type(chart_data.chart_type)
        cd = _build_xy_data(chart_data)

    chart_shape: Any = slide.shapes.add_chart(xl_type, x, y, cx, cy, cd)
    chart = chart_shape.chart

    _apply_display_options(
        chart,
        title=chart_data.title,
        show_legend=chart_data.show_legend,
        show_data_labels=chart_data.show_data_labels,
    )

    element_id = _new_element_id()
    chart_shape.name = element_id  # → OOXML <p:cNvPr name="oasp-chart-..."> (relocatable token)
    return chart_shape, element_id


def _insert_result(
    element_id: str,
    slide_index: int,
    chart_data: CategoricalChartData | ScatterChartData,
    chart_shape: Any,
) -> dict[str, Any]:
    return {
        "elementId": element_id,
        "slideIndex": slide_index,
        "chartType": chart_data.chart_type,
        "seriesCount": len(chart_data.series),
        **_geometry_dict(chart_shape),
    }


def _modify_chart_in_prs(
    slide: Any,
    shape: Any,
    element_id: str,
    update: CategoricalChartUpdate | ScatterChartUpdate,
) -> tuple[str, list[str]]:
    """Apply a partial update to a located chart shape (no save). Returns (elementId, updatedFields).

    Type changes (cross-variant or same-variant) are delete-and-recreate at the
    same geometry; the recreated shape inherits the same opaque id (legacy ids
    are migrated to a fresh opaque id). Pure data/display updates mutate in place.
    """
    chart = shape.chart

    # Determine current chartType string (best-effort) for variant comparison.
    reverse = {v: k for k, v in _CHART_TYPE_MAP.items()}
    current_type = reverse.get(chart.chart_type) or "ColumnClustered"

    target_type = update.chart_type
    target_is_categorical = _is_categorical(target_type)
    current_is_categorical = _is_categorical(current_type)

    updated_fields: list[str] = []

    cross_variant = target_is_categorical != current_is_categorical
    same_variant_type_change = (target_type != current_type) and not cross_variant

    if cross_variant:
        # Cross-variant switch needs full new series. Rebuild from scratch at the same geometry.
        if isinstance(update, CategoricalChartUpdate):
            if update.series is None or update.categories is None:
                raise ChartEngineError(
                    code=ErrorCode.INVALID_CHART_DATA,
                    message=(
                        "Cross-variant switch to categorical requires both 'categories' and 'series' "
                        "(scatter shape has no compatible data)"
                    ),
                )
            new_data = CategoricalChartData(
                chartType=update.chart_type,
                categories=update.categories,
                series=update.series,
                title=update.title,
                showLegend=update.show_legend,
                showDataLabels=update.show_data_labels,
            )
            _validate_categorical(new_data)
            cd: Any = _build_categorical_data(new_data)
        else:
            if update.series is None:
                raise ChartEngineError(
                    code=ErrorCode.INVALID_CHART_DATA,
                    message="Cross-variant switch to scatter requires 'series' with points",
                )
            new_scatter = ScatterChartData(
                chartType=update.chart_type,
                series=update.series,
                title=update.title,
                showLegend=update.show_legend,
                showDataLabels=update.show_data_labels,
            )
            _validate_scatter(new_scatter)
            cd = _build_xy_data(new_scatter)

        # Capture geometry, drop old shape, add new chart at same coords.
        x: Any = Emu(int(shape.left or 0))
        y: Any = Emu(int(shape.top or 0))
        cx: Any = Emu(int(shape.width or 0))
        cy: Any = Emu(int(shape.height or 0))
        sp = shape._element
        sp.getparent().remove(sp)
        new_xl = _xl_chart_type(target_type)
        new_shape = slide.shapes.add_chart(new_xl, x, y, cx, cy, cd)
        chart = new_shape.chart
        element_id = element_id if _is_opaque_id(element_id) else _new_element_id()
        new_shape.name = element_id  # preserve / migrate the opaque id onto the recreated shape
        updated_fields.extend(["chartType", "series"])
        if isinstance(update, CategoricalChartUpdate) and update.categories is not None:
            updated_fields.append("categories")
        _apply_display_options(
            chart,
            title=update.title,
            show_legend=update.show_legend,
            show_data_labels=update.show_data_labels,
            title_explicit_none=False,
        )
        if update.title is not None:
            updated_fields.append("title")
        if update.show_legend is not None:
            updated_fields.append("showLegend")
        if update.show_data_labels is not None:
            updated_fields.append("showDataLabels")
    else:
        # Same-variant.
        # python-pptx does not expose ``chart.chart_type`` as a setter, so a chart-type
        # change inside the same variant (e.g. Pie → ColumnClustered) is also handled by
        # delete-and-recreate at the same geometry — same strategy as cross-variant.
        if same_variant_type_change:
            current_data = _extract_chart_data(chart)
            if isinstance(update, CategoricalChartUpdate):
                new_categories = (
                    update.categories if update.categories is not None else current_data.get("categories", [])
                )
                if update.series is not None:
                    new_data_full = CategoricalChartData(
                        chartType=update.chart_type,
                        categories=list(new_categories),
                        series=update.series,
                        title=update.title,
                        showLegend=update.show_legend,
                        showDataLabels=update.show_data_labels,
                    )
                else:
                    existing_series_raw = current_data.get("series", [])
                    existing_series = [
                        {"name": s.get("name", ""), "values": s.get("values", [])} for s in existing_series_raw
                    ]
                    new_data_full = CategoricalChartData.model_validate(
                        {
                            "chartType": update.chart_type,
                            "categories": list(new_categories),
                            "series": existing_series,
                            "title": update.title,
                            "showLegend": update.show_legend,
                            "showDataLabels": update.show_data_labels,
                        }
                    )
                _validate_categorical(new_data_full)
                cd = _build_categorical_data(new_data_full)
            else:
                # Same-variant scatter type change is impossible (only one scatter type),
                # but keep the branch for symmetry / future variants.
                if update.series is None:
                    series_raw = current_data.get("series", [])
                    points_per_series = [
                        [ScatterPoint(x=p["x"], y=p["y"]) for p in s.get("points", [])] for s in series_raw
                    ]
                    fallback_series = [
                        ScatterSeries(
                            name=series_raw[i].get("name", ""), points=points_per_series[i] or [ScatterPoint(x=0, y=0)]
                        )
                        for i in range(len(series_raw))
                    ] or [ScatterSeries(name="series1", points=[ScatterPoint(x=0, y=0)])]
                else:
                    fallback_series = list(update.series)
                new_scatter = ScatterChartData(
                    chartType="Scatter",
                    series=fallback_series,
                    title=update.title,
                    showLegend=update.show_legend,
                    showDataLabels=update.show_data_labels,
                )
                _validate_scatter(new_scatter)
                cd = _build_xy_data(new_scatter)

            x = Emu(int(shape.left or 0))
            y = Emu(int(shape.top or 0))
            cx = Emu(int(shape.width or 0))
            cy = Emu(int(shape.height or 0))
            sp = shape._element
            sp.getparent().remove(sp)
            new_xl = _xl_chart_type(update.chart_type)
            new_shape = slide.shapes.add_chart(new_xl, x, y, cx, cy, cd)
            chart = new_shape.chart
            element_id = element_id if _is_opaque_id(element_id) else _new_element_id()
            new_shape.name = element_id  # preserve / migrate the opaque id onto the recreated shape
            updated_fields.append("chartType")
            if isinstance(update, CategoricalChartUpdate):
                if update.categories is not None:
                    updated_fields.append("categories")
                if update.series is not None:
                    updated_fields.append("series")
            elif update.series is not None:
                updated_fields.append("series")

        elif isinstance(update, CategoricalChartUpdate):
            # Same chart type, possibly new data — replace_data() handles in-place.
            if update.categories is not None or update.series is not None:
                current_data = _extract_chart_data(chart)
                new_categories = (
                    update.categories if update.categories is not None else current_data.get("categories", [])
                )
                if update.series is not None:
                    new_data_full = CategoricalChartData(
                        chartType=update.chart_type,
                        categories=list(new_categories),
                        series=update.series,
                        title=update.title,
                        showLegend=update.show_legend,
                        showDataLabels=update.show_data_labels,
                    )
                else:
                    existing_series_raw = current_data.get("series", [])
                    existing_series = [
                        {"name": s.get("name", ""), "values": s.get("values", [])} for s in existing_series_raw
                    ]
                    new_data_full = CategoricalChartData.model_validate(
                        {
                            "chartType": update.chart_type,
                            "categories": list(new_categories),
                            "series": existing_series,
                            "title": update.title,
                            "showLegend": update.show_legend,
                            "showDataLabels": update.show_data_labels,
                        }
                    )
                _validate_categorical(new_data_full)
                cd = _build_categorical_data(new_data_full)
                chart.replace_data(cd)
                if update.categories is not None:
                    updated_fields.append("categories")
                if update.series is not None:
                    updated_fields.append("series")

        else:  # ScatterChartUpdate, same-variant
            if update.series is not None:
                new_scatter = ScatterChartData(
                    chartType="Scatter",
                    series=update.series,
                    title=update.title,
                    showLegend=update.show_legend,
                    showDataLabels=update.show_data_labels,
                )
                _validate_scatter(new_scatter)
                cd = _build_xy_data(new_scatter)
                chart.replace_data(cd)
                updated_fields.append("series")

        title_explicit_none = update.title is None and "title" in update.model_fields_set
        _apply_display_options(
            chart,
            title=update.title,
            show_legend=update.show_legend,
            show_data_labels=update.show_data_labels,
            title_explicit_none=title_explicit_none,
        )
        if "title" in update.model_fields_set:
            updated_fields.append("title")
        if update.show_legend is not None:
            updated_fields.append("showLegend")
        if update.show_data_labels is not None:
            updated_fields.append("showDataLabels")

    return element_id, updated_fields


# ---------------------------------------------------------------------------
# base64 single-slide package helpers (client round-trip)
# ---------------------------------------------------------------------------


def _encode_prs_b64(prs: Any) -> str:
    """Serialize a ``Presentation`` to a base64-encoded .pptx package string."""
    buf = io.BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _load_prs_from_b64(slide_b64: str) -> Any:
    """Decode a base64 single-slide package into a ``Presentation``. Raises 4002 if invalid."""
    try:
        raw = base64.b64decode(slide_b64, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise ChartEngineError(
            code=ErrorCode.INVALID_PARAM,
            message="Invalid base64 slide package",
        ) from exc
    try:
        return Presentation(io.BytesIO(raw))
    except Exception as exc:  # python-pptx raises PackageNotFoundError / BadZipFile / KeyError
        raise ChartEngineError(
            code=ErrorCode.INVALID_PARAM,
            message=f"Not a valid .pptx package: {exc}",
        ) from exc


def _first_slide(prs: Any) -> Any:
    if len(prs.slides) == 0:
        raise ChartEngineError(
            code=ErrorCode.INVALID_PARAM,
            message="Slide package contains no slides",
        )
    return prs.slides[0]


# ---------------------------------------------------------------------------
# In-memory base64 blocking implementations (open document, OASP 0.3.0)
# ---------------------------------------------------------------------------


def _generate_chart_slide_b64_blocking(
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None,
) -> dict[str, Any]:
    """Build a fresh standalone 1-slide deck containing the chart → base64 (no input package)."""
    prs = Presentation()  # default template, zero slides
    slide = prs.slides.add_slide(prs.slide_layouts[_BLANK_LAYOUT_INDEX])
    chart_shape, element_id = _add_chart_to_slide(prs, slide, chart_data, options)
    result = _insert_result(element_id, 0, chart_data, chart_shape)
    result["slideBase64"] = _encode_prs_b64(prs)
    return result


def _insert_chart_into_slide_b64_blocking(
    slide_b64: str,
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None,
) -> dict[str, Any]:
    """Add a chart onto the single slide of a base64 package → base64."""
    prs = _load_prs_from_b64(slide_b64)
    slide = _first_slide(prs)
    chart_shape, element_id = _add_chart_to_slide(prs, slide, chart_data, options)
    result = _insert_result(element_id, 0, chart_data, chart_shape)
    result["slideBase64"] = _encode_prs_b64(prs)
    return result


def _get_chart_from_slide_b64_blocking(slide_b64: str, element_id: str) -> dict[str, Any]:
    """Read a chart's data from a base64 single-slide package."""
    prs = _load_prs_from_b64(slide_b64)
    _slide, shape, slide_idx = _find_chart(prs, element_id, 0)
    return {
        "elementId": element_id,
        "slideIndex": slide_idx,
        "chart": _extract_chart_data(shape.chart),
        **_geometry_dict(shape),
    }


def _update_chart_in_slide_b64_blocking(
    slide_b64: str,
    element_id: str,
    update: CategoricalChartUpdate | ScatterChartUpdate,
) -> dict[str, Any]:
    """Apply a partial update to a chart in a base64 single-slide package → base64."""
    prs = _load_prs_from_b64(slide_b64)
    slide, shape, _slide_idx = _find_chart(prs, element_id, 0)
    returned_id, updated_fields = _modify_chart_in_prs(slide, shape, element_id, update)
    return {
        "elementId": returned_id,
        "chartType": update.chart_type,
        "updatedFields": updated_fields,
        "slideBase64": _encode_prs_b64(prs),
    }


# ---------------------------------------------------------------------------
# Async wrappers — in-memory base64 (no disk access)
# ---------------------------------------------------------------------------


async def generate_chart_slide_base64(
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None = None,
) -> dict[str, Any]:
    """Generate a standalone single-slide .pptx containing ``chart_data`` → base64.

    Used by the open-document "insert → new page" path (no existing page needed).
    Returns ``{slideBase64, elementId, slideIndex, chartType, seriesCount, left, top, width, height}``.
    """
    _validate_chart_data(chart_data)
    return await asyncio.to_thread(_generate_chart_slide_b64_blocking, chart_data, options)


async def insert_chart_into_slide_base64(
    slide_base64: str,
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None = None,
) -> dict[str, Any]:
    """Add ``chart_data`` onto the single slide of ``slide_base64`` → base64.

    Used by the open-document "insert → existing page" path (the Add-In exports the
    live slide via ``Slide.exportAsBase64`` and re-inserts the returned package).
    """
    _validate_chart_data(chart_data)
    return await asyncio.to_thread(_insert_chart_into_slide_b64_blocking, slide_base64, chart_data, options)


async def get_chart_from_slide_base64(slide_base64: str, element_id: str) -> dict[str, Any]:
    """Read a chart's data from a base64 single-slide package (open-document get)."""
    return await asyncio.to_thread(_get_chart_from_slide_b64_blocking, slide_base64, element_id)


async def update_chart_in_slide_base64(
    slide_base64: str,
    element_id: str,
    update: CategoricalChartUpdate | ScatterChartUpdate,
) -> dict[str, Any]:
    """Apply a partial update to a chart in a base64 single-slide package → base64 (open-document update)."""
    return await asyncio.to_thread(_update_chart_in_slide_b64_blocking, slide_base64, element_id, update)
