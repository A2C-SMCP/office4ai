"""
OOXML chart engine for OASP /ppt chart events (insert / get / update).

PowerPoint Office.js does not expose chart creation or data-update APIs (see
office-js#5463). The OASP protocol therefore routes ppt:insert:chart /
ppt:get:chart / ppt:update:chart through this Server-side engine, which mutates
the .pptx OOXML directly via ``python-pptx``.

Operational constraints (mirrored from the OASP events-ppt admonition):

- The Add-In must ``save()`` the document before calling — unsaved client edits
  will be overwritten when this engine rewrites the .pptx on disk.
- Concurrent chart calls on the same document must be serialized (see
  ``document_lock.DocumentLockManager``); two simultaneous writes corrupt the
  OOXML package.
- Latency is dominated by disk I/O and python-pptx parsing — typically >1s on
  multi-MB decks.

Element identity
================
We expose ``elementId = "chart-<slide_index>-<shape_id>"`` where ``shape_id``
is the python-pptx integer (== OOXML ``<p:nvSpPr><p:cNvPr id="...">``). The
shape id is only unique within a slide, so we prefix the slide index to make
the wire identifier globally addressable. Backwards-compatible parsing also
accepts the legacy ``"chart-<shape_id>"`` form (best-effort lookup across all
slides).
"""

from __future__ import annotations

import asyncio
import math
import urllib.parse
from dataclasses import dataclass
from pathlib import Path
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
# URI ↔ filesystem path
# ---------------------------------------------------------------------------


def _uri_to_path(document_uri: str) -> Path:
    """Convert a ``file://`` URI to a local Path. Raises 3001 if invalid."""
    parsed = urllib.parse.urlparse(document_uri)
    if parsed.scheme != "file":
        raise ChartEngineError(
            code=ErrorCode.DOCUMENT_NOT_FOUND,
            message=f"Only file:// URIs are supported, got: {document_uri!r}",
        )
    # urllib gives us an unquoted path with a leading slash even on macOS/Linux.
    raw = urllib.parse.unquote(parsed.path)
    path = Path(raw)
    if not path.exists():
        raise ChartEngineError(
            code=ErrorCode.DOCUMENT_NOT_FOUND,
            message=f"Document not found: {path}",
        )
    return path


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


# ---------------------------------------------------------------------------
# python-pptx helpers — chart construction / extraction / mutation
# ---------------------------------------------------------------------------

_DEFAULT_WIDTH_PT = 480.0
_DEFAULT_HEIGHT_PT = 320.0


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


def _parse_element_id(element_id: str) -> tuple[int | None, int]:
    """Parse ``chart-<slide>-<shape>`` (preferred) or legacy ``chart-<shape>``.

    Returns ``(slide_index_or_None, shape_id)``. Raises 3010 on malformed input.
    """
    if not element_id.startswith("chart-"):
        raise ChartEngineError(
            code=ErrorCode.ELEMENT_NOT_FOUND,
            message=f"Invalid chart elementId format: {element_id!r}",
        )
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


def _format_element_id(slide_index: int, shape_id: int) -> str:
    return f"chart-{slide_index}-{shape_id}"


def _find_chart(prs: Any, element_id: str, hint_index: int | None = None) -> tuple[Any, Any, int]:
    """Return (slide, chart_shape, slide_index) for the chart with this elementId.

    Raises 3010 ELEMENT_NOT_FOUND if no chart shape matches.
    """
    parsed_slide, shape_id = _parse_element_id(element_id)
    slides = list(prs.slides)
    target_slide = parsed_slide if parsed_slide is not None else hint_index

    # Search the qualified slide first, then fall back to the rest.
    if target_slide is not None and 0 <= target_slide < len(slides):
        slides_iter = [slides[target_slide]] + [s for i, s in enumerate(slides) if i != target_slide]
    else:
        slides_iter = slides

    for slide in slides_iter:
        for shape in slide.shapes:
            if not getattr(shape, "has_chart", False):
                continue
            if int(shape.shape_id) == shape_id:
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


# ---------------------------------------------------------------------------
# Public engine functions
# ---------------------------------------------------------------------------


def _insert_chart_blocking(
    document_path: Path,
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None,
) -> dict[str, Any]:
    prs = Presentation(str(document_path))

    slide_index_target = options.slide_index if options and options.slide_index is not None else None
    if slide_index_target is None:
        slide_index_target = 0
    if slide_index_target < 0 or slide_index_target >= len(prs.slides):
        raise ChartEngineError(
            code=ErrorCode.INVALID_PARAM,
            message=(f"slideIndex {slide_index_target} out of range [0, {len(prs.slides) - 1}]"),
        )
    slide = prs.slides[slide_index_target]

    x, y, cx, cy = _resolve_geometry(prs, options)

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

    prs.save(str(document_path))

    element_id = _format_element_id(slide_index_target, int(chart_shape.shape_id))
    return {
        "elementId": element_id,
        "slideIndex": slide_index_target,
        "chartType": chart_data.chart_type,
        "seriesCount": len(chart_data.series),
        "left": Emu(int(chart_shape.left or 0)).pt,
        "top": Emu(int(chart_shape.top or 0)).pt,
        "width": Emu(int(chart_shape.width or 0)).pt,
        "height": Emu(int(chart_shape.height or 0)).pt,
    }


def _get_chart_blocking(document_path: Path, element_id: str, hint_index: int | None) -> dict[str, Any]:
    prs = Presentation(str(document_path))
    _slide, shape, slide_idx = _find_chart(prs, element_id, hint_index)
    chart = shape.chart
    return {
        "elementId": element_id,
        "slideIndex": slide_idx,
        "chart": _extract_chart_data(chart),
        "left": Emu(int(shape.left or 0)).pt,
        "top": Emu(int(shape.top or 0)).pt,
        "width": Emu(int(shape.width or 0)).pt,
        "height": Emu(int(shape.height or 0)).pt,
    }


def _is_categorical(chart_type: str) -> bool:
    return chart_type in _CATEGORICAL_TYPES


def _update_chart_blocking(
    document_path: Path,
    element_id: str,
    update: CategoricalChartUpdate | ScatterChartUpdate,
) -> dict[str, Any]:
    prs = Presentation(str(document_path))
    slide, shape, slide_idx = _find_chart(prs, element_id, None)
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
        element_id = _format_element_id(slide_idx, int(new_shape.shape_id))
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
            element_id = _format_element_id(slide_idx, int(new_shape.shape_id))
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

    prs.save(str(document_path))

    return {
        "elementId": element_id,
        "chartType": target_type,
        "updatedFields": updated_fields,
    }


# ---------------------------------------------------------------------------
# Async wrappers (offload blocking I/O to a worker thread)
# ---------------------------------------------------------------------------


async def insert_chart(
    document_uri: str,
    chart_data: CategoricalChartData | ScatterChartData,
    options: ChartInsertOptions | None,
) -> dict[str, Any]:
    """Insert a new chart into the .pptx at ``document_uri``.

    Raises ``ChartEngineError`` (carrying an OASP code) on validation / IO failure.
    """
    path = _uri_to_path(document_uri)
    if isinstance(chart_data, CategoricalChartData):
        _validate_categorical(chart_data)
    else:
        _validate_scatter(chart_data)
    return await asyncio.to_thread(_insert_chart_blocking, path, chart_data, options)


async def get_chart(document_uri: str, element_id: str, slide_index: int | None = None) -> dict[str, Any]:
    """Read an existing chart's data into the OASP ``ChartData`` wire shape."""
    path = _uri_to_path(document_uri)
    return await asyncio.to_thread(_get_chart_blocking, path, element_id, slide_index)


async def update_chart(
    document_uri: str,
    element_id: str,
    update: CategoricalChartUpdate | ScatterChartUpdate,
) -> dict[str, Any]:
    """Apply a partial update to an existing chart."""
    path = _uri_to_path(document_uri)
    return await asyncio.to_thread(_update_chart_blocking, path, element_id, update)
