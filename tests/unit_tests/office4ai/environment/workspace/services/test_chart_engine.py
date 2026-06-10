"""Unit tests for the OOXML chart engine (insert / get / update)."""

from __future__ import annotations

import base64
import io
import math
from pathlib import Path
from typing import Any

import pytest
from pptx import Presentation

from office4ai.environment.workspace.dtos.common import ErrorCode
from office4ai.environment.workspace.dtos.ppt import (
    CategoricalChartData,
    CategoricalChartUpdate,
    CategoricalSeries,
    ChartInsertOptions,
    ScatterChartData,
    ScatterChartUpdate,
    ScatterPoint,
    ScatterSeries,
)
from office4ai.environment.workspace.services import chart_engine
from office4ai.environment.workspace.services.chart_engine import ChartEngineError


@pytest.fixture
def empty_pptx(tmp_path: Path) -> Path:
    """A blank 2-slide deck."""
    path = tmp_path / "deck.pptx"
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[5])
    prs.slides.add_slide(prs.slide_layouts[5])
    prs.save(str(path))
    return path


# ---------------------------------------------------------------------------
# insert_chart
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_insert_categorical_chart_returns_element_id(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["Jan", "Feb", "Mar"],
        series=[CategoricalSeries(name="Rev", values=[100, 200, 150])],
        title="Q1",
    )
    result = await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0))
    assert result["elementId"].startswith("oasp-chart-")
    assert result["chartType"] == "ColumnClustered"
    assert result["seriesCount"] == 1
    assert result["slideIndex"] == 0


@pytest.mark.asyncio
async def test_insert_uses_explicit_geometry(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="Pie",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    result = await chart_engine.insert_chart(
        empty_pptx.as_uri(),
        chart,
        ChartInsertOptions(slideIndex=1, left=72, top=36, width=400, height=300),
    )
    assert math.isclose(result["left"], 72, abs_tol=0.5)
    assert math.isclose(result["top"], 36, abs_tol=0.5)
    assert math.isclose(result["width"], 400, abs_tol=0.5)
    assert math.isclose(result["height"], 300, abs_tol=0.5)
    assert result["slideIndex"] == 1


@pytest.mark.asyncio
async def test_insert_scatter_chart(empty_pptx: Path) -> None:
    chart = ScatterChartData(
        chartType="Scatter",
        series=[
            ScatterSeries(
                name="ads",
                points=[ScatterPoint(x=1, y=2), ScatterPoint(x=3, y=4)],
            )
        ],
    )
    result = await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0))
    assert result["chartType"] == "Scatter"
    assert result["seriesCount"] == 1


@pytest.mark.asyncio
async def test_insert_categorical_dimension_mismatch_raises_3015(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="Pie",
        categories=["A", "B", "C"],
        series=[CategoricalSeries(name="x", values=[1, 2])],  # 2 != 3
    )
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.insert_chart(empty_pptx.as_uri(), chart, None)
    assert exc.value.code == ErrorCode.INVALID_CHART_DATA


@pytest.mark.asyncio
async def test_insert_scatter_non_finite_raises_3015(empty_pptx: Path) -> None:
    chart = ScatterChartData(
        chartType="Scatter",
        series=[ScatterSeries(name="s", points=[ScatterPoint(x=float("inf"), y=1)])],
    )
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.insert_chart(empty_pptx.as_uri(), chart, None)
    assert exc.value.code == ErrorCode.INVALID_CHART_DATA


@pytest.mark.asyncio
async def test_insert_unknown_document_uri_raises_3001() -> None:
    chart = CategoricalChartData(
        chartType="Line",
        categories=["A"],
        series=[CategoricalSeries(name="s", values=[1])],
    )
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.insert_chart("file:///nonexistent/missing.pptx", chart, None)
    assert exc.value.code == ErrorCode.DOCUMENT_NOT_FOUND


@pytest.mark.asyncio
async def test_insert_slide_index_out_of_range_raises_4002(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="Line",
        categories=["A"],
        series=[CategoricalSeries(name="s", values=[1])],
    )
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=99))
    assert exc.value.code == ErrorCode.INVALID_PARAM


# ---------------------------------------------------------------------------
# get_chart
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_get_categorical_round_trip(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["Q1", "Q2"],
        series=[
            CategoricalSeries(name="Rev", values=[100, 200]),
            CategoricalSeries(name="Cost", values=[80, 90]),
        ],
        title="Year",
        showLegend=True,
    )
    inserted = await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0))
    eid = inserted["elementId"]

    got = await chart_engine.get_chart(empty_pptx.as_uri(), eid)
    assert got["elementId"] == eid
    assert got["chart"]["chartType"] == "ColumnClustered"
    assert got["chart"]["categories"] == ["Q1", "Q2"]
    assert got["chart"]["title"] == "Year"
    assert len(got["chart"]["series"]) == 2
    assert got["chart"]["series"][0]["name"] == "Rev"
    assert got["chart"]["series"][0]["values"] == [100.0, 200.0]


@pytest.mark.asyncio
async def test_get_scatter_round_trip(empty_pptx: Path) -> None:
    chart = ScatterChartData(
        chartType="Scatter",
        series=[
            ScatterSeries(
                name="ads",
                points=[ScatterPoint(x=1, y=10), ScatterPoint(x=2, y=20), ScatterPoint(x=3, y=40)],
            )
        ],
        title="adv-vs-sales",
    )
    inserted = await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0))
    got = await chart_engine.get_chart(empty_pptx.as_uri(), inserted["elementId"])
    assert got["chart"]["chartType"] == "Scatter"
    assert got["chart"]["title"] == "adv-vs-sales"
    pts = got["chart"]["series"][0]["points"]
    assert pts == [{"x": 1.0, "y": 10.0}, {"x": 2.0, "y": 20.0}, {"x": 3.0, "y": 40.0}]


@pytest.mark.asyncio
async def test_get_unknown_element_raises_3010(empty_pptx: Path) -> None:
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.get_chart(empty_pptx.as_uri(), "chart-99999")
    assert exc.value.code == ErrorCode.ELEMENT_NOT_FOUND


@pytest.mark.asyncio
async def test_get_invalid_element_id_format_raises_3010(empty_pptx: Path) -> None:
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.get_chart(empty_pptx.as_uri(), "shape-1")
    assert exc.value.code == ErrorCode.ELEMENT_NOT_FOUND


# ---------------------------------------------------------------------------
# update_chart
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_update_title_only(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
        title="Original",
    )
    eid = (await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0)))["elementId"]

    upd = CategoricalChartUpdate(chartType="ColumnClustered", title="Revised")
    result = await chart_engine.update_chart(empty_pptx.as_uri(), eid, upd)
    assert "title" in result["updatedFields"]

    got = await chart_engine.get_chart(empty_pptx.as_uri(), eid)
    assert got["chart"]["title"] == "Revised"


@pytest.mark.asyncio
async def test_update_replace_categorical_series(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    eid = (await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0)))["elementId"]

    upd = CategoricalChartUpdate(
        chartType="ColumnClustered",
        categories=["X", "Y", "Z"],
        series=[CategoricalSeries(name="t", values=[10, 20, 30])],
    )
    result = await chart_engine.update_chart(empty_pptx.as_uri(), eid, upd)
    assert "series" in result["updatedFields"]
    assert "categories" in result["updatedFields"]

    got = await chart_engine.get_chart(empty_pptx.as_uri(), eid)
    assert got["chart"]["categories"] == ["X", "Y", "Z"]
    assert got["chart"]["series"][0]["values"] == [10.0, 20.0, 30.0]


@pytest.mark.asyncio
async def test_update_categorical_dimension_mismatch_raises_3015(empty_pptx: Path) -> None:
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    eid = (await chart_engine.insert_chart(empty_pptx.as_uri(), chart, None))["elementId"]

    upd = CategoricalChartUpdate(
        chartType="ColumnClustered",
        categories=["X", "Y", "Z"],
        series=[CategoricalSeries(name="t", values=[10, 20])],  # 2 != 3
    )
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.update_chart(empty_pptx.as_uri(), eid, upd)
    assert exc.value.code == ErrorCode.INVALID_CHART_DATA


@pytest.mark.asyncio
async def test_cross_variant_categorical_to_scatter(empty_pptx: Path) -> None:
    cat = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    eid = (await chart_engine.insert_chart(empty_pptx.as_uri(), cat, ChartInsertOptions(slideIndex=0)))["elementId"]

    upd = ScatterChartUpdate(
        chartType="Scatter",
        series=[ScatterSeries(name="conv", points=[ScatterPoint(x=1, y=10), ScatterPoint(x=2, y=20)])],
    )
    result = await chart_engine.update_chart(empty_pptx.as_uri(), eid, upd)
    assert result["chartType"] == "Scatter"
    assert "chartType" in result["updatedFields"]
    assert "series" in result["updatedFields"]
    new_eid = result["elementId"]

    got = await chart_engine.get_chart(empty_pptx.as_uri(), new_eid)
    assert got["chart"]["chartType"] == "Scatter"


@pytest.mark.asyncio
async def test_cross_variant_scatter_to_categorical_requires_categories_and_series(
    empty_pptx: Path,
) -> None:
    sc = ScatterChartData(
        chartType="Scatter",
        series=[ScatterSeries(name="ads", points=[ScatterPoint(x=1, y=2)])],
    )
    eid = (await chart_engine.insert_chart(empty_pptx.as_uri(), sc, ChartInsertOptions(slideIndex=0)))["elementId"]

    # Only chartType — no series, no categories — should fail with 3015.
    upd = CategoricalChartUpdate(chartType="ColumnClustered")
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.update_chart(empty_pptx.as_uri(), eid, upd)
    assert exc.value.code == ErrorCode.INVALID_CHART_DATA


@pytest.mark.asyncio
async def test_update_unknown_element_raises_3010(empty_pptx: Path) -> None:
    upd = CategoricalChartUpdate(chartType="ColumnClustered", title="x")
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.update_chart(empty_pptx.as_uri(), "chart-99999", upd)
    assert exc.value.code == ErrorCode.ELEMENT_NOT_FOUND


# ---------------------------------------------------------------------------
# elementId opacity (oasp-chart-<uuid> written to cNvPr/@name) — #13
# ---------------------------------------------------------------------------


@pytest.mark.asyncio
async def test_insert_element_id_is_opaque_and_written_to_shape_name(empty_pptx: Path) -> None:
    """Inserted chart's elementId is opaque and persisted as the shape name (cNvPr/@name)."""
    chart = CategoricalChartData(
        chartType="Line",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    eid = (await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0)))["elementId"]
    assert eid.startswith("oasp-chart-")
    # The opaque id must survive save()/reload as the shape's OOXML name.
    prs = Presentation(str(empty_pptx))
    names = [shape.name for slide in prs.slides for shape in slide.shapes if getattr(shape, "has_chart", False)]
    assert eid in names


@pytest.mark.asyncio
async def test_get_relocates_by_opaque_name_not_native_id(empty_pptx: Path) -> None:
    """After a second chart shifts native ids, get still finds the first chart by its opaque name."""
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
        title="first",
    )
    first = (await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0)))["elementId"]
    # Insert a second chart on the same slide (native shape ids advance).
    await chart_engine.insert_chart(empty_pptx.as_uri(), chart, ChartInsertOptions(slideIndex=0))

    got = await chart_engine.get_chart(empty_pptx.as_uri(), first)
    assert got["elementId"] == first
    assert got["chart"]["title"] == "first"


@pytest.mark.asyncio
async def test_legacy_chart_element_id_still_resolves(tmp_path: Path) -> None:
    """A 0.2.0 chart (default python-pptx name, no opaque id) resolves via legacy chart-<slide>-<shape>."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Pt

    path = tmp_path / "legacy.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["A", "B"]
    cd.add_series("s", (1, 2))
    shape: Any = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(10), Pt(10), Pt(300), Pt(200), cd)  # type: ignore[arg-type]
    legacy_eid = f"chart-0-{int(shape.shape_id)}"  # not written to shape.name
    prs.save(str(path))

    got = await chart_engine.get_chart(path.as_uri(), legacy_eid)
    assert got["chart"]["chartType"] == "ColumnClustered"
    assert got["chart"]["categories"] == ["A", "B"]


@pytest.mark.asyncio
async def test_legacy_chart_migrates_to_opaque_id_on_recreate(tmp_path: Path) -> None:
    """A 0.2.0 (legacy-named) chart updated via a cross-variant recreate is migrated to a fresh opaque id."""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Pt

    path = tmp_path / "legacy.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    cd = CategoryChartData()
    cd.categories = ["A", "B"]
    cd.add_series("s", (1, 2))
    shape: Any = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Pt(10), Pt(10), Pt(300), Pt(200), cd)  # type: ignore[arg-type]
    legacy_eid = f"chart-0-{int(shape.shape_id)}"  # default python-pptx name, NOT an opaque id
    prs.save(str(path))

    upd = ScatterChartUpdate(
        chartType="Scatter",
        series=[ScatterSeries(name="p", points=[ScatterPoint(x=1, y=2), ScatterPoint(x=3, y=4)])],
    )
    result = await chart_engine.update_chart(path.as_uri(), legacy_eid, upd)
    new_eid = result["elementId"]
    assert new_eid.startswith("oasp-chart-")  # migrated off the legacy form
    assert new_eid != legacy_eid

    # The recreated chart now carries the opaque id in its OOXML name and is located by it.
    got = await chart_engine.get_chart(path.as_uri(), new_eid)
    assert got["chart"]["chartType"] == "Scatter"
    prs_reloaded = Presentation(str(path))
    names = [s.name for sl in prs_reloaded.slides for s in sl.shapes if getattr(s, "has_chart", False)]
    assert new_eid in names


@pytest.mark.asyncio
async def test_update_preserves_opaque_element_id_across_recreate(empty_pptx: Path) -> None:
    """Cross-variant update recreates the shape but keeps the same opaque elementId."""
    cat = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    eid = (await chart_engine.insert_chart(empty_pptx.as_uri(), cat, ChartInsertOptions(slideIndex=0)))["elementId"]

    upd = ScatterChartUpdate(
        chartType="Scatter",
        series=[ScatterSeries(name="p", points=[ScatterPoint(x=1, y=2), ScatterPoint(x=3, y=4)])],
    )
    result = await chart_engine.update_chart(empty_pptx.as_uri(), eid, upd)
    assert result["elementId"] == eid  # stable across delete-and-recreate

    got = await chart_engine.get_chart(empty_pptx.as_uri(), eid)
    assert got["chart"]["chartType"] == "Scatter"


# ---------------------------------------------------------------------------
# Path B — in-memory base64 single-slide packages — #13
# ---------------------------------------------------------------------------


def _blank_slide_base64() -> str:
    """A base64-encoded single-slide .pptx (mimics the Add-In's Slide.exportAsBase64)."""
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])
    buf = io.BytesIO()
    prs.save(buf)
    return base64.b64encode(buf.getvalue()).decode("ascii")


@pytest.mark.asyncio
async def test_generate_standalone_slide_base64_round_trip() -> None:
    """生成独立单页: in-memory deck → base64 → load → parse back to ChartData (no disk)."""
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["Q1", "Q2", "Q3"],
        series=[CategoricalSeries(name="Rev", values=[10, 20, 30])],
        title="Quarterly",
        showLegend=True,
    )
    result = await chart_engine.generate_chart_slide_base64(chart, ChartInsertOptions(width=400, height=300))
    assert result["elementId"].startswith("oasp-chart-")
    assert result["slideIndex"] == 0
    assert result["chartType"] == "ColumnClustered"
    assert result["seriesCount"] == 1
    b64 = result["slideBase64"]

    # The package is a valid, self-contained 1-slide deck.
    prs = Presentation(io.BytesIO(base64.b64decode(b64)))
    assert len(prs.slides) == 1

    # Read the chart back out of the package by its opaque id.
    got = await chart_engine.get_chart_from_slide_base64(b64, result["elementId"])
    assert got["elementId"] == result["elementId"]
    assert got["chart"]["chartType"] == "ColumnClustered"
    assert got["chart"]["categories"] == ["Q1", "Q2", "Q3"]
    assert got["chart"]["title"] == "Quarterly"
    assert got["chart"]["series"][0]["values"] == [10.0, 20.0, 30.0]


@pytest.mark.asyncio
async def test_generate_standalone_scatter_slide_base64() -> None:
    chart = ScatterChartData(
        chartType="Scatter",
        series=[ScatterSeries(name="ads", points=[ScatterPoint(x=1, y=10), ScatterPoint(x=2, y=20)])],
        title="xy",
    )
    result = await chart_engine.generate_chart_slide_base64(chart)
    got = await chart_engine.get_chart_from_slide_base64(result["slideBase64"], result["elementId"])
    assert got["chart"]["chartType"] == "Scatter"
    assert got["chart"]["series"][0]["points"] == [{"x": 1.0, "y": 10.0}, {"x": 2.0, "y": 20.0}]


@pytest.mark.asyncio
async def test_insert_chart_into_existing_slide_base64() -> None:
    """改单页(插现有页): load exported slide base64 → add chart → back to base64."""
    slide_b64 = _blank_slide_base64()
    chart = CategoricalChartData(
        chartType="Pie",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[3, 7])],
    )
    result = await chart_engine.insert_chart_into_slide_base64(slide_b64, chart, ChartInsertOptions(left=20, top=20))
    assert result["elementId"].startswith("oasp-chart-")
    assert "slideBase64" in result

    got = await chart_engine.get_chart_from_slide_base64(result["slideBase64"], result["elementId"])
    assert got["chart"]["chartType"] == "Pie"
    assert got["chart"]["series"][0]["values"] == [3.0, 7.0]


@pytest.mark.asyncio
async def test_update_chart_in_slide_base64_round_trip() -> None:
    """改单页(改图): generate → update title in base64 → re-read; opaque id stays stable."""
    chart = CategoricalChartData(
        chartType="BarClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
        title="Old",
    )
    gen = await chart_engine.generate_chart_slide_base64(chart)
    eid = gen["elementId"]

    upd = CategoricalChartUpdate(chartType="BarClustered", title="New")
    result = await chart_engine.update_chart_in_slide_base64(gen["slideBase64"], eid, upd)
    assert result["elementId"] == eid
    assert "title" in result["updatedFields"]

    got = await chart_engine.get_chart_from_slide_base64(result["slideBase64"], eid)
    assert got["chart"]["title"] == "New"


@pytest.mark.asyncio
async def test_update_data_in_slide_base64() -> None:
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    gen = await chart_engine.generate_chart_slide_base64(chart)
    upd = CategoricalChartUpdate(
        chartType="ColumnClustered",
        categories=["X", "Y", "Z"],
        series=[CategoricalSeries(name="t", values=[10, 20, 30])],
    )
    result = await chart_engine.update_chart_in_slide_base64(gen["slideBase64"], gen["elementId"], upd)
    got = await chart_engine.get_chart_from_slide_base64(result["slideBase64"], gen["elementId"])
    assert got["chart"]["categories"] == ["X", "Y", "Z"]
    assert got["chart"]["series"][0]["values"] == [10.0, 20.0, 30.0]


@pytest.mark.asyncio
async def test_update_cross_variant_in_slide_base64_round_trip() -> None:
    """改单页(改图,跨 variant): base64 deck → recreate Column→Scatter → re-encode → reload locates by same id."""
    chart = CategoricalChartData(
        chartType="ColumnClustered",
        categories=["A", "B"],
        series=[CategoricalSeries(name="s", values=[1, 2])],
    )
    gen = await chart_engine.generate_chart_slide_base64(chart)
    eid = gen["elementId"]

    upd = ScatterChartUpdate(
        chartType="Scatter",
        series=[ScatterSeries(name="p", points=[ScatterPoint(x=1, y=2), ScatterPoint(x=3, y=4)])],
    )
    result = await chart_engine.update_chart_in_slide_base64(gen["slideBase64"], eid, upd)
    assert result["elementId"] == eid  # opaque id preserved across delete-and-recreate in the package
    assert result["chartType"] == "Scatter"

    # The re-encoded package reloads and the same opaque id still locates the (now Scatter) chart.
    got = await chart_engine.get_chart_from_slide_base64(result["slideBase64"], eid)
    assert got["chart"]["chartType"] == "Scatter"
    assert got["chart"]["series"][0]["points"] == [{"x": 1.0, "y": 2.0}, {"x": 3.0, "y": 4.0}]


@pytest.mark.asyncio
async def test_get_from_base64_unknown_element_raises_3010() -> None:
    slide_b64 = _blank_slide_base64()
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.get_chart_from_slide_base64(slide_b64, "oasp-chart-does-not-exist")
    assert exc.value.code == ErrorCode.ELEMENT_NOT_FOUND


@pytest.mark.asyncio
async def test_load_malformed_base64_raises_4002() -> None:
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.get_chart_from_slide_base64("not-valid-base64-@@@", "oasp-chart-x")
    assert exc.value.code == ErrorCode.INVALID_PARAM


@pytest.mark.asyncio
async def test_load_valid_base64_but_not_pptx_raises_4002() -> None:
    junk = base64.b64encode(b"this is not a pptx package").decode("ascii")
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.get_chart_from_slide_base64(junk, "oasp-chart-x")
    assert exc.value.code == ErrorCode.INVALID_PARAM


@pytest.mark.asyncio
async def test_insert_into_base64_with_no_slides_raises_4002() -> None:
    """A package with zero slides cannot host a chart-into-existing-page insert."""
    prs = Presentation()  # default template has zero slides
    buf = io.BytesIO()
    prs.save(buf)
    empty_b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    chart = CategoricalChartData(
        chartType="Line",
        categories=["A"],
        series=[CategoricalSeries(name="s", values=[1])],
    )
    with pytest.raises(ChartEngineError) as exc:
        await chart_engine.insert_chart_into_slide_base64(empty_b64, chart, None)
    assert exc.value.code == ErrorCode.INVALID_PARAM
