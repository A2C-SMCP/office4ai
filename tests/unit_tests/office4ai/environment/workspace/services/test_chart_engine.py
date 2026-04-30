"""Unit tests for the OOXML chart engine (insert / get / update)."""

from __future__ import annotations

import math
from pathlib import Path

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
    assert result["elementId"].startswith("chart-")
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
