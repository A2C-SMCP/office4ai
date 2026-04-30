"""Unit tests for PPT chart DTOs (OASP /ppt Draft, v0.2.0)."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import ErrorCode, request_registry
from office4ai.environment.workspace.dtos.ppt import (
    CategoricalChartData,
    CategoricalChartUpdate,
    CategoricalSeries,
    ChartInsertOptions,
    PptGetChartRequest,
    PptInsertChartRequest,
    PptUpdateChartRequest,
    ScatterChartData,
    ScatterChartUpdate,
    ScatterPoint,
    ScatterSeries,
)


class TestErrorCodeRegistry:
    def test_invalid_chart_data_code(self) -> None:
        assert ErrorCode.INVALID_CHART_DATA == "3015"

    def test_request_registry_has_chart_events(self) -> None:
        assert request_registry.contains("ppt:insert:chart")
        assert request_registry.contains("ppt:get:chart")
        assert request_registry.contains("ppt:update:chart")


class TestCategoricalChartData:
    def test_valid_categorical_data(self) -> None:
        data = CategoricalChartData(
            chartType="ColumnClustered",
            categories=["A", "B"],
            series=[CategoricalSeries(name="x", values=[1, 2])],
            title="t",
            showLegend=True,
            showDataLabels=False,
        )
        assert data.chart_type == "ColumnClustered"
        assert data.title == "t"
        # populate_by_name accepts snake_case alongside alias
        from_snake = CategoricalChartData.model_validate(
            {
                "chart_type": "Pie",
                "categories": ["A"],
                "series": [{"name": "x", "values": [1]}],
            }
        )
        assert from_snake.chart_type == "Pie"

    def test_to_payload_emits_camel_case(self) -> None:
        data = CategoricalChartData(
            chartType="Line",
            categories=["A"],
            series=[CategoricalSeries(name="x", values=[1])],
            showLegend=True,
        )
        dumped = data.model_dump(by_alias=True, exclude_none=True)
        assert dumped["chartType"] == "Line"
        assert dumped["showLegend"] is True
        assert "show_legend" not in dumped

    def test_categorical_series_min_one(self) -> None:
        with pytest.raises(ValidationError):
            CategoricalChartData(chartType="Pie", categories=["A"], series=[])


class TestScatterChartData:
    def test_valid_scatter_data(self) -> None:
        data = ScatterChartData(
            chartType="Scatter",
            series=[ScatterSeries(name="x", points=[ScatterPoint(x=1, y=2)])],
        )
        assert data.chart_type == "Scatter"

    def test_scatter_points_min_one(self) -> None:
        with pytest.raises(ValidationError):
            ScatterChartData(
                chartType="Scatter",
                series=[ScatterSeries(name="x", points=[])],
            )


class TestPptInsertChartRequest:
    def test_categorical_request_round_trip(self) -> None:
        req = PptInsertChartRequest.build(
            document_uri="file:///t.pptx",
            chart={
                "chartType": "BarClustered",
                "categories": ["A"],
                "series": [{"name": "s", "values": [1]}],
            },
        )
        payload = req.to_payload()
        assert payload["documentUri"] == "file:///t.pptx"
        assert payload["chart"]["chartType"] == "BarClustered"

    def test_scatter_request_discriminator_routes_correctly(self) -> None:
        req = PptInsertChartRequest.build(
            document_uri="file:///t.pptx",
            chart={
                "chartType": "Scatter",
                "series": [{"name": "s", "points": [{"x": 1, "y": 2}]}],
            },
        )
        payload = req.to_payload()
        assert payload["chart"]["chartType"] == "Scatter"
        assert "categories" not in payload["chart"]

    def test_invalid_chart_type_rejected(self) -> None:
        with pytest.raises(ValidationError):
            PptInsertChartRequest.build(
                document_uri="file:///t.pptx",
                chart={"chartType": "Bogus", "categories": [], "series": []},
            )

    def test_options_optional(self) -> None:
        req = PptInsertChartRequest.build(
            document_uri="file:///t.pptx",
            chart={
                "chartType": "Pie",
                "categories": ["A", "B"],
                "series": [{"name": "s", "values": [1, 2]}],
            },
        )
        assert req.options is None


class TestPptGetChartRequest:
    def test_minimal_request(self) -> None:
        req = PptGetChartRequest.build(document_uri="file:///t.pptx", element_id="chart-3")
        payload = req.to_payload()
        assert payload["elementId"] == "chart-3"
        assert "slideIndex" not in payload  # excluded when None

    def test_with_slide_hint(self) -> None:
        req = PptGetChartRequest.build(document_uri="file:///t.pptx", element_id="chart-3", slide_index=2)
        payload = req.to_payload()
        assert payload["slideIndex"] == 2


class TestPptUpdateChartRequest:
    def test_categorical_update_partial(self) -> None:
        req = PptUpdateChartRequest.build(
            document_uri="file:///t.pptx",
            element_id="chart-3",
            chart={"chartType": "ColumnClustered", "title": "New"},
        )
        payload = req.to_payload()
        assert payload["chart"]["chartType"] == "ColumnClustered"
        assert payload["chart"]["title"] == "New"

    def test_cross_variant_update(self) -> None:
        req = PptUpdateChartRequest.build(
            document_uri="file:///t.pptx",
            element_id="chart-3",
            chart={"chartType": "Scatter", "series": [{"name": "x", "points": [{"x": 1, "y": 2}]}]},
        )
        payload = req.to_payload()
        assert payload["chart"]["chartType"] == "Scatter"


class TestChartUpdateDiscriminator:
    def test_categorical_update_constructs(self) -> None:
        upd = CategoricalChartUpdate(chartType="Pie", title="x")
        assert upd.chart_type == "Pie"

    def test_scatter_update_constructs(self) -> None:
        upd = ScatterChartUpdate(chartType="Scatter", title="x")
        assert upd.chart_type == "Scatter"


class TestChartInsertOptionsDefaults:
    def test_all_optional(self) -> None:
        opts = ChartInsertOptions()
        assert opts.slide_index is None
        assert opts.left is None
        assert opts.width is None

    def test_validation_positive_dimensions(self) -> None:
        with pytest.raises(ValidationError):
            ChartInsertOptions(width=-10)
        with pytest.raises(ValidationError):
            ChartInsertOptions(height=0)

    def test_payload_uses_camel_case(self) -> None:
        opts = ChartInsertOptions(slideIndex=1, width=400)
        dumped = opts.model_dump(by_alias=True, exclude_none=True)
        assert dumped == {"slideIndex": 1, "width": 400}
