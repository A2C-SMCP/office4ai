"""ppt_insert_chart MCP Tool (OASP /ppt Draft, v0.2.0 — Server OOXML)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import DocumentStatus
from office4ai.environment.workspace.dtos.ppt import (
    CategoricalChartData,
    ChartInsertOptions,
    ScatterChartData,
)
from office4ai.environment.workspace.services import chart_engine, chart_router
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager


class PptInsertChartInput(BaseModel):
    """MCP 输入模型：在指定 .pptx 中插入图表（Server 端 OOXML 离线处理）。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/deck.pptx)")
    chart: CategoricalChartData | ScatterChartData = Field(
        ...,
        description=(
            "Chart payload (discriminated by chartType). "
            "Categorical (ColumnClustered/Bar/Line/Pie/...): requires categories + series.values. "
            "Scatter: requires series with points (x, y)."
        ),
    )
    options: ChartInsertOptions | None = Field(
        default=None,
        description="Geometry & target slide (optional). Defaults: slideIndex=current, 480x320pt centered.",
    )


class PptInsertChartTool(BaseTool):
    """Insert a chart on a slide via OOXML rewriting (OASP /ppt Draft)."""

    @property
    def name(self) -> str:
        return "ppt_insert_chart"

    @property
    def description(self) -> str:
        return (
            "Insert a chart (column/bar/line/pie/scatter/...) into a PowerPoint slide "
            "(OASP /ppt Draft, dual-path — all chart OOXML is built Server-side). "
            "When the document is CLOSED, the Server edits the .pptx on disk; when it is OPEN in "
            "PowerPoint via the Add-In, the Server applies the chart to the live slide through a "
            "client round-trip. While the Add-In round-trip handler is not yet available it may "
            "return 3003 (close the document, then retry); this lifts automatically once it ships. "
            "Expect >1s latency. After an on-disk write the document must be reopened to render the new "
            "chart (an MCP resource_updated notification is fired to /ppt subscribers). "
            "ChartType discriminates the schema: categorical types (ColumnClustered/ColumnStacked/"
            "BarClustered/Line/LineMarkers/Pie/Doughnut/Area/Radar) require categories + "
            "series[].values; Scatter requires series[].points = [{x, y}]. "
            "Mismatched dimensions trigger 3015 INVALID_CHART_DATA."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptInsertChartInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "insert:chart"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptInsertChartInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Dual-path route (#15): CONNECTED → client round-trip; DISCONNECTED → on-disk OOXML."""
        try:
            validated = self.validate_input(arguments, PptInsertChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        connected = self.workspace.get_document_status(document_uri) == DocumentStatus.CONNECTED
        async with document_lock_manager.acquire(document_uri):
            try:
                if connected:
                    result_data = await chart_router.insert_chart_path_b(
                        self.workspace, document_uri, validated.chart, validated.options
                    )
                else:
                    result_data = await chart_engine.insert_chart(
                        document_uri=document_uri,
                        chart_data=validated.chart,
                        options=validated.options,
                    )
            except chart_router.PathBUnavailable:
                # Reactive degradation: path B not available yet → flipped 3003 guidance.
                return {"success": False, "error": chart_router.DEGRADE_MESSAGE_WRITE}
            except ChartEngineError as e:
                return {"success": False, "error": str(e)}
            except Exception as e:  # noqa: BLE001 - surface unexpected I/O errors as 3004
                return {"success": False, "error": f"3004: {e}"}

        # On-disk writes change the file (reopen to render); live round-trips already updated it.
        result_data["requiresReload"] = not connected
        self.workspace.update_last_activity(
            document_uri=document_uri,
            tool_name=self.name,
            result_data=result_data,
        )
        self.workspace.notify_resource_updated(["window://office4ai/ppt", "window://office4ai"])
        return {"success": True, "data": result_data}
