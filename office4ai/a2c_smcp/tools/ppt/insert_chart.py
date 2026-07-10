"""ppt_insert_chart MCP Tool (OASP /ppt Draft — Server OOXML, connected-only)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.resources.per_file_window import affected_window_uris
from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import DocumentStatus
from office4ai.environment.workspace.dtos.ppt import (
    CategoricalChartData,
    ChartInsertOptions,
    ScatterChartData,
)
from office4ai.environment.workspace.services import chart_router
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager


class PptInsertChartInput(BaseModel):
    """MCP 输入模型：在打开的文档中插入图表（Server 端 OOXML，连接态客户端往返）。"""

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
        description="Geometry & target slide (optional). Defaults: slideIndex=0, 480x320pt centered.",
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
            "(OASP /ppt Draft — all chart OOXML is built Server-side). Requires the target "
            "document OPEN in PowerPoint via the Add-In: the Server applies the chart to the live "
            "slide through a client round-trip. For a CLOSED .pptx this returns 3003 — build the "
            "chart offline with the authoring pipeline (office_run_script + python-pptx) instead. "
            "While the Add-In round-trip handler is not yet available a CONNECTED call may also "
            "return 3003 (close the document and use office_run_script); this lifts automatically "
            "once it ships. Expect >1s latency. The live round-trip updates the open document in "
            "place (an MCP resource_updated notification is fired to /ppt subscribers). "
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
        """Connected-only (#68): CONNECTED → live client round-trip; DISCONNECTED → refuse."""
        try:
            validated = self.validate_input(arguments, PptInsertChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        # Offline chart work is served by the authoring pipeline, not this tool (Path A removed, #68).
        if self.workspace.get_document_status(document_uri) != DocumentStatus.CONNECTED:
            return {"success": False, "error": chart_router.OFFLINE_MESSAGE}
        async with document_lock_manager.acquire(document_uri):
            try:
                result_data = await chart_router.insert_chart_path_b(
                    self.workspace, document_uri, validated.chart, validated.options
                )
            except chart_router.PathBUnavailable:
                # Reactive degradation: live round-trip not available yet → offline-authoring guidance.
                return {"success": False, "error": chart_router.DEGRADE_MESSAGE}
            except ChartEngineError as e:
                return {"success": False, "error": str(e)}
            except Exception as e:  # noqa: BLE001 - surface unexpected errors as 3004
                return {"success": False, "error": f"3004: {e}"}

        # The live round-trip already updated the open document → no reopen needed.
        result_data["requiresReload"] = False
        self.workspace.update_last_activity(
            document_uri=document_uri,
            tool_name=self.name,
            result_data=result_data,
        )
        # 通知该文件的 per-file 窗口（内容变了）；不通知根索引——插图表未改变已连接文件集。
        self.workspace.notify_resource_updated(affected_window_uris("/ppt", document_uri))
        return {"success": True, "data": result_data}
