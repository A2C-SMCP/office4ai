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
from office4ai.environment.workspace.services import chart_engine
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager

# Empirically verified (manual_tests/ppt/test_chart_e2e.py --mode conflict):
# when the Add-In holds the .pptx open in PowerPoint, the next save() flushes
# PowerPoint's in-memory model to disk and SILENTLY OVERWRITES any chart this
# tool just wrote. Until OASP defines a ppt:notify:reload event, refuse the
# write up-front and surface a clear error so the LLM can prompt the user.
_CONNECTED_REJECT_MESSAGE = (
    "3003: Document is currently open in PowerPoint via the Add-In. "
    "Server-side OOXML chart writes will be silently overwritten by the next "
    "PowerPoint save (verified empirically — see docs/manual_tests/"
    "ppt_chart_v0.2.0.md). Please ask the user to close the document in "
    "PowerPoint, then retry. A future ppt:notify:reload OASP event will lift "
    "this restriction."
)


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
            "(OASP /ppt Draft, Server-side OOXML path — bypasses Office.js). "
            "REFUSES with 3003 if the document is currently open in PowerPoint via the Add-In — "
            "PowerPoint's in-memory model would silently overwrite the chart on the next save. "
            "When 3003 is returned, ask the user to close the document in PowerPoint, then retry. "
            "Expect >1s latency. After success the document needs to be reopened to render the new chart "
            "(an MCP resource_updated notification is fired to /ppt subscribers). "
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
        """Override to route through the OOXML chart engine instead of Socket.IO."""
        try:
            validated = self.validate_input(arguments, PptInsertChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        if self.workspace.get_document_status(document_uri) == DocumentStatus.CONNECTED:
            return {"success": False, "error": _CONNECTED_REJECT_MESSAGE}
        async with document_lock_manager.acquire(document_uri):
            try:
                result_data = await chart_engine.insert_chart(
                    document_uri=document_uri,
                    chart_data=validated.chart,
                    options=validated.options,
                )
            except ChartEngineError as e:
                return {"success": False, "error": str(e)}
            except Exception as e:  # noqa: BLE001 - surface unexpected I/O errors as 3004
                return {"success": False, "error": f"3004: {e}"}

        result_data["requiresReload"] = True
        self.workspace.update_last_activity(
            document_uri=document_uri,
            tool_name=self.name,
            result_data=result_data,
        )
        self.workspace.notify_resource_updated(["window://office4ai/ppt", "window://office4ai"])
        return {"success": True, "data": result_data}
