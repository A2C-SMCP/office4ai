"""ppt_update_chart MCP Tool (OASP /ppt Draft — Server OOXML, connected-only)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.resources.per_file_window import affected_window_uris
from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import DocumentStatus
from office4ai.environment.workspace.dtos.ppt import (
    CategoricalChartUpdate,
    ScatterChartUpdate,
)
from office4ai.environment.workspace.services import chart_router
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager


class PptUpdateChartInput(BaseModel):
    """MCP 输入模型：更新已存在图表（Server 端 OOXML，连接态客户端往返）。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/deck.pptx)")
    element_id: str | int = Field(
        ...,
        alias="elementId",
        description="Chart element ID (opaque string, e.g. 'oasp-chart-<uuid>')",
    )
    chart: CategoricalChartUpdate | ScatterChartUpdate = Field(
        ...,
        description=(
            "Update payload (discriminated by chartType — REQUIRED). "
            "Same-variant change (e.g. Pie → ColumnClustered): categories/series can be omitted "
            "to keep the originals. Cross-variant change (categorical → Scatter, or vice versa): "
            "MUST supply new series matching the target shape, otherwise 3015 INVALID_CHART_DATA. "
            "Tip: call ppt_get_chart first and reuse its chartType to pick the right variant."
        ),
    )
    slide_index: int | None = Field(
        default=None,
        alias="slideIndex",
        description=(
            "Slide index (0-based) of the chart. REQUIRED — the live round-trip exports that "
            "slide to apply the update. Reuse the slideIndex returned by ppt_get_chart."
        ),
        ge=0,
    )

    model_config = {"populate_by_name": True}

    @field_validator("element_id", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptUpdateChartTool(BaseTool):
    """Update an existing chart's data, type, title or display options (OASP /ppt Draft)."""

    @property
    def name(self) -> str:
        return "ppt_update_chart"

    @property
    def description(self) -> str:
        return (
            "Update an existing PowerPoint chart in-place via OOXML rewriting "
            "(OASP /ppt Draft — all chart OOXML is built Server-side). Requires the target document "
            "OPEN in PowerPoint via the Add-In: the Server applies the change to the live slide through "
            "a client round-trip (pass slideIndex — reuse the one from ppt_get_chart). For a CLOSED "
            ".pptx this returns 3003 — edit the chart offline with the authoring pipeline "
            "(office_run_script + python-pptx) instead. While the Add-In round-trip handler is not yet "
            "available a CONNECTED call may also return 3003 (close the document and use "
            "office_run_script); this lifts automatically once it ships. "
            "Expect >1s latency; concurrent updates on the same document are serialized. "
            "The live round-trip updates the open document in place (an MCP resource_updated notification is fired). "
            "chartType is the REQUIRED discriminator — call ppt_get_chart first to read it. "
            "title=null deletes the title; explicit null is preserved on the wire. "
            "Cross-variant switches (e.g. Pie → Scatter) MUST include compatible series, "
            "otherwise 3015 INVALID_CHART_DATA. Categorical: series[].values length must match "
            "categories length."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptUpdateChartInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "update:chart"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptUpdateChartInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Connected-only (#68): CONNECTED live round-trip; DISCONNECTED → refuse."""
        try:
            validated = self.validate_input(arguments, PptUpdateChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        element_id = str(validated.element_id)
        # Offline chart work is served by the authoring pipeline, not this tool (Path A removed, #68).
        if self.workspace.get_document_status(document_uri) != DocumentStatus.CONNECTED:
            return {"success": False, "error": chart_router.OFFLINE_MESSAGE}
        # The live round-trip must export a specific slide; without slideIndex it cannot route.
        if validated.slide_index is None:
            return {"success": False, "error": chart_router.LIVE_NEEDS_SLIDE_INDEX}
        async with document_lock_manager.acquire(document_uri):
            try:
                result_data = await chart_router.update_chart_path_b(
                    self.workspace, document_uri, element_id, validated.chart, validated.slide_index
                )
            except chart_router.PathBUnavailable:
                # Reactive degradation: live round-trip not available yet → offline-authoring guidance.
                return {"success": False, "error": chart_router.DEGRADE_MESSAGE}
            except ChartEngineError as e:
                return {"success": False, "error": str(e)}
            except Exception as e:  # noqa: BLE001
                return {"success": False, "error": f"3004: {e}"}

        result_data["requiresReload"] = False
        self.workspace.update_last_activity(
            document_uri=document_uri,
            tool_name=self.name,
            result_data=result_data,
        )
        # 通知该文件的 per-file 窗口（内容变了）；不通知根索引——改图表未改变已连接文件集。
        self.workspace.notify_resource_updated(affected_window_uris("/ppt", document_uri))
        return {"success": True, "data": result_data}
