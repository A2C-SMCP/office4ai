"""ppt_update_chart MCP Tool (OASP /ppt Draft, v0.2.0 — Server OOXML)."""

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
from office4ai.environment.workspace.services import chart_engine, chart_router
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager


class PptUpdateChartInput(BaseModel):
    """MCP 输入模型：更新已存在图表（Server 端 OOXML 离线处理）。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/deck.pptx)")
    element_id: str | int = Field(
        ...,
        alias="elementId",
        description="Chart element ID (format: 'chart-<shape_id>')",
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
            "Slide index (0-based) of the chart. Optional on disk (elementId locates the chart), "
            "but REQUIRED to update a chart while the document is open in PowerPoint — the live "
            "round-trip exports that slide. Reuse the slideIndex returned by ppt_get_chart."
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
            "(OASP /ppt Draft, dual-path — all chart OOXML is built Server-side). "
            "When the document is CLOSED, the Server edits the .pptx on disk; when it is OPEN in "
            "PowerPoint via the Add-In, the Server applies the change to the live slide through a "
            "client round-trip (pass slideIndex — reuse the one from ppt_get_chart). While the Add-In "
            "round-trip handler is not yet available it may return 3003 (close the document, then "
            "retry); this lifts automatically once it ships. "
            "Expect >1s latency; concurrent updates on the same document are serialized. "
            "After an on-disk write an MCP resource_updated notification is fired; reopen the deck to see it. "
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
        """Dual-path route (#15): CONNECTED → client round-trip; DISCONNECTED → on-disk OOXML."""
        try:
            validated = self.validate_input(arguments, PptUpdateChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        element_id = str(validated.element_id)
        connected = self.workspace.get_document_status(document_uri) == DocumentStatus.CONNECTED
        async with document_lock_manager.acquire(document_uri):
            try:
                if connected:
                    # Path B needs the slide to export; without it we cannot safely route an
                    # open-document write (on-disk would be overwritten) → degrade up-front.
                    if validated.slide_index is None:
                        return {"success": False, "error": chart_router.DEGRADE_MESSAGE_WRITE}
                    result_data = await chart_router.update_chart_path_b(
                        self.workspace, document_uri, element_id, validated.chart, validated.slide_index
                    )
                else:
                    result_data = await chart_engine.update_chart(
                        document_uri=document_uri,
                        element_id=element_id,
                        update=validated.chart,
                    )
            except chart_router.PathBUnavailable:
                # Reactive degradation: path B not available yet → flipped 3003 guidance.
                return {"success": False, "error": chart_router.DEGRADE_MESSAGE_WRITE}
            except ChartEngineError as e:
                return {"success": False, "error": str(e)}
            except Exception as e:  # noqa: BLE001
                return {"success": False, "error": f"3004: {e}"}

        result_data["requiresReload"] = not connected
        self.workspace.update_last_activity(
            document_uri=document_uri,
            tool_name=self.name,
            result_data=result_data,
        )
        # 通知该文件的 per-file 窗口（内容变了）；不通知根索引——改图表未改变已连接文件集。
        self.workspace.notify_resource_updated(affected_window_uris("/ppt", document_uri))
        return {"success": True, "data": result_data}
