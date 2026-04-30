"""ppt_update_chart MCP Tool (OASP /ppt Draft, v0.2.0 — Server OOXML)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import DocumentStatus
from office4ai.environment.workspace.dtos.ppt import (
    CategoricalChartUpdate,
    ScatterChartUpdate,
)
from office4ai.environment.workspace.services import chart_engine
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager

# Same fail-loud rationale as insert_chart — see that file for the empirical link.
_CONNECTED_REJECT_MESSAGE = (
    "3003: Document is currently open in PowerPoint via the Add-In. "
    "Server-side OOXML chart writes will be silently overwritten by the next "
    "PowerPoint save (verified empirically — see docs/manual_tests/"
    "ppt_chart_v0.2.0.md). Please ask the user to close the document in "
    "PowerPoint, then retry. A future ppt:notify:reload OASP event will lift "
    "this restriction."
)


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
            "(OASP /ppt Draft, Server-side path — bypasses Office.js). "
            "REFUSES with 3003 if the document is currently open in PowerPoint via the Add-In — "
            "PowerPoint's in-memory model would silently overwrite the update on the next save. "
            "When 3003 is returned, ask the user to close the document in PowerPoint, then retry. "
            "Expect >1s latency; concurrent updates on the same document are serialized. "
            "After success an MCP resource_updated notification is fired; reopen the deck to see the change. "
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
        """Override to route through the OOXML chart engine instead of Socket.IO."""
        try:
            validated = self.validate_input(arguments, PptUpdateChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        element_id = str(validated.element_id)
        if self.workspace.get_document_status(document_uri) == DocumentStatus.CONNECTED:
            return {"success": False, "error": _CONNECTED_REJECT_MESSAGE}
        async with document_lock_manager.acquire(document_uri):
            try:
                result_data = await chart_engine.update_chart(
                    document_uri=document_uri,
                    element_id=element_id,
                    update=validated.chart,
                )
            except ChartEngineError as e:
                return {"success": False, "error": str(e)}
            except Exception as e:  # noqa: BLE001
                return {"success": False, "error": f"3004: {e}"}

        result_data["requiresReload"] = True
        self.workspace.update_last_activity(
            document_uri=document_uri,
            tool_name=self.name,
            result_data=result_data,
        )
        self.workspace.notify_resource_updated(["window://office4ai/ppt", "window://office4ai"])
        return {"success": True, "data": result_data}
