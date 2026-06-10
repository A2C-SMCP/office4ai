"""ppt_get_chart MCP Tool (OASP /ppt Draft, v0.2.0 — Server OOXML)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import DocumentStatus
from office4ai.environment.workspace.services import chart_engine, chart_router
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager


class PptGetChartInput(BaseModel):
    """MCP 输入模型：读取已存在图表的当前数据（Server 端 OOXML 离线处理）。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/deck.pptx)")
    element_id: str | int = Field(
        ...,
        alias="elementId",
        description=("Chart element ID returned by ppt_insert_chart (format: 'chart-<shape_id>'). Required."),
    )
    slide_index: int | None = Field(
        default=None,
        alias="slideIndex",
        description="Optional slide hint to speed up lookup (0-based); elementId remains authoritative.",
        ge=0,
    )

    model_config = {"populate_by_name": True}

    # OF4AI-8: LLM may infer numeric IDs as int; coerce to str so engine can parse "chart-N".
    @field_validator("element_id", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptGetChartTool(BaseTool):
    """Read an existing chart's data (OASP /ppt Draft, Server OOXML)."""

    @property
    def name(self) -> str:
        return "ppt_get_chart"

    @property
    def description(self) -> str:
        return (
            "Read an existing PowerPoint chart's logical data and display options "
            "(OASP /ppt Draft, dual-path — chart OOXML is parsed Server-side). "
            "When the document is OPEN in PowerPoint and slideIndex is supplied, the Server reads the "
            "live (possibly unsaved) slide via a client round-trip; otherwise it reads the .pptx on disk. "
            "Returns chartType, categories/points, series, title, showLegend, showDataLabels, "
            "and geometry (left/top/width/height in points). Inspect chartType FIRST: "
            "categorical types expose 'categories' + 'series[].values'; Scatter exposes 'series[].points'. "
            "Recommended call before ppt_update_chart so chartType (and slideIndex) can be supplied."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetChartInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:chart"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetChartInput

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """Dual-path route (#15): CONNECTED + slideIndex → live round-trip; else on-disk read.

        Read-only, so reactive degradation falls back to the on-disk read (at worst stale data)
        rather than erroring — the same behaviour as reading a connected document in 0.2.0.
        """
        try:
            validated = self.validate_input(arguments, PptGetChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        element_id = str(validated.element_id)
        connected = self.workspace.get_document_status(document_uri) == DocumentStatus.CONNECTED
        async with document_lock_manager.acquire(document_uri):
            try:
                if connected and validated.slide_index is not None:
                    try:
                        result_data = await chart_router.get_chart_path_b(
                            self.workspace, document_uri, element_id, validated.slide_index
                        )
                    except chart_router.PathBUnavailable:
                        # Read-only degradation: fall back to the on-disk read.
                        result_data = await chart_engine.get_chart(
                            document_uri=document_uri,
                            element_id=element_id,
                            slide_index=validated.slide_index,
                        )
                else:
                    result_data = await chart_engine.get_chart(
                        document_uri=document_uri,
                        element_id=element_id,
                        slide_index=validated.slide_index,
                    )
            except ChartEngineError as e:
                return {"success": False, "error": str(e)}
            except Exception as e:  # noqa: BLE001
                return {"success": False, "error": f"3004: {e}"}

        self.workspace.update_last_activity(
            document_uri=document_uri,
            tool_name=self.name,
            result_data=result_data,
        )
        return {"success": True, "data": result_data}
