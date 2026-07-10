"""ppt_get_chart MCP Tool (OASP /ppt Draft — Server OOXML, connected-only)."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field, field_validator

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import DocumentStatus
from office4ai.environment.workspace.services import chart_router
from office4ai.environment.workspace.services.chart_engine import ChartEngineError
from office4ai.environment.workspace.services.document_lock import document_lock_manager


class PptGetChartInput(BaseModel):
    """MCP 输入模型：读取已存在图表的当前数据（Server 端 OOXML，连接态客户端往返）。"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/deck.pptx)")
    element_id: str | int = Field(
        ...,
        alias="elementId",
        description=(
            "Chart element ID returned by ppt_insert_chart (opaque string, e.g. 'oasp-chart-<uuid>'). Required."
        ),
    )
    slide_index: int | None = Field(
        default=None,
        alias="slideIndex",
        description="Slide index (0-based) of the chart's slide. REQUIRED — the live round-trip exports that slide to read it.",
        ge=0,
    )

    model_config = {"populate_by_name": True}

    # OF4AI-8: LLM may infer numeric IDs as int; coerce to str so engine can parse "chart-N".
    @field_validator("element_id", mode="before")
    @classmethod
    def _coerce_element_id(cls, v: Any) -> Any:
        return str(v) if isinstance(v, int) else v


class PptGetChartTool(BaseTool):
    """Read an existing chart's data (OASP /ppt Draft, Server OOXML, connected-only)."""

    @property
    def name(self) -> str:
        return "ppt_get_chart"

    @property
    def description(self) -> str:
        return (
            "Read an existing PowerPoint chart's logical data and display options "
            "(OASP /ppt Draft — chart OOXML is parsed Server-side). Requires the target document OPEN "
            "in PowerPoint via the Add-In, with slideIndex supplied: the Server reads the live (possibly "
            "unsaved) slide via a client round-trip. For a CLOSED .pptx this returns 3003 — read the "
            "chart offline with the authoring pipeline (office_run_script + python-pptx) instead. "
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
        """Connected-only (#68): CONNECTED + slideIndex → live round-trip read; else refuse."""
        try:
            validated = self.validate_input(arguments, PptGetChartInput)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        document_uri = validated.document_uri
        element_id = str(validated.element_id)
        # Offline chart reads are served by the authoring pipeline, not this tool (Path A removed, #68).
        if self.workspace.get_document_status(document_uri) != DocumentStatus.CONNECTED:
            return {"success": False, "error": chart_router.OFFLINE_MESSAGE}
        if validated.slide_index is None:
            return {"success": False, "error": chart_router.LIVE_NEEDS_SLIDE_INDEX}
        async with document_lock_manager.acquire(document_uri):
            try:
                result_data = await chart_router.get_chart_path_b(
                    self.workspace, document_uri, element_id, validated.slide_index
                )
            except chart_router.PathBUnavailable:
                # Live read path not available yet → offline-authoring guidance (no on-disk fallback, #68).
                return {"success": False, "error": chart_router.DEGRADE_MESSAGE}
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
