"""ppt_get_slide_screenshot MCP Tool"""

from __future__ import annotations

from typing import Any, ClassVar, Literal

from pydantic import BaseModel, Field

from office4ai.a2c_smcp.tools.base import BaseTool
from office4ai.environment.workspace.base import OfficeObs
from office4ai.environment.workspace.dtos.ppt import ScreenshotOptions


class PptGetSlideScreenshotInput(BaseModel):
    """MCP 输入模型: 获取幻灯片截图"""

    document_uri: str = Field(..., description="Target document URI (e.g. file:///path/to/presentation.pptx)")
    slideIndex: int = Field(..., description="Slide index (0-based)", ge=0)
    options: ScreenshotOptions | None = Field(None, description="Screenshot options (format, quality)")


class PptGetSlideScreenshotTool(BaseTool):
    """获取幻灯片截图"""

    # OASP format → MCP image MIME type (office4ai #42)
    # 协议 (ScreenshotOptions.format) 只发 png/jpeg; "jpg" 为防御性容错键, 协议层不会命中
    _FORMAT_MIME_MAP: ClassVar[dict[str, str]] = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg"}

    @property
    def name(self) -> str:
        return "ppt_get_slide_screenshot"

    @property
    def description(self) -> str:
        return (
            "Get a screenshot of a specific slide as a Base64-encoded image. "
            "Supports PNG and JPEG formats with configurable quality. "
            "Use this to visually inspect a slide's current appearance."
        )

    @property
    def input_schema(self) -> dict[str, Any]:
        return PptGetSlideScreenshotInput.model_json_schema()

    @property
    def category(self) -> Literal["word", "ppt", "excel"]:
        return "ppt"

    @property
    def event_name(self) -> str:
        return "get:slideScreenshot"

    @property
    def input_model(self) -> type[BaseModel]:
        return PptGetSlideScreenshotInput

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """获取类工具: 返回截图信息"""
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        fmt = obs.data.get("format", "unknown")
        base64_data = obs.data.get("base64", "")
        content = f"Screenshot ({fmt}): {len(base64_data)} chars base64"
        return {"success": True, "content": content, "data": obs.data}

    def to_mcp_content(self, result: dict[str, Any]) -> list[dict[str, Any]]:
        """截图以 MCP ``image`` 内容类型回传, 进入消费方视觉通道而非文本 prompt (office4ai #42)。

        成功且格式已知 → 单个 ``image`` 块 (base64 落在 ``data`` 字段, 不入任何 text)。
        失败 / 缺 base64 / 未知格式 → 保守回退默认 ``text`` 块, 既保留错误信息又不发错误 MIME。
        """
        # 失败: 沿用默认 text (业务返回为 {success, error}, 本就不含 base64)
        if not result.get("success"):
            return super().to_mcp_content(result)
        data = result.get("data") or {}
        base64_data = data.get("base64") or ""
        fmt = str(data.get("format") or "").lower()
        mime = self._FORMAT_MIME_MAP.get(fmt)
        if base64_data and mime is not None:
            return [{"type": "image", "data": base64_data, "mimeType": mime}]
        # 成功但无法作为 image (缺 base64 / 未知格式): 降级 text, 但**绝不**回灌 base64,
        # 否则未知格式分支会让 #42 根因 (base64 内联文本上下文) 无声复发。
        suppressed = (
            f"Screenshot unavailable as image (format={fmt or 'unknown'}, {len(base64_data)} chars base64 suppressed)"
        )
        return [{"type": "text", "text": suppressed}]
