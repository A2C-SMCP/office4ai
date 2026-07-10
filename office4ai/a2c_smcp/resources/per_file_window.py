"""per-file window 资源（W4b-1 / #64）：每个已连接文件 = 一个独立 ``window://`` 资源。

取代旧的 per-type 聚合窗口（``window://office4ai/word|ppt``）。base_uri 形如
``window://office4ai/{word|ppt}/{doc_id}``，``doc_id`` 由 ``document_uri`` 稳定编码派生
（同一文件跨刷新 URI 稳定；文件名保可读 + sha1 短摘要消歧）。excel 暂不投射
（W4b-3 / #66，blocked-by milestone #3）。

契约锚点：``docs/milestone4_authoring_desktop_spec.md`` §「W4 交互契约（#69 体验门控定稿）」契约②。
"""

from __future__ import annotations

import asyncio
import hashlib
from abc import abstractmethod
from typing import Any
from urllib.parse import quote, urlencode

from loguru import logger

from office4ai.a2c_smcp.resources.base import BaseResource, parse_window_uri_params
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import normalize_document_uri

# Socket.IO namespace → window 类型段。excel 缺席 = W4b-3 延后（连接时不投射窗口）。
WINDOW_TYPE_BY_NAMESPACE: dict[str, str] = {"/word": "word", "/ppt": "ppt"}

# 根索引 URI（连接文件集变化时通知其订阅者重读）
ROOT_WINDOW_URI = "window://office4ai"


def encode_doc_id(document_uri: str) -> str:
    """把 ``document_uri`` 编码成稳定、URL-safe、可读的 window 路径段。

    文件名（basename）保可读 + 归一 uri 的 sha1 短摘要消歧（同名不同目录不撞车）。
    ``document_uri`` 由 connection_manager 归一化，故编码稳定可复现。
    """
    name = document_uri.rsplit("/", 1)[-1] or document_uri
    digest = hashlib.sha1(document_uri.encode("utf-8")).hexdigest()[:8]
    return f"{quote(name, safe='')}-{digest}"


def per_file_window_base_uri(namespace: str, document_uri: str) -> str | None:
    """``(namespace, document_uri)`` → per-file window ``base_uri``；不支持的 namespace 返回 ``None``。"""
    wtype = WINDOW_TYPE_BY_NAMESPACE.get(namespace)
    if wtype is None:
        return None
    return f"window://office4ai/{wtype}/{encode_doc_id(document_uri)}"


def affected_window_uris(namespace: str, document_uri: str, *, include_root: bool = False) -> list[str]:
    """某文件被修改时受影响的 window URI：其 per-file 窗口（可选 + 根索引）。

    ``document_uri`` 会先归一化以匹配注册时的 per-file 窗口坐标系。namespace 无对应窗口
    （如 excel，W4b-3 前）时不含 per-file 项。工具自通知与 call_tool 的 affected 计算共用本函数。
    """
    base = per_file_window_base_uri(namespace, normalize_document_uri(document_uri))
    uris: list[str] = []
    if base is not None:
        uris.append(base)
    if include_root:
        uris.append(ROOT_WINDOW_URI)
    return uris


class PerFileWindowResource(BaseResource):
    """单个已连接文件的窗口资源基类（承载 URI/命名/参数，渲染交子类）。"""

    FETCH_TIMEOUT = 3  # 秒

    def __init__(
        self,
        workspace: OfficeWorkspace,
        document_uri: str,
        namespace: str,
        priority: int = 50,
        fullscreen: bool = False,
    ) -> None:
        if not isinstance(priority, int) or not (0 <= priority <= 100):
            raise ValueError(f"priority must be int in [0, 100], got: {priority}")
        wtype = WINDOW_TYPE_BY_NAMESPACE.get(namespace)
        if wtype is None:
            raise ValueError(f"per-file window 不支持的 namespace: {namespace}")
        self.workspace = workspace
        self.document_uri = document_uri
        self.namespace = namespace
        self._wtype = wtype
        self._base_uri = f"window://office4ai/{wtype}/{encode_doc_id(document_uri)}"
        self._priority = priority
        self._fullscreen = fullscreen

    # ── BaseResource implementation ──

    @property
    def uri(self) -> str:
        query = urlencode(
            {
                "priority": str(self._priority),
                "fullscreen": "true" if self._fullscreen else "false",
            }
        )
        return f"{self._base_uri}?{query}"

    @property
    def base_uri(self) -> str:
        return self._base_uri

    @property
    def filename(self) -> str:
        return self.document_uri.rsplit("/", 1)[-1] or self.document_uri

    @property
    def name(self) -> str:
        return f"{self._wtype.upper()} · {self.filename}"

    @property
    def description(self) -> str:
        return f"{self._wtype} 文档窗口（单文件）：{self.document_uri}"

    @property
    def mime_type(self) -> str:
        return "text/plain"

    def update_from_uri(self, uri: str) -> None:
        self._priority, self._fullscreen = parse_window_uri_params(
            uri, self._priority, self._fullscreen, log_prefix=f"{self._wtype} file window"
        )

    @abstractmethod
    async def read(self) -> str:  # pragma: no cover
        raise NotImplementedError

    # ── Internal helpers ──

    async def _fetch_with_timeout(self, event: str, params: dict[str, Any]) -> dict[str, Any] | None:
        """通用 3s 超时拉取本文件数据，失败/超时返回 None。"""
        try:
            response = await asyncio.wait_for(
                self.workspace.emit_to_document(self.document_uri, event, params),
                timeout=self.FETCH_TIMEOUT,
            )
            if response.get("success"):
                data: dict[str, Any] = response.get("data", {})
                return data
            return None
        except (asyncio.TimeoutError, ValueError) as e:
            # 3.10：asyncio.wait_for 抛 asyncio.TimeoutError，它在 3.10 不是内置 TimeoutError；
            # 3.11+ 两者已别名统一。捕获 asyncio.TimeoutError 在两版本都正确（否则超时不降级）。
            logger.warning(f"Fetch failed for {event} on {self.document_uri}: {e}")
            return None


class WordFileWindowResource(PerFileWindowResource):
    """单个 Word 文件窗口：渲染该文件的元数据 + 可见内容。"""

    async def read(self) -> str:
        lines: list[str] = [f"# Word 文档: {self.filename}", "", f"- URI: {self.document_uri}", ""]

        stats, content = await asyncio.gather(
            self._fetch_with_timeout("word:get:documentStats", {"document_uri": self.document_uri}),
            self._fetch_with_timeout("word:get:visibleContent", {"document_uri": self.document_uri}),
        )

        lines.append("## 元数据")
        if stats is not None:
            page_count = stats.get("pageCount", "N/A")
            word_count = stats.get("wordCount", 0)
            paragraph_count = stats.get("paragraphCount", "N/A")
            word_count_str = f"{word_count:,}" if isinstance(word_count, int) else str(word_count)
            lines.append(f"- 总页数: {page_count}")
            lines.append(f"- 总字数: {word_count_str}")
            lines.append(f"- 段落数: {paragraph_count}")
        else:
            lines.append("[元数据不可用: 请求超时]")

        lines.append("")
        lines.append("## 当前可见内容")
        if content is not None:
            text = content.get("text", "")
            lines.append(text if text else "(空)")
        else:
            lines.append("[可见内容不可用: 请求超时]")

        return "\n".join(lines)


class PptFileWindowResource(PerFileWindowResource):
    """单个 PPT 文件窗口：渲染该文件的元数据 + 当前幻灯片 ±N 张摘要。"""

    DEFAULT_RANGE = 2  # ±N slides

    def __init__(
        self,
        workspace: OfficeWorkspace,
        document_uri: str,
        namespace: str,
        priority: int = 50,
        fullscreen: bool = False,
    ) -> None:
        super().__init__(workspace, document_uri, namespace, priority, fullscreen)
        self._range = self.DEFAULT_RANGE

    def update_from_uri(self, uri: str) -> None:
        super().update_from_uri(uri)
        from urllib.parse import parse_qs, urlparse

        params = parse_qs(urlparse(uri).query)
        if "range" in params:
            try:
                new_range = int(params["range"][0])
                if 0 <= new_range <= 10:
                    if new_range != self._range:
                        logger.debug(f"PPT file window range: {self._range} -> {new_range}")
                        self._range = new_range
                else:
                    logger.warning(f"Invalid range value in URI: {new_range}, must be in [0, 10]")
            except (ValueError, IndexError) as e:
                logger.warning(f"Failed to parse range from URI: {e}")

    async def read(self) -> str:
        lines: list[str] = [f"# PPT 文档: {self.filename}", "", f"- URI: {self.document_uri}", ""]

        pres_info = await self._fetch_with_timeout("ppt:get:slideInfo", {"document_uri": self.document_uri})
        if pres_info is None:
            lines.append("## 元数据")
            lines.append("[元数据不可用: 请求超时]")
            return "\n".join(lines)

        slide_count = pres_info.get("slideCount", 0)
        current_index = pres_info.get("currentSlideIndex", 0)
        dimensions = pres_info.get("dimensions", {})
        width = dimensions.get("width", "?")
        height = dimensions.get("height", "?")
        aspect_ratio = dimensions.get("aspectRatio", "?")

        lines.append("## 元数据")
        lines.append(f"- 总张数: {slide_count}")
        lines.append(f"- 尺寸: {width}×{height} pt ({aspect_ratio})")
        lines.append(f"- 当前幻灯片: 第 {current_index + 1} 张")

        if slide_count > 0:
            start = max(0, current_index - self._range)
            end = min(slide_count - 1, current_index + self._range)

            lines.append("")
            lines.append(f"## 幻灯片摘要 (第 {start + 1}-{end + 1} 张)")

            tasks = [
                self._fetch_with_timeout(
                    "ppt:get:slideInfo",
                    {"document_uri": self.document_uri, "slide_index": i},
                )
                for i in range(start, end + 1)
            ]
            slide_results = await asyncio.gather(*tasks)

            for idx, slide_data in enumerate(slide_results):
                i = start + idx
                is_current = i == current_index
                marker = "➡️ " if is_current else ""
                current_label = " (当前)" if is_current else ""

                if slide_data is not None:
                    slide_info = slide_data.get("slideInfo", {})
                    title = slide_info.get("title", f"幻灯片 {i + 1}")
                    elements = slide_data.get("elements", [])
                    notes = slide_info.get("notes", "")

                    lines.append("")
                    lines.append(f"### {marker}第 {i + 1} 张: {title}{current_label}")

                    if elements:
                        type_counts: dict[str, int] = {}
                        for elem in elements:
                            elem_type = elem.get("type", "未知")
                            type_counts[elem_type] = type_counts.get(elem_type, 0) + 1
                        elem_str = ", ".join(f"{t}×{c}" for t, c in type_counts.items())
                        lines.append(f"- 元素: {elem_str}")
                    else:
                        lines.append("- 元素: (无)")

                    lines.append(f"- 备注: {notes if notes else '(无)'}")
                else:
                    lines.append("")
                    lines.append(f"### {marker}第 {i + 1} 张{current_label}")
                    lines.append("[幻灯片信息不可用: 请求超时]")

        return "\n".join(lines)


def create_per_file_window(
    workspace: OfficeWorkspace,
    document_uri: str,
    namespace: str,
    *,
    priority: int = 50,
    fullscreen: bool = False,
) -> PerFileWindowResource | None:
    """按 namespace 造对应 per-file window；excel/未知 namespace 返回 ``None``（W4b-3 延后）。"""
    wtype = WINDOW_TYPE_BY_NAMESPACE.get(namespace)
    if wtype == "word":
        return WordFileWindowResource(workspace, document_uri, namespace, priority, fullscreen)
    if wtype == "ppt":
        return PptFileWindowResource(workspace, document_uri, namespace, priority, fullscreen)
    return None
