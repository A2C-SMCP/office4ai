"""window://office4ai Resource — Office Workspace 根索引资源."""

from __future__ import annotations

from urllib.parse import urlencode

from office4ai.a2c_smcp.resources.base import BaseResource, parse_window_uri_params
from office4ai.a2c_smcp.resources.per_file_window import WINDOW_TYPE_BY_NAMESPACE, per_file_window_base_uri
from office4ai.environment.workspace.office_workspace import OfficeWorkspace
from office4ai.environment.workspace.socketio.services.connection_manager import connection_manager


class WindowResource(BaseResource):
    """
    Office Workspace 根索引资源

    通过 ``window://office4ai`` 向 AI Agent 展示子资源索引总览。W4b-1（#64）起改为
    **每文件一个子窗口**（per-file window 取代 per-type 聚合），逐条列出已连接文件的
    独立 ``window://`` 子资源。

    渲染示例（有文档连接时）::

        # Office 工作区

        ## 子资源
        - window://office4ai/word/report.docx-1a2b3c4d — WORD · report.docx
        - window://office4ai/ppt/deck.pptx-5e6f7a8b — PPT · deck.pptx

    渲染示例（无文档连接时）::

        # Office 工作区

        ## 子资源
        暂无文档连接，等待 Office Add-In 接入。
    """

    def __init__(self, workspace: OfficeWorkspace, priority: int = 0, fullscreen: bool = False) -> None:
        if not isinstance(priority, int) or not (0 <= priority <= 100):
            raise ValueError(f"priority must be int in [0, 100], got: {priority}")
        self.workspace = workspace
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
        return f"window://office4ai?{query}"

    @property
    def base_uri(self) -> str:
        return "window://office4ai"

    @property
    def name(self) -> str:
        return "Office 工作区"

    @property
    def description(self) -> str:
        return "Office 工作区根索引，逐条列出每个已连接文件的独立 window 子资源。"

    @property
    def mime_type(self) -> str:
        return "text/plain"

    async def read(self) -> str:
        return self._render()

    def update_from_uri(self, uri: str) -> None:
        self._priority, self._fullscreen = parse_window_uri_params(
            uri, self._priority, self._fullscreen, log_prefix="Window resource"
        )

    # ── Rendering ──

    def _render(self) -> str:
        clients = connection_manager.get_all_clients()

        # 每文件一个子窗口（按 per-file base_uri 去重）；word/ppt/excel 均投射（W4b-1 + W4b-3）。
        windows: dict[str, tuple[str, str]] = {}  # base_uri -> (wtype, document_uri)
        for c in clients:
            base_uri = per_file_window_base_uri(c.namespace, c.document_uri)
            if base_uri is None:
                continue
            windows[base_uri] = (WINDOW_TYPE_BY_NAMESPACE[c.namespace], c.document_uri)

        lines: list[str] = ["# Office 工作区", "", "## 子资源"]

        if windows:
            for base_uri in sorted(windows):
                wtype, doc_uri = windows[base_uri]
                filename = doc_uri.rsplit("/", 1)[-1] or doc_uri
                lines.append(f"- {base_uri} — {wtype.upper()} · {filename}")
        else:
            lines.append("暂无文档连接，等待 Office Add-In 接入。")

        return "\n".join(lines)
