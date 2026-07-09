# filename: base.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any
from urllib.parse import parse_qs, urlparse

from loguru import logger
from mcp.server.lowlevel.helper_types import ReadResourceContents
from mcp.types import Resource


def parse_window_uri_params(
    uri: str,
    current_priority: int,
    current_fullscreen: bool,
    log_prefix: str = "Window resource",
) -> tuple[int, bool]:
    """Parse priority and fullscreen from a window:// URI query string.

    Returns (new_priority, new_fullscreen) with unchanged values for invalid/missing params.
    """
    parsed = urlparse(uri)
    params = parse_qs(parsed.query)

    priority = current_priority
    fullscreen = current_fullscreen

    if "priority" in params:
        try:
            val = int(params["priority"][0])
            if 0 <= val <= 100:
                if val != priority:
                    logger.debug(f"{log_prefix} priority: {priority} -> {val}")
                    priority = val
            else:
                logger.warning(f"Invalid priority value in URI: {val}, must be in [0, 100]")
        except (ValueError, IndexError) as e:
            logger.warning(f"Failed to parse priority from URI: {e}")

    if "fullscreen" in params:
        try:
            fs_str = params["fullscreen"][0].lower()
            new_fs: bool | None = None
            if fs_str in {"true", "1", "yes", "on"}:
                new_fs = True
            elif fs_str in {"false", "0", "no", "off"}:
                new_fs = False
            else:
                logger.warning(f"Invalid fullscreen value in URI: {fs_str}, ignoring")

            if new_fs is not None and new_fs != fullscreen:
                logger.debug(f"{log_prefix} fullscreen: {fullscreen} -> {new_fs}")
                fullscreen = new_fs
        except (IndexError, AttributeError) as e:
            logger.warning(f"Failed to parse fullscreen from URI: {e}")

    return priority, fullscreen


class BaseResource(ABC):
    @property
    @abstractmethod
    def uri(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def base_uri(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def name(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def description(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    @abstractmethod
    def mime_type(self) -> str:  # pragma: no cover
        raise NotImplementedError

    @property
    def meta(self) -> dict[str, Any] | None:
        """Optional MCP ``Resource._meta`` payload.

        Default ``None`` (window resources). ``skill://`` roots override to carry
        ``{"source": "resources", ...}`` so the A2C Computer treats them as SKILL
        roots to materialize (skill.md §3 / §11).
        """
        return None

    def list_entries(self) -> list[Resource]:
        """MCP ``resources/list`` entries this resource contributes.

        Default: a single ``Resource``. Multi-file resources (e.g. ``SkillResource``
        in ``resources`` source mode) override to expand into a root + sub-resources.
        """
        payload: dict[str, Any] = {
            "uri": self.uri,
            "name": self.name,
            "description": self.description,
            "mimeType": self.mime_type,
        }
        meta = self.meta
        if meta is not None:
            payload["_meta"] = meta
        return [Resource.model_validate(payload)]

    def update_from_uri(self, uri: str) -> None:
        return

    @abstractmethod
    async def read(self) -> str:  # pragma: no cover
        raise NotImplementedError

    async def read_content(self, uri: str) -> list[ReadResourceContents]:
        """MCP ``resources/read`` content for ``uri``.

        Default wraps :meth:`read` as a single text block with this resource's mime.
        Multi-file resources override to serve per-relative-path text/binary content.
        """
        text = await self.read()
        return [ReadResourceContents(content=text, mime_type=self.mime_type)]
