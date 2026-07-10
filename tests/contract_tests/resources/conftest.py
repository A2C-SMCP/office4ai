"""Contract test fixtures for per-file window resources (W4b-1 / #64)."""

from __future__ import annotations

from collections.abc import Callable

import pytest_asyncio

from office4ai.a2c_smcp.resources.per_file_window import PptFileWindowResource, WordFileWindowResource
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


@pytest_asyncio.fixture
async def make_word_file_window(
    workspace: OfficeWorkspace,
) -> Callable[[str], WordFileWindowResource]:
    def _make(document_uri: str) -> WordFileWindowResource:
        return WordFileWindowResource(workspace, document_uri, "/word")

    return _make


@pytest_asyncio.fixture
async def make_ppt_file_window(
    workspace: OfficeWorkspace,
) -> Callable[[str], PptFileWindowResource]:
    def _make(document_uri: str) -> PptFileWindowResource:
        return PptFileWindowResource(workspace, document_uri, "/ppt")

    return _make
