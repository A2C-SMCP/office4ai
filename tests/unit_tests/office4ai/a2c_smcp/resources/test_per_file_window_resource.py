"""per-file window 资源单测（W4b-1 / #64）。

覆盖：编码/URI 助手（encode_doc_id / per_file_window_base_uri / affected_window_uris /
create_per_file_window）+ Word/PPT 单文件渲染（元数据 + 内容，含超时降级）。
"""

from __future__ import annotations

import asyncio
from unittest.mock import AsyncMock

import pytest

from office4ai.a2c_smcp.resources.per_file_window import (
    ROOT_WINDOW_URI,
    ExcelFileWindowResource,
    PptFileWindowResource,
    WordFileWindowResource,
    affected_window_uris,
    create_per_file_window,
    encode_doc_id,
    per_file_window_base_uri,
)
from office4ai.environment.workspace.office_workspace import OfficeWorkspace


class TestHelpers:
    def test_encode_doc_id_readable_and_stable(self) -> None:
        uri = "file:///Users/foo/report.docx"
        first = encode_doc_id(uri)
        assert first.startswith("report.docx-")
        assert encode_doc_id(uri) == first  # 稳定：同 uri 同 id

    def test_encode_doc_id_disambiguates_same_name(self) -> None:
        a = encode_doc_id("file:///a/report.docx")
        b = encode_doc_id("file:///b/report.docx")
        assert a != b  # 同名不同目录不撞车（hash 后缀消歧）

    def test_encode_doc_id_url_safe(self) -> None:
        # 含空格 → 百分号编码，路径段仍合法
        assert " " not in encode_doc_id("file:///a/my report.docx")

    def test_base_uri_word_ppt_excel(self) -> None:
        assert per_file_window_base_uri("/word", "file:///a/x.docx").startswith("window://office4ai/word/")
        assert per_file_window_base_uri("/ppt", "file:///a/y.pptx").startswith("window://office4ai/ppt/")
        assert per_file_window_base_uri("/excel", "file:///a/z.xlsx").startswith("window://office4ai/excel/")  # W4b-3
        assert per_file_window_base_uri("/unknown", "file:///a/z.txt") is None

    def test_affected_window_uris(self) -> None:
        uri = "file:///a/y.pptx"
        assert affected_window_uris("/ppt", uri) == [per_file_window_base_uri("/ppt", uri)]
        assert affected_window_uris("/ppt", uri, include_root=True) == [
            per_file_window_base_uri("/ppt", uri),
            ROOT_WINDOW_URI,
        ]
        # excel 现有 per-file 窗口（W4b-3 / #66）
        xlsx = "file:///a/z.xlsx"
        assert affected_window_uris("/excel", xlsx) == [per_file_window_base_uri("/excel", xlsx)]
        # 未知 namespace → 仅根（include_root）或空
        assert affected_window_uris("/unknown", "file:///a/z.txt") == []
        assert affected_window_uris("/unknown", "file:///a/z.txt", include_root=True) == [ROOT_WINDOW_URI]

    def test_create_per_file_window_dispatch(self, workspace: OfficeWorkspace) -> None:
        assert isinstance(create_per_file_window(workspace, "file:///a/x.docx", "/word"), WordFileWindowResource)
        assert isinstance(create_per_file_window(workspace, "file:///a/y.pptx", "/ppt"), PptFileWindowResource)
        assert isinstance(create_per_file_window(workspace, "file:///a/z.xlsx", "/excel"), ExcelFileWindowResource)
        assert create_per_file_window(workspace, "file:///a/z.txt", "/unknown") is None


class TestWordFileWindow:
    def _resource(self, workspace: OfficeWorkspace) -> WordFileWindowResource:
        return WordFileWindowResource(workspace, "file:///tmp/report.docx", "/word")

    def test_metadata(self, workspace: OfficeWorkspace) -> None:
        r = self._resource(workspace)
        assert r.base_uri == per_file_window_base_uri("/word", "file:///tmp/report.docx")
        assert r.uri.startswith(r.base_uri + "?")
        assert "priority=50" in r.uri
        assert r.name == "WORD · report.docx"
        assert r.document_uri == "file:///tmp/report.docx"

    def test_invalid_namespace_rejected(self, workspace: OfficeWorkspace) -> None:
        with pytest.raises(ValueError, match="不支持的 namespace"):
            WordFileWindowResource(workspace, "file:///tmp/x.txt", "/unknown")

    @pytest.mark.asyncio
    async def test_read_renders_stats_and_content(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            assert document_uri == "file:///tmp/report.docx"
            if "documentStats" in event:
                return {"success": True, "data": {"pageCount": 5, "wordCount": 1200, "paragraphCount": 20}}
            return {"success": True, "data": {"text": "Hello World"}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        content = await self._resource(workspace).read()

        assert "Word 文档: report.docx" in content
        assert "总页数: 5" in content
        assert "总字数: 1,200" in content
        assert "Hello World" in content

    @pytest.mark.asyncio
    async def test_read_stats_timeout_degrades(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "documentStats" in event:
                await asyncio.sleep(10)  # cancelled by timeout
            return {"success": True, "data": {"text": "Body"}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        r = self._resource(workspace)
        r.FETCH_TIMEOUT = 0.1
        content = await r.read()

        assert "元数据不可用" in content
        assert "Body" in content

    @pytest.mark.asyncio
    async def test_read_empty_content(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "documentStats" in event:
                return {"success": True, "data": {"pageCount": 1, "wordCount": 0, "paragraphCount": 0}}
            return {"success": True, "data": {"text": ""}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        content = await self._resource(workspace).read()
        assert "(空)" in content

    @pytest.mark.asyncio
    async def test_concurrent_fetch(self, workspace: OfficeWorkspace) -> None:
        async def slow_emit(document_uri: str, event: str, data: dict) -> dict:
            await asyncio.sleep(0.2)
            if "documentStats" in event:
                return {"success": True, "data": {"pageCount": 1, "wordCount": 10, "paragraphCount": 1}}
            return {"success": True, "data": {"text": "content"}}

        workspace.emit_to_document = AsyncMock(side_effect=slow_emit)
        start = asyncio.get_event_loop().time()
        content = await self._resource(workspace).read()
        elapsed = asyncio.get_event_loop().time() - start
        assert elapsed < 0.35, f"expected concurrent (<0.35s), got {elapsed:.2f}s"
        assert "content" in content


class TestPptFileWindow:
    def _resource(self, workspace: OfficeWorkspace) -> PptFileWindowResource:
        return PptFileWindowResource(workspace, "file:///tmp/deck.pptx", "/ppt")

    def test_metadata_and_range_param(self, workspace: OfficeWorkspace) -> None:
        r = self._resource(workspace)
        assert r.base_uri == per_file_window_base_uri("/ppt", "file:///tmp/deck.pptx")
        assert r.name == "PPT · deck.pptx"
        assert r._range == PptFileWindowResource.DEFAULT_RANGE
        r.update_from_uri(r.base_uri + "?priority=80&fullscreen=true&range=1")
        assert r._range == 1
        assert "priority=80" in r.uri

    @pytest.mark.asyncio
    async def test_read_renders_slides(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            assert document_uri == "file:///tmp/deck.pptx"
            if "slide_index" not in data:
                return {
                    "success": True,
                    "data": {
                        "slideCount": 3,
                        "currentSlideIndex": 1,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                }
            idx = data["slide_index"]
            return {
                "success": True,
                "data": {
                    "slideInfo": {"title": f"Slide {idx + 1}", "notes": ""},
                    "elements": [{"type": "TextBox"}],
                },
            }

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        content = await self._resource(workspace).read()

        assert "PPT 文档: deck.pptx" in content
        assert "总张数: 3" in content
        assert "当前幻灯片: 第 2 张" in content
        assert "➡️" in content  # current slide marker
        assert "TextBox" in content

    @pytest.mark.asyncio
    async def test_read_meta_timeout(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            await asyncio.sleep(10)
            return {"success": True, "data": {}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        r = self._resource(workspace)
        r.FETCH_TIMEOUT = 0.1
        content = await r.read()
        assert "元数据不可用" in content


class TestExcelFileWindow:
    """W4b-3（#66）：Excel 单文件窗口渲染（工作簿摘要 + 工作表 + 当前选区）。"""

    def _resource(self, workspace: OfficeWorkspace) -> ExcelFileWindowResource:
        return ExcelFileWindowResource(workspace, "file:///tmp/data.xlsx", "/excel")

    def test_metadata(self, workspace: OfficeWorkspace) -> None:
        r = self._resource(workspace)
        assert r.base_uri == per_file_window_base_uri("/excel", "file:///tmp/data.xlsx")
        assert r.base_uri.startswith("window://office4ai/excel/")
        assert r.name == "EXCEL · data.xlsx"
        assert r.document_uri == "file:///tmp/data.xlsx"

    @pytest.mark.asyncio
    async def test_read_renders_workbook_and_selection(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            assert document_uri == "file:///tmp/data.xlsx"
            if "workbookInfo" in event:
                # wire 契约（OASP SheetInfo）：camelCase isActive/isHidden（与 contract 工厂同源）。
                return {
                    "success": True,
                    "data": {
                        "fileName": "data.xlsx",
                        "activeSheet": "Sheet1",
                        "sheets": [
                            {"name": "Sheet1", "index": 0, "isActive": True, "isHidden": False},
                            {"name": "Hidden", "index": 1, "isActive": False, "isHidden": True},
                            {"name": "Plain", "index": 2, "isActive": False, "isHidden": False},
                        ],
                    },
                }
            return {"success": True, "data": {"address": "A1:C5", "rowCount": 5, "columnCount": 3}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        content = await self._resource(workspace).read()

        assert "Excel 文档: data.xlsx" in content
        assert "工作表数: 3" in content
        assert "活动工作表: Sheet1" in content
        assert "- Sheet1 (当前)" in content
        assert "- Hidden (隐藏)" in content
        assert "- Plain\n" in content  # 非活动/非隐藏 → 无后缀标记
        assert "地址: A1:C5" in content
        assert "尺寸: 5×3" in content

    @pytest.mark.asyncio
    async def test_read_no_worksheets(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "workbookInfo" in event:
                return {"success": True, "data": {"fileName": "empty.xlsx", "activeSheet": "?", "sheets": []}}
            return {"success": True, "data": {"address": "A1", "rowCount": 1, "columnCount": 1}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        content = await self._resource(workspace).read()
        assert "(无工作表)" in content

    @pytest.mark.asyncio
    async def test_read_workbook_timeout_degrades(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "workbookInfo" in event:
                await asyncio.sleep(10)  # cancelled by timeout
            return {"success": True, "data": {"address": "B2", "rowCount": 1, "columnCount": 1}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        r = self._resource(workspace)
        r.FETCH_TIMEOUT = 0.1
        content = await r.read()

        assert "元数据不可用" in content
        assert "工作表列表不可用" in content
        assert "地址: B2" in content  # selection 仍渲染（独立并发拉取）

    @pytest.mark.asyncio
    async def test_read_selection_timeout_degrades(self, workspace: OfficeWorkspace) -> None:
        async def mock_emit(document_uri: str, event: str, data: dict) -> dict:
            if "selectedRange" in event:
                await asyncio.sleep(10)
            return {"success": True, "data": {"fileName": "d.xlsx", "activeSheet": "S", "sheets": [{"name": "S"}]}}

        workspace.emit_to_document = AsyncMock(side_effect=mock_emit)
        r = self._resource(workspace)
        r.FETCH_TIMEOUT = 0.1
        content = await r.read()

        assert "工作表数: 1" in content
        assert "选区信息不可用" in content
