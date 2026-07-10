"""PptFileWindowResource 契约测试 — 真实 Socket.IO + MockAddInClient (W4b-1 / #64)."""

from __future__ import annotations

import asyncio
import time
from collections.abc import Callable

import pytest

from office4ai.a2c_smcp.resources.per_file_window import PptFileWindowResource
from tests.contract_tests.mock_addin.client import MockAddInClient


@pytest.mark.asyncio
@pytest.mark.contract
class TestPptFileWindowContract:
    async def test_read_fetches_presentation_and_slides(
        self,
        make_ppt_file_window: Callable[[str], PptFileWindowResource],
    ) -> None:
        """完整链路：slideInfo 无参 + 带参响应 → 元数据 + slide 摘要渲染。"""
        doc_uri = "file:///tmp/contract_ppt_test.pptx"
        resource = make_ppt_file_window(doc_uri)

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/ppt",
            client_id="contract_ppt_client_slides",
            document_uri=doc_uri,
        )

        def slideinfo_response(req):
            slide_index = req.get("slideIndex")
            if slide_index is None:
                return {
                    "requestId": req["requestId"],
                    "success": True,
                    "data": {
                        "slideCount": 10,
                        "currentSlideIndex": 3,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                    "timestamp": time.time(),
                    "duration": 10,
                }
            else:
                return {
                    "requestId": req["requestId"],
                    "success": True,
                    "data": {
                        "slideInfo": {
                            "title": f"Slide {slide_index + 1}",
                            "notes": f"Note for slide {slide_index + 1}" if slide_index % 2 == 0 else "",
                        },
                        "elements": [
                            {"type": "TextBox", "id": f"tb-{slide_index}"},
                            {"type": "Image", "id": f"img-{slide_index}"},
                        ],
                    },
                    "timestamp": time.time(),
                    "duration": 10,
                }

        client.register_response("ppt:get:slideInfo", slideinfo_response)

        await client.connect()
        try:
            content = await resource.read()

            assert "PPT 文档: contract_ppt_test.pptx" in content
            assert "总张数: 10" in content
            assert "当前幻灯片: 第 4 张" in content
            assert "➡️" in content
            assert "元素:" in content or "TextBox" in content
        finally:
            await client.disconnect()

    async def test_serial_slide_fetch_count(
        self,
        make_ppt_file_window: Callable[[str], PptFileWindowResource],
    ) -> None:
        """验证调用次数: 1 次无参 + (2*range+1) 次带参 (max bounded by total)."""
        doc_uri = "file:///tmp/contract_ppt_count.pptx"
        resource = make_ppt_file_window(doc_uri)

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/ppt",
            client_id="contract_ppt_client_count",
            document_uri=doc_uri,
        )

        def slideinfo_response(req):
            slide_index = req.get("slideIndex")
            if slide_index is None:
                return {
                    "requestId": req["requestId"],
                    "success": True,
                    "data": {
                        "slideCount": 10,
                        "currentSlideIndex": 5,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                    "timestamp": time.time(),
                    "duration": 10,
                }
            return {
                "requestId": req["requestId"],
                "success": True,
                "data": {
                    "slideInfo": {"title": f"Slide {slide_index + 1}", "notes": ""},
                    "elements": [],
                },
                "timestamp": time.time(),
                "duration": 10,
            }

        client.register_response("ppt:get:slideInfo", slideinfo_response)

        await client.connect()
        try:
            resource._range = 2

            await resource.read()

            # 1 (无参) + 5 (slides 3,4,5,6,7) = 6 calls
            assert len(client.received_events) == 6
        finally:
            await client.disconnect()

    async def test_partial_slide_timeout(
        self,
        make_ppt_file_window: Callable[[str], PptFileWindowResource],
    ) -> None:
        """部分 slide 超时 → 超时的显示错误, 其余正常。"""
        doc_uri = "file:///tmp/contract_ppt_partial.pptx"
        resource = make_ppt_file_window(doc_uri)

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/ppt",
            client_id="contract_ppt_client_partial",
            document_uri=doc_uri,
        )

        async def slideinfo_response(req):
            slide_index = req.get("slideIndex")
            if slide_index is None:
                return {
                    "requestId": req["requestId"],
                    "success": True,
                    "data": {
                        "slideCount": 5,
                        "currentSlideIndex": 2,
                        "dimensions": {"width": 960, "height": 540, "aspectRatio": "16:9"},
                    },
                    "timestamp": time.time(),
                    "duration": 10,
                }
            if slide_index == 1:
                await asyncio.sleep(5)  # Exceeds 3s timeout
            return {
                "requestId": req["requestId"],
                "success": True,
                "data": {
                    "slideInfo": {"title": f"Slide {slide_index + 1}", "notes": ""},
                    "elements": [{"type": "TextBox", "id": "t1"}],
                },
                "timestamp": time.time(),
                "duration": 10,
            }

        client.register_response("ppt:get:slideInfo", slideinfo_response)

        await client.connect()
        try:
            content = await resource.read()

            # Slide 1 (index 1) should be degraded
            assert "幻灯片信息不可用" in content
            # Other slides should render normally
            assert "Slide 1" in content  # index 0
        finally:
            await client.disconnect()
