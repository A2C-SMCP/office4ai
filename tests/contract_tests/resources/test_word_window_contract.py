"""WordFileWindowResource 契约测试 — 真实 Socket.IO + MockAddInClient (W4b-1 / #64)."""

from __future__ import annotations

import time
from collections.abc import Callable

import pytest

from office4ai.a2c_smcp.resources.per_file_window import WordFileWindowResource
from tests.contract_tests.mock_addin.client import MockAddInClient


@pytest.mark.asyncio
@pytest.mark.contract
class TestWordFileWindowContract:
    async def test_read_fetches_stats_and_content(
        self,
        make_word_file_window: Callable[[str], WordFileWindowResource],
    ) -> None:
        """完整链路：MockAddInClient 响应 stats + visibleContent → per-file read() 渲染正确。"""
        doc_uri = "file:///tmp/contract_word_test.docx"
        resource = make_word_file_window(doc_uri)

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_word_client_stats",
            document_uri=doc_uri,
        )

        client.register_response(
            "word:get:documentStats",
            lambda req: {
                "requestId": req["requestId"],
                "success": True,
                "data": {"pageCount": 5, "wordCount": 1200, "paragraphCount": 20},
                "timestamp": time.time(),
                "duration": 10,
            },
        )
        client.register_response(
            "word:get:visibleContent",
            lambda req: {
                "requestId": req["requestId"],
                "success": True,
                "data": {"text": "Hello World\nContract test paragraph"},
                "timestamp": time.time(),
                "duration": 10,
            },
        )

        await client.connect()
        try:
            content = await resource.read()

            assert "Word 文档: contract_word_test.docx" in content
            assert "总页数: 5" in content
            assert "1,200" in content
            assert "Hello World" in content
        finally:
            await client.disconnect()

    async def test_read_timeout_degradation(
        self,
        make_word_file_window: Callable[[str], WordFileWindowResource],
    ) -> None:
        """超时降级：5s 延迟响应 → 3s 超时 → 降级渲染。"""
        import asyncio

        doc_uri = "file:///tmp/contract_word_timeout.docx"
        resource = make_word_file_window(doc_uri)

        client = MockAddInClient(
            server_url="http://127.0.0.1:3003",
            namespace="/word",
            client_id="contract_word_client_timeout",
            document_uri=doc_uri,
        )

        async def slow_response(req):
            await asyncio.sleep(5)
            return {
                "requestId": req["requestId"],
                "success": True,
                "data": {"pageCount": 1, "wordCount": 10, "paragraphCount": 1},
                "timestamp": time.time(),
                "duration": 5000,
            }

        client.register_response("word:get:documentStats", slow_response)
        client.register_response("word:get:visibleContent", slow_response)

        await client.connect()
        try:
            content = await resource.read()

            assert "Word 文档: contract_word_timeout.docx" in content
            assert "不可用" in content or "超时" in content
        finally:
            await client.disconnect()
