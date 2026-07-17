"""Contract tests —— {word,ppt,excel}:run:script 在线逃生舱（issue #87 / oasp#18）.

走**真实**端到端链路（不 mock RPC 链的任何一环）：
    {ns}_run_script MCP 工具.execute()
      → BaseTool.execute（派生 server_timeout + 构建 OfficeAction）
      → OfficeWorkspace.execute → wrap_request（DTO 校验 + by_alias camelCase）
      → AsyncServer.call() → MockAddInClient ack（回 ScriptResult / 回带 details 的错误）
      → format_result / format_wire_error 摊平

三命名空间同构，参数化覆盖。重点核对：
  1. wire 请求为 camelCase（script/args/timeoutMs），snake→camel 转换在真链路上成立；
  2. 成功回 ScriptResult 信封被纯中转透传（result/logs/durationMs/logsTruncated）；
  3. 失败回带 phase/fault/officeCode/stack/logs 的 details 经 format_wire_error 逐字段存活。

OASP Spec: conventions.md#run-script / data-structures.md §script-execution
"""

from __future__ import annotations

import asyncio
from collections.abc import Callable
from typing import Any

import pytest

from office4ai.a2c_smcp.tools.excel import ExcelRunScriptTool
from office4ai.a2c_smcp.tools.ppt import PptRunScriptTool
from office4ai.a2c_smcp.tools.word import WordRunScriptTool
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

pytestmark = [pytest.mark.asyncio, pytest.mark.contract]

SERVER_URL = "http://127.0.0.1:3003"

# (tool_cls, namespace, category, document_uri)
NS_SPECS = [
    (WordRunScriptTool, "/word", "word", "file:///tmp/contract_run_script.docx"),
    (PptRunScriptTool, "/ppt", "ppt", "file:///tmp/contract_run_script.pptx"),
    (ExcelRunScriptTool, "/excel", "excel", "file:///tmp/contract_run_script.xlsx"),
]


async def _run_tool(
    workspace: OfficeWorkspace,
    client_factory: Any,
    tool_cls: type,
    namespace: str,
    category: str,
    document_uri: str,
    arguments: dict[str, Any],
    response_factory: Callable[[dict], dict],
) -> tuple[dict, list[tuple[str, dict]]]:
    """一次经真实 MCP 工具的 run:script round-trip，返回 (tool_result, received_events)。"""
    wire_event = f"{category}:run:script"
    client = client_factory(
        server_url=SERVER_URL,
        namespace=namespace,
        client_id=f"contract_run_script_{category}",
        document_uri=document_uri,
    )
    client.register_response(wire_event, response_factory)  # 必须在 connect() 之前
    await client.connect()
    try:
        tool = tool_cls(workspace)
        result = await tool.execute({"document_uri": document_uri, **arguments})
        return result, client.received_events
    finally:
        await client.disconnect()


@pytest.mark.parametrize("tool_cls,namespace,category,document_uri", NS_SPECS)
async def test_run_script_success_roundtrip(
    workspace, mock_word_client_factory, tool_cls, namespace, category, document_uri
) -> None:
    """成功路径：wire 载荷 camelCase + ScriptResult 信封纯中转透传。"""
    captured: dict = {}

    def factory(request: dict) -> dict:
        captured.update(request)
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": {"result": {"n": 3}, "logs": ["hello"], "durationMs": 12, "logsTruncated": False},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    result, received = await _run_tool(
        workspace,
        mock_word_client_factory,
        tool_cls,
        namespace,
        category,
        document_uri,
        {"script": "return 3;", "args": {"k": "v"}, "timeout_ms": 70000},
        factory,
    )

    # wire 请求：snake→camel（script/args/timeoutMs），documentUri 注入
    assert captured["script"] == "return 3;"
    assert captured["args"] == {"k": "v"}
    assert captured["timeoutMs"] == 70000  # camelCase alias 上线缆
    assert "timeout_ms" not in captured  # 无 snake 泄漏
    assert captured["documentUri"] == document_uri

    # MockAddIn 确实收到了对应 wire 事件
    assert received and received[0][0] == f"{category}:run:script"

    # 纯中转：ScriptResult 信封原样透传（最小/透传返回约定）
    assert result["success"] is True
    assert result["data"] == {"result": {"n": 3}, "logs": ["hello"], "durationMs": 12, "logsTruncated": False}


@pytest.mark.parametrize("tool_cls,namespace,category,document_uri", NS_SPECS)
async def test_run_script_error_details_survive(
    workspace, mock_word_client_factory, tool_cls, namespace, category, document_uri
) -> None:
    """失败路径：rich details（phase/fault/officeCode/stack/logs）经 format_wire_error 逐字段存活。"""

    def factory(request: dict) -> dict:
        return {
            "requestId": request["requestId"],
            "success": False,
            "error": {
                "code": "3004",
                "message": "Script threw",
                "details": {
                    "phase": "execute",
                    "fault": "GeneralException",
                    "officeCode": "GeneralException",
                    "stack": "Error: boom\n  at eval",
                    "logs": ["s1", "s2"],
                },
            },
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    result, _ = await _run_tool(
        workspace,
        mock_word_client_factory,
        tool_cls,
        namespace,
        category,
        document_uri,
        {"script": "throw new Error('boom');"},
        factory,
    )

    assert result["success"] is False
    err = result["error"]
    # 透传零裁剪：code + message + 每个 details 字段都可反解析
    assert "3004" in err and "Script threw" in err
    for token in ("phase", "execute", "officeCode", "GeneralException", "stack", "s1", "s2"):
        assert token in err, f"details token {token!r} lost in flattened error: {err}"


async def test_run_script_optional_fields_omitted_on_wire(workspace, mock_word_client_factory) -> None:
    """未给 args/timeout_ms 时 wire 上不出现（Add-In 取默认档；exclude_none）。"""
    captured: dict = {}

    def factory(request: dict) -> dict:
        captured.update(request)
        return {
            "requestId": request["requestId"],
            "success": True,
            "data": {"result": None, "logs": [], "durationMs": 1, "logsTruncated": False},
            "timestamp": int(asyncio.get_event_loop().time() * 1000),
        }

    await _run_tool(
        workspace,
        mock_word_client_factory,
        WordRunScriptTool,
        "/word",
        "word",
        "file:///tmp/contract_run_script.docx",
        {"script": "return 1;"},
        factory,
    )

    assert captured["script"] == "return 1;"
    assert "args" not in captured
    assert "timeoutMs" not in captured
