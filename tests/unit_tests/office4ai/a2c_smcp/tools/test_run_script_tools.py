"""{word,ppt,excel}_run_script 在线逃生舱工具单元测试（issue #87 / oasp#18）.

覆盖：元数据 + requires_connection（W4a 收敛）+ ToolAnnotations（destructive/openWorld）
+ 逃生舱描述（首句强制优先 typed / 两准入条件 / 与离线 office_run_script 互指）
+ server_timeout 派生 + execute 构建正确 OfficeAction（含 server_timeout_ms）+ 中转透传。
"""

from __future__ import annotations

from unittest.mock import AsyncMock, MagicMock

import pytest

from office4ai.a2c_smcp.tools.excel import ExcelRunScriptTool
from office4ai.a2c_smcp.tools.ppt import PptRunScriptTool
from office4ai.a2c_smcp.tools.run_script_base import (
    DEFAULT_SCRIPT_TIMEOUT_MS,
    SERVER_TIMEOUT_GRACE_MS,
)
from office4ai.a2c_smcp.tools.word import WordRunScriptTool
from office4ai.environment.workspace.base import OfficeObs

TOOL_SPECS = [
    (WordRunScriptTool, "word_run_script", "word", "Word"),
    (PptRunScriptTool, "ppt_run_script", "ppt", "PowerPoint"),
    (ExcelRunScriptTool, "excel_run_script", "excel", "Excel"),
]


@pytest.fixture
def mock_workspace():
    workspace = MagicMock()
    workspace.execute = AsyncMock(return_value=OfficeObs(success=True, data={}))
    return workspace


class TestMetadata:
    @pytest.mark.parametrize("cls,name,category,_host", TOOL_SPECS)
    def test_identity(self, mock_workspace, cls, name, category, _host) -> None:
        tool = cls(mock_workspace)
        assert tool.name == name
        assert tool.category == category
        assert tool.event_name == "run:script"
        # category + event_name 拼出 OASP 全名（workspace.execute 用 f"{category}:{action_name}"）
        assert f"{tool.category}:{tool.event_name}" == f"{category}:run:script"

    @pytest.mark.parametrize("cls,name,category,_host", TOOL_SPECS)
    def test_requires_connection_gated(self, mock_workspace, cls, name, category, _host) -> None:
        # 平台 category 派生 → True：随 Add-In 连接经 W4a 收敛（区别于常驻 office_run_script）
        assert cls(mock_workspace).requires_connection is True

    @pytest.mark.parametrize("cls,name,category,_host", TOOL_SPECS)
    def test_input_schema(self, mock_workspace, cls, name, category, _host) -> None:
        props = cls(mock_workspace).input_schema["properties"]
        assert {"document_uri", "script", "args", "timeout_ms"} <= set(props)


class TestAnnotations:
    @pytest.mark.parametrize("cls,name,category,_host", TOOL_SPECS)
    def test_destructive_and_openworld(self, mock_workspace, cls, name, category, _host) -> None:
        ann = cls(mock_workspace).annotations
        assert ann is not None
        assert ann.destructiveHint is True  # 可写文档、非原子/不可回滚
        assert ann.openWorldHint is True  # 脚本非沙箱、可 fetch 任意第三方


class TestEscapeHatchDescription:
    @pytest.mark.parametrize("cls,name,category,host", TOOL_SPECS)
    def test_first_sentence_prefers_typed(self, mock_workspace, cls, name, category, host) -> None:
        desc = cls(mock_workspace).description
        assert desc.startswith("ESCAPE HATCH — ALWAYS prefer the typed")
        assert host in desc

    @pytest.mark.parametrize("cls,name,category,host", TOOL_SPECS)
    def test_two_admission_conditions(self, mock_workspace, cls, name, category, host) -> None:
        desc = cls(mock_workspace).description
        # 仅两准入条件：能力缺口 / confirmed bug
        assert "no typed" in desc
        assert "confirmed bug" in desc
        assert "do NOT use run:script for convenience" in desc

    @pytest.mark.parametrize("cls,name,category,host", TOOL_SPECS)
    def test_cross_references_offline_channel(self, mock_workspace, cls, name, category, host) -> None:
        desc = cls(mock_workspace).description
        # 与离线 office_run_script 互指防混淆（在线 Office.js vs 离线 Python 沙箱）
        assert "office_run_script" in desc
        assert "OFFLINE" in desc and "ONLINE" in desc


class TestServerTimeoutDerivation:
    def test_default_when_absent(self, mock_workspace) -> None:
        tool = WordRunScriptTool(mock_workspace)
        assert tool.derive_server_timeout_ms({}) == DEFAULT_SCRIPT_TIMEOUT_MS + SERVER_TIMEOUT_GRACE_MS

    def test_custom_plus_grace(self, mock_workspace) -> None:
        tool = ExcelRunScriptTool(mock_workspace)
        assert tool.derive_server_timeout_ms({"timeout_ms": 120000}) == 120000 + SERVER_TIMEOUT_GRACE_MS

    @pytest.mark.parametrize("bad", [0, -5, None, "x"])
    def test_invalid_falls_back_to_default(self, mock_workspace, bad) -> None:
        tool = PptRunScriptTool(mock_workspace)
        expected = DEFAULT_SCRIPT_TIMEOUT_MS + SERVER_TIMEOUT_GRACE_MS
        assert tool.derive_server_timeout_ms({"timeout_ms": bad}) == expected


class TestExecuteRelay:
    """execute 走 BaseTool 默认路径 → OfficeAction(category:run:script, server_timeout_ms) 中转。"""

    @pytest.mark.asyncio
    async def test_builds_action_with_derived_timeout(self, mock_workspace) -> None:
        tool = WordRunScriptTool(mock_workspace)
        await tool.execute({"document_uri": "file:///a.docx", "script": "return 1;", "timeout_ms": 90000})
        mock_workspace.execute.assert_awaited_once()
        action = mock_workspace.execute.call_args[0][0]
        assert action.category == "word"
        assert action.action_name == "run:script"
        assert action.params["document_uri"] == "file:///a.docx"
        assert action.params["script"] == "return 1;"
        assert action.params["timeout_ms"] == 90000
        # 派生 server_timeout = 90000 + grace（不用全局 30s，见 issue #87 超时派生）
        assert action.server_timeout_ms == 90000 + SERVER_TIMEOUT_GRACE_MS

    @pytest.mark.asyncio
    async def test_default_timeout_when_omitted(self, mock_workspace) -> None:
        tool = ExcelRunScriptTool(mock_workspace)
        await tool.execute({"document_uri": "file:///a.xlsx", "script": "return 1;"})
        action = mock_workspace.execute.call_args[0][0]
        assert action.server_timeout_ms == DEFAULT_SCRIPT_TIMEOUT_MS + SERVER_TIMEOUT_GRACE_MS
        assert "timeout_ms" not in action.params  # exclude_none

    @pytest.mark.asyncio
    async def test_passes_through_result(self, mock_workspace) -> None:
        mock_workspace.execute.return_value = OfficeObs(
            success=True, data={"result": 42, "logs": ["hi"], "durationMs": 3, "logsTruncated": False}
        )
        tool = PptRunScriptTool(mock_workspace)
        out = await tool.execute({"document_uri": "file:///a.pptx", "script": "return 42;"})
        # 纯中转：Add-In 的 ScriptResult data 原样透传（最小/透传返回约定）
        assert out == {"success": True, "data": {"result": 42, "logs": ["hi"], "durationMs": 3, "logsTruncated": False}}
