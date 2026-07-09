"""office_run_script MCP 工具单元测试 | office_run_script tool unit tests (milestone #4 · S1).

聚焦工具接入契约（元数据 + execute override 收敛为 5 键返回契约 + 输入校验），
沙箱行为本身在 test_runtime.py 覆盖。
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock

from office4ai.a2c_smcp.tools.authoring import OfficeRunScriptTool


def _tool() -> OfficeRunScriptTool:
    # 本工具 execute 全程不触碰 workspace，用 MagicMock 满足 BaseTool.__init__ 即可
    return OfficeRunScriptTool(MagicMock())


def test_metadata() -> None:
    tool = _tool()
    assert tool.name == "office_run_script"
    assert tool.category == "authoring"
    assert tool.requires_connection is False
    schema = tool.input_schema
    assert "script" in schema["properties"]
    assert schema["required"] == ["script"]


async def test_execute_happy_path_returns_contract(tmp_path: Path) -> None:
    tool = _tool()
    result = await tool.execute(
        {"script": "open('note.txt','w').write('ok')\n", "work_dir": str(tmp_path)},
    )

    assert set(result.keys()) == {"ok", "path", "summary", "logs", "stderr"}
    assert result["ok"] is True
    assert Path(result["path"]) == tmp_path
    assert (tmp_path / "note.txt").exists()


async def test_execute_missing_script_is_invalid(tmp_path: Path) -> None:
    tool = _tool()
    result = await tool.execute({"work_dir": str(tmp_path)})

    assert result["ok"] is False
    assert "invalid input" in result["stderr"]
    assert set(result.keys()) == {"ok", "path", "summary", "logs", "stderr"}
