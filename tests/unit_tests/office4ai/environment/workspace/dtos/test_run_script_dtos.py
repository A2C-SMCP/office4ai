"""run:script DTO 单元测试 —— ScriptResult + RunScriptRequest 家族（issue #87 / oasp#18）.

覆盖：三命名空间子类自动注册（base 抽象不注册）、snake_case↔camelCase 别名双向、
to_payload 上线缆全 camelCase、ScriptResult 解析/回落。
"""

from __future__ import annotations

import pytest

from office4ai.environment.workspace.dtos.common import (
    ExcelRunScriptRequest,
    PptRunScriptRequest,
    RunScriptRequest,
    ScriptResult,
    WordRunScriptRequest,
    request_registry,
)


class TestRegistration:
    """三具体子类经 event_name 自动注册；抽象基类不注册。"""

    @pytest.mark.parametrize(
        "event,cls",
        [
            ("word:run:script", WordRunScriptRequest),
            ("ppt:run:script", PptRunScriptRequest),
            ("excel:run:script", ExcelRunScriptRequest),
        ],
    )
    def test_concrete_registered(self, event: str, cls: type) -> None:
        assert request_registry.contains(event)
        assert request_registry.get(event) is cls
        assert cls.event_name == event

    def test_base_is_abstract_not_registered(self) -> None:
        # 抽象基类 event_name 留空 → 不进注册表（否则空 event 会污染 wrap_request 查找）
        assert RunScriptRequest.event_name == ""
        assert not request_registry.contains("")

    def test_shared_structure_single_definition(self) -> None:
        # 三子类共享同一结构：字段集合一致（宿主无关执行信封，不应拆三份）
        base_fields = set(RunScriptRequest.model_fields)
        for cls in (WordRunScriptRequest, PptRunScriptRequest, ExcelRunScriptRequest):
            assert set(cls.model_fields) == base_fields


class TestWireRoundTrip:
    """snake_case（Python 内部）↔ camelCase（wire）双向 + to_payload 全 camelCase。"""

    def test_build_and_payload_camelcase(self) -> None:
        req = ExcelRunScriptRequest.build(
            document_uri="file:///a.xlsx",
            script="return 1;",
            args={"x": 1},
            timeout_ms=70000,
        )
        payload = req.to_payload()
        assert payload["script"] == "return 1;"
        assert payload["args"] == {"x": 1}
        assert payload["timeoutMs"] == 70000  # snake_case field -> camelCase alias
        assert payload["documentUri"] == "file:///a.xlsx"
        assert "requestId" in payload and "timestamp" in payload
        # wire 上只有 camelCase，无 snake_case 泄漏
        assert "timeout_ms" not in payload
        assert "document_uri" not in payload

    def test_accepts_both_snake_and_camel_input(self) -> None:
        # populate_by_name=True：输入既接受 snake_case 也接受 camelCase alias
        snake = WordRunScriptRequest(requestId="r1", documentUri="file:///a.docx", script="s", timeout_ms=1000)
        camel = WordRunScriptRequest(requestId="r2", documentUri="file:///a.docx", script="s", timeoutMs=1000)
        assert snake.timeout_ms == camel.timeout_ms == 1000

    def test_optional_fields_omitted_when_none(self) -> None:
        # exclude_none：未给 args/timeout_ms 时 wire 上不出现（Add-In 取默认档）
        payload = PptRunScriptRequest.build(document_uri="file:///a.pptx", script="s").to_payload()
        assert payload["script"] == "s"
        assert "args" not in payload
        assert "timeoutMs" not in payload

    def test_script_is_required(self) -> None:
        with pytest.raises(ValueError):
            WordRunScriptRequest.build(document_uri="file:///a.docx")  # 缺 script


class TestScriptResult:
    """ScriptResult wire 契约锚点：camelCase 解析 + 回落。"""

    def test_parse_from_wire(self) -> None:
        sr = ScriptResult.model_validate(
            {"result": {"n": 3}, "logs": ["a", "b"], "durationMs": 12, "logsTruncated": True}
        )
        assert sr.result == {"n": 3}
        assert sr.logs == ["a", "b"]
        assert sr.duration_ms == 12
        assert sr.logs_truncated is True

    def test_dump_by_alias_camelcase(self) -> None:
        sr = ScriptResult(result=None, logs=[], durationMs=0, logsTruncated=False)
        dumped = sr.model_dump(by_alias=True)
        assert dumped["durationMs"] == 0
        assert dumped["logsTruncated"] is False
        assert "duration_ms" not in dumped

    def test_result_null_and_defaults(self) -> None:
        # result 允许 null（无返回值）、logs 默认空、logsTruncated 默认 False
        sr = ScriptResult(durationMs=5)
        assert sr.result is None
        assert sr.logs == []
        assert sr.logs_truncated is False
