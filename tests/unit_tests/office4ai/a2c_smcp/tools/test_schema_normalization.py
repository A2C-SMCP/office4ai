"""Schema 内联测试 | Schema inlining tests (office4ai #37)

回归保护: MCP 工具下发给模型的 JSON Schema 必须**自包含**——把所有 ``$ref`` 解引用
内联、删除 ``$defs``。真机定位发现 a2c 客户端把 ``inputSchema`` 原样透传给模型 (GLM-5),
而该 function-calling 栈不支持工具参数里的 ``$ref``/``$defs``: 凡走 ``$ref`` 的嵌套对象
字段模型都看不到内部结构, 必填字段遂被误填成 JSON 字符串。内联后字段结构显式可见。

- ``inline_schema_refs`` 纯函数行为 (内联 / 删 $defs / 环回退 / 幂等)
- 全部工具经内联后不含任何 ``$ref`` / ``$defs``
- 受影响的 11 个必填嵌套对象字段被展开为显式 ``properties``、描述保留、必填语义不变
- 对照组 (可选字段) 仍在、仍可空, 只是结构被内联
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

import office4ai.a2c_smcp.tools.excel  # noqa: F401  (注册 BaseTool 子类)
import office4ai.a2c_smcp.tools.ppt  # noqa: F401
import office4ai.a2c_smcp.tools.word  # noqa: F401
from office4ai.a2c_smcp.tools.base import BaseTool, inline_schema_refs

# Issue #37 列出的 11 个受影响工具及其失败字段
AFFECTED = {
    "excel_add_conditional_format": "rule",
    "excel_set_range_format": "format",
    "excel_update_chart": "properties",
    "ppt_insert_image": "image",
    "ppt_insert_table": "options",
    "ppt_update_element": "updates",
    "ppt_update_text_box": "updates",
    "ppt_update_image": "image",
    "word_insert_image": "image",
    "word_insert_table": "options",
    "word_replace_selection": "content",
}


def _all_tool_classes() -> list[type[BaseTool]]:
    seen: set[type[BaseTool]] = set()

    def walk(cls: type[BaseTool]) -> None:
        for sub in cls.__subclasses__():
            seen.add(sub)
            walk(sub)

    walk(BaseTool)
    return sorted((c for c in seen if not getattr(c, "__abstractmethods__", None)), key=lambda c: c.__name__)


def _has_ref(node: object) -> bool:
    if isinstance(node, dict):
        if "$ref" in node:
            return True
        return any(_has_ref(v) for v in node.values())
    if isinstance(node, list):
        return any(_has_ref(v) for v in node)
    return False


def _bad_ref_nodes(node: object) -> list[list[str]]:
    """所有「同时含 ``$ref`` 和其他键」的节点的兄弟键列表 (#37 不变量)。"""
    found: list[list[str]] = []
    if isinstance(node, dict):
        if "$ref" in node and len(node) > 1:
            found.append(sorted(k for k in node if k != "$ref"))
        for v in node.values():
            found += _bad_ref_nodes(v)
    elif isinstance(node, list):
        for v in node:
            found += _bad_ref_nodes(v)
    return found


def _tool_instances() -> dict[str, BaseTool]:
    out: dict[str, BaseTool] = {}
    for cls in _all_tool_classes():
        try:
            inst = cls(MagicMock())
        except Exception:
            continue
        out[inst.name] = inst
    return out


# ---------------------------------------------------------------------------
# inline_schema_refs 纯函数
# ---------------------------------------------------------------------------


class TestInlineSchemaRefs:
    def test_bare_ref_with_sibling_is_inlined(self) -> None:
        schema = {
            "type": "object",
            "properties": {"f": {"$ref": "#/$defs/X", "description": "field-level"}},
            "$defs": {"X": {"type": "object", "properties": {"a": {"type": "string"}}}},
        }
        out = inline_schema_refs(schema)
        assert "$defs" not in out
        assert out["properties"]["f"] == {
            "type": "object",
            "properties": {"a": {"type": "string"}},
            "description": "field-level",  # 字段级兄弟键覆盖/保留
        }

    def test_field_level_description_overrides_def_description(self) -> None:
        schema = {
            "properties": {"f": {"$ref": "#/$defs/X", "description": "field-level"}},
            "$defs": {"X": {"type": "object", "description": "def-level"}},
        }
        out = inline_schema_refs(schema)
        assert out["properties"]["f"]["description"] == "field-level"

    def test_nested_refs_inlined_recursively(self) -> None:
        schema = {
            "properties": {"f": {"$ref": "#/$defs/Outer"}},
            "$defs": {
                "Outer": {"type": "object", "properties": {"inner": {"$ref": "#/$defs/Inner"}}},
                "Inner": {"type": "integer"},
            },
        }
        out = inline_schema_refs(schema)
        assert not _has_ref(out)
        assert out["properties"]["f"]["properties"]["inner"] == {"type": "integer"}

    def test_defs_dropped_when_fully_inlined(self) -> None:
        schema = {"properties": {"f": {"$ref": "#/$defs/X"}}, "$defs": {"X": {"type": "string"}}}
        out = inline_schema_refs(schema)
        assert "$defs" not in out
        assert not _has_ref(out)

    def test_no_ref_passthrough(self) -> None:
        schema = {"type": "object", "properties": {"a": {"type": "string"}}}
        assert inline_schema_refs(schema) == schema

    def test_idempotent(self) -> None:
        schema = {
            "properties": {"f": {"$ref": "#/$defs/X", "description": "d"}},
            "$defs": {"X": {"type": "object", "properties": {"a": {"type": "string"}}}},
        }
        once = inline_schema_refs(schema)
        assert inline_schema_refs(once) == once

    def test_cyclic_ref_falls_back_to_defs_retained_and_isolated(self) -> None:
        """环引用无法完全内联: 保留 $defs, 且残留 $ref 不带兄弟键 (#37 不变量仍成立)。"""
        schema = {
            "properties": {"node": {"$ref": "#/$defs/Node", "description": "d"}},
            "$defs": {
                "Node": {"type": "object", "properties": {"child": {"$ref": "#/$defs/Node"}}},
            },
        }
        out = inline_schema_refs(schema)
        assert "$defs" in out  # 无法完全内联 → 保留
        assert _has_ref(out)  # 环 ref 残留
        assert _bad_ref_nodes(out) == []  # 但没有「$ref + 兄弟键」


# ---------------------------------------------------------------------------
# 全量工具不变量 (等价于 list_tools() 下发形态)
# ---------------------------------------------------------------------------


class TestToolSchemaInvariant:
    def test_inlined_schemas_have_no_ref_or_defs(self) -> None:
        """经内联后, 任何工具 schema 都应自包含: 无 $ref、无 $defs。"""
        offenders: dict[str, str] = {}
        for name, inst in _tool_instances().items():
            inlined = inline_schema_refs(inst.input_schema)
            if _has_ref(inlined):
                offenders[name] = "still has $ref"
            elif "$defs" in inlined:
                offenders[name] = "still has $defs"
        assert offenders == {}, f"内联后仍非自包含: {offenders}"

    def test_raw_schema_uses_ref_for_affected_tools(self) -> None:
        """文档化根因: 修复前 (raw input_schema) 这 11 个必填字段都走 $ref。"""
        tools = _tool_instances()
        for name, field in AFFECTED.items():
            assert name in tools, f"工具未注册: {name}"
            prop = tools[name].input_schema["properties"][field]
            assert "$ref" in prop, f"{name}.{field} 预期为 $ref (raw), 实际: {prop}"


# ---------------------------------------------------------------------------
# 受影响工具: 内联结果 + 语义保持
# ---------------------------------------------------------------------------


class TestAffectedToolsInlining:
    @pytest.mark.parametrize("name,field", sorted(AFFECTED.items()))
    def test_affected_field_inlined_to_explicit_properties(self, name: str, field: str) -> None:
        inst = _tool_instances()[name]
        raw_prop = inst.input_schema["properties"][field]

        inlined = inline_schema_refs(inst.input_schema)
        prop = inlined["properties"][field]

        assert not _has_ref(prop), f"{name}.{field} 内联后不应再含 $ref"
        # 嵌套对象被展开为显式结构 (properties 或至少 type)
        assert "properties" in prop or prop.get("type") == "object"
        # 字段级 description 保留
        if "description" in raw_prop:
            assert prop.get("description") == raw_prop["description"]
        # 整份 schema 自包含
        assert "$defs" not in inlined

    @pytest.mark.parametrize("name,field", sorted(AFFECTED.items()))
    def test_field_stays_required(self, name: str, field: str) -> None:
        """内联不改可选性: 必填字段仍在 required。"""
        inst = _tool_instances()[name]
        inlined = inline_schema_refs(inst.input_schema)
        assert field in inlined.get("required", [])


# ---------------------------------------------------------------------------
# 对照组: 可选字段仍在、仍可空, 只是被内联
# ---------------------------------------------------------------------------


class TestControlGroupStillNullable:
    def test_optional_field_inlined_but_still_nullable(self) -> None:
        """ppt_insert_shape.options (可选) 内联后无 $ref, 但仍保留可空 (null 分支 + default)。"""
        inst = _tool_instances()["ppt_insert_shape"]
        inlined = inline_schema_refs(inst.input_schema)
        prop = inlined["properties"]["options"]
        assert not _has_ref(prop)
        assert "options" not in inlined.get("required", [])  # 仍为可选
        assert prop.get("default") is None
        assert {"type": "null"} in prop["anyOf"]  # 可空分支保留


# ---------------------------------------------------------------------------
# 集成: list_tools() 下发的 inputSchema 已内联 (守护 server.py 接线)
# ---------------------------------------------------------------------------


class TestListToolsWiresInlining:
    async def test_list_tools_emits_self_contained_schema(self) -> None:
        import os
        from unittest.mock import patch

        from mcp.types import ListToolsRequest

        from office4ai.a2c_smcp.config import MCPServerConfig
        from office4ai.a2c_smcp.server import BaseMCPServer

        class _FakeTool:
            name = "fake_required_object_tool"
            description = "d"
            category = "ppt"
            input_schema = {
                "type": "object",
                "properties": {"x": {"$ref": "#/$defs/Y", "description": "d"}},
                "$defs": {"Y": {"type": "object", "properties": {"a": {"type": "string"}}}},
                "required": ["x"],
            }

        class _Server(BaseMCPServer):
            def _register_tools(self) -> None:
                self.tools["fake_required_object_tool"] = _FakeTool()  # type: ignore[assignment]

            def _register_resources(self) -> None:
                pass

        with patch.dict(os.environ, {}, clear=True):
            server = _Server(MCPServerConfig(), "test-server")

        handler = server.server.request_handlers[ListToolsRequest]
        result = await handler(ListToolsRequest(method="tools/list"))
        schema = result.root.tools[0].inputSchema
        assert not _has_ref(schema)
        assert "$defs" not in schema
        assert schema["properties"]["x"] == {
            "type": "object",
            "properties": {"a": {"type": "string"}},
            "description": "d",
        }
