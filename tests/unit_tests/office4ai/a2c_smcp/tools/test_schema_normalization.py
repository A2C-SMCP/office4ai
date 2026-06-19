"""Schema 归一化测试 | Schema normalization tests (office4ai #37)

回归保护: 必填嵌套对象参数的 JSON Schema 必须把 ``$ref`` 隔离 (不带兄弟键), 否则
function-calling / MCP 预处理层无法内联展开, 模型会把对象误填成 JSON 字符串。

- ``normalize_ref_siblings`` 纯函数行为 (含幂等性)
- 全部工具经归一化后不存在「``$ref`` + 兄弟键」节点 (等价于 ``list_tools()`` 下发形态)
- 对照组 (已隔离的可选字段) 不被改动、必填语义不变
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest

import office4ai.a2c_smcp.tools.excel  # noqa: F401  (注册 BaseTool 子类)
import office4ai.a2c_smcp.tools.ppt  # noqa: F401
import office4ai.a2c_smcp.tools.word  # noqa: F401
from office4ai.a2c_smcp.tools.base import BaseTool, normalize_ref_siblings

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
    # 仅保留可实例化的具体工具
    return sorted((c for c in seen if not getattr(c, "__abstractmethods__", None)), key=lambda c: c.__name__)


def _bad_ref_nodes(node: object, path: str = "") -> list[tuple[str, list[str]]]:
    """返回所有「同时含 ``$ref`` 和其他键」的节点 (path, 兄弟键列表)。"""
    found: list[tuple[str, list[str]]] = []
    if isinstance(node, dict):
        if "$ref" in node and len(node) > 1:
            found.append((path, sorted(k for k in node if k != "$ref")))
        for k, v in node.items():
            found += _bad_ref_nodes(v, f"{path}.{k}")
    elif isinstance(node, list):
        for i, v in enumerate(node):
            found += _bad_ref_nodes(v, f"{path}[{i}]")
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
# normalize_ref_siblings 纯函数
# ---------------------------------------------------------------------------


class TestNormalizeRefSiblings:
    def test_bare_ref_with_sibling_is_isolated(self) -> None:
        out = normalize_ref_siblings({"$ref": "#/$defs/X", "description": "d"})
        assert out == {"anyOf": [{"$ref": "#/$defs/X"}], "description": "d"}

    def test_lone_ref_untouched(self) -> None:
        assert normalize_ref_siblings({"$ref": "#/$defs/X"}) == {"$ref": "#/$defs/X"}

    def test_anyof_isolated_ref_untouched(self) -> None:
        node = {
            "anyOf": [{"$ref": "#/$defs/X"}, {"type": "null"}],
            "default": None,
            "description": "d",
        }
        assert normalize_ref_siblings(node) == node

    def test_array_items_ref_untouched(self) -> None:
        node = {"type": "array", "items": {"$ref": "#/$defs/X"}, "description": "d"}
        assert normalize_ref_siblings(node) == node

    def test_nested_bare_ref_is_isolated(self) -> None:
        node = {"properties": {"f": {"$ref": "#/$defs/X", "description": "d"}}}
        out = normalize_ref_siblings(node)
        assert out["properties"]["f"] == {"anyOf": [{"$ref": "#/$defs/X"}], "description": "d"}

    def test_ref_sibling_inside_defs_is_isolated(self) -> None:
        """全树递归: $defs 内部的「$ref + 兄弟键」也要被隔离 (现有工具暂无此形态, 钉死能力)。"""
        node = {
            "$defs": {"Outer": {"properties": {"inner": {"$ref": "#/$defs/Y", "title": "t"}}}},
            "properties": {"x": {"type": "string"}},
        }
        out = normalize_ref_siblings(node)
        assert out["$defs"]["Outer"]["properties"]["inner"] == {"anyOf": [{"$ref": "#/$defs/Y"}], "title": "t"}

    def test_idempotent(self) -> None:
        node = {"$ref": "#/$defs/X", "description": "d"}
        once = normalize_ref_siblings(node)
        assert normalize_ref_siblings(once) == once

    def test_no_ref_passthrough(self) -> None:
        node = {"type": "object", "properties": {"a": {"type": "string"}}}
        assert normalize_ref_siblings(node) == node


# ---------------------------------------------------------------------------
# 全量工具不变量 (等价于 list_tools() 下发形态)
# ---------------------------------------------------------------------------


class TestToolSchemaInvariant:
    def test_normalized_schemas_have_no_bare_ref_siblings(self) -> None:
        """经归一化后, 任何工具 schema 树中都不应存在「$ref + 兄弟键」节点。"""
        offenders: dict[str, list[tuple[str, list[str]]]] = {}
        for name, inst in _tool_instances().items():
            normalized = normalize_ref_siblings(inst.input_schema)
            bad = _bad_ref_nodes(normalized)
            if bad:
                offenders[name] = bad
        assert offenders == {}, f"归一化后仍存在裸 $ref + 兄弟键: {offenders}"

    def test_raw_schema_exhibits_the_bug_for_affected_tools(self) -> None:
        """文档化根因: 修复前 (raw input_schema) 这 11 个工具确实是裸 $ref + 兄弟键。"""
        tools = _tool_instances()
        for name, field in AFFECTED.items():
            assert name in tools, f"工具未注册: {name}"
            prop = tools[name].input_schema["properties"][field]
            assert "$ref" in prop and len(prop) > 1, f"{name}.{field} 预期为裸 $ref + 兄弟键 (raw), 实际: {prop}"


# ---------------------------------------------------------------------------
# 受影响工具: 归一化结果 + 语义保持
# ---------------------------------------------------------------------------


class TestAffectedToolsNormalization:
    @pytest.mark.parametrize("name,field", sorted(AFFECTED.items()))
    def test_affected_field_isolated_and_preserves_description(self, name: str, field: str) -> None:
        inst = _tool_instances()[name]
        raw_prop = inst.input_schema["properties"][field]
        ref = raw_prop["$ref"]

        normalized = normalize_ref_siblings(inst.input_schema)
        prop = normalized["properties"][field]

        # $ref 被隔离进 anyOf 单分支
        assert prop["anyOf"] == [{"$ref": ref}]
        assert "$ref" not in prop  # 顶层不再有裸 $ref
        # description 平移保留
        assert prop.get("description") == raw_prop.get("description")
        # $defs 仍在, 引用可解析
        defname = ref.rsplit("/", 1)[-1]
        assert defname in normalized["$defs"]

    @pytest.mark.parametrize("name,field", sorted(AFFECTED.items()))
    def test_field_stays_required(self, name: str, field: str) -> None:
        """归一化不得把必填字段改成可选 (不引入 null 分支 / default)。"""
        inst = _tool_instances()[name]
        normalized = normalize_ref_siblings(inst.input_schema)
        assert field in normalized.get("required", [])
        prop = normalized["properties"][field]
        assert "default" not in prop
        assert {"type": "null"} not in prop["anyOf"]


# ---------------------------------------------------------------------------
# 对照组: 已正常工作的工具不被改动
# ---------------------------------------------------------------------------


class TestControlGroupUnchanged:
    def test_optional_field_structure_unchanged(self) -> None:
        """ppt_insert_shape.options (可选, anyOf+null) 经归一化后结构不变。"""
        inst = _tool_instances()["ppt_insert_shape"]
        raw = inst.input_schema["properties"]["options"]
        normalized = normalize_ref_siblings(inst.input_schema)["properties"]["options"]
        assert normalized == raw
        assert {"type": "null"} in normalized["anyOf"]


# ---------------------------------------------------------------------------
# 集成: list_tools() 下发的 inputSchema 已归一化 (守护 server.py 接线)
# ---------------------------------------------------------------------------


class TestListToolsWiresNormalization:
    async def test_list_tools_emits_normalized_schema(self) -> None:
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
                "$defs": {"Y": {"type": "object"}},
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
        assert schema["properties"]["x"] == {"anyOf": [{"$ref": "#/$defs/Y"}], "description": "d"}
