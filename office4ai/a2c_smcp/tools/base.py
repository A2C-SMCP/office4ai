# filename: base.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Literal, TypeVar

from loguru import logger
from pydantic import BaseModel

from office4ai.environment.workspace.base import OfficeAction, OfficeObs
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

T = TypeVar("T", bound=BaseModel)


def _schema_has_ref(node: Any) -> bool:
    """是否仍残留 ``$ref`` | Whether any ``$ref`` remains."""
    if isinstance(node, dict):
        if "$ref" in node:
            return True
        return any(_schema_has_ref(v) for v in node.values())
    if isinstance(node, list):
        return any(_schema_has_ref(v) for v in node)
    return False


def _isolate_ref_siblings(node: dict[str, Any]) -> dict[str, Any]:
    """``{"$ref": X, ...兄弟键}`` -> ``{"anyOf": [{"$ref": X}], ...兄弟键}`` (回退用)。"""
    ref = node["$ref"]
    siblings = {k: v for k, v in node.items() if k != "$ref"}
    if not siblings:
        return node
    return {"anyOf": [{"$ref": ref}], **siblings}


def _inline_refs(node: Any, defs: dict[str, Any], stack: tuple[str, ...]) -> Any:
    """递归把 ``$ref`` 解引用为 ``$defs`` 中的定义并内联展开。

    遇到 ``{"$ref": "#/$defs/X", ...兄弟键}``: 解析 X 的定义、递归内联其内部 ``$ref``,
    再把兄弟键 (如字段级 ``description``) 覆盖合并到展开结果上 (字段级描述优先于定义级)。
    遇到环或未知目标: 退回 ``_isolate_ref_siblings`` (避免裸 ``$ref`` + 兄弟键, 见 #37)。
    """
    if isinstance(node, dict):
        if "$ref" in node:
            name = node["$ref"].rsplit("/", 1)[-1]
            if name in stack or name not in defs:  # 环 / 未知目标: 无法内联, 退回隔离
                return _isolate_ref_siblings({k: _inline_refs(v, defs, stack) for k, v in node.items()})
            resolved = _inline_refs(defs[name], defs, (*stack, name))
            merged = dict(resolved) if isinstance(resolved, dict) else resolved
            if isinstance(merged, dict):
                for k, v in node.items():
                    if k != "$ref":
                        merged[k] = _inline_refs(v, defs, stack)
            return merged
        return {k: _inline_refs(v, defs, stack) for k, v in node.items()}
    if isinstance(node, list):
        return [_inline_refs(v, defs, stack) for v in node]
    return node


def inline_schema_refs(schema: Any) -> Any:
    """内联 JSON Schema 的全部 ``$ref``, 产出自包含 (无 ``$ref``/``$defs``) 的 schema。

    office4ai #37: MCP 工具的**必填**嵌套对象参数被模型误填成 JSON 字符串。根因经真机定位:
    a2c 客户端把工具 ``inputSchema`` **原样透传**给模型 (不解析 ``$ref``), 而该 function-calling
    栈不支持工具参数里的 ``$ref``/``$defs`` —— 凡走 ``$ref`` 的字段 (Pydantic 对嵌套对象一律
    产出 ``$ref``) 模型都看不到内部结构, 必填字段遂退化为填 JSON 字符串; 可选字段"看似正常"
    只是因为模型常省略不填。

    最初尝试把「裸 ``$ref`` + 兄弟键」改写为 ``anyOf: [{$ref}]`` 隔离 ``$ref``, 真机验证仍失败
    (``$ref`` 本身就不被支持)。故改为**完全内联**: 把每个 ``$ref`` 解引用展开成显式
    ``{type, properties, ...}``, 并删除顶层 ``$defs``, 让每个工具 schema 自包含、无任何 ``$ref``。

    字段级兄弟键 (如 ``description``) 在内联时覆盖合并到定义之上。对**无环**模型 (本项目现状)
    完全内联、丢弃 ``$defs``; 极端**环引用**模型无法完全内联时, 保留 ``$defs`` 并对残留
    ``$ref`` 做兄弟键隔离 (退回 #37 的次优形态), 避免重新引入裸 ``$ref`` + 兄弟键。
    """
    if not isinstance(schema, dict):
        return schema
    defs = schema.get("$defs", {})
    body = {k: v for k, v in schema.items() if k != "$defs"}
    inlined = _inline_refs(body, defs, ())
    if isinstance(inlined, dict) and _schema_has_ref(inlined):
        # 环 / 未知引用: 无法完全内联, 保留 $defs 供解析
        inlined["$defs"] = defs
    return inlined


class BaseTool(ABC):
    """
    声明式工具基类 | Declarative Tool Base Class

    子类只需声明元数据 (name, description, category, event_name, input_model),
    通用执行逻辑由基类 execute() 提供。

    Subclasses only need to declare metadata; generic execution logic is provided
    by the base class execute() method.
    """

    def __init__(self, workspace: OfficeWorkspace) -> None:
        self.workspace = workspace

    # ── 子类必须声明的元数据 | Metadata subclass must declare ──

    @property
    @abstractmethod
    def name(self) -> str:  # pragma: no cover
        """工具名称, 如 'word_insert_text' | Tool name"""
        raise NotImplementedError

    @property
    @abstractmethod
    def description(self) -> str:  # pragma: no cover
        """工具描述, 面向 AI 的自然语言说明 | Tool description for AI"""
        raise NotImplementedError

    @property
    @abstractmethod
    def input_schema(self) -> dict[str, Any]:  # pragma: no cover
        """JSON Schema, 通常由 InputModel.model_json_schema() 生成"""
        raise NotImplementedError

    @property
    @abstractmethod
    def category(self) -> Literal["word", "ppt", "excel"]:  # pragma: no cover
        """平台类别: 'word' | 'ppt' | 'excel' | Platform category"""
        raise NotImplementedError

    @property
    @abstractmethod
    def event_name(self) -> str:  # pragma: no cover
        """Socket.IO 事件名, 如 'insert:text' | Socket.IO event name"""
        raise NotImplementedError

    @property
    @abstractmethod
    def input_model(self) -> type[BaseModel]:  # pragma: no cover
        """Pydantic InputModel 类 | Pydantic InputModel class"""
        raise NotImplementedError

    # ── 通用执行逻辑 | Generic execution logic ──

    async def execute(self, arguments: dict[str, Any]) -> dict[str, Any]:
        """
        通用执行流程 | Generic execution flow:
        1. 验证输入 | Validate input
        2. 提取 document_uri 和业务参数 | Extract document_uri and business params
        3. 构建 OfficeAction | Build OfficeAction
        4. 调用 workspace.execute() | Call workspace.execute()
        5. 格式化返回 (hook) | Format result (hook)
        """
        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, self.input_model)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 提取参数
        params = validated.model_dump(exclude_none=True)
        document_uri = params.pop("document_uri")

        # 3. 构建 OfficeAction
        action = OfficeAction(
            category=self.category,
            action_name=self.event_name,
            params={"document_uri": document_uri, **params},
        )

        # 4. 执行
        try:
            obs = await self.workspace.execute(action)
        except Exception as e:
            logger.exception(f"工具执行失败 | Tool execution failed: {self.name}")
            return {"success": False, "error": str(e)}

        # 4.5 活动追踪 | Activity tracking
        if obs.success:
            self.workspace.update_last_activity(
                document_uri=document_uri,
                tool_name=self.name,
                result_data=obs.data,
            )

        # 5. 格式化返回 (hook)
        return self.format_result(obs)

    def format_result(self, obs: OfficeObs) -> dict[str, Any]:
        """
        默认返回格式化 hook. 子类可 override.
        Default result formatting hook. Subclass can override.

        默认行为: 返回 JSON 结构 ``{success, data}`` (失败 ``{success, error}``).
        获取类工具可 override 返回纯文本/Markdown 摘要 (``{success, content, data}``).

        写操作「最小返回」约定 (office4ai #26 决策, Plan A):
            写操作工具沿用本默认实现，透传 Add-In/协议定义的**最小返回**——只回操作锚点
            (如 ``set:range`` → ``{address}``)，**不**回写入后的数据快照。这与可组合性不
            冲突：可组合性由「独立读工具 + Agent 编排」承载 (dev_plan「最小但可组合的动作
            单元」)，而非由胖返回值承载。Agent 若需写后状态，应再调一次对应读工具
            (write-then-read)。决策与读工具映射表见
            ``docs/discussions/excel-minimal-return-decision.md``。
        """
        if not obs.success:
            return {"success": False, "error": obs.error or "Unknown error"}
        return {"success": True, "data": obs.data}

    def to_mcp_content(self, result: dict[str, Any]) -> list[dict[str, Any]]:
        """将 ``execute()`` 的结果映射为 MCP 内容块 | Map ``execute()`` result to MCP content blocks.

        默认行为: 整个结果序列化为单个 ``text`` 块, 与历史行为逐字一致。

        需要返回大体积二进制/图片载荷的工具 (如截图) **必须** override 本方法, 改发 MCP
        ``image`` 等内容类型, 否则十余万字符 base64 会被原样内联进 LLM 文本上下文, 撑爆
        token 触发周期性压缩与死循环 (office4ai #42)。``call_tool`` 据此分发, 不再硬编码 text。
        """
        return [{"type": "text", "text": str(result)}]

    def validate_input(self, arguments: dict[str, Any], model: type[T]) -> T:
        """Pydantic 输入验证 | Pydantic input validation"""
        return model.model_validate(arguments)
