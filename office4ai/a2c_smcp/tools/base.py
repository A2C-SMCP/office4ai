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


def normalize_ref_siblings(schema: Any) -> Any:
    """归一化 JSON Schema, 隔离「裸 ``$ref`` + 兄弟键」节点 | Isolate bare ``$ref`` carrying sibling keys.

    office4ai #37: Pydantic v2 对**必填**嵌套对象字段 (``Field(...)``) 产出
    ``{"$ref": "#/$defs/X", "description": "..."}`` —— ``$ref`` 与兄弟键同层。按 JSON Schema
    (Draft-07 / OpenAPI 3.0) 语义, ``$ref`` 与兄弟键并存时兄弟键被忽略, function-calling / MCP
    的 schema 预处理层据此无法正确内联展开, 模型拿不到内部字段结构, 退化为把对象误填成 JSON 字符串。

    能正常工作的可选字段 (如 ``ppt_insert_shape.options``) 之所以正常, 是因为 Pydantic 把
    ``$ref`` 隔离进 ``anyOf`` 独立分支 (分支内 ``$ref`` 无兄弟键)。本函数对必填字段复刻同样的隔离::

        {"$ref": X, "description": Y}  ->  {"anyOf": [{"$ref": X}], "description": Y}

    必填语义保留 (不加 ``null`` 分支、不加 ``default``)。变换是**幂等**的, 且对已隔离的
    ``$ref`` (``anyOf``/``items`` 内, 或 ``{"$ref": ...}`` 独占) 是 no-op —— 仅当某对象
    同时含 ``$ref`` 和其他键时才改写。``$defs`` 原样保留, 引用可解析。

    选用 ``anyOf`` (而非等价的 OpenAPI 惯用 ``allOf``) 是为了与本项目可选字段既有形态
    (``anyOf: [{$ref}, {type: null}]``) 对称, 隔离效果两者相同。
    """
    if isinstance(schema, dict):
        if "$ref" in schema and len(schema) > 1:
            ref = schema["$ref"]
            siblings = {k: normalize_ref_siblings(v) for k, v in schema.items() if k != "$ref"}
            return {"anyOf": [{"$ref": ref}], **siblings}
        return {k: normalize_ref_siblings(v) for k, v in schema.items()}
    if isinstance(schema, list):
        return [normalize_ref_siblings(v) for v in schema]
    return schema


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

    def validate_input(self, arguments: dict[str, Any], model: type[T]) -> T:
        """Pydantic 输入验证 | Pydantic input validation"""
        return model.model_validate(arguments)
