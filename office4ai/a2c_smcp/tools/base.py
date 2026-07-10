# filename: base.py
# @Time    : 2025/12/18 16:07
# @Author  : JQQ
# @Email   : jqq1716@gmail.com
# @Software: PyCharm

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import Any, Literal, TypeVar, cast

from loguru import logger
from pydantic import BaseModel

from office4ai.environment.workspace.base import OfficeAction, OfficeObs
from office4ai.environment.workspace.office_workspace import OfficeWorkspace

T = TypeVar("T", bound=BaseModel)


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
    def category(self) -> Literal["word", "ppt", "excel", "authoring"]:  # pragma: no cover
        """平台类别: 'word' | 'ppt' | 'excel' | 'authoring' | Platform category

        'authoring' 标记不绑定单一 Office 平台的 standalone 工具（如 office_run_script），
        不投射到任一文件窗口资源，且 requires_connection 默认为 False（常驻，不随连接收敛）。
        """
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

    # ── 动态工具收敛（W4a / #63）| Dynamic tool convergence ──

    @property
    def requires_connection(self) -> bool:
        """本工具是否依赖 Add-In 连接（W4a 动态收敛按此过滤 list_tools）。

        默认按 ``category`` 派生：``word``/``ppt``/``excel`` 平台工具依赖对应 Add-In 连接
        （``True``）；``authoring`` 等 standalone 工具（如 ``office_run_script``）常驻
        （``False``），无连接也暴露。chart 工具（``category='ppt'``）经此默认即 ``True``——
        即便有 Path A 离线能力，也按 W4a 二元模型随连接收敛；Path A 代码删除见 F1(#68)。
        子类如需背离 category 语义可 override 本属性。
        """
        return self.category != "authoring"

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
        # 0. 守卫：默认 execute() 只服务平台工具（word/ppt/excel）。standalone 工具
        #    （category='authoring' 等）必须 override execute()——否则在此清晰失败，
        #    而非在下方构造 OfficeAction 时抛未捕获的 pydantic ValidationError。
        if self.category not in ("word", "ppt", "excel"):
            return {
                "success": False,
                "error": f"{self.name}: base execute() supports only word/ppt/excel tools; "
                "standalone tools must override execute()",
            }

        # 1. 验证输入
        try:
            validated = self.validate_input(arguments, self.input_model)
        except ValueError as e:
            return {"success": False, "error": str(e)}

        # 2. 提取参数
        params = validated.model_dump(exclude_none=True)
        document_uri = params.pop("document_uri")

        # 3. 构建 OfficeAction
        #    本默认 execute 只被平台工具（word/ppt/excel）走到；standalone 工具
        #    （如 office_run_script，category='authoring'）会完全 override execute，
        #    不会到这里。故此处收敛为三值 Literal 安全。
        action = OfficeAction(
            category=cast(Literal["word", "ppt", "excel"], self.category),
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
