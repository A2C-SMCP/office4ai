"""
Base Workspace Module

定义 Workspace 环境的基础接口和数据模型。
"""

from abc import ABC, abstractmethod
from enum import Enum
from typing import Any, Literal

from pydantic import BaseModel

# ============================================================================
# Data Models
# ============================================================================


class OfficeAction(BaseModel):
    """
    统一动作格式，通过 Workspace.execute() 执行

    Attributes:
        category: Office 应用类型 (word/excel/ppt)
        action_name: 操作名称 (如 "insert:text")
        params: 操作参数
        server_timeout_ms: 可选的 Server 侧 ack 超时覆写（毫秒）。为 None 时 workspace 用
            全局默认 ``request_timeout``。长脚本类事件（``{ns}:run:script``，issue #87）据此
            派生 ``(timeoutMs ?? 60000) + GRACE`` 覆写，避免全局 30s 先于脚本超时挂断、
            在 Server 层重演幽灵副作用。语义：脚本时长执法者是 Add-In，Server 超时仅失联兜底。
    """

    category: Literal["word", "excel", "ppt"]
    action_name: str
    params: dict[str, Any]
    server_timeout_ms: int | None = None


class OfficeObs(BaseModel):
    """
    统一观察格式，Workspace 执行动作后返回

    Attributes:
        success: 操作是否成功
        data: 结构化结果数据
        error: 错误信息 (如果失败)
        metadata: 元数据 (执行耗时、文档版本等)
    """

    success: bool
    data: dict[str, Any]
    error: str | None = None
    metadata: dict[str, Any] = {}


class DocumentStatus(Enum):
    """
    文档连接状态

    Attributes:
        CONNECTED: 文档已连接且活跃
        DISCONNECTED: 文档已断开
        UNKNOWN: 状态未知 (需要探测)
    """

    CONNECTED = "connected"
    DISCONNECTED = "disconnected"
    UNKNOWN = "unknown"


# ============================================================================
# Base Workspace
# ============================================================================


class BaseWorkspace(ABC):
    """
    Workspace 基类，负责文档会话管理

    Workspace 是一个工作会话管理器，类似于 VSCode 的 Workspace 概念。
    它负责：
    1. 管理 Office Add-In 的 Socket.IO 连接
    2. 提供统一动作接口 (execute)
    3. 维护文档状态
    4. 路由动作到对应的 Add-In
    """

    @abstractmethod
    async def execute(self, action: OfficeAction) -> OfficeObs:
        """
        执行统一动作接口

        Args:
            action: Office 动作对象

        Returns:
            OfficeObs: 执行结果
        """
        pass

    @abstractmethod
    def get_document_status(self, document_uri: str) -> DocumentStatus:
        """
        获取文档状态

        Args:
            document_uri: 文档 URI

        Returns:
            DocumentStatus: 文档连接状态
        """
        pass

    @abstractmethod
    async def emit_to_document(self, document_uri: str, event: str, data: dict[str, Any]) -> dict[str, Any]:
        """
        向指定文档发送 Socket.IO 事件

        Args:
            document_uri: 目标文档 URI
            event: 事件名称 (如 "word:get:selectedContent")
            data: 事件数据

        Returns:
            dict: Add-In 返回的响应数据

        Raises:
            ValueError: 如果文档未连接
            TimeoutError: 如果请求超时
        """
        pass

    @abstractmethod
    async def start(self) -> None:
        """
        启动 Workspace

        初始化 Socket.IO 服务器，开始监听连接
        """
        pass

    @abstractmethod
    async def stop(self) -> None:
        """
        停止 Workspace

        关闭 Socket.IO 服务器，清理资源
        """
        pass

    @property
    @abstractmethod
    def is_running(self) -> bool:
        """
        Workspace 是否正在运行

        Returns:
            bool: True 如果正在运行
        """
        pass
