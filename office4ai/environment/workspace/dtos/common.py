"""
Common Socket.IO DTOs

Defines base structures for Socket.IO communication between Workspace and Add-In clients.
"""

import threading
import uuid
from datetime import datetime
from typing import Any, ClassVar, Optional, Self

from pydantic import BaseModel, ConfigDict, Field


class Singleton(type):
    """
    Thread-safe singleton metaclass.

    Ensures that only one instance of a class exists per process.
    Uses double-checked locking for thread safety.

    Example:
        >>> class MyClass(metaclass=Singleton):
        ...     pass
        >>> a = MyClass()
        >>> b = MyClass()
        >>> a is b  # True
    """

    _instances: dict[type, Any] = {}
    _lock: threading.Lock = threading.Lock()

    def __call__(cls, *args: Any, **kwargs: Any) -> Any:
        """Create or return the singleton instance"""
        if cls not in cls._instances:
            with cls._lock:
                # Double-checked locking pattern
                if cls not in cls._instances:
                    cls._instances[cls] = super().__call__(*args, **kwargs)
        return cls._instances[cls]


class RequestRegistry(metaclass=Singleton):
    """
    Global registry for request DTO classes.

    Implements automatic registration via BaseRequest.__init_subclass__.
    Singleton pattern ensures uniqueness across the process lifecycle.

    Uses Singleton metaclass for thread-safe singleton implementation.

    Example:
        >>> from office4ai.environment.workspace.dtos.common import request_registry
        >>> dto_class = request_registry.get("word:get:selectedContent")
        >>> is_registered = request_registry.contains("word:get:selectedContent")
        >>> all_events = request_registry.all_events()
    """

    def __init__(self) -> None:
        """Initialize the registry (called once by singleton metaclass)"""
        self._registry: dict[str, type[BaseRequest]] = {}

    def register(self, event: str, cls: type["BaseRequest"]) -> None:
        """
        Register a request DTO class for an event.

        Args:
            event: Event name (e.g., "word:get:selectedContent")
            cls: Request DTO class

        Raises:
            ValueError: If event is already registered (prevents accidental overwrites)
        """
        if event in self._registry:
            raise ValueError(f"Event '{event}' is already registered with {self._registry[event].__name__}")
        self._registry[event] = cls

    def get(self, event: str) -> type["BaseRequest"] | None:
        """
        Get the DTO class for an event.

        Args:
            event: Event name

        Returns:
            DTO class if found, None otherwise
        """
        return self._registry.get(event)

    def contains(self, event: str) -> bool:
        """
        Check if an event is registered.

        Args:
            event: Event name

        Returns:
            True if registered, False otherwise
        """
        return event in self._registry

    def all_events(self) -> list[str]:
        """
        Get all registered event names.

        Returns:
            Sorted list of event names
        """
        return sorted(self._registry.keys())


# Global registry singleton
request_registry = RequestRegistry()


class SocketIOBaseModel(BaseModel):
    """
    Base model for all Socket.IO DTOs.

    Provides automatic snake_case ↔ camelCase conversion for protocol compliance:
    - Internal Python: snake_case (PEP 8 compliant)
    - External JSON: camelCase (Socket.IO protocol)

    Example:
        >>> class MyRequest(SocketIOBaseModel):
        ...     request_id: str = Field(alias="requestId")
        ...
        >>> # Internal: obj.request_id
        >>> # External: {"requestId": "..."}
    """

    model_config: ClassVar[ConfigDict] = ConfigDict(
        populate_by_name=True,  # Accept both alias and field name
    )


class BaseRequest(SocketIOBaseModel):
    """
    Base structure for Server → Client requests.

    When Workspace sends a command to an Add-In, all requests inherit from this.

    Uses Pydantic aliases for protocol compliance:
    - Internal: snake_case (PEP 8 compliant)
    - External: camelCase (Socket.IO protocol)

    Auto-registration:
        Subclasses with a non-empty event_name ClassVar are automatically
        registered to the global request_registry upon class definition.
    """

    # Subclasses must override this to enable auto-registration
    event_name: ClassVar[str] = ""  # Empty string means abstract base class

    request_id: str = Field(
        ...,
        alias="requestId",
        description="Unique request identifier for matching responses",
    )
    document_uri: str = Field(
        ...,
        alias="documentUri",
        description="Document URI (file:///path/to.docx)",
    )
    timestamp: int | None = Field(
        default_factory=lambda: int(datetime.now().timestamp() * 1000),
        alias="timestamp",
        description="Client timestamp in milliseconds",
    )

    def __init_subclass__(cls, **kwargs: Any) -> None:
        """
        Auto-register subclasses with non-empty event_name.

        This hook is called when a subclass is defined, enabling automatic
        registration without manual boilerplate.
        """
        super().__init_subclass__(**kwargs)
        # Only register concrete subclasses with an event_name
        if cls.event_name:
            request_registry.register(cls.event_name, cls)

    @classmethod
    def build(cls, document_uri: str, **business_params: Any) -> Self:
        """
        Build a request instance with auto-generated request_id.

        This is the recommended way to create request instances when you
        have the DTO class available.

        Args:
            document_uri: Document URI (file:///path/to.docx)
            **business_params: Business-specific parameters

        Returns:
            Request instance with auto-generated request_id and timestamp

        Example:
            >>> request = WordGetSelectedContentRequest.build(
            ...     document_uri="file:///test.docx",
            ...     options={"includeText": True}
            ... )
            >>> print(request.request_id)  # Auto-generated UUID
            >>> payload = request.to_payload()  # camelCase dict
        """
        # Use alias names for Pydantic fields
        return cls(
            requestId=str(uuid.uuid4()),
            documentUri=document_uri,
            **business_params,
        )

    @classmethod
    def from_event(cls, event: str, document_uri: str, **business_params: Any) -> "BaseRequest":
        """
        Build a request instance by event name (function-style interface).

        This is a compatibility wrapper for the old wrap_request() function.
        Useful when you only have the event name as a string.

        Args:
            event: Event name (e.g., "word:get:selectedContent")
            document_uri: Document URI (file:///path/to.docx)
            **business_params: Business-specific parameters

        Returns:
            Request instance with auto-generated request_id and timestamp

        Raises:
            RequestWrapperError: If event is not registered

        Example:
            >>> from office4ai.environment.workspace.socketio.request_wrapper import RequestWrapperError
            >>> try:
            ...     request = BaseRequest.from_event(
            ...         "word:get:selectedContent",
            ...         "file:///test.docx",
            ...         options={"includeText": True}
            ...     )
            ... except RequestWrapperError as e:
            ...     print(f"Unknown event: {e}")
        """
        dto_class = request_registry.get(event)
        if not dto_class:
            # Import here to avoid circular dependency
            from office4ai.environment.workspace.socketio.request_wrapper import RequestWrapperError

            raise RequestWrapperError(f"Unknown event '{event}'. Not registered in request_registry.")
        return dto_class.build(document_uri=document_uri, **business_params)

    def to_payload(self) -> dict[str, Any]:
        """
        Convert to camelCase JSON payload for Socket.IO transmission.

        Returns:
            dict with camelCase keys, ready for JSON serialization

        Example:
            >>> request = WordGetSelectedContentRequest.build(
            ...     document_uri="file:///test.docx",
            ...     options={"includeText": True}
            ... )
            >>> payload = request.to_payload()
            >>> print(payload)
            {
                "requestId": "a1b2c3d4-e5f6-4a5b-8c7d-9e0f1a2b3c4d",
                "documentUri": "file:///test.docx",
                "timestamp": 1234567890000,
                "options": {"includeText": True}
            }
        """
        return self.model_dump(by_alias=True, exclude_none=True)


class BaseResponse(SocketIOBaseModel):
    """
    Base structure for Client → Server responses.

    When Add-In responds to a Workspace command, all responses inherit from this.

    Uses Pydantic aliases for protocol compliance:
    - Internal: snake_case (PEP 8 compliant)
    - External: camelCase (Socket.IO protocol)
    """

    request_id: str = Field(
        ...,
        alias="requestId",
        description="Request ID being responded to",
    )
    success: bool = Field(..., alias="success", description="Whether the operation succeeded")
    data: dict[str, Any] | None = Field(
        default=None,
        alias="data",
        description="Response data",
    )
    error: Optional["ErrorResponse"] = Field(
        default=None,
        alias="error",
        description="Error details if failed",
    )
    timestamp: int = Field(
        ...,
        alias="timestamp",
        description="Server timestamp in milliseconds",
    )
    duration: int | None = Field(
        default=None,
        alias="duration",
        description="Operation duration in milliseconds",
    )


class ErrorResponse(SocketIOBaseModel):
    """
    Standardized error information.

    Uses Pydantic aliases for protocol compliance.
    """

    code: str = Field(
        ...,
        alias="code",
        description="Error code (e.g., '3000')",
    )
    message: str = Field(
        ...,
        alias="message",
        description="Human-readable error message",
    )
    details: dict[str, Any] | None = Field(
        default=None,
        alias="details",
        description="Additional error details",
    )


class ErrorCode:
    """
    Standard error codes for Socket.IO communication.

    Code ranges:
    - 1xxx: General errors
    - 2xxx: Authentication errors
    - 3xxx: Office API errors
    - 4xxx: Validation errors
    """

    # General errors (1xxx)
    UNKNOWN = "1000"
    INVALID_REQUEST = "1001"
    TIMEOUT = "1002"
    NOT_IMPLEMENTED = "1003"
    INTERNAL_ERROR = "1004"
    RATE_LIMITED = "1005"

    # Authentication errors (2xxx)
    UNAUTHORIZED = "2000"
    TOKEN_EXPIRED = "2001"
    INVALID_TOKEN = "2002"
    HANDSHAKE_FAILED = "2003"
    SESSION_INVALID = "2004"
    CONNECTION_LOST = "2005"
    PROTOCOL_VERSION_MISMATCH = "2006"  # OASP 0.3.0: oaspVersion incompatible at handshake

    # Document errors (3xxx)
    DOCUMENT_ERROR = "3000"
    DOCUMENT_NOT_FOUND = "3001"
    SELECTION_EMPTY = "3002"
    DOCUMENT_READ_ONLY = "3003"
    OPERATION_FAILED = "3004"
    # RESOURCE_NOT_ACCESSIBLE("3005") 已退役 (oasp#23)：无指涉对象——全协议无任何请求参数引用
    # 可拉取的外部资源（documentUri 已由 3001/3003 覆盖，图片等载荷一律 inline base64 自包含），
    # 自初始提交起从未被任何事件引用。不留别名。
    # ⚠️ 唯一潜在反例 oasp#26：若 imageInfo.data 补齐显式请求参数且需 materialize 外链图片，
    # 该退役前提失效、可能须重新引入——跟进该 issue 时先回看此处。
    CONTENT_TOO_LARGE = "3006"
    # 3007/3008/3012 经 oasp#23 裁决接线（详见 error-handling.md 各码「触发场景」小节）：
    # 3007 = 事件可用但**这一种格式**不受支持（换格式即可成功；区别于 4002 线缆层不可解码、
    #        3016 整个能力不可用）；输入侧转码重发、输出侧改请求另一格式
    # 3008 = 序号相对**当前文档状态**无效（重读状态后原值重试；区别于 4004 静态声明边界）。
    #        ⚠️ 过渡期：规范层效力仅及 ppt:{delete,goto}:slide，全量清扫见 oasp#24
    # 3012 = 以搜索文本定位的操作零匹配**且零元无良定义结果**（仅 word:insert:comment
    #        searchText 模式）；replace:text/select:text/find:values 的零元是正常 success
    FORMAT_NOT_SUPPORTED = "3007"
    POSITION_INVALID = "3008"
    RANGE_INVALID = "3009"
    ELEMENT_NOT_FOUND = "3010"
    STYLE_NOT_FOUND = "3011"
    SEARCH_NO_MATCH = "3012"
    NO_TABLE_AT_CURSOR = "3013"
    ALREADY_MERGED = "3014"
    INVALID_CHART_DATA = "3015"
    API_NOT_SUPPORTED = "3016"  # OASP 0.3.0: required capability/requirement set unavailable on client/platform
    FORMULA_ERROR = "3017"  # OASP oasp#17 (Unreleased): formula syntax error or invalid reference
    DATA_TYPE_MISMATCH = (
        "3018"  # OASP oasp#17 (Unreleased): apply-time value/type incompatible (≠ wire param type 4003)
    )
    # OFFICE_API_ERROR("3999") 已退役 (oasp#20)：非注册码且点名实现技术，语义并入 3000 DOCUMENT_ERROR，勿再引入

    # Validation errors (4xxx)
    VALIDATION_ERROR = "4000"
    MISSING_PARAM = "4001"
    INVALID_PARAM = "4002"
    INVALID_PARAM_TYPE = "4003"
    PARAM_OUT_OF_RANGE = "4004"


# ---------------------------------------------------------------------------
# Script execution envelope —— {excel,word,ppt}:run:script（OASP oasp#18，0.5.0 起，issue #87）
#
# 「封装层逃生舱」的宿主无关执行信封：承载「一段 JS 源码 + 其产出」。宿主差异（Excel /
# Word / PowerPoint 各自的对象模型）全部落在脚本内部注入的 ``context`` 上、不进入信封，
# 故三命名空间**共享同一结构**（与 WordFont ≠ PptFont 的有意零映射方向相反、各自适用）。
# 因此这两个结构与 BaseRequest / BaseResponse 同属**传输层**，定义在 common.py。
# 规范：oasp data-structures §脚本执行 / conventions §run-script。
# ---------------------------------------------------------------------------


class ScriptResult(SocketIOBaseModel):
    """``{ns}:run:script`` 成功响应的 ``data`` 结构（三命名空间共享）。

    office4ai 为**纯中转**：成功响应的 ``data`` 原样透传给上层 Agent，本结构提供 wire
    契约的**类型锚点**（与 TS ``ScriptResult`` 严格同步），供消费者/测试按需解析校验，
    而非在中转链上强制反序列化（沿用写操作「最小/透传返回」约定）。

    注意：与 ``office4ai.office.authoring.runtime.ScriptResult``（离线沙箱 5 键契约）
    同名但**不同结构**——本结构是**在线** Office.js 执行的 wire 响应，那个是**离线**
    Python 沙箱的落盘结果，两通道独立、勿混淆（见 ``office_run_script`` vs
    ``{ns}_run_script`` 工具描述互指）。

    大小限制（UTF-8 字节，超限由 Add-In 侧执法）：``result`` ≤ 512KB、``logs`` ≤ 100KB
    （超限置 ``logs_truncated=True``）——见 conventions §run-script。
    """

    result: Any = Field(
        default=None,
        alias="result",
        description="脚本 return 的值，已序列化为纯 JSON；无返回值时为 null",
    )
    logs: list[str] = Field(
        default_factory=list,
        alias="logs",
        description="脚本内 console.* 的输出，按调用顺序",
    )
    duration_ms: int = Field(
        ...,
        alias="durationMs",
        description="脚本执行耗时（毫秒）",
    )
    logs_truncated: bool = Field(
        default=False,
        alias="logsTruncated",
        description="日志是否因超限（logs ≤ 100KB）被截断",
    )


class RunScriptRequest(BaseRequest):
    """``{ns}:run:script`` 请求信封基类（宿主无关，抽象——不自注册）。

    三命名空间的请求结构**完全相同**，仅 ``event_name`` 不同。子类
    （``ExcelRunScriptRequest`` / ``WordRunScriptRequest`` / ``PptRunScriptRequest``）
    各声明自己的 ``event_name`` ClassVar，经 ``BaseRequest.__init_subclass__`` 自动注册。

    ``args`` / ``result`` 不做深度校验（纯 JSON 由调用链结构性保证）。
    ``timeout_ms`` 是**脚本执行**超时（由 Add-In 执法，超时回 ``1002``）；Server 侧的
    ack 超时另按 ``server_timeout = (timeoutMs ?? 60000) + GRACE`` 派生（见工具层），
    使 Add-In 成为超时执法者、Server 超时仅失联兜底。
    """

    # 抽象基类：event_name 留空 → 不自注册（仅三个具体子类注册）
    event_name: ClassVar[str] = ""

    script: str = Field(
        ...,
        alias="script",
        description="待执行 JS 源码；以 async 函数体语义执行，可用 return 返回结果",
    )
    args: dict[str, Any] | None = Field(
        default=None,
        alias="args",
        description="注入脚本的参数，脚本内经全局 `args` 读取；必须可 JSON 序列化",
    )
    timeout_ms: int | None = Field(
        default=None,
        alias="timeoutMs",
        description="脚本执行超时（毫秒）；缺省取「脚本执行」档默认值（60000），无硬上限",
    )


class ExcelRunScriptRequest(RunScriptRequest):
    """``excel:run:script`` —— Excel Add-In 注入 ``Excel.RequestContext`` 执行 JS。"""

    event_name: ClassVar[str] = "excel:run:script"


class WordRunScriptRequest(RunScriptRequest):
    """``word:run:script`` —— Word Add-In 注入 ``Word.RequestContext`` 执行 JS。"""

    event_name: ClassVar[str] = "word:run:script"


class PptRunScriptRequest(RunScriptRequest):
    """``ppt:run:script`` —— PowerPoint Add-In 注入 ``PowerPoint.RequestContext`` 执行 JS。"""

    event_name: ClassVar[str] = "ppt:run:script"


# Forward reference resolution
BaseResponse.model_rebuild()
