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
    RESOURCE_NOT_ACCESSIBLE = "3005"
    CONTENT_TOO_LARGE = "3006"
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
    OFFICE_API_ERROR = "3999"

    # Validation errors (4xxx)
    VALIDATION_ERROR = "4000"
    MISSING_PARAM = "4001"
    INVALID_PARAM = "4002"
    INVALID_PARAM_TYPE = "4003"
    PARAM_OUT_OF_RANGE = "4004"


# Forward reference resolution
BaseResponse.model_rebuild()
