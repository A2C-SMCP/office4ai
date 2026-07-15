"""
Test common DTOs

测试通用 DTO 验证。
"""

import pytest
from pydantic import ValidationError

from office4ai.environment.workspace.dtos.common import (
    BaseRequest,
    BaseResponse,
    ErrorCode,
    ErrorResponse,
)


class TestBaseRequest:
    """Test BaseRequest DTO"""

    def test_valid_request(self) -> None:
        """Test creating valid request"""
        request = BaseRequest(requestId="req_123", documentUri="file:///test.docx")

        assert request.request_id == "req_123"
        assert request.document_uri == "file:///test.docx"
        assert isinstance(request.timestamp, int)

    def test_missing_request_id(self) -> None:
        """Test validation fails without requestId"""
        with pytest.raises(ValidationError) as exc_info:
            BaseRequest(documentUri="file:///test.docx")

        assert "requestId" in str(exc_info.value)

    def test_missing_document_uri(self) -> None:
        """Test validation fails without documentUri"""
        with pytest.raises(ValidationError) as exc_info:
            BaseRequest(requestId="req_123")

        assert "documentUri" in str(exc_info.value)

    def test_custom_timestamp(self) -> None:
        """Test custom timestamp"""
        ts = 1234567890000
        request = BaseRequest(requestId="req_123", documentUri="file:///test.docx", timestamp=ts)

        assert request.timestamp == ts


class TestBaseResponse:
    """Test BaseResponse DTO"""

    def test_success_response(self) -> None:
        """Test creating success response"""
        response = BaseResponse(requestId="req_123", success=True, data={"text": "Hello"}, timestamp=1234567890000)

        assert response.request_id == "req_123"
        assert response.success is True
        assert response.data == {"text": "Hello"}
        assert response.error is None

    def test_error_response(self) -> None:
        """Test creating error response"""
        error = ErrorResponse(code=ErrorCode.DOCUMENT_ERROR, message="Operation failed")

        response = BaseResponse(requestId="req_123", success=False, error=error, timestamp=1234567890000)

        assert response.success is False
        assert response.error.code == "3000"
        assert response.error.message == "Operation failed"

    def test_response_with_duration(self) -> None:
        """Test response with operation duration"""
        response = BaseResponse(requestId="req_123", success=True, timestamp=1234567890000, duration=150)

        assert response.duration == 150


class TestErrorResponse:
    """Test ErrorResponse DTO"""

    def test_error_response(self) -> None:
        """Test creating error response"""
        error = ErrorResponse(
            code=ErrorCode.VALIDATION_ERROR,
            message="Invalid parameter",
            details={"field": "text", "issue": "too long"},
        )

        assert error.code == "4000"
        assert error.message == "Invalid parameter"
        assert error.details == {"field": "text", "issue": "too long"}

    def test_error_without_details(self) -> None:
        """Test error without optional details"""
        error = ErrorResponse(code=ErrorCode.UNKNOWN, message="Something went wrong")

        assert error.details is None


class TestErrorCode:
    """Test ErrorCode constants"""

    def test_general_errors(self) -> None:
        """Test general error codes (1xxx)"""
        assert ErrorCode.UNKNOWN == "1000"
        assert ErrorCode.INVALID_REQUEST == "1001"
        assert ErrorCode.TIMEOUT == "1002"

    def test_document_errors(self) -> None:
        """Test document error codes (3xxx)"""
        assert ErrorCode.DOCUMENT_ERROR == "3000"
        assert ErrorCode.DOCUMENT_NOT_FOUND == "3001"
        assert ErrorCode.SELECTION_EMPTY == "3002"

    def test_excel_convergence_errors(self) -> None:
        """oasp#17 /excel 错误码收敛：5xxx 退役后新增的两个通用码"""
        assert ErrorCode.FORMULA_ERROR == "3017"
        assert ErrorCode.DATA_TYPE_MISMATCH == "3018"

    def test_validation_errors(self) -> None:
        """Test validation error codes (4xxx)"""
        assert ErrorCode.VALIDATION_ERROR == "4000"
        assert ErrorCode.MISSING_PARAM == "4001"
