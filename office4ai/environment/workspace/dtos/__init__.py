"""
Workspace Socket.IO DTOs

This package contains data transfer objects for Socket.IO communication
between office4ai (Workspace) and Office Add-In clients.

Architecture:
- These DTOs are internal to the Workspace environment
- They define the protocol for Add-In ↔ Workspace communication
- NOT to be confused with office4ai/dtos/ which are for A2C-SMCP protocol
"""

from office4ai.environment.workspace.dtos.common import (
    BaseRequest,
    BaseResponse,
    ErrorCode,
    ErrorResponse,
)

# Excel DTOs (OASP 0.3.0 events-excel.md, milestone #3 — #18 read slice + #19 Range + #20 Format)
from office4ai.environment.workspace.dtos.excel import (
    BorderFormat,
    ConditionalFormatRule,
    ExcelAddConditionalFormatRequest,
    ExcelClearConditionalFormatRequest,
    ExcelClearRangeRequest,
    ExcelCopyRangeRequest,
    ExcelDeleteRangeRequest,
    ExcelGetRangeFormatRequest,
    ExcelGetRangeRequest,
    ExcelGetSelectedRangeRequest,
    ExcelGetWorkbookInfoRequest,
    ExcelGetWorksheetInfoRequest,
    ExcelInsertRangeRequest,
    ExcelMergeCellsRequest,
    ExcelSetFormulaRequest,
    ExcelSetRangeFormatRequest,
    ExcelSetRangeRequest,
    ExcelUnmergeCellsRequest,
    GetRangeData,
    GetRangeFormatData,
    RangeBorders,
    RangeFillInfo,
    RangeFontInfo,
    RangeFormatInfo,
    RangeOperationResult,
    SelectedRangeInfo,
    SetAlignmentFormat,
    SetFillFormat,
    SetFontFormat,
    SetRangeFormatOptions,
    SheetInfo,
    UsedRangeInfo,
    WorkbookInfo,
    WorksheetInfo,
)

# PPT DTOs
from office4ai.environment.workspace.dtos.ppt import (
    PptDeleteSlideRequest,
    PptGetCurrentSlideElementsRequest,
    PptGetSlideElementsRequest,
    PptGetSlideScreenshotRequest,
    PptInsertImageRequest,
    PptInsertShapeRequest,
    PptInsertTableRequest,
    PptInsertTextRequest,
    PptMoveSlideRequest,
    PptUpdateTextBoxRequest,
)

# Word DTOs
from office4ai.environment.workspace.dtos.word import (
    SelectionInfo,
    SelectTextResult,
    WordAppendTextRequest,
    WordExportContentRequest,
    WordGetDocumentStatsRequest,
    WordGetDocumentStructureRequest,
    WordGetSelectedContentRequest,
    WordGetSelectionRequest,
    WordGetSelectionResponse,
    WordGetVisibleContentRequest,
    WordInsertEquationRequest,
    WordInsertImageRequest,
    WordInsertTableRequest,
    WordInsertTextRequest,
    WordInsertTOCRequest,
    WordReplaceSelectionRequest,
    WordReplaceTextRequest,
    WordSelectTextRequest,
)

__all__ = [
    # Common
    "BaseRequest",
    "BaseResponse",
    "ErrorResponse",
    "ErrorCode",
    # Word
    "WordGetSelectedContentRequest",
    "WordGetSelectionRequest",
    "WordGetSelectionResponse",
    "WordGetVisibleContentRequest",
    "WordGetDocumentStructureRequest",
    "WordGetDocumentStatsRequest",
    "WordInsertTextRequest",
    "WordReplaceSelectionRequest",
    "WordReplaceTextRequest",
    "WordSelectTextRequest",
    "SelectTextResult",
    "WordAppendTextRequest",
    "WordInsertImageRequest",
    "WordInsertTableRequest",
    "WordInsertEquationRequest",
    "WordInsertTOCRequest",
    "WordExportContentRequest",
    "SelectionInfo",
    # Excel read slice (#18)
    "ExcelGetWorkbookInfoRequest",
    "ExcelGetWorksheetInfoRequest",
    "ExcelGetSelectedRangeRequest",
    "SheetInfo",
    "UsedRangeInfo",
    "WorkbookInfo",
    "WorksheetInfo",
    "SelectedRangeInfo",
    # Excel Range CRUD + 公式 (#19)
    "ExcelGetRangeRequest",
    "ExcelSetRangeRequest",
    "ExcelClearRangeRequest",
    "ExcelCopyRangeRequest",
    "ExcelDeleteRangeRequest",
    "ExcelInsertRangeRequest",
    "ExcelSetFormulaRequest",
    "RangeFontInfo",
    "RangeFillInfo",
    "RangeFormatInfo",
    "GetRangeData",
    "RangeOperationResult",
    # Excel Format / 条件格式 / 合并单元格 (#20)
    "ExcelGetRangeFormatRequest",
    "ExcelSetRangeFormatRequest",
    "ExcelAddConditionalFormatRequest",
    "ExcelClearConditionalFormatRequest",
    "ExcelMergeCellsRequest",
    "ExcelUnmergeCellsRequest",
    "GetRangeFormatData",
    "BorderFormat",
    "RangeBorders",
    "SetFontFormat",
    "SetFillFormat",
    "SetAlignmentFormat",
    "SetRangeFormatOptions",
    "ConditionalFormatRule",
    # PPT
    "PptGetCurrentSlideElementsRequest",
    "PptGetSlideElementsRequest",
    "PptGetSlideScreenshotRequest",
    "PptInsertTextRequest",
    "PptInsertImageRequest",
    "PptInsertTableRequest",
    "PptInsertShapeRequest",
    "PptDeleteSlideRequest",
    "PptMoveSlideRequest",
    "PptUpdateTextBoxRequest",
]
