"""
Word Socket.IO DTOs

Defines data structures for Word-specific Socket.IO events.
"""

from typing import Any, ClassVar, Literal, Optional

from pydantic import Field

from .common import BaseRequest, ErrorResponse, SocketIOBaseModel

# ============================================================================
# Content Retrieval DTOs
# ============================================================================


class WordGetStylesRequest(BaseRequest):
    """
    Request to get all available styles from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:get:styles"

    options: Optional["GetStylesOptions"] = Field(
        default=None,
        alias="options",
        description="Style retrieval options",
    )


class WordGetSelectionRequest(BaseRequest):
    """
    Request to get selection information from Word document.

    Lightweight query that returns position information only.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/36569100
    """

    event_name: ClassVar[str] = "word:get:selection"


class SelectionInfo(SocketIOBaseModel):
    """
    Selection position information.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/36569100
    """

    is_empty: bool = Field(
        ...,
        alias="isEmpty",
        description="Whether the selection is empty (cursor point, no highlighted text)",
    )
    type: Literal["NoSelection", "InsertionPoint", "Normal"] = Field(
        ...,
        alias="type",
        description="Selection type",
    )
    start: int | None = Field(
        default=None,
        alias="start",
        description="Start position (character offset from document beginning)",
    )
    end: int | None = Field(
        default=None,
        alias="end",
        description="End position (character offset from document beginning)",
    )
    text: str | None = Field(
        default=None,
        alias="text",
        description="Selected text content",
    )


class WordGetSelectionResponse(SocketIOBaseModel):
    """
    Response for word:get:selection operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/36569100
    """

    data: SelectionInfo | None = Field(
        default=None,
        alias="data",
        description="Selection position information",
    )


class WordGetSelectedContentRequest(BaseRequest):
    """
    Request to get selected content from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:get:selectedContent"

    options: Optional["GetContentOptions"] = Field(
        default=None,
        alias="options",
        description="Content retrieval options",
    )


class WordGetSelectedContentResponse(SocketIOBaseModel):
    """
    Response for word:get:selectedContent operation.

    Uses Pydantic aliases for protocol compliance.
    """

    text: str = Field(..., alias="text", description="Selected text content")
    elements: list["AnyContentElement"] = Field(
        default_factory=list,
        alias="elements",
        description="Content elements (text, images, tables)",
    )
    metadata: Optional["ContentMetadata"] = Field(
        default=None,
        alias="metadata",
        description="Content metadata",
    )


class WordGetVisibleContentRequest(BaseRequest):
    """
    Request to get visible content from Word document.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30736386
    """

    event_name: ClassVar[str] = "word:get:visibleContent"

    options: Optional["GetContentOptions"] = Field(
        default=None,
        alias="options",
        description="Content retrieval options",
    )


class WordGetVisibleContentResponse(SocketIOBaseModel):
    """
    Response for word:get:visibleContent operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30736386
    """

    text: str = Field(..., alias="text", description="Visible text content")
    elements: list["AnyContentElement"] = Field(
        default_factory=list,
        alias="elements",
        description="Content elements (text, images, tables)",
    )
    metadata: Optional["ContentMetadata"] = Field(
        default=None,
        alias="metadata",
        description="Content metadata",
    )


class WordGetDocumentStructureRequest(BaseRequest):
    """
    Request to get Word document structure.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30769153
    """

    event_name: ClassVar[str] = "word:get:documentStructure"


class DocumentStructure(SocketIOBaseModel):
    """
    Document structure information.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30769153
    """

    paragraph_count: int = Field(
        ...,
        alias="paragraphCount",
        description="Number of paragraphs in the document",
    )
    table_count: int = Field(
        ...,
        alias="tableCount",
        description="Number of tables in the document",
    )
    image_count: int = Field(
        ...,
        alias="imageCount",
        description="Number of images in the document",
    )
    section_count: int = Field(
        ...,
        alias="sectionCount",
        description="Number of sections in the document",
    )


class WordGetDocumentStructureResponse(SocketIOBaseModel):
    """
    Response for word:get:documentStructure operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30769153
    """

    data: DocumentStructure = Field(
        ...,
        alias="data",
        description="Document structure information",
    )


class WordGetDocumentStatsRequest(BaseRequest):
    """
    Request to get Word document statistics.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30375938
    """

    event_name: ClassVar[str] = "word:get:documentStats"


class DocumentStats(SocketIOBaseModel):
    """
    Document statistics information.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30375938
    """

    word_count: int = Field(
        ...,
        alias="wordCount",
        description="Number of words in the document",
    )
    character_count: int = Field(
        ...,
        alias="characterCount",
        description="Number of characters in the document (including spaces and punctuation)",
    )
    paragraph_count: int = Field(
        ...,
        alias="paragraphCount",
        description="Number of paragraphs in the document (including empty paragraphs)",
    )


class WordGetDocumentStatsResponse(SocketIOBaseModel):
    """
    Response for word:get:documentStats operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30375938
    """

    data: DocumentStats = Field(
        ...,
        alias="data",
        description="Document statistics information",
    )


class GetContentOptions(SocketIOBaseModel):
    """
    Options for content retrieval operations.

    Uses Pydantic aliases for protocol compliance.
    """

    include_text: bool = Field(
        default=True,
        alias="includeText",
        description="Include text content",
    )
    include_images: bool = Field(
        default=True,
        alias="includeImages",
        description="Include images",
    )
    include_tables: bool = Field(
        default=True,
        alias="includeTables",
        description="Include tables",
    )
    include_content_controls: bool = Field(
        default=False,
        alias="includeContentControls",
        description="Include content controls",
    )
    detailed_metadata: bool = Field(
        default=False,
        alias="detailedMetadata",
        description="Include detailed metadata",
    )
    max_text_length: int | None = Field(
        default=None,
        alias="maxTextLength",
        description="Maximum text length",
    )


class ContentMetadata(SocketIOBaseModel):
    """
    Metadata for content retrieval results.

    Uses Pydantic aliases for protocol compliance.
    """

    is_empty: bool = Field(
        ...,
        alias="isEmpty",
        description="Whether the content is empty",
    )
    character_count: int = Field(
        ...,
        alias="characterCount",
        description="Number of characters in the content",
    )


class AnyContentElement(SocketIOBaseModel):
    """
    Generic content element for different content types.

    Uses Pydantic aliases for protocol compliance.
    """

    type: Literal["text", "image", "table", "other"] = Field(
        ...,
        alias="type",
        description="Element type",
    )
    content: dict[str, Any] = Field(
        ...,
        alias="content",
        description="Element content (varies by type)",
    )


# ============================================================================
# Text Operation DTOs
# ============================================================================


class WordInsertTextRequest(BaseRequest):
    """
    Request to insert text into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:text"

    text: str = Field(..., alias="text", description="Text to insert")
    location: Literal["Cursor", "Start", "End"] = Field(
        default="Cursor",
        alias="location",
        description="Insertion location",
    )
    format: Optional["TextFormat"] = Field(
        default=None,
        alias="format",
        description="Text formatting",
    )


class WordReplaceSelectionRequest(BaseRequest):
    """
    Request to replace selected content in Word document.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30605313
    """

    event_name: ClassVar[str] = "word:replace:selection"

    content: "ReplaceContent" = Field(
        ...,
        alias="content",
        description="Replacement content (text or images must be provided)",
    )


class WordReplaceSelectionResponse(SocketIOBaseModel):
    """
    Response for word:replace:selection operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30605313
    """

    replaced: bool = Field(..., alias="replaced", description="Whether content was replaced")
    character_count: int = Field(..., alias="characterCount", description="Number of characters replaced")


class WordReplaceTextRequest(BaseRequest):
    """
    Request to find and replace text in Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:replace:text"

    search_text: str = Field(
        ...,
        alias="searchText",
        description="Text to search for (max 255 characters, enforced by Word.js API)",
        max_length=255,
    )
    replace_text: str = Field(
        ...,
        alias="replaceText",
        description="Replacement text",
    )
    options: Optional["ReplaceOptions"] = Field(
        default=None,
        alias="options",
        description="Replace options",
    )
    format: Optional["TextFormat"] = Field(
        default=None,
        alias="format",
        description="Text formatting to apply to the replaced text",
    )


class WordSelectTextRequest(BaseRequest):
    """
    Request to search and select text in Word document.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/42467331
    """

    event_name: ClassVar[str] = "word:select:text"

    search_text: str = Field(
        ...,
        alias="searchText",
        description="Text to search for (max 255 characters, enforced by Word.js API)",
        max_length=255,
    )
    search_options: Optional["SelectTextSearchOptions"] = Field(
        default=None,
        alias="searchOptions",
        description="Search options (matchCase, matchWholeWord, matchWildcards)",
    )
    selection_mode: Literal["select", "start", "end"] = Field(
        default="select",
        alias="selectionMode",
        description="Selection mode: select/highlight text, start cursor at beginning, or end cursor at end",
    )
    select_index: int = Field(
        default=1,
        alias="selectIndex",
        description="Which match to select (1-based, default: 1)",
        ge=1,
    )


class SelectTextSearchOptions(SocketIOBaseModel):
    """
    Search options for word:select:text operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/42467331
    """

    match_case: bool = Field(
        default=False,
        alias="matchCase",
        description="Match case (case-sensitive search)",
    )
    match_whole_word: bool = Field(
        default=False,
        alias="matchWholeWord",
        description="Match whole word only",
    )
    match_wildcards: bool = Field(
        default=False,
        alias="matchWildcards",
        description="Use wildcards in search pattern",
    )


class SelectTextResult(SocketIOBaseModel):
    """
    Result for word:select:text operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/42467331
    """

    success: bool = Field(
        ...,
        alias="success",
        description="Whether text was successfully selected",
    )
    match_count: int = Field(
        ...,
        alias="matchCount",
        description="Total number of matches found",
    )
    selected_index: int = Field(
        ...,
        alias="selectedIndex",
        description="Index of the selected match (1-based)",
    )
    selected_text: str = Field(
        ...,
        alias="selectedText",
        description="Text that was selected",
    )
    selection_info: Optional["SelectionInfo"] = Field(
        default=None,
        alias="selectionInfo",
        description="Detailed selection information",
    )


class WordSelectTextResponse(SocketIOBaseModel):
    """
    Response for word:select:text operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/42467331
    """

    request_id: str = Field(..., alias="requestId", description="Request ID being responded to")
    success: bool = Field(..., alias="success", description="Whether the operation succeeded")
    data: SelectTextResult | None = Field(
        default=None,
        alias="data",
        description="Selection result with matchCount, selectedIndex, selectedText, and selectionInfo",
    )
    error: Optional["ErrorResponse"] = Field(
        default=None,
        alias="error",
        description="Error details if failed",
    )
    timestamp: int = Field(..., alias="timestamp", description="Server timestamp in milliseconds")


class WordAppendTextRequest(BaseRequest):
    """
    Request to append text to Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:append:text"

    text: str = Field(..., alias="text", description="Text to append")
    location: Literal["Start", "End"] = Field(
        default="End",
        alias="location",
        description="Append location",
    )
    format: Optional["TextFormat"] = Field(
        default=None,
        alias="format",
        description="Text formatting",
    )


class TextFormat(SocketIOBaseModel):
    """
    Text formatting options.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/29753356/word+insert+text

    Priority Rule (Important):
        - Direct formatting (bold/italic/fontSize/etc) takes precedence over styleName
        - If any direct format fields are provided, styleName is ignored
        - If only styleName is provided, Word style is applied
        - If neither is provided, default formatting is used

    Examples:
        # ❌ Not recommended: styleName will be ignored when direct format is present
        format = {"bold": True, "style_name": "Heading 1"}  # Only bold takes effect

        # ✅ Recommended: Use Word style only
        format = {"style_name": "Heading 1"}  # Apply Heading 1 style

        # ✅ Recommended: Use direct format only
        format = {"bold": True, "color": "#FF0000"}  # Precise format control
    """

    bold: bool | None = Field(default=None, alias="bold", description="Bold text")
    italic: bool | None = Field(
        default=None,
        alias="italic",
        description="Italic text",
    )
    font_size: int | None = Field(
        default=None,
        alias="fontSize",
        description="Font size",
    )
    font_name: str | None = Field(
        default=None,
        alias="fontName",
        description="Font name",
    )
    color: str | None = Field(
        default=None,
        alias="color",
        description="Font color (hex)",
    )
    underline: (
        Literal[
            "Mixed",
            "None",
            "Hidden",
            "DotLine",
            "Single",
            "Word",
            "Double",
            "Thick",
            "Dotted",
            "DottedHeavy",
            "DashLine",
            "DashLineHeavy",
            "DashLineLong",
            "DashLineLongHeavy",
            "DotDashLine",
            "DotDashLineHeavy",
            "TwoDotDashLine",
            "TwoDotDashLineHeavy",
            "Wave",
            "WaveHeavy",
            "WaveDouble",
        ]
        | None
    ) = Field(
        default=None,
        alias="underline",
        description="Underline type (Word.UnderlineType)",
    )
    style_name: str | None = Field(
        default=None,
        alias="styleName",
        description="Word style name (e.g., 'Heading 1', 'Normal', 'Title')",
    )


class ReplaceContent(SocketIOBaseModel):
    """
    Content for replacement operation.

    Uses Pydantic aliases for protocol compliance.
    """

    text: str | None = Field(
        default=None,
        alias="text",
        description="Replacement text",
    )
    images: list[dict[str, Any]] | None = Field(
        default=None,
        alias="images",
        description="Replacement images",
    )
    format: Optional["TextFormat"] = Field(
        default=None,
        alias="format",
        description="Text formatting",
    )


class ReplaceOptions(SocketIOBaseModel):
    """
    Options for find and replace operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30801921
    """

    match_case: bool = Field(
        default=False,
        alias="matchCase",
        description="Match case",
    )
    match_whole_word: bool = Field(
        default=False,
        alias="matchWholeWord",
        description="Match whole word",
    )
    replace_all: bool = Field(
        default=False,
        alias="replaceAll",
        description="Replace all occurrences",
    )


class ReplaceTextResult(SocketIOBaseModel):
    """
    Result for find and replace operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30801921
    """

    replace_count: int = Field(
        ...,
        alias="replaceCount",
        description="Number of replacements made",
    )


class WordReplaceTextResponse(SocketIOBaseModel):
    """
    Response for word:replace:text operation.

    Uses Pydantic aliases for protocol compliance.

    Confluence Spec: https://turingfocus.atlassian.net/wiki/pages/30801921
    """

    request_id: str = Field(..., alias="requestId", description="Request ID being responded to")
    success: bool = Field(..., alias="success", description="Whether the operation succeeded")
    data: ReplaceTextResult | None = Field(
        default=None,
        alias="data",
        description="Replace result with replaceCount",
    )
    error: Optional["ErrorResponse"] = Field(
        default=None,
        alias="error",
        description="Error details if failed",
    )
    timestamp: int = Field(..., alias="timestamp", description="Server timestamp in milliseconds")


# ============================================================================
# Multimedia Operation DTOs
# ============================================================================


class WordInsertImageRequest(BaseRequest):
    """
    Request to insert image into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:image"

    image: "ImageData" = Field(..., alias="image", description="Image data")
    location: Optional["InsertLocation"] = Field(
        default=None,
        alias="location",
        description="Insertion location",
    )
    wrap_type: Literal["Inline", "Square", "Tight", "Behind", "InFront"] | None = Field(
        default="Inline",
        alias="wrapType",
        description="Text wrapping type",
    )


class WordInsertTableRequest(BaseRequest):
    """
    Request to insert table into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:table"

    options: "TableInsertOptions" = Field(
        ...,
        alias="options",
        description="Table insertion options",
    )


class WordInsertEquationRequest(BaseRequest):
    """
    Request to insert equation into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:equation"

    latex: str = Field(..., alias="latex", description="LaTeX equation string")
    options: Optional["EquationOptions"] = Field(
        default=None,
        alias="options",
        description="Equation options",
    )


class ImageData(SocketIOBaseModel):
    """
    Image data for insertion.

    Uses Pydantic aliases for protocol compliance.
    """

    base64: str = Field(..., alias="base64", description="Base64 encoded image")
    width: int | None = Field(
        default=None,
        alias="width",
        description="Image width",
    )
    height: int | None = Field(
        default=None,
        alias="height",
        description="Image height",
    )
    alt_text: str | None = Field(
        default=None,
        alias="altText",
        description="Alternative text",
    )


class InsertLocation(SocketIOBaseModel):
    """
    Insertion location details.

    Uses Pydantic aliases for protocol compliance.
    """

    type: Literal["Cursor", "Start", "End", "BeforeBookmark", "AfterBookmark"] = Field(
        ...,
        alias="type",
        description="Location type",
    )
    bookmark_name: str | None = Field(
        default=None,
        alias="bookmarkName",
        description="Bookmark name if applicable",
    )


class TableInsertOptions(SocketIOBaseModel):
    """
    Table insertion options.

    Uses Pydantic aliases for protocol compliance.

    OASP Spec: https://a2c-smcp.github.io/oasp-protocol/specification/events-word/#wordinserttable
    """

    rows: int = Field(..., alias="rows", description="Number of rows", ge=1)
    columns: int = Field(..., alias="columns", description="Number of columns", ge=1)
    data: list[list[str]] | None = Field(
        default=None,
        alias="data",
        description="Table data",
    )
    style: str | None = Field(
        default=None,
        alias="style",
        description="Table style name",
    )
    insert_location: Literal["Start", "End", "Before", "After", "Replace"] | None = Field(
        default=None,
        alias="insertLocation",
        description=(
            "Where to insert the table relative to the current selection. "
            "Start/End: document start/end; Before/After: relative to current selection/cursor paragraph; "
            "Replace: replace current selection. When omitted, Add-In defaults to 'End' (backward compatible)."
        ),
    )


class EquationOptions(SocketIOBaseModel):
    """
    Equation insertion options.

    Uses Pydantic aliases for protocol compliance.
    """

    inline: bool = Field(
        default=True,
        alias="inline",
        description="Inline equation",
    )


# ============================================================================
# Advanced Feature DTOs
# ============================================================================


class WordInsertTOCRequest(BaseRequest):
    """
    Request to insert table of contents into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:toc"

    options: Optional["TOCOptions"] = Field(
        default=None,
        alias="options",
        description="TOC options",
    )


class WordExportContentRequest(BaseRequest):
    """
    Request to export content from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:export:content"

    format: Literal["text", "html", "markdown"] = Field(
        ...,
        alias="format",
        description="Export format",
    )
    options: Optional["ExportOptions"] = Field(
        default=None,
        alias="options",
        description="Export options",
    )


class TOCOptions(SocketIOBaseModel):
    """
    Table of contents options.

    Uses Pydantic aliases for protocol compliance.
    """

    max_level: int = Field(
        default=3,
        alias="maxLevel",
        description="Maximum heading level",
        ge=1,
        le=9,
    )
    styles: list[str] | None = Field(
        default=None,
        alias="styles",
        description="Heading styles to include",
    )


class ExportOptions(SocketIOBaseModel):
    """
    Content export options.

    Uses Pydantic aliases for protocol compliance.
    """

    include_images: bool = Field(
        default=False,
        alias="includeImages",
        description="Include images in export",
    )
    include_tables: bool = Field(
        default=True,
        alias="includeTables",
        description="Include tables in export",
    )


# ============================================================================
# Style DTOs
# ============================================================================


class GetStylesOptions(SocketIOBaseModel):
    """
    Options for retrieving styles from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    include_built_in: bool = Field(
        default=True,
        alias="includeBuiltIn",
        description="Include built-in styles",
    )
    include_custom: bool = Field(
        default=True,
        alias="includeCustom",
        description="Include custom styles",
    )
    include_unused: bool = Field(
        default=False,
        alias="includeUnused",
        description="Include unused styles",
    )
    detailed_info: bool = Field(
        default=False,
        alias="detailedInfo",
        description="Include detailed information like description",
    )


class StyleInfo(SocketIOBaseModel):
    """
    Information about a Word style.

    Uses Pydantic aliases for protocol compliance.
    """

    name: str = Field(..., alias="name", description="Style name (localized)")
    type: Literal["Paragraph", "Character", "Table", "List"] = Field(
        ...,
        alias="type",
        description="Style type",
    )
    built_in: bool = Field(..., alias="builtIn", description="Whether it's a built-in style")
    in_use: bool = Field(..., alias="inUse", description="Whether the style is used in document")
    description: str | None = Field(
        default=None,
        alias="description",
        description="Style description (only when detailedInfo=true)",
    )


class StylesResult(SocketIOBaseModel):
    """
    Result containing styles from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    styles: list[StyleInfo] = Field(
        ...,
        alias="styles",
        description="List of styles",
    )


# ============================================================================
# Comment DTOs
# ============================================================================


class GetCommentsOptions(SocketIOBaseModel):
    """
    Options for retrieving comments from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    include_resolved: bool = Field(
        default=False,
        alias="includeResolved",
        description="Include resolved comments",
    )
    include_replies: bool = Field(
        default=True,
        alias="includeReplies",
        description="Include reply threads",
    )
    include_associated_text: bool = Field(
        default=False,
        alias="includeAssociatedText",
        description="Include the text associated with each comment",
    )
    detailed_metadata: bool = Field(
        default=False,
        alias="detailedMetadata",
        description="Include detailed metadata (author email, creation date)",
    )
    max_text_length: int | None = Field(
        default=None,
        alias="maxTextLength",
        description="Maximum length of associated text to return",
    )


class CommentReplyData(SocketIOBaseModel):
    """
    Data for a comment reply.

    Uses Pydantic aliases for protocol compliance.
    """

    id: str = Field(..., alias="id", description="Reply ID")
    content: str = Field(..., alias="content", description="Reply text content")
    author_name: str | None = Field(
        default=None,
        alias="authorName",
        description="Reply author name",
    )
    author_email: str | None = Field(
        default=None,
        alias="authorEmail",
        description="Reply author email",
    )
    creation_date: str | None = Field(
        default=None,
        alias="creationDate",
        description="Reply creation date (ISO 8601 format)",
    )


class CommentData(SocketIOBaseModel):
    """
    Data for a document comment.

    Uses Pydantic aliases for protocol compliance.
    """

    id: str = Field(..., alias="id", description="Comment ID")
    content: str = Field(..., alias="content", description="Comment text content")
    author_name: str | None = Field(
        default=None,
        alias="authorName",
        description="Comment author name",
    )
    author_email: str | None = Field(
        default=None,
        alias="authorEmail",
        description="Comment author email",
    )
    creation_date: str | None = Field(
        default=None,
        alias="creationDate",
        description="Comment creation date (ISO 8601 format)",
    )
    resolved: bool = Field(
        default=False,
        alias="resolved",
        description="Whether the comment is resolved",
    )
    associated_text: str | None = Field(
        default=None,
        alias="associatedText",
        description="Text associated with the comment",
    )
    replies: list["CommentReplyData"] | None = Field(
        default=None,
        alias="replies",
        description="Reply threads for this comment",
    )


class InsertCommentSearchOptions(SocketIOBaseModel):
    """
    Search options for targeting comment insertion by text search.

    Uses Pydantic aliases for protocol compliance.
    """

    match_case: bool = Field(
        default=False,
        alias="matchCase",
        description="Match case (case-sensitive search)",
    )
    match_whole_word: bool = Field(
        default=False,
        alias="matchWholeWord",
        description="Match whole word only",
    )


class InsertCommentTarget(SocketIOBaseModel):
    """
    Target location for inserting a comment.

    Uses Pydantic aliases for protocol compliance.
    """

    type: Literal["selection", "searchText"] = Field(
        ...,
        alias="type",
        description="Target type: 'selection' (current selection) or 'searchText' (find text first)",
    )
    search_text: str | None = Field(
        default=None,
        alias="searchText",
        description="Text to search for (required when type is 'searchText')",
    )
    search_options: InsertCommentSearchOptions | None = Field(
        default=None,
        alias="searchOptions",
        description="Search options (only used when type is 'searchText')",
    )


class WordGetCommentsRequest(BaseRequest):
    """
    Request to get comments from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:get:comments"

    options: Optional["GetCommentsOptions"] = Field(
        default=None,
        alias="options",
        description="Comment retrieval options",
    )


class WordGetCommentsResponse(SocketIOBaseModel):
    """
    Response for word:get:comments operation.

    Uses Pydantic aliases for protocol compliance.
    """

    comments: list[CommentData] = Field(
        default_factory=list,
        alias="comments",
        description="List of comments in the document",
    )


class WordInsertCommentRequest(BaseRequest):
    """
    Request to insert a comment into Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:insert:comment"

    text: str = Field(..., alias="text", description="Comment text content")
    target: Optional["InsertCommentTarget"] = Field(
        default=None,
        alias="target",
        description="Target location (defaults to current selection)",
    )


class WordDeleteCommentRequest(BaseRequest):
    """
    Request to delete a comment from Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:delete:comment"

    comment_id: str = Field(
        ...,
        alias="commentId",
        description="ID of the comment to delete",
    )


class WordReplyCommentRequest(BaseRequest):
    """
    Request to reply to a comment in Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:reply:comment"

    comment_id: str = Field(
        ...,
        alias="commentId",
        description="ID of the comment to reply to",
    )
    text: str = Field(..., alias="text", description="Reply text content")


class WordResolveCommentRequest(BaseRequest):
    """
    Request to resolve or unresolve a comment in Word document.

    Uses Pydantic aliases for protocol compliance.
    """

    event_name: ClassVar[str] = "word:resolve:comment"

    comment_id: str = Field(
        ...,
        alias="commentId",
        description="ID of the comment to resolve/unresolve",
    )
    resolved: bool = Field(
        default=True,
        alias="resolved",
        description="True to resolve, False to unresolve",
    )


# Resolve forward references
GetCommentsOptions.model_rebuild()
CommentReplyData.model_rebuild()
CommentData.model_rebuild()
InsertCommentSearchOptions.model_rebuild()
InsertCommentTarget.model_rebuild()
WordGetCommentsResponse.model_rebuild()
SelectionInfo.model_rebuild()
WordGetSelectionResponse.model_rebuild()
GetContentOptions.model_rebuild()
TextFormat.model_rebuild()
ReplaceContent.model_rebuild()
ReplaceOptions.model_rebuild()
ReplaceTextResult.model_rebuild()
WordReplaceTextResponse.model_rebuild()
SelectTextSearchOptions.model_rebuild()
SelectTextResult.model_rebuild()
WordSelectTextResponse.model_rebuild()
ImageData.model_rebuild()
InsertLocation.model_rebuild()
TableInsertOptions.model_rebuild()
EquationOptions.model_rebuild()
TOCOptions.model_rebuild()
ExportOptions.model_rebuild()
GetStylesOptions.model_rebuild()
StyleInfo.model_rebuild()
StylesResult.model_rebuild()
WordReplaceSelectionResponse.model_rebuild()
ContentMetadata.model_rebuild()
AnyContentElement.model_rebuild()
WordGetSelectedContentResponse.model_rebuild()
WordGetVisibleContentResponse.model_rebuild()
DocumentStructure.model_rebuild()
WordGetDocumentStructureResponse.model_rebuild()
DocumentStats.model_rebuild()
WordGetDocumentStatsResponse.model_rebuild()
