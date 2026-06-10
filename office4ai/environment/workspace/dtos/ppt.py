"""
PowerPoint Socket.IO DTOs

Defines data structures for PowerPoint-specific Socket.IO events.

Field naming convention (aligned with Word DTOs):
- Python field names: snake_case (PEP 8)
- Wire format aliases: camelCase (OASP protocol)
- populate_by_name=True: accept both snake_case and camelCase on input
- model_dump(by_alias=True): always output camelCase for Socket.IO transmission
"""

from typing import Annotated, ClassVar, Literal, Optional

from pydantic import Field

from .common import BaseRequest, SocketIOBaseModel

# ============================================================================
# Chart Type Enumerations (OASP /ppt Draft, v0.2.0)
# ============================================================================
# Chart types align with `Excel.ChartType` so callers can cast directly.
# Two shape variants:
#   - Categorical (9 types): use `categories` + `series.values`
#   - Scatter (1 type): use `series.points` (each point is { x, y })

CategoricalChartType = Literal[
    "ColumnClustered",
    "ColumnStacked",
    "BarClustered",
    "Line",
    "LineMarkers",
    "Pie",
    "Doughnut",
    "Area",
    "Radar",
]

ScatterChartType = Literal["Scatter"]

ChartType = Literal[
    "ColumnClustered",
    "ColumnStacked",
    "BarClustered",
    "Line",
    "LineMarkers",
    "Pie",
    "Doughnut",
    "Area",
    "Radar",
    "Scatter",
]

# ============================================================================
# Content Retrieval DTOs
# ============================================================================


class PptGetCurrentSlideElementsRequest(BaseRequest):
    """
    Request to get current slide elements.
    """

    event_name: ClassVar[str] = "ppt:get:currentSlideElements"

    pass


class PptGetSlideElementsRequest(BaseRequest):
    """
    Request to get elements from a specific slide.
    """

    event_name: ClassVar[str] = "ppt:get:slideElements"

    slide_index: int = Field(..., alias="slideIndex", description="Slide index (0-based)", ge=0)
    options: Optional["SlideElementsOptions"] = Field(default=None, description="Elements retrieval options")


class PptGetSlideScreenshotRequest(BaseRequest):
    """
    Request to get slide screenshot.
    """

    event_name: ClassVar[str] = "ppt:get:slideScreenshot"

    slide_index: int = Field(..., alias="slideIndex", description="Slide index (0-based)", ge=0)
    options: Optional["ScreenshotOptions"] = Field(default=None, description="Screenshot options")


class SlideElementsOptions(SocketIOBaseModel):
    """
    Options for slide elements retrieval.
    """

    include_text: bool = Field(default=True, alias="includeText", description="Include text elements")
    include_images: bool = Field(default=True, alias="includeImages", description="Include image elements")
    include_shapes: bool = Field(default=True, alias="includeShapes", description="Include shape elements")
    include_tables: bool = Field(default=True, alias="includeTables", description="Include table elements")
    include_charts: bool = Field(default=True, alias="includeCharts", description="Include chart elements")


class ScreenshotOptions(SocketIOBaseModel):
    """
    Screenshot generation options.
    """

    format: Literal["png", "jpeg"] = Field(default="png", description="Image format")
    quality: int | None = Field(default=None, description="Image quality (0-100), JPEG only", ge=0, le=100)


# ============================================================================
# Content Insertion DTOs
# ============================================================================


class PptInsertTextRequest(BaseRequest):
    """
    Request to insert text into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:text"

    text: str = Field(..., description="Text to insert")
    options: Optional["TextInsertOptions"] = Field(default=None, description="Insertion options")


class PptInsertImageRequest(BaseRequest):
    """
    Request to insert image into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:image"

    image: "SlideImageData" = Field(..., description="Image data")
    options: Optional["ElementInsertOptions"] = Field(default=None, description="Insertion options")


class PptInsertTableRequest(BaseRequest):
    """
    Request to insert table into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:table"

    options: "SlideTableInsertOptions" = Field(..., description="Table insertion options")


class PptInsertShapeRequest(BaseRequest):
    """
    Request to insert shape into PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:insert:shape"

    shape_type: Literal[
        "Rectangle",
        "RoundedRectangle",
        "Circle",
        "Oval",
        "Triangle",
        "Line",
        "Arrow",
        "Star",
        "TextBox",
    ] = Field(..., alias="shapeType", description="Shape type")
    options: Optional["ShapeInsertOptions"] = Field(default=None, description="Insertion options")


class TextInsertOptions(SocketIOBaseModel):
    """
    Text insertion options.
    """

    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    width: float | None = Field(default=None, description="Width (points)")
    height: float | None = Field(default=None, description="Height (points)")
    font_size: int | None = Field(default=None, alias="fontSize", description="Font size")
    font_name: str | None = Field(default=None, alias="fontName", description="Font name")
    color: str | None = Field(default=None, description="Font color (hex)")
    fill_color: str | None = Field(default=None, alias="fillColor", description="Fill color (hex)")


class SlideImageData(SocketIOBaseModel):
    """
    Image data for slide insertion.
    """

    base64: str = Field(..., description="Base64 encoded image")


class ElementInsertOptions(SocketIOBaseModel):
    """
    Common element insertion options.
    """

    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    width: float | None = Field(default=None, description="Width (points)")
    height: float | None = Field(default=None, description="Height (points)")


class SlideTableInsertOptions(SocketIOBaseModel):
    """
    Table insertion options for slides.
    """

    rows: int = Field(..., description="Number of rows", ge=1)
    columns: int = Field(..., description="Number of columns", ge=1)
    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    data: list[list[str]] | None = Field(default=None, description="Table data")


class ShapeInsertOptions(SocketIOBaseModel):
    """
    Shape insertion options.
    """

    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (default: current)")
    left: float | None = Field(default=None, description="Left position (points)")
    top: float | None = Field(default=None, description="Top position (points)")
    width: float | None = Field(default=None, description="Width (points)")
    height: float | None = Field(default=None, description="Height (points)")
    fill_color: str | None = Field(default=None, alias="fillColor", description="Fill color (hex)")
    border_color: str | None = Field(default=None, alias="borderColor", description="Border color (hex)")
    border_width: float | None = Field(default=None, alias="borderWidth", description="Border width (points)")
    text: str | None = Field(default=None, description="Shape text")


# ============================================================================
# Slide Management DTOs
# ============================================================================


class PptDeleteSlideRequest(BaseRequest):
    """
    Request to delete a slide.
    """

    event_name: ClassVar[str] = "ppt:delete:slide"

    slide_index: int = Field(..., alias="slideIndex", description="Slide index (0-based)", ge=0)


class PptMoveSlideRequest(BaseRequest):
    """
    Request to move a slide to a new position.
    """

    event_name: ClassVar[str] = "ppt:move:slide"

    from_index: int = Field(..., alias="fromIndex", description="Current slide index (0-based)", ge=0)
    to_index: int = Field(..., alias="toIndex", description="Target slide index (0-based)", ge=0)


# ============================================================================
# Update Operation DTOs
# ============================================================================


class PptUpdateTextBoxRequest(BaseRequest):
    """
    Request to update a text box in PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:update:textBox"

    element_id: str = Field(..., alias="elementId", description="Element ID to update")
    updates: "TextBoxUpdates" = Field(..., description="Updates to apply")


class TextBoxUpdates(SocketIOBaseModel):
    """
    Text box update fields.
    """

    text: str | None = Field(default=None, description="New text content")
    font_size: int | None = Field(default=None, alias="fontSize", description="Font size")
    font_name: str | None = Field(default=None, alias="fontName", description="Font name")
    color: str | None = Field(default=None, description="Font color (hex)")
    fill_color: str | None = Field(default=None, alias="fillColor", description="Fill color (hex)")
    bold: bool | None = Field(default=None, description="Bold text")
    italic: bool | None = Field(default=None, description="Italic text")


class PptGetSlideInfoRequest(BaseRequest):
    """
    Request to get presentation/slide basic info.
    """

    event_name: ClassVar[str] = "ppt:get:slideInfo"

    slide_index: int | None = Field(
        default=None, alias="slideIndex", description="Slide index (0-based), optional", ge=0
    )


class SlideLayoutsOptions(SocketIOBaseModel):
    """
    Options for slide layouts retrieval.
    """

    include_placeholders: bool = Field(
        default=True, alias="includePlaceholders", description="Include placeholder details"
    )


class PptGetSlideLayoutsRequest(BaseRequest):
    """
    Request to get available slide layout templates.
    """

    event_name: ClassVar[str] = "ppt:get:slideLayouts"

    options: Optional["SlideLayoutsOptions"] = Field(default=None, description="Layouts retrieval options")


class ImageUpdateOptions(SocketIOBaseModel):
    """
    Options for image update.
    """

    keep_dimensions: bool = Field(default=True, alias="keepDimensions", description="Keep original dimensions")
    width: float | None = Field(default=None, description="New width (points)")
    height: float | None = Field(default=None, description="New height (points)")


class PptUpdateImageRequest(BaseRequest):
    """
    Request to replace image content in PowerPoint slide.
    """

    event_name: ClassVar[str] = "ppt:update:image"

    element_id: str = Field(..., alias="elementId", description="Image element ID to update")
    image: "SlideImageData" = Field(..., description="New image data")
    options: Optional["ImageUpdateOptions"] = Field(default=None, description="Update options")


class TableCellUpdate(SocketIOBaseModel):
    """
    Single table cell update.
    """

    row_index: int = Field(..., alias="rowIndex", description="Row index (0-based)", ge=0)
    column_index: int = Field(..., alias="columnIndex", description="Column index (0-based)", ge=0)
    text: str = Field(..., description="New text content")


class PptUpdateTableCellRequest(BaseRequest):
    """
    Request to update specific table cells.
    """

    event_name: ClassVar[str] = "ppt:update:tableCell"

    element_id: str = Field(..., alias="elementId", description="Table element ID")
    cells: list["TableCellUpdate"] = Field(..., description="Cells to update", min_length=1)


class RowUpdate(SocketIOBaseModel):
    """
    Row-level batch update.
    """

    row_index: int = Field(..., alias="rowIndex", description="Row index (0-based)", ge=0)
    values: list[str] = Field(..., description="Values for each column in the row")


class ColumnUpdate(SocketIOBaseModel):
    """
    Column-level batch update.
    """

    column_index: int = Field(..., alias="columnIndex", description="Column index (0-based)", ge=0)
    values: list[str] = Field(..., description="Values for each row in the column")


class PptUpdateTableRowColumnRequest(BaseRequest):
    """
    Request to batch update table by row or column.
    """

    event_name: ClassVar[str] = "ppt:update:tableRowColumn"

    element_id: str = Field(..., alias="elementId", description="Table element ID")
    rows: list["RowUpdate"] | None = Field(default=None, description="Row updates")
    columns: list["ColumnUpdate"] | None = Field(default=None, description="Column updates")


class CellFormat(SocketIOBaseModel):
    """
    Format for a specific table cell.
    """

    row_index: int = Field(..., alias="rowIndex", description="Row index (0-based)", ge=0)
    column_index: int = Field(..., alias="columnIndex", description="Column index (0-based)", ge=0)
    background_color: str | None = Field(default=None, alias="backgroundColor", description="Background color (hex)")
    font_size: int | None = Field(default=None, alias="fontSize", description="Font size")
    font_color: str | None = Field(default=None, alias="fontColor", description="Font color (hex)")
    bold: bool | None = Field(default=None, description="Bold text")
    italic: bool | None = Field(default=None, description="Italic text")
    horizontal_alignment: Literal["Left", "Center", "Right"] | None = Field(
        default=None, alias="horizontalAlignment", description="Horizontal alignment"
    )
    vertical_alignment: Literal["Top", "Middle", "Bottom"] | None = Field(
        default=None, alias="verticalAlignment", description="Vertical alignment"
    )


class RowFormat(SocketIOBaseModel):
    """
    Format for a table row.
    """

    row_index: int = Field(..., alias="rowIndex", description="Row index (0-based)", ge=0)
    height: float | None = Field(default=None, description="Row height (points)")
    background_color: str | None = Field(default=None, alias="backgroundColor", description="Background color (hex)")
    font_size: int | None = Field(default=None, alias="fontSize", description="Font size")


class ColumnFormat(SocketIOBaseModel):
    """
    Format for a table column.
    """

    column_index: int = Field(..., alias="columnIndex", description="Column index (0-based)", ge=0)
    width: float | None = Field(default=None, description="Column width (points)")
    background_color: str | None = Field(default=None, alias="backgroundColor", description="Background color (hex)")
    font_size: int | None = Field(default=None, alias="fontSize", description="Font size")


class PptUpdateTableFormatRequest(BaseRequest):
    """
    Request to update table formatting.
    """

    event_name: ClassVar[str] = "ppt:update:tableFormat"

    element_id: str = Field(..., alias="elementId", description="Table element ID")
    cell_formats: list["CellFormat"] | None = Field(default=None, alias="cellFormats", description="Cell-level formats")
    row_formats: list["RowFormat"] | None = Field(default=None, alias="rowFormats", description="Row-level formats")
    column_formats: list["ColumnFormat"] | None = Field(
        default=None, alias="columnFormats", description="Column-level formats"
    )


class ElementUpdates(SocketIOBaseModel):
    """
    Geometric property updates for an element.
    """

    left: float | None = Field(default=None, description="New X position (points)")
    top: float | None = Field(default=None, description="New Y position (points)")
    width: float | None = Field(default=None, description="New width (points)")
    height: float | None = Field(default=None, description="New height (points)")
    rotation: float | None = Field(default=None, description="New rotation angle (degrees, 0-360)")


class PptUpdateElementRequest(BaseRequest):
    """
    Request to update element position/size/rotation.
    """

    event_name: ClassVar[str] = "ppt:update:element"

    element_id: str = Field(..., alias="elementId", description="Element ID to update")
    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (0-based)", ge=0)
    updates: "ElementUpdates" = Field(..., description="Geometric updates")


class PptDeleteElementRequest(BaseRequest):
    """
    Request to delete element(s) from a slide.
    """

    event_name: ClassVar[str] = "ppt:delete:element"

    element_id: str | None = Field(default=None, alias="elementId", description="Single element ID to delete")
    element_ids: list[str] | None = Field(default=None, alias="elementIds", description="Batch element IDs to delete")
    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (0-based)", ge=0)


class PptReorderElementRequest(BaseRequest):
    """
    Request to adjust element z-order.
    """

    event_name: ClassVar[str] = "ppt:reorder:element"

    element_id: str = Field(..., alias="elementId", description="Element ID to reorder")
    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (0-based)", ge=0)
    action: Literal["bringToFront", "sendToBack", "bringForward", "sendBackward"] = Field(
        ..., description="Reorder action"
    )


class AddSlideOptions(SocketIOBaseModel):
    """
    Options for adding a new slide.
    """

    insert_index: int | None = Field(
        default=None, alias="insertIndex", description="Insert position index (0-based)", ge=0
    )
    layout: str | None = Field(default=None, description="Layout name (e.g. 'Title Slide', 'Blank')")


class PptAddSlideRequest(BaseRequest):
    """
    Request to add a new slide.
    """

    event_name: ClassVar[str] = "ppt:add:slide"

    options: Optional["AddSlideOptions"] = Field(default=None, description="Slide options")


class PptGotoSlideRequest(BaseRequest):
    """
    Request to jump to a specific slide.
    """

    event_name: ClassVar[str] = "ppt:goto:slide"

    slide_index: int = Field(..., alias="slideIndex", description="Target slide index (0-based)", ge=0)


# ============================================================================
# Chart DTOs (OASP /ppt Draft, v0.2.0)
# ============================================================================
# Server-handled OOXML path: these events are intercepted by office4ai Server
# (python-pptx) and do NOT round-trip through Office.js. See OASP events-ppt
# admonition for save()/reload constraints.


class CategoricalSeries(SocketIOBaseModel):
    """A single data series in a categorical chart (one value per category)."""

    name: str = Field(..., description="Series name shown in the legend")
    values: list[float] = Field(
        ...,
        description="Numeric Y values; length must equal CategoricalChartData.categories.length",
    )
    color: str | None = Field(default=None, description="Optional series color (hex, e.g. '#1F4E79')")


class ScatterPoint(SocketIOBaseModel):
    """One (x, y) sample in a scatter series."""

    x: float = Field(..., description="X coordinate (must be finite)")
    y: float = Field(..., description="Y coordinate (must be finite)")


class ScatterSeries(SocketIOBaseModel):
    """A single data series in a scatter chart (each point carries its own (x, y))."""

    name: str = Field(..., description="Series name shown in the legend")
    points: list[ScatterPoint] = Field(..., description="Sample points; length ≥ 1", min_length=1)
    color: str | None = Field(default=None, description="Optional series color (hex, e.g. '#1F4E79')")


class CategoricalChartData(SocketIOBaseModel):
    """Logical data + display options for categorical charts (column / bar / line / pie / ...)."""

    chart_type: CategoricalChartType = Field(
        ...,
        alias="chartType",
        description="Discriminator selecting the categorical shape",
    )
    categories: list[str] = Field(..., description="Discrete X-axis labels (e.g. ['Jan','Feb','Mar'])")
    series: list[CategoricalSeries] = Field(..., description="Data series (≥ 1)", min_length=1)
    title: str | None = Field(default=None, description="Optional chart title")
    show_legend: bool | None = Field(default=None, alias="showLegend", description="Show legend (default true)")
    show_data_labels: bool | None = Field(
        default=None, alias="showDataLabels", description="Show data labels (default false)"
    )


class ScatterChartData(SocketIOBaseModel):
    """Logical data + display options for scatter charts (continuous X, no categories)."""

    chart_type: ScatterChartType = Field(
        ...,
        alias="chartType",
        description="Discriminator literal 'Scatter'",
    )
    series: list[ScatterSeries] = Field(..., description="Data series (≥ 1)", min_length=1)
    title: str | None = Field(default=None, description="Optional chart title")
    show_legend: bool | None = Field(default=None, alias="showLegend", description="Show legend (default true)")
    show_data_labels: bool | None = Field(
        default=None, alias="showDataLabels", description="Show data labels (default false)"
    )


# Discriminated union over chart_type → schema shape.
# Pydantic / FastMCP idiom: Annotated[Union[...], Field(discriminator=<alias>)].
# The discriminator uses the wire alias 'chartType' since populate_by_name=True and the
# nested models declare alias='chartType' on the discriminator field.
ChartData = Annotated[
    CategoricalChartData | ScatterChartData,
    Field(discriminator="chart_type"),
]


class CategoricalChartUpdate(SocketIOBaseModel):
    """Partial update payload for categorical charts (chartType is required as discriminator)."""

    chart_type: CategoricalChartType = Field(..., alias="chartType", description="Required discriminator")
    categories: list[str] | None = Field(default=None, description="Replace categories (full overwrite)")
    series: list[CategoricalSeries] | None = Field(default=None, description="Replace series (full overwrite)")
    title: str | None = Field(default=None, description="New title; explicit null deletes title")
    show_legend: bool | None = Field(default=None, alias="showLegend")
    show_data_labels: bool | None = Field(default=None, alias="showDataLabels")


class ScatterChartUpdate(SocketIOBaseModel):
    """Partial update payload for scatter charts (chartType literal 'Scatter')."""

    chart_type: ScatterChartType = Field(..., alias="chartType", description="Required discriminator")
    series: list[ScatterSeries] | None = Field(default=None, description="Replace series (full overwrite)")
    title: str | None = Field(default=None, description="New title; explicit null deletes title")
    show_legend: bool | None = Field(default=None, alias="showLegend")
    show_data_labels: bool | None = Field(default=None, alias="showDataLabels")


ChartUpdate = Annotated[
    CategoricalChartUpdate | ScatterChartUpdate,
    Field(discriminator="chart_type"),
]


class ChartInsertOptions(SocketIOBaseModel):
    """Geometry + target slide for chart insertion (all optional)."""

    slide_index: int | None = Field(default=None, alias="slideIndex", description="Slide index (default current)", ge=0)
    left: float | None = Field(default=None, description="X position in points (default centered)")
    top: float | None = Field(default=None, description="Y position in points (default centered)")
    width: float | None = Field(default=None, description="Width in points (default 480)", gt=0)
    height: float | None = Field(default=None, description="Height in points (default 320)", gt=0)


class PptInsertChartRequest(BaseRequest):
    """Request to insert a new chart on a slide (Server-handled OOXML)."""

    event_name: ClassVar[str] = "ppt:insert:chart"

    chart: ChartData = Field(..., description="Chart data (discriminated by chartType)")
    options: Optional["ChartInsertOptions"] = Field(default=None, description="Geometry + target slide")


class PptGetChartRequest(BaseRequest):
    """Request to read an existing chart's data (Server-handled OOXML)."""

    event_name: ClassVar[str] = "ppt:get:chart"

    element_id: str = Field(..., alias="elementId", description="Chart element ID")
    slide_index: int | None = Field(
        default=None,
        alias="slideIndex",
        description="Optional slide hint; elementId is authoritative",
        ge=0,
    )


class PptUpdateChartRequest(BaseRequest):
    """Request to update an existing chart (Server-handled OOXML)."""

    event_name: ClassVar[str] = "ppt:update:chart"

    element_id: str = Field(..., alias="elementId", description="Chart element ID")
    chart: ChartUpdate = Field(..., description="Update payload (discriminated by chartType)")


# ============================================================================
# OOXML Carrier DTOs (OASP /ppt Draft, v0.3.0)
# ============================================================================
# Low-level whole-slide OOXML transport primitives backing the dual-path chart
# implementation (open-document Office.js round-trip). These are transport
# primitives, NOT for routine AI use.
# Spec: events-ppt.md §ppt:get:slideOoxml / §ppt:insert:slidesOoxml.

# camelCase on the wire (consistent with other ppt events); the Add-In maps to the
# PascalCase PowerPoint.InsertSlideFormatting enum internally.
SlideFormatting = Literal["keepSourceFormatting", "useDestinationTheme"]


class PptGetSlideOoxmlRequest(BaseRequest):
    """Export one slide's live OOXML as a base64 .pptx package (ppt:get:slideOoxml).

    Reflects the document's live state (including unsaved local edits). The response
    carries an opaque slideId for subsequent insert:slidesOoxml positioning/replacement.
    """

    event_name: ClassVar[str] = "ppt:get:slideOoxml"

    slide_index: int = Field(..., alias="slideIndex", description="Target slide index (0-based)", ge=0)


class PptInsertSlidesOoxmlRequest(BaseRequest):
    """Apply an OOXML slide package (>=1 slide), optionally replacing + repositioning (ppt:insert:slidesOoxml).

    When replace_slide_id / final_slide_index are provided, the [insert -> delete old -> move]
    composite runs as one sequential best-effort round-trip (NOT atomic; no rollback guarantee).
    """

    event_name: ClassVar[str] = "ppt:insert:slidesOoxml"

    base64: str = Field(
        ...,
        description="Slide package (>=1 slide) as base64 .pptx (no data URL prefix)",
    )
    formatting: SlideFormatting | None = Field(
        default=None,
        description="Keep source formatting or apply destination theme (consumer default: keepSourceFormatting)",
    )
    target_slide_index: int | None = Field(
        default=None,
        alias="targetSlideIndex",
        description="Insert after this 0-based index; default = end of document",
        ge=0,
    )
    replace_slide_id: str | None = Field(
        default=None,
        alias="replaceSlideId",
        description="Opaque slideId (from ppt:get:slideOoxml) to delete after insert (round-trip)",
    )
    final_slide_index: int | None = Field(
        default=None,
        alias="finalSlideIndex",
        description="Move the inserted slide to this 0-based index to reposition (round-trip)",
        ge=0,
    )


# Resolve forward references
SlideElementsOptions.model_rebuild()
ScreenshotOptions.model_rebuild()
TextInsertOptions.model_rebuild()
SlideImageData.model_rebuild()
ElementInsertOptions.model_rebuild()
SlideTableInsertOptions.model_rebuild()
ShapeInsertOptions.model_rebuild()
TextBoxUpdates.model_rebuild()
SlideLayoutsOptions.model_rebuild()
ImageUpdateOptions.model_rebuild()
TableCellUpdate.model_rebuild()
RowUpdate.model_rebuild()
ColumnUpdate.model_rebuild()
CellFormat.model_rebuild()
RowFormat.model_rebuild()
ColumnFormat.model_rebuild()
ElementUpdates.model_rebuild()
AddSlideOptions.model_rebuild()
CategoricalSeries.model_rebuild()
ScatterPoint.model_rebuild()
ScatterSeries.model_rebuild()
CategoricalChartData.model_rebuild()
ScatterChartData.model_rebuild()
CategoricalChartUpdate.model_rebuild()
ScatterChartUpdate.model_rebuild()
ChartInsertOptions.model_rebuild()
