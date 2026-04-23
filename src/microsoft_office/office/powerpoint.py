"""PowerPoint COM automation logic.

Provides comprehensive control over PowerPoint presentations via COM,
including slide management, shape operations, text formatting, charts,
tables, animations, transitions, and export capabilities.
"""

from microsoft_office.com_utils import ensure_absolute_path, get_or_create_app, rgb

# ============================================================
# Constants
# ============================================================

# Slide layout constants
LAYOUT_TITLE = 1
LAYOUT_TITLE_AND_CONTENT = 2
LAYOUT_BLANK = 7

# Text alignment (PpParagraphAlignment)
PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
PP_ALIGN_RIGHT = 3
PP_ALIGN_JUSTIFY = 4

# Shape types (MsoAutoShapeType)
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_ISOSCELES_TRIANGLE = 7
MSO_SHAPE_OVAL = 9

# Gradient direction (MsoGradientStyle)
MSO_GRADIENT_HORIZONTAL = 1
MSO_GRADIENT_VERTICAL = 2
MSO_GRADIENT_DIAGONAL_UP = 3
MSO_GRADIENT_DIAGONAL_DOWN = 4
MSO_GRADIENT_FROM_CORNER = 5
MSO_GRADIENT_FROM_CENTER = 7

# Connector types (MsoConnectorType)
MSO_CONNECTOR_STRAIGHT = 1
MSO_CONNECTOR_ELBOW = 2
MSO_CONNECTOR_CURVE = 3

# Dash styles (MsoLineDashStyle)
MSO_LINE_SOLID = 1
MSO_LINE_SQUARE_DOT = 2
MSO_LINE_DASH = 3
MSO_LINE_DASH_DOT = 4

# Arrow head style
MSO_ARROWHEAD_NONE = 1
MSO_ARROWHEAD_TRIANGLE = 2

# Bullet type
PP_BULLET_UNNUMBERED = 1

# Z-order
MSO_SEND_TO_BACK = 1

# Orientation for textbox
MSO_TEXT_ORIENTATION_HORIZONTAL = 1

# Chart types (XlChartType)
XL_CHART_AREA = 1
XL_CHART_LINE = 4
XL_CHART_PIE = 5
XL_CHART_COLUMN_CLUSTERED = 51
XL_CHART_COLUMN_STACKED = 52
XL_CHART_BAR_CLUSTERED = 57

# Transition types (PpEntryEffect)
PP_TRANSITION_NONE = 0
PP_TRANSITION_CUT = 257
PP_TRANSITION_DISSOLVE = 1537
PP_TRANSITION_FADE = 3844
PP_TRANSITION_PUSH = 3845
PP_TRANSITION_WIPE = 3847
PP_TRANSITION_SPLIT = 3848

# Animation effect types (MsoAnimEffect)
MSO_ANIM_EFFECT_APPEAR = 1
MSO_ANIM_EFFECT_FLY = 2
MSO_ANIM_EFFECT_FADE = 10
MSO_ANIM_EFFECT_WIPE = 22

# Animation trigger (MsoAnimTriggerType)
MSO_ANIM_TRIGGER_ON_CLICK = 1
MSO_ANIM_TRIGGER_WITH_PREVIOUS = 2
MSO_ANIM_TRIGGER_AFTER_PREVIOUS = 3

# Save format
PP_SAVE_AS_PDF = 32

# Shadow type
MSO_SHADOW_21 = 1


# ============================================================
# Internal helpers
# ============================================================

def _get_app():
    """Get or create a PowerPoint COM application instance."""
    return get_or_create_app("PowerPoint.Application")


def _find_body_shape(slide):
    """Find the body (non-title) text shape on a slide.

    Iterates through all shapes on the slide and returns the first
    shape with a text frame that is not the title placeholder.
    """
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if not shape.HasTextFrame:
            continue
        try:
            if slide.Shapes.HasTitle and shape.Name == slide.Shapes.Title.Name:
                continue
        except Exception:
            pass
        if shape.HasTextFrame:
            return shape
    return None


# ============================================================
# Basic Operations
# ============================================================

def create_presentation(file_path: str | None = None) -> dict:
    """Create a new PowerPoint presentation.

    Args:
        file_path: Optional path to save the new presentation to.

    Returns:
        Dict with slide_count and name.
    """
    app = _get_app()
    prs = app.Presentations.Add()
    if file_path:
        prs.SaveAs(ensure_absolute_path(file_path))
    return {"slide_count": prs.Slides.Count, "name": prs.Name}


def open_presentation(file_path: str) -> dict:
    """Open an existing PowerPoint presentation.

    Args:
        file_path: Path to the .pptx file to open.

    Returns:
        Dict with slide_count and name.
    """
    app = _get_app()
    prs = app.Presentations.Open(ensure_absolute_path(file_path))
    return {"slide_count": prs.Slides.Count, "name": prs.Name}


def save_presentation(file_path: str | None = None) -> dict:
    """Save the active presentation.

    Args:
        file_path: Optional new path for SaveAs. If None, saves in place.

    Returns:
        Dict with name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    if file_path:
        prs.SaveAs(ensure_absolute_path(file_path))
    else:
        prs.Save()
    return {"name": prs.Name}


def close_presentation() -> dict:
    """Close the active presentation without saving.

    Returns:
        Dict with the name of the closed presentation.
    """
    app = _get_app()
    prs = app.ActivePresentation
    name = prs.Name
    prs.Close()
    return {"name": name}


def get_info() -> dict:
    """Get information about the active presentation.

    Returns:
        Dict with name, slide_count, and a list of slides with their text content.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slides = []
    for i in range(1, prs.Slides.Count + 1):
        slide = prs.Slides(i)
        shapes_text = []
        for j in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(j)
            if shape.HasTextFrame:
                text = shape.TextFrame.TextRange.Text
                if text.strip():
                    shapes_text.append(text)
        slides.append({"slide_number": i, "texts": shapes_text})
    return {"name": prs.Name, "slide_count": prs.Slides.Count, "slides": slides}


def add_slide(layout: int = LAYOUT_TITLE_AND_CONTENT, position: int | None = None) -> dict:
    """Add a new slide to the active presentation.

    Args:
        layout: Slide layout type (1=Title, 2=TitleAndContent, 7=Blank).
        position: Position to insert at. Defaults to end.

    Returns:
        Dict with slide_number and layout.
    """
    app = _get_app()
    prs = app.ActivePresentation
    pos = position if position else prs.Slides.Count + 1
    try:
        custom_layout = prs.SlideMaster.CustomLayouts(layout)
        slide = prs.Slides.AddSlide(pos, custom_layout)
    except Exception:
        slide = prs.Slides.Add(pos, layout)
    return {"slide_number": slide.SlideIndex, "layout": layout}


def delete_slide(slide_number: int) -> dict:
    """Delete a slide by its number.

    Args:
        slide_number: 1-based slide index.

    Returns:
        Dict with deleted slide number and remaining count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    prs.Slides(slide_number).Delete()
    return {"deleted": slide_number, "remaining": prs.Slides.Count}


def move_slide(from_position: int, to_position: int) -> dict:
    """Move a slide from one position to another.

    Args:
        from_position: Current 1-based position.
        to_position: Target 1-based position.

    Returns:
        Dict with from and to positions.
    """
    app = _get_app()
    prs = app.ActivePresentation
    prs.Slides(from_position).MoveTo(to_position)
    return {"from": from_position, "to": to_position}


def duplicate_slide(slide_number: int) -> dict:
    """Duplicate a slide.

    Args:
        slide_number: 1-based index of the slide to duplicate.

    Returns:
        Dict with original and new slide numbers.
    """
    app = _get_app()
    prs = app.ActivePresentation
    dup = prs.Slides(slide_number).Duplicate()
    new_index = dup(1).SlideIndex
    return {"slide_number": slide_number, "new_slide_number": new_index}


def set_title(slide_number: int, title: str) -> dict:
    """Set the title text of a slide.

    Args:
        slide_number: 1-based slide index.
        title: Title text to set.

    Returns:
        Dict with slide_number and title.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    if slide.Shapes.HasTitle:
        slide.Shapes.Title.TextFrame.TextRange.Text = title
    else:
        shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, 50, 20, 600, 50)
        shape.TextFrame.TextRange.Text = title
    return {"slide_number": slide_number, "title": title}


def set_body(slide_number: int, body: str) -> dict:
    """Set the body text of a slide.

    Args:
        slide_number: 1-based slide index.
        body: Body text to set.

    Returns:
        Dict with slide_number and body.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    body_shape = _find_body_shape(slide)
    if body_shape:
        body_shape.TextFrame.TextRange.Text = body
    else:
        shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, 50, 100, 600, 350)
        shape.TextFrame.TextRange.Text = body
    return {"slide_number": slide_number, "body": body}


def get_slide_content(slide_number: int) -> dict:
    """Get all shape content from a slide.

    Args:
        slide_number: 1-based slide index.

    Returns:
        Dict with slide_number and a list of shapes with their properties.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        info = {"name": shape.Name, "index": i, "has_text": bool(shape.HasTextFrame)}
        if shape.HasTextFrame:
            info["text"] = shape.TextFrame.TextRange.Text
        info["left"] = shape.Left
        info["top"] = shape.Top
        info["width"] = shape.Width
        info["height"] = shape.Height
        shapes.append(info)
    return {"slide_number": slide_number, "shapes": shapes}


# ============================================================
# Text Formatting
# ============================================================

def set_text_format(
    slide_number: int,
    shape_index: int,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_name: str | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
) -> dict:
    """Set text formatting on all text in a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        bold: Whether text should be bold.
        italic: Whether text should be italic.
        font_size: Font size in points.
        font_name: Font family name.
        font_color_rgb: Tuple of (R, G, B) values 0-255.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if not shape.HasTextFrame:
        return {"error": f"Shape {shape_index} has no text frame"}
    font = shape.TextFrame.TextRange.Font
    if bold is not None:
        font.Bold = bold
    if italic is not None:
        font.Italic = italic
    if font_size is not None:
        font.Size = font_size
    if font_name is not None:
        font.Name = font_name
    if font_color_rgb is not None:
        font.Color.RGB = rgb(*font_color_rgb)
    return {"slide_number": slide_number, "shape_index": shape_index}


def set_shape_text(slide_number: int, shape_index: int, text: str) -> dict:
    """Set text content of any shape that has a text frame.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        text: Text to set.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if not shape.HasTextFrame:
        return {"error": f"Shape {shape_index} has no text frame"}
    shape.TextFrame.TextRange.Text = text
    return {"slide_number": slide_number, "shape_index": shape_index}


# ============================================================
# Background
# ============================================================

def set_background_solid(slide_number: int, r: int, g: int, b: int) -> dict:
    """Set a solid color background on a slide.

    Args:
        slide_number: 1-based slide index.
        r: Red component (0-255).
        g: Green component (0-255).
        b: Blue component (0-255).

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    slide.FollowMasterBackground = False
    slide.Background.Fill.Solid()
    slide.Background.Fill.ForeColor.RGB = rgb(r, g, b)
    return {"slide_number": slide_number}


def set_background_gradient(
    slide_number: int,
    color1_rgb: tuple[int, int, int],
    color2_rgb: tuple[int, int, int],
    direction: int = MSO_GRADIENT_HORIZONTAL,
) -> dict:
    """Set a two-color gradient background on a slide.

    Args:
        slide_number: 1-based slide index.
        color1_rgb: Start color as (R, G, B).
        color2_rgb: End color as (R, G, B).
        direction: Gradient direction (1=Horizontal, 2=Vertical, 3=DiagonalUp,
                   4=DiagonalDown, 5=FromCorner, 7=FromCenter).

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    slide.FollowMasterBackground = False
    fill = slide.Background.Fill
    fill.TwoColorGradient(direction, 1)
    fill.ForeColor.RGB = rgb(*color1_rgb)
    fill.BackColor.RGB = rgb(*color2_rgb)
    return {"slide_number": slide_number}


# ============================================================
# Comprehensive Design
# ============================================================

def design_slide(
    slide_number: int,
    bg_rgb: tuple[int, int, int] | None = None,
    bg_gradient: tuple[tuple[int, int, int], tuple[int, int, int], int] | None = None,
    title_font_name: str | None = None,
    title_font_size: float | None = None,
    title_bold: bool | None = None,
    title_color_rgb: tuple[int, int, int] | None = None,
    title_alignment: int | None = None,
    body_font_name: str | None = None,
    body_font_size: float | None = None,
    body_color_rgb: tuple[int, int, int] | None = None,
    body_line_spacing: float | None = None,
    body_alignment: int | None = None,
) -> dict:
    """Apply comprehensive design settings to a single slide.

    Configures background, title formatting, and body formatting in one call.

    Args:
        slide_number: 1-based slide index.
        bg_rgb: Solid background color as (R, G, B).
        bg_gradient: Gradient as ((R1,G1,B1), (R2,G2,B2), direction).
        title_font_name: Font name for the title.
        title_font_size: Font size for the title in points.
        title_bold: Whether the title should be bold.
        title_color_rgb: Title font color as (R, G, B).
        title_alignment: Title paragraph alignment.
        body_font_name: Font name for body text.
        body_font_size: Font size for body text in points.
        body_color_rgb: Body font color as (R, G, B).
        body_line_spacing: Line spacing multiplier for body text.
        body_alignment: Body paragraph alignment.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    # Background
    slide.FollowMasterBackground = False
    if bg_gradient:
        c1, c2, direction = bg_gradient
        fill = slide.Background.Fill
        fill.TwoColorGradient(direction, 1)
        fill.ForeColor.RGB = rgb(*c1)
        fill.BackColor.RGB = rgb(*c2)
    elif bg_rgb:
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = rgb(*bg_rgb)

    # Title formatting
    if slide.Shapes.HasTitle:
        title_shape = slide.Shapes.Title
        font = title_shape.TextFrame.TextRange.Font
        if title_font_name:
            font.Name = title_font_name
        if title_font_size:
            font.Size = title_font_size
        if title_bold is not None:
            font.Bold = title_bold
        if title_color_rgb:
            font.Color.RGB = rgb(*title_color_rgb)
        if title_alignment:
            title_shape.TextFrame.TextRange.ParagraphFormat.Alignment = title_alignment

    # Body formatting
    body_shape = _find_body_shape(slide)
    if body_shape:
        tf = body_shape.TextFrame.TextRange
        if body_font_name:
            tf.Font.Name = body_font_name
        if body_font_size:
            tf.Font.Size = body_font_size
        if body_color_rgb:
            tf.Font.Color.RGB = rgb(*body_color_rgb)
        if body_line_spacing:
            tf.ParagraphFormat.SpaceWithin = body_line_spacing
        if body_alignment:
            tf.ParagraphFormat.Alignment = body_alignment

    return {"slide_number": slide_number}


# ============================================================
# Shape Operations
# ============================================================

def add_shape(
    slide_number: int,
    shape_type: int,
    left: float, top: float, width: float, height: float,
    fill_rgb: tuple[int, int, int] | None = None,
    fill_transparency: float = 0.0,
    line_rgb: tuple[int, int, int] | None = None,
    line_weight: float | None = None,
    shadow: bool = False,
    rotation: float = 0.0,
    z_order_back: bool = False,
) -> dict:
    """Add an auto shape with full styling options.

    Args:
        slide_number: 1-based slide index.
        shape_type: MsoAutoShapeType (1=Rect, 5=RoundedRect, 9=Oval).
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.
        fill_rgb: Fill color as (R, G, B). None for no fill.
        fill_transparency: Fill transparency 0.0-1.0.
        line_rgb: Line color as (R, G, B). None for no line.
        line_weight: Line weight in points.
        shadow: Whether to add a default drop shadow.
        rotation: Rotation angle in degrees.
        z_order_back: Whether to send shape to back.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes.AddShape(shape_type, left, top, width, height)

    # Fill
    if fill_rgb:
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = rgb(*fill_rgb)
        if fill_transparency > 0:
            shape.Fill.Transparency = fill_transparency
    else:
        shape.Fill.Background()

    # Line
    if line_rgb:
        shape.Line.Visible = True
        shape.Line.ForeColor.RGB = rgb(*line_rgb)
        if line_weight:
            shape.Line.Weight = line_weight
    else:
        shape.Line.Visible = False

    # Shadow
    if shadow:
        shape.Shadow.Visible = True
        shape.Shadow.Type = MSO_SHADOW_21
        shape.Shadow.Blur = 8
        shape.Shadow.OffsetX = 3
        shape.Shadow.OffsetY = 3
        shape.Shadow.Transparency = 0.6
        shape.Shadow.ForeColor.RGB = rgb(0, 0, 0)

    # Rotation
    if rotation:
        shape.Rotation = rotation

    # Z-order
    if z_order_back:
        shape.ZOrder(MSO_SEND_TO_BACK)

    return {"slide_number": slide_number, "shape_name": shape.Name}


def add_shape_with_gradient(
    slide_number: int,
    shape_type: int,
    left: float, top: float, width: float, height: float,
    color1_rgb: tuple[int, int, int],
    color2_rgb: tuple[int, int, int],
    gradient_direction: int = MSO_GRADIENT_HORIZONTAL,
    transparency: float = 0.0,
    shadow: bool = False,
    z_order_back: bool = False,
) -> dict:
    """Add a shape with a two-color gradient fill.

    Args:
        slide_number: 1-based slide index.
        shape_type: MsoAutoShapeType (1=Rect, 5=RoundedRect, 9=Oval).
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.
        color1_rgb: Start gradient color as (R, G, B).
        color2_rgb: End gradient color as (R, G, B).
        gradient_direction: MsoGradientStyle direction constant.
        transparency: Fill transparency 0.0-1.0.
        shadow: Whether to add a drop shadow.
        z_order_back: Whether to send shape to back.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes.AddShape(shape_type, left, top, width, height)

    fill = shape.Fill
    fill.TwoColorGradient(gradient_direction, 1)
    fill.ForeColor.RGB = rgb(*color1_rgb)
    fill.BackColor.RGB = rgb(*color2_rgb)
    if transparency > 0:
        fill.Transparency = transparency

    shape.Line.Visible = False

    if shadow:
        shape.Shadow.Visible = True
        shape.Shadow.Blur = 10
        shape.Shadow.OffsetX = 4
        shape.Shadow.OffsetY = 4
        shape.Shadow.Transparency = 0.5
        shape.Shadow.ForeColor.RGB = rgb(0, 0, 0)

    if z_order_back:
        shape.ZOrder(MSO_SEND_TO_BACK)

    return {"slide_number": slide_number, "shape_name": shape.Name}


def add_textbox(
    slide_number: int,
    left: float, top: float, width: float, height: float,
    text: str,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    bold: bool = False,
    italic: bool = False,
    alignment: int = PP_ALIGN_LEFT,
    fill_rgb: tuple[int, int, int] | None = None,
    fill_transparency: float = 0.0,
) -> dict:
    """Add a text box with full formatting options.

    Args:
        slide_number: 1-based slide index.
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.
        text: Text content.
        font_name: Font family name.
        font_size: Font size in points.
        font_color_rgb: Font color as (R, G, B).
        bold: Whether text should be bold.
        italic: Whether text should be italic.
        alignment: Paragraph alignment (1=Left, 2=Center, 3=Right).
        fill_rgb: Background fill color as (R, G, B).
        fill_transparency: Fill transparency 0.0-1.0.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)

    tf = shape.TextFrame
    tf.WordWrap = True
    tr = tf.TextRange
    tr.Text = text
    tr.ParagraphFormat.Alignment = alignment

    font = tr.Font
    if font_name:
        font.Name = font_name
    if font_size:
        font.Size = font_size
    if font_color_rgb:
        font.Color.RGB = rgb(*font_color_rgb)
    font.Bold = bold
    font.Italic = italic

    if fill_rgb:
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = rgb(*fill_rgb)
        if fill_transparency > 0:
            shape.Fill.Transparency = fill_transparency
    else:
        shape.Fill.Background()

    shape.Line.Visible = False

    return {"slide_number": slide_number, "shape_name": shape.Name}


def add_line(
    slide_number: int,
    x1: float, y1: float, x2: float, y2: float,
    color_rgb: tuple[int, int, int] = (0, 0, 0),
    weight: float = 1.0,
    dash_style: int = MSO_LINE_SOLID,
) -> dict:
    """Add a line to a slide.

    Args:
        slide_number: 1-based slide index.
        x1: Start X coordinate in points.
        y1: Start Y coordinate in points.
        x2: End X coordinate in points.
        y2: End Y coordinate in points.
        color_rgb: Line color as (R, G, B).
        weight: Line weight in points.
        dash_style: Dash style (1=Solid, 2=SquareDot, 3=Dash, 4=DashDot).

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    connector = slide.Shapes.AddLine(x1, y1, x2, y2)
    connector.Line.ForeColor.RGB = rgb(*color_rgb)
    connector.Line.Weight = weight
    connector.Line.DashStyle = dash_style
    return {"slide_number": slide_number, "shape_name": connector.Name}


def add_connector(
    slide_number: int,
    connector_type: int,
    begin_x: float, begin_y: float,
    end_x: float, end_y: float,
    color_rgb: tuple[int, int, int] | None = None,
    weight: float = 1.0,
    dash_style: int = MSO_LINE_SOLID,
    begin_arrow: bool = False,
    end_arrow: bool = False,
) -> dict:
    """Add a connector shape between two points.

    Args:
        slide_number: 1-based slide index.
        connector_type: Connector type (1=Straight, 2=Elbow, 3=Curve).
        begin_x: Start X coordinate in points.
        begin_y: Start Y coordinate in points.
        end_x: End X coordinate in points.
        end_y: End Y coordinate in points.
        color_rgb: Line color as (R, G, B). None for default.
        weight: Line weight in points.
        dash_style: Dash style (1=Solid, 2=SquareDot, 3=Dash, 4=DashDot).
        begin_arrow: Whether to add an arrowhead at the start.
        end_arrow: Whether to add an arrowhead at the end.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes.AddConnector(connector_type, begin_x, begin_y, end_x, end_y)

    if color_rgb:
        shape.Line.ForeColor.RGB = rgb(*color_rgb)
    shape.Line.Weight = weight
    shape.Line.DashStyle = dash_style

    if begin_arrow:
        shape.Line.BeginArrowheadStyle = MSO_ARROWHEAD_TRIANGLE
    if end_arrow:
        shape.Line.EndArrowheadStyle = MSO_ARROWHEAD_TRIANGLE

    return {"slide_number": slide_number, "shape_name": shape.Name}


def clear_extra_shapes(slide_number: int) -> dict:
    """Remove all shapes except title and body placeholders.

    Args:
        slide_number: 1-based slide index.

    Returns:
        Dict with slide_number and count of deleted shapes.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    title_name = None
    body_name = None
    if slide.Shapes.HasTitle:
        title_name = slide.Shapes.Title.Name
    body_shape = _find_body_shape(slide)
    if body_shape:
        body_name = body_shape.Name

    to_delete = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if shape.Name != title_name and shape.Name != body_name:
            to_delete.append(shape.Name)

    for name in reversed(to_delete):
        for i in range(1, slide.Shapes.Count + 1):
            if slide.Shapes(i).Name == name:
                slide.Shapes(i).Delete()
                break

    return {"slide_number": slide_number, "deleted": len(to_delete)}


def reposition_shape(
    slide_number: int,
    shape_index: int,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> dict:
    """Reposition and/or resize a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        left: New left position in points.
        top: New top position in points.
        width: New width in points.
        height: New height in points.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if left is not None:
        shape.Left = left
    if top is not None:
        shape.Top = top
    if width is not None:
        shape.Width = width
    if height is not None:
        shape.Height = height
    return {"slide_number": slide_number, "shape_name": shape.Name}


# ============================================================
# Shape Effects
# ============================================================

def set_shape_shadow(
    slide_number: int,
    shape_index: int,
    blur: float = 8,
    offset_x: float = 3,
    offset_y: float = 3,
    transparency: float = 0.6,
    color_rgb: tuple[int, int, int] = (0, 0, 0),
) -> dict:
    """Add a drop shadow effect to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        blur: Shadow blur radius in points.
        offset_x: Horizontal shadow offset in points.
        offset_y: Vertical shadow offset in points.
        transparency: Shadow transparency 0.0-1.0.
        color_rgb: Shadow color as (R, G, B).

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shadow = shape.Shadow
    shadow.Visible = True
    shadow.Type = MSO_SHADOW_21
    shadow.Blur = blur
    shadow.OffsetX = offset_x
    shadow.OffsetY = offset_y
    shadow.Transparency = transparency
    shadow.ForeColor.RGB = rgb(*color_rgb)
    return {"slide_number": slide_number}


def set_shape_glow(
    slide_number: int,
    shape_index: int,
    color_rgb: tuple[int, int, int],
    radius: float = 10.0,
    transparency: float = 0.0,
) -> dict:
    """Add a glow effect to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        color_rgb: Glow color as (R, G, B).
        radius: Glow radius in points.
        transparency: Glow transparency 0.0-1.0.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    glow = shape.Glow
    glow.Color.RGB = rgb(*color_rgb)
    glow.Radius = radius
    glow.Transparency = transparency
    return {"slide_number": slide_number}


def set_shape_reflection(
    slide_number: int,
    shape_index: int,
    reflection_type: int = 1,
) -> dict:
    """Add a reflection effect to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        reflection_type: Reflection preset type (1-9).

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.Reflection.Type = reflection_type
    return {"slide_number": slide_number}


def set_shape_soft_edges(
    slide_number: int,
    shape_index: int,
    radius: float = 10.0,
) -> dict:
    """Apply soft edges effect to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        radius: Soft edge radius in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.SoftEdge.Type = 0  # Custom
    shape.SoftEdge.Radius = radius
    return {"slide_number": slide_number}


def set_shape_3d(
    slide_number: int,
    shape_index: int,
    bevel_type: int = 3,
    bevel_width: float = 6,
    bevel_height: float = 6,
    depth: float = 0,
    material: int = 1,
) -> dict:
    """Apply 3D bevel effect to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        bevel_type: Bevel preset (3=Circle, 6=RelaxedInset).
        bevel_width: Bevel width in points.
        bevel_height: Bevel height in points.
        depth: Extrusion depth in points.
        material: Lighting material preset.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    three_d = shape.ThreeD
    three_d.BevelTopType = bevel_type
    three_d.BevelTopDepth = bevel_height
    three_d.BevelTopInset = bevel_width
    if depth > 0:
        three_d.Depth = depth
    three_d.PresetMaterial = material
    return {"slide_number": slide_number}


# ============================================================
# Bullets
# ============================================================

def set_bullets(
    slide_number: int,
    bullet_char: int = 8226,
    bullet_color_rgb: tuple[int, int, int] | None = None,
    bullet_size: float | None = None,
    exclude_texts: list[str] | None = None,
) -> dict:
    """Set bullet formatting on body text paragraphs.

    Args:
        slide_number: 1-based slide index.
        bullet_char: Unicode character code for bullet (default 8226 = bullet).
        bullet_color_rgb: Bullet color as (R, G, B).
        bullet_size: Bullet relative size (e.g. 1.0 = 100%).
        exclude_texts: List of paragraph texts to exclude from bullet formatting.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    body_shape = _find_body_shape(slide)
    if not body_shape:
        return {"error": "No body shape found"}

    exclude = set(exclude_texts or [])
    tf = body_shape.TextFrame
    for p_idx in range(1, tf.TextRange.Paragraphs().Count + 1):
        para = tf.TextRange.Paragraphs(p_idx)
        text = para.Text.strip()
        if text and text not in exclude:
            pf = para.ParagraphFormat
            pf.Bullet.Visible = True
            pf.Bullet.Type = PP_BULLET_UNNUMBERED
            pf.Bullet.Character = bullet_char
            if bullet_color_rgb:
                pf.Bullet.Font.Color.RGB = rgb(*bullet_color_rgb)
            if bullet_size:
                pf.Bullet.RelativeSize = bullet_size
        else:
            para.ParagraphFormat.Bullet.Visible = False

    return {"slide_number": slide_number}


# ============================================================
# Tables
# ============================================================

def add_table(
    slide_number: int,
    rows: int, cols: int,
    left: float, top: float, width: float, height: float,
    data: list[list[str]] | None = None,
    header_fill_rgb: tuple[int, int, int] | None = None,
    header_font_color_rgb: tuple[int, int, int] | None = None,
    border_color_rgb: tuple[int, int, int] | None = None,
) -> dict:
    """Add a table to a slide.

    Args:
        slide_number: 1-based slide index.
        rows: Number of rows.
        cols: Number of columns.
        left: Left position in points.
        top: Top position in points.
        width: Table width in points.
        height: Table height in points.
        data: 2D list of cell values. Row-major, 0-indexed.
        header_fill_rgb: Fill color for the first row as (R, G, B).
        header_font_color_rgb: Font color for the first row as (R, G, B).
        border_color_rgb: Border color for all cells as (R, G, B).

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
    table = shape.Table

    # Populate data
    if data:
        for r_idx, row in enumerate(data):
            for c_idx, value in enumerate(row):
                if r_idx < rows and c_idx < cols:
                    table.Cell(r_idx + 1, c_idx + 1).Shape.TextFrame.TextRange.Text = str(value)

    # Header styling
    if header_fill_rgb:
        for c_idx in range(1, cols + 1):
            table.Cell(1, c_idx).Shape.Fill.Solid()
            table.Cell(1, c_idx).Shape.Fill.ForeColor.RGB = rgb(*header_fill_rgb)

    if header_font_color_rgb:
        for c_idx in range(1, cols + 1):
            table.Cell(1, c_idx).Shape.TextFrame.TextRange.Font.Color.RGB = rgb(*header_font_color_rgb)

    # Border styling
    if border_color_rgb:
        color_val = rgb(*border_color_rgb)
        for r_idx in range(1, rows + 1):
            for c_idx in range(1, cols + 1):
                cell = table.Cell(r_idx, c_idx)
                for edge in range(1, 5):  # 1=Left, 2=Right, 3=Top, 4=Bottom
                    border = cell.Borders(edge)
                    border.ForeColor.RGB = color_val
                    border.Weight = 1.0

    return {"slide_number": slide_number, "shape_name": shape.Name}


# ============================================================
# Charts
# ============================================================

def add_chart(
    slide_number: int,
    chart_type: int,
    left: float, top: float, width: float, height: float,
    categories: list[str],
    series_data: list[dict],
    title: str | None = None,
    has_legend: bool = True,
) -> dict:
    """Add a chart to a slide.

    Args:
        slide_number: 1-based slide index.
        chart_type: XlChartType (51=ColumnClustered, 4=Line, 5=Pie,
                    57=BarClustered, 1=Area, 52=ColumnStacked).
        left: Left position in points.
        top: Top position in points.
        width: Chart width in points.
        height: Chart height in points.
        categories: List of category labels.
        series_data: List of dicts with "name" (str) and "values" (list[float]).
        title: Optional chart title.
        has_legend: Whether to show the legend.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    shape = slide.Shapes.AddChart2(-1, chart_type, left, top, width, height)
    chart = shape.Chart

    # Write data to the chart's embedded workbook
    wb = chart.ChartData.Workbook
    ws = wb.Worksheets(1)

    # Clear existing data
    ws.Cells.Clear()

    # Write categories (column A, starting from row 2)
    for i, cat in enumerate(categories):
        ws.Cells(i + 2, 1).Value = cat

    # Write series
    for s_idx, series in enumerate(series_data):
        col = s_idx + 2
        ws.Cells(1, col).Value = series["name"]
        for v_idx, value in enumerate(series["values"]):
            ws.Cells(v_idx + 2, col).Value = value

    # Set chart data range
    num_rows = len(categories) + 1
    num_cols = len(series_data) + 1
    last_col_letter = chr(ord("A") + num_cols - 1)
    data_range = f"A1:{last_col_letter}{num_rows}"
    chart.SetSourceData(ws.Range(data_range))

    # Close the workbook to release it
    wb.Close(False)

    # Chart options
    if title:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    else:
        chart.HasTitle = False

    chart.HasLegend = has_legend

    return {"slide_number": slide_number, "shape_name": shape.Name}


# ============================================================
# Speaker Notes
# ============================================================

def set_speaker_notes(slide_number: int, notes_text: str) -> dict:
    """Set speaker notes for a slide.

    Args:
        slide_number: 1-based slide index.
        notes_text: Notes text content.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    slide.NotesPage.Shapes(2).TextFrame.TextRange.Text = notes_text
    return {"slide_number": slide_number}


def get_speaker_notes(slide_number: int) -> dict:
    """Get speaker notes from a slide.

    Args:
        slide_number: 1-based slide index.

    Returns:
        Dict with slide_number and notes text.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    notes_text = slide.NotesPage.Shapes(2).TextFrame.TextRange.Text
    return {"slide_number": slide_number, "notes": notes_text}


# ============================================================
# Export
# ============================================================

def export_to_pdf(output_path: str) -> dict:
    """Export the active presentation to PDF.

    Args:
        output_path: File path for the output PDF.

    Returns:
        Dict with the output path.
    """
    app = _get_app()
    prs = app.ActivePresentation
    abs_path = ensure_absolute_path(output_path)
    prs.SaveAs(abs_path, PP_SAVE_AS_PDF)
    return {"output_path": abs_path}


# ============================================================
# Transitions & Animations
# ============================================================

def set_slide_transition(
    slide_number: int,
    transition_type: int = PP_TRANSITION_NONE,
    duration: float = 1.0,
    advance_on_click: bool = True,
    advance_time: float | None = None,
) -> dict:
    """Set slide transition effect.

    Args:
        slide_number: 1-based slide index.
        transition_type: PpEntryEffect constant (0=None, 3844=Fade, 3845=Push,
                         3847=Wipe, 3848=Split, 257=Cut, 1537=Dissolve).
        duration: Transition duration in seconds.
        advance_on_click: Whether to advance on mouse click.
        advance_time: Optional auto-advance time in seconds.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    transition = slide.SlideShowTransition
    transition.EntryEffect = transition_type
    transition.Duration = duration
    transition.AdvanceOnClick = advance_on_click
    if advance_time is not None:
        transition.AdvanceOnTime = True
        transition.AdvanceTime = advance_time
    else:
        transition.AdvanceOnTime = False
    return {"slide_number": slide_number}


def add_animation(
    slide_number: int,
    shape_index: int,
    effect_type: int,
    trigger: int = MSO_ANIM_TRIGGER_ON_CLICK,
    duration: float = 0.5,
    delay: float = 0.0,
) -> dict:
    """Add an animation effect to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        effect_type: MsoAnimEffect (1=Appear, 2=Fly, 10=Fade, 22=Wipe).
        trigger: Trigger type (1=OnClick, 2=WithPrevious, 3=AfterPrevious).
        duration: Animation duration in seconds.
        delay: Delay before animation starts in seconds.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    timeline = slide.TimeLine
    effect = timeline.MainSequence.AddEffect(shape, effect_type)
    effect.Timing.TriggerType = trigger
    effect.Timing.Duration = duration
    effect.Timing.TriggerDelayTime = delay
    return {"slide_number": slide_number}


# ============================================================
# Grouping
# ============================================================

def group_shapes(slide_number: int, shape_indices: list[int]) -> dict:
    """Group multiple shapes together.

    Args:
        slide_number: 1-based slide index.
        shape_indices: List of 1-based shape indices to group.

    Returns:
        Dict with slide_number and the group shape name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    names = []
    for idx in shape_indices:
        names.append(slide.Shapes(idx).Name)
    group = slide.Shapes.Range(names).Group()
    return {"slide_number": slide_number, "group_name": group.Name}


def ungroup_shapes(slide_number: int, shape_index: int) -> dict:
    """Ungroup a grouped shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based index of the group shape to ungroup.

    Returns:
        Dict with slide_number and count of resulting shapes.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    result = shape.Ungroup()
    return {"slide_number": slide_number, "ungrouped_count": result.Count}


# ============================================================
# Images
# ============================================================

def add_image(
    slide_number: int,
    file_path: str,
    left: float, top: float,
    width: float | None = None,
    height: float | None = None,
) -> dict:
    """Insert an image onto a slide.

    Args:
        slide_number: 1-based slide index.
        file_path: Path to the image file.
        left: Left position in points.
        top: Top position in points.
        width: Width in points. None to use original size.
        height: Height in points. None to use original size.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    abs_path = ensure_absolute_path(file_path)
    if width and height:
        shape = slide.Shapes.AddPicture(
            abs_path, LinkToFile=False, SaveWithDocument=True,
            Left=left, Top=top, Width=width, Height=height,
        )
    else:
        shape = slide.Shapes.AddPicture(
            abs_path, LinkToFile=False, SaveWithDocument=True,
            Left=left, Top=top,
        )
    return {"slide_number": slide_number, "shape_name": shape.Name}


def crop_image(
    slide_number: int,
    shape_index: int,
    crop_left: float = 0,
    crop_right: float = 0,
    crop_top: float = 0,
    crop_bottom: float = 0,
) -> dict:
    """Crop an image shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the picture.
        crop_left: Amount to crop from left in points.
        crop_right: Amount to crop from right in points.
        crop_top: Amount to crop from top in points.
        crop_bottom: Amount to crop from bottom in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    pf = shape.PictureFormat
    pf.CropLeft = crop_left
    pf.CropRight = crop_right
    pf.CropTop = crop_top
    pf.CropBottom = crop_bottom
    return {"slide_number": slide_number}


def set_image_effects(
    slide_number: int,
    shape_index: int,
    brightness: float | None = None,
    contrast: float | None = None,
) -> dict:
    """Adjust brightness and contrast of an image.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the picture.
        brightness: Brightness adjustment (-1.0 to 1.0).
        contrast: Contrast adjustment (-1.0 to 1.0).

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    pf = shape.PictureFormat
    if brightness is not None:
        pf.Brightness = brightness
    if contrast is not None:
        pf.Contrast = contrast
    return {"slide_number": slide_number}


# ============================================================
# Slide Size & Sections
# ============================================================

def set_slide_size(width: float, height: float) -> dict:
    """Set the slide dimensions.

    Args:
        width: Slide width in points (standard 16:9 = 960, 4:3 = 720).
        height: Slide height in points (standard = 540).

    Returns:
        Dict with width and height.
    """
    app = _get_app()
    prs = app.ActivePresentation
    prs.PageSetup.SlideWidth = width
    prs.PageSetup.SlideHeight = height
    return {"width": width, "height": height}


def add_section(name: str, slide_number: int) -> dict:
    """Add a named section starting at a specific slide.

    Args:
        name: Section name.
        slide_number: 1-based slide index where the section begins.

    Returns:
        Dict with section name and slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    prs.SectionProperties.AddSection(slide_number - 1, name)
    return {"name": name, "slide_number": slide_number}


def get_sections() -> dict:
    """Get all sections in the active presentation.

    Returns:
        Dict with a list of sections, each having name, index, and slide count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    sp = prs.SectionProperties
    sections = []
    for i in range(1, sp.Count + 1):
        sections.append({
            "index": i,
            "name": sp.Name(i),
            "slide_count": sp.SlidesCount(i),
        })
    return {"sections": sections}


# ============================================================
# Shape Management
# ============================================================

def align_shapes(slide_number: int, shape_indices: list[int], alignment: str) -> dict:
    """Align multiple shapes on a slide.

    Args:
        slide_number: 1-based slide index.
        shape_indices: List of 1-based shape indices to align.
        alignment: Alignment type ("left", "center", "right", "top", "middle", "bottom").

    Returns:
        Dict with slide_number and alignment.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    shapes = [slide.Shapes(i) for i in shape_indices]

    if alignment == "left":
        min_left = min(s.Left for s in shapes)
        for s in shapes:
            s.Left = min_left
    elif alignment == "center":
        centers = [(s.Left + s.Width / 2) for s in shapes]
        avg_center = sum(centers) / len(centers)
        for s in shapes:
            s.Left = avg_center - s.Width / 2
    elif alignment == "right":
        max_right = max(s.Left + s.Width for s in shapes)
        for s in shapes:
            s.Left = max_right - s.Width
    elif alignment == "top":
        min_top = min(s.Top for s in shapes)
        for s in shapes:
            s.Top = min_top
    elif alignment == "middle":
        middles = [(s.Top + s.Height / 2) for s in shapes]
        avg_middle = sum(middles) / len(middles)
        for s in shapes:
            s.Top = avg_middle - s.Height / 2
    elif alignment == "bottom":
        max_bottom = max(s.Top + s.Height for s in shapes)
        for s in shapes:
            s.Top = max_bottom - s.Height

    return {"slide_number": slide_number, "alignment": alignment}


def distribute_shapes(slide_number: int, shape_indices: list[int], direction: str) -> dict:
    """Distribute shapes evenly on a slide.

    Args:
        slide_number: 1-based slide index.
        shape_indices: List of 1-based shape indices to distribute.
        direction: "horizontal" or "vertical".

    Returns:
        Dict with slide_number and direction.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    shapes = [slide.Shapes(i) for i in shape_indices]
    if len(shapes) < 3:
        return {"slide_number": slide_number, "direction": direction}

    if direction == "horizontal":
        shapes.sort(key=lambda s: s.Left)
        min_left = shapes[0].Left
        max_right = shapes[-1].Left + shapes[-1].Width
        total_shape_width = sum(s.Width for s in shapes)
        gap = (max_right - min_left - total_shape_width) / (len(shapes) - 1)
        current_left = min_left
        for s in shapes:
            s.Left = current_left
            current_left += s.Width + gap
    elif direction == "vertical":
        shapes.sort(key=lambda s: s.Top)
        min_top = shapes[0].Top
        max_bottom = shapes[-1].Top + shapes[-1].Height
        total_shape_height = sum(s.Height for s in shapes)
        gap = (max_bottom - min_top - total_shape_height) / (len(shapes) - 1)
        current_top = min_top
        for s in shapes:
            s.Top = current_top
            current_top += s.Height + gap

    return {"slide_number": slide_number, "direction": direction}


def set_shape_z_order(slide_number: int, shape_index: int, order: str) -> dict:
    """Set the z-order of a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        order: "front", "back", "forward", "backward".

    Returns:
        Dict with slide_number and order.
    """
    # msoBringToFront=0, msoSendToBack=1, msoBringForward=2, msoSendBackward=3
    order_map = {"front": 0, "back": 1, "forward": 2, "backward": 3}
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.ZOrder(order_map[order])
    return {"slide_number": slide_number, "order": order}


def duplicate_shape(slide_number: int, shape_index: int) -> dict:
    """Duplicate a shape on the same slide.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.

    Returns:
        Dict with slide_number and new_shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    new_shape = shape.Duplicate()(1)
    return {"slide_number": slide_number, "new_shape_name": new_shape.Name}


def delete_shape(slide_number: int, shape_index: int) -> dict:
    """Delete a shape from a slide.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.

    Returns:
        Dict with slide_number and deleted shape name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    name = shape.Name
    shape.Delete()
    return {"slide_number": slide_number, "shape_name": name}


def get_shape_properties(slide_number: int, shape_index: int) -> dict:
    """Get detailed properties of a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.

    Returns:
        Dict with shape properties.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)

    props = {
        "slide_number": slide_number,
        "shape_index": shape_index,
        "name": shape.Name,
        "left": shape.Left,
        "top": shape.Top,
        "width": shape.Width,
        "height": shape.Height,
        "rotation": shape.Rotation,
        "shape_type": shape.Type,
        "has_text": bool(shape.HasTextFrame),
    }

    if shape.HasTextFrame:
        props["text"] = shape.TextFrame.TextRange.Text

    try:
        if shape.Fill.Visible:
            props["fill_type"] = shape.Fill.Type
    except Exception:
        pass

    return props


def copy_shape_to_slide(source_slide: int, shape_index: int, target_slide: int) -> dict:
    """Copy a shape from one slide to another.

    Args:
        source_slide: 1-based source slide index.
        shape_index: 1-based shape index on the source slide.
        target_slide: 1-based target slide index.

    Returns:
        Dict with source_slide, target_slide, and new_shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    src = prs.Slides(source_slide)
    dst = prs.Slides(target_slide)
    shape = src.Shapes(shape_index)
    shape.Copy()
    new_shape = dst.Shapes.Paste()(1)
    return {
        "source_slide": source_slide,
        "target_slide": target_slide,
        "new_shape_name": new_shape.Name,
    }


def set_shape_size(slide_number: int, shape_index: int, width: float, height: float) -> dict:
    """Resize a shape independently.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        width: New width in points.
        height: New height in points.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.LockAspectRatio = False
    shape.Width = width
    shape.Height = height
    return {"slide_number": slide_number, "shape_name": shape.Name}


# ============================================================
# Advanced Text Effects
# ============================================================

def set_text_shadow(
    slide_number: int,
    shape_index: int,
    blur_radius: float = 5,
    distance: float = 3,
    angle: float = 45,
    color_rgb: tuple[int, int, int] = (0, 0, 0),
    transparency: float = 0.6,
) -> dict:
    """Add shadow effect to text in a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        blur_radius: Shadow blur radius.
        distance: Shadow distance.
        angle: Shadow angle in degrees.
        color_rgb: Shadow color as (R, G, B).
        transparency: Shadow transparency 0.0-1.0.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    font = shape.TextFrame.TextRange.Font
    shadow = font.Shadow
    shadow.Visible = True
    shadow.Blur = blur_radius
    shadow.OffsetX = distance * 0.707  # approximate cos(45)
    shadow.OffsetY = distance * 0.707  # approximate sin(45)
    shadow.Transparency = transparency / 100 if transparency > 1 else transparency
    shadow.ForeColor.RGB = rgb(*color_rgb)
    return {"slide_number": slide_number}


def set_text_glow(
    slide_number: int,
    shape_index: int,
    radius: float = 5,
    color_rgb: tuple[int, int, int] = (255, 255, 0),
    transparency: float = 0.4,
) -> dict:
    """Add glow effect to text in a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        radius: Glow radius.
        color_rgb: Glow color as (R, G, B).
        transparency: Glow transparency 0.0-1.0.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    font = shape.TextFrame.TextRange.Font
    glow = font.Glow
    glow.Color.RGB = rgb(*color_rgb)
    glow.Radius = radius
    glow.Transparency = transparency / 100 if transparency > 1 else transparency
    return {"slide_number": slide_number}


def set_text_outline(
    slide_number: int,
    shape_index: int,
    color_rgb: tuple[int, int, int] = (0, 0, 0),
    weight: float = 1.0,
) -> dict:
    """Add outline/stroke to text in a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        color_rgb: Outline color as (R, G, B).
        weight: Outline weight in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    font = shape.TextFrame.TextRange.Font
    line = font.Line
    line.Visible = True
    line.ForeColor.RGB = rgb(*color_rgb)
    line.Weight = weight
    return {"slide_number": slide_number}


def set_text_gradient_fill(
    slide_number: int,
    shape_index: int,
    color1_rgb: tuple[int, int, int] = (0, 0, 255),
    color2_rgb: tuple[int, int, int] = (255, 0, 255),
    angle: float = 0,
) -> dict:
    """Apply gradient fill to text characters.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        color1_rgb: Start gradient color as (R, G, B).
        color2_rgb: End gradient color as (R, G, B).
        angle: Gradient angle in degrees.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    font = shape.TextFrame.TextRange.Font
    fill = font.Fill
    fill.TwoColorGradient(MSO_GRADIENT_HORIZONTAL, 1)
    fill.ForeColor.RGB = rgb(*color1_rgb)
    fill.BackColor.RGB = rgb(*color2_rgb)
    return {"slide_number": slide_number}


def add_rich_textbox(
    slide_number: int,
    left: float, top: float, width: float, height: float,
    runs: list[dict],
) -> dict:
    """Add a textbox with per-run formatting.

    Args:
        slide_number: 1-based slide index.
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.
        runs: List of dicts with keys: text, font_size, bold, italic, font_color (R,G,B tuple), font_name.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height)

    tf = shape.TextFrame
    tf.WordWrap = True

    # Build full text
    full_text = "".join(run["text"] for run in runs)
    tf.TextRange.Text = full_text

    # Apply per-run formatting
    pos = 1  # 1-based index in TextRange
    for run in runs:
        length = len(run["text"])
        if length == 0:
            continue
        tr = tf.TextRange.Characters(pos, length)
        if "font_size" in run and run["font_size"] is not None:
            tr.Font.Size = run["font_size"]
        if "bold" in run and run["bold"] is not None:
            tr.Font.Bold = run["bold"]
        if "italic" in run and run["italic"] is not None:
            tr.Font.Italic = run["italic"]
        if "font_name" in run and run["font_name"] is not None:
            tr.Font.Name = run["font_name"]
        if "font_color" in run and run["font_color"] is not None:
            c = run["font_color"]
            tr.Font.Color.RGB = rgb(*c)
        pos += length

    shape.Line.Visible = False

    return {"slide_number": slide_number, "shape_name": shape.Name}


# ============================================================
# Advanced Fills
# ============================================================

def set_shape_pattern_fill(
    slide_number: int,
    shape_index: int,
    pattern_type: int,
    fore_color_rgb: tuple[int, int, int],
    back_color_rgb: tuple[int, int, int],
) -> dict:
    """Apply a pattern fill to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        pattern_type: msoPattern enum value.
        fore_color_rgb: Foreground color as (R, G, B).
        back_color_rgb: Background color as (R, G, B).

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.Fill.Patterned(pattern_type)
    shape.Fill.ForeColor.RGB = rgb(*fore_color_rgb)
    shape.Fill.BackColor.RGB = rgb(*back_color_rgb)
    return {"slide_number": slide_number}


def set_shape_texture_fill(
    slide_number: int,
    shape_index: int,
    texture_path: str,
) -> dict:
    """Apply a texture fill from an image file to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        texture_path: Path to the texture image file.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    abs_path = ensure_absolute_path(texture_path)
    shape.Fill.UserTextured(abs_path)
    return {"slide_number": slide_number}


def set_shape_picture_fill(
    slide_number: int,
    shape_index: int,
    image_path: str,
    stretch: bool = True,
) -> dict:
    """Apply a picture fill to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        image_path: Path to the image file.
        stretch: Whether to stretch the image to fill.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    abs_path = ensure_absolute_path(image_path)
    shape.Fill.UserPicture(abs_path)
    return {"slide_number": slide_number}


# ============================================================
# Freeform Shapes
# ============================================================

def add_freeform_shape(
    slide_number: int,
    points: list[list[float]],
    fill_color_rgb: tuple[int, int, int] | None = None,
    line_color_rgb: tuple[int, int, int] | None = None,
    line_weight: float = 1.0,
    closed: bool = True,
) -> dict:
    """Add a freeform shape from a list of points.

    Args:
        slide_number: 1-based slide index.
        points: List of [x, y] coordinate pairs in points.
        fill_color_rgb: Fill color as (R, G, B). None for no fill.
        line_color_rgb: Line color as (R, G, B). None for no line.
        line_weight: Line weight in points.
        closed: Whether to close the shape path.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    if len(points) < 2:
        raise ValueError("At least 2 points are required")

    # Start building freeform - msoEditingAuto = 0
    builder = slide.Shapes.BuildFreeform(0, points[0][0], points[0][1])
    for pt in points[1:]:
        builder.AddNodes(0, 0, pt[0], pt[1])  # msoSegmentLine=0, msoEditingAuto=0

    if closed and len(points) > 2:
        builder.AddNodes(0, 0, points[0][0], points[0][1])

    shape = builder.ConvertToShape()

    if fill_color_rgb:
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = rgb(*fill_color_rgb)
    else:
        shape.Fill.Background()

    if line_color_rgb:
        shape.Line.Visible = True
        shape.Line.ForeColor.RGB = rgb(*line_color_rgb)
        shape.Line.Weight = line_weight
    else:
        shape.Line.Visible = False

    return {"slide_number": slide_number, "shape_name": shape.Name}


# ============================================================
# Advanced Table
# ============================================================

def format_table_cell(
    slide_number: int,
    shape_index: int,
    row: int,
    col: int,
    fill_color_rgb: tuple[int, int, int] | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    alignment: int | None = None,
) -> dict:
    """Format a specific table cell.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the table.
        row: 1-based row index.
        col: 1-based column index.
        fill_color_rgb: Cell fill color as (R, G, B).
        font_color_rgb: Font color as (R, G, B).
        font_size: Font size in points.
        bold: Whether text should be bold.
        alignment: Paragraph alignment (1=Left, 2=Center, 3=Right).

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    cell = shape.Table.Cell(row, col)

    if fill_color_rgb:
        cell.Shape.Fill.Solid()
        cell.Shape.Fill.ForeColor.RGB = rgb(*fill_color_rgb)

    tr = cell.Shape.TextFrame.TextRange
    if font_color_rgb:
        tr.Font.Color.RGB = rgb(*font_color_rgb)
    if font_size is not None:
        tr.Font.Size = font_size
    if bold is not None:
        tr.Font.Bold = bold
    if alignment is not None:
        tr.ParagraphFormat.Alignment = alignment

    return {"slide_number": slide_number}


def merge_table_cells(
    slide_number: int,
    shape_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> dict:
    """Merge a range of table cells.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the table.
        start_row: 1-based starting row.
        start_col: 1-based starting column.
        end_row: 1-based ending row.
        end_col: 1-based ending column.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    table = shape.Table
    cell1 = table.Cell(start_row, start_col)
    cell2 = table.Cell(end_row, end_col)
    cell1.Merge(cell2)
    return {"slide_number": slide_number}


def set_table_border(
    slide_number: int,
    shape_index: int,
    border_type: str,
    color_rgb: tuple[int, int, int] | None = None,
    weight: float | None = None,
) -> dict:
    """Set table border style.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the table.
        border_type: "all", "outside", "inside", "none".
        color_rgb: Border color as (R, G, B).
        weight: Border weight in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    table = shape.Table
    rows = table.Rows.Count
    cols = table.Columns.Count

    color_val = rgb(*color_rgb) if color_rgb else rgb(0, 0, 0)
    w = weight if weight else 1.0

    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            cell = table.Cell(r, c)
            # edges: 1=Left, 2=Right, 3=Top, 4=Bottom
            for edge in range(1, 5):
                border = cell.Borders(edge)
                is_outside = (
                    (edge == 1 and c == 1) or
                    (edge == 2 and c == cols) or
                    (edge == 3 and r == 1) or
                    (edge == 4 and r == rows)
                )
                is_inside = not is_outside

                if border_type == "none":
                    border.Visible = False
                elif border_type == "all":
                    border.Visible = True
                    border.ForeColor.RGB = color_val
                    border.Weight = w
                elif border_type == "outside" and is_outside:
                    border.Visible = True
                    border.ForeColor.RGB = color_val
                    border.Weight = w
                elif border_type == "outside" and is_inside:
                    border.Visible = False
                elif border_type == "inside" and is_inside:
                    border.Visible = True
                    border.ForeColor.RGB = color_val
                    border.Weight = w
                elif border_type == "inside" and is_outside:
                    border.Visible = False

    return {"slide_number": slide_number}


# ============================================================
# Slide Master & Layout
# ============================================================

def get_slide_layouts() -> dict:
    """Get available slide layouts.

    Returns:
        Dict with a list of layouts, each having index and name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    layouts = []
    try:
        for i in range(1, prs.SlideMaster.CustomLayouts.Count + 1):
            layout = prs.SlideMaster.CustomLayouts(i)
            layouts.append({"index": i, "name": layout.Name})
    except Exception:
        # Fallback for presentations without custom layouts
        pass
    return {"layouts": layouts}


def apply_slide_layout(slide_number: int, layout_index: int) -> dict:
    """Apply a layout to a slide.

    Args:
        slide_number: 1-based slide index.
        layout_index: 1-based layout index from get_slide_layouts.

    Returns:
        Dict with slide_number and layout_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    layout = prs.SlideMaster.CustomLayouts(layout_index)
    slide.CustomLayout = layout
    return {"slide_number": slide_number, "layout_index": layout_index}


def set_slide_number_visibility(visible: bool = True, start_number: int = 1) -> dict:
    """Show or hide slide numbers.

    Args:
        visible: Whether slide numbers should be visible.
        start_number: Starting slide number.

    Returns:
        Dict with visible and start_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    prs.PageSetup.FirstSlideNumber = start_number
    # Toggle header/footer slide number visibility
    for i in range(1, prs.Slides.Count + 1):
        slide = prs.Slides(i)
        slide.HeadersFooters.SlideNumber.Visible = visible
    return {"visible": visible, "start_number": start_number}


# ============================================================
# Media
# ============================================================

def add_video(
    slide_number: int,
    file_path: str,
    left: float, top: float, width: float, height: float,
) -> dict:
    """Embed a video on a slide.

    Args:
        slide_number: 1-based slide index.
        file_path: Path to the video file.
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    abs_path = ensure_absolute_path(file_path)
    shape = slide.Shapes.AddMediaObject2(abs_path, False, True, left, top, width, height)
    return {"slide_number": slide_number, "shape_name": shape.Name}


def add_audio(
    slide_number: int,
    file_path: str,
    left: float = 0,
    top: float = 0,
    play_across_slides: bool = False,
) -> dict:
    """Embed an audio file on a slide.

    Args:
        slide_number: 1-based slide index.
        file_path: Path to the audio file.
        left: Left position in points.
        top: Top position in points.
        play_across_slides: Whether to play across slides.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    abs_path = ensure_absolute_path(file_path)
    shape = slide.Shapes.AddMediaObject2(abs_path, False, True, left, top)

    if play_across_slides:
        # Set to play across slides using animation settings
        anim = slide.TimeLine.MainSequence.AddEffect(
            shape, 1  # msoAnimEffectMediaPlay
        )
        anim.EffectInformation.PlaySettings.PlayAcrossSlides = True

    return {"slide_number": slide_number, "shape_name": shape.Name}


# ============================================================
# Advanced Animation
# ============================================================

def add_motion_path(
    slide_number: int,
    shape_index: int,
    path_type: str,
    duration: float = 1.0,
) -> dict:
    """Add a motion path animation to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        path_type: "line", "arc", "circle", "diamond", "custom".
        duration: Animation duration in seconds.

    Returns:
        Dict with slide_number.
    """
    # MsoAnimEffect motion path constants
    path_map = {
        "line": 63,      # msoAnimEffectPathRight
        "arc": 54,       # msoAnimEffectPathArcDown
        "circle": 56,    # msoAnimEffectPathCircle
        "diamond": 58,   # msoAnimEffectPathDiamond
        "custom": 63,    # msoAnimEffectPathRight as fallback
    }
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    effect_type = path_map.get(path_type, 63)
    effect = slide.TimeLine.MainSequence.AddEffect(shape, effect_type)
    effect.Timing.Duration = duration
    return {"slide_number": slide_number}


def set_animation_trigger(
    slide_number: int,
    shape_index: int,
    trigger_type: str,
    trigger_shape_index: int | None = None,
) -> dict:
    """Set the trigger type for a shape's animation.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        trigger_type: "on_click", "with_previous", "after_previous".
        trigger_shape_index: Optional shape index for click trigger.

    Returns:
        Dict with slide_number.
    """
    trigger_map = {
        "on_click": MSO_ANIM_TRIGGER_ON_CLICK,
        "with_previous": MSO_ANIM_TRIGGER_WITH_PREVIOUS,
        "after_previous": MSO_ANIM_TRIGGER_AFTER_PREVIOUS,
    }
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)

    # Find the animation effect for this shape
    timeline = slide.TimeLine
    for i in range(1, timeline.MainSequence.Count + 1):
        effect = timeline.MainSequence(i)
        if effect.Shape.Name == shape.Name:
            effect.Timing.TriggerType = trigger_map[trigger_type]
            break

    return {"slide_number": slide_number}


def set_animation_order(
    slide_number: int,
    effect_index: int,
    new_position: int,
) -> dict:
    """Reorder an animation effect in the sequence.

    Args:
        slide_number: 1-based slide index.
        effect_index: 1-based current index of the effect.
        new_position: 1-based new position in the sequence.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    effect = slide.TimeLine.MainSequence(effect_index)
    effect.MoveAfter(slide.TimeLine.MainSequence(new_position))
    return {"slide_number": slide_number}


# ============================================================
# Advanced Shape Effects
# ============================================================

def set_shape_border(
    slide_number: int,
    shape_index: int,
    color_rgb: tuple[int, int, int] | None = None,
    weight: float | None = None,
    dash_style: str | None = None,
) -> dict:
    """Set border (line) style on a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        color_rgb: Border color as (R, G, B).
        weight: Border weight in points.
        dash_style: "solid", "dash", "dot", "dash_dot".

    Returns:
        Dict with slide_number.
    """
    dash_map = {
        "solid": MSO_LINE_SOLID,
        "dash": MSO_LINE_DASH,
        "dot": MSO_LINE_SQUARE_DOT,
        "dash_dot": MSO_LINE_DASH_DOT,
    }
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.Line.Visible = True
    if color_rgb:
        shape.Line.ForeColor.RGB = rgb(*color_rgb)
    if weight is not None:
        shape.Line.Weight = weight
    if dash_style and dash_style in dash_map:
        shape.Line.DashStyle = dash_map[dash_style]
    return {"slide_number": slide_number}


# ============================================================
# Export (Single Slide)
# ============================================================

def export_slide_as_image(
    slide_number: int,
    file_path: str,
    width: int = 1920,
    height: int = 1080,
) -> dict:
    """Export a single slide as an image file.

    Args:
        slide_number: 1-based slide index.
        file_path: Output file path (PNG or JPG).
        width: Image width in pixels.
        height: Image height in pixels.

    Returns:
        Dict with slide_number and file_path.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    abs_path = ensure_absolute_path(file_path)
    slide.Export(abs_path, "PNG", width, height)
    return {"slide_number": slide_number, "file_path": abs_path}


# ============================================================
# Shape Text Advanced
# ============================================================

def set_shape_text_vertical(
    slide_number: int,
    shape_index: int,
    orientation: str,
) -> dict:
    """Set text orientation in a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        orientation: "horizontal", "vertical", "vertical270", "stacked".

    Returns:
        Dict with slide_number.
    """
    # msoTextOrientationHorizontal=1, msoTextOrientationDownward=2 (vertical),
    # msoTextOrientationUpward=3 (vertical270), msoTextOrientationVerticalFarEast=4 (stacked)
    orient_map = {
        "horizontal": 1,
        "vertical": 2,
        "vertical270": 3,
        "stacked": 4,
    }
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.TextFrame.Orientation = orient_map.get(orientation, 1)
    return {"slide_number": slide_number}


def set_shape_text_margin(
    slide_number: int,
    shape_index: int,
    left: float = 7.2,
    top: float = 3.6,
    right: float = 7.2,
    bottom: float = 3.6,
) -> dict:
    """Set internal text margins for a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        left: Left margin in points.
        top: Top margin in points.
        right: Right margin in points.
        bottom: Bottom margin in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    tf = shape.TextFrame
    tf.MarginLeft = left
    tf.MarginTop = top
    tf.MarginRight = right
    tf.MarginBottom = bottom
    return {"slide_number": slide_number}


def set_shape_autofit(
    slide_number: int,
    shape_index: int,
    autofit_type: str,
) -> dict:
    """Set text auto-fit behavior for a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        autofit_type: "none", "shrink", "resize_shape".

    Returns:
        Dict with slide_number.
    """
    # ppAutoSizeNone=0, ppAutoSizeShrinkToFit=1 (shrink), ppAutoSizeShapeToFitText=2
    autofit_map = {"none": 0, "shrink": 1, "resize_shape": 2}
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.TextFrame.AutoSize = autofit_map.get(autofit_type, 0)
    return {"slide_number": slide_number}


# ============================================================
# Slide Manipulation (Extended)
# ============================================================

def set_slide_background_image(slide_number: int, image_path: str) -> dict:
    """Set a background image for a slide.

    Args:
        slide_number: 1-based slide index.
        image_path: Path to the image file.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    abs_path = ensure_absolute_path(image_path)
    slide.FollowMasterBackground = False
    slide.Background.Fill.UserPicture(abs_path)
    return {"slide_number": slide_number}


def get_slide_count() -> dict:
    """Return the total number of slides.

    Returns:
        Dict with slide_count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    return {"slide_count": prs.Slides.Count}


def clear_slide(slide_number: int) -> dict:
    """Remove all shapes from a slide (clean slate).

    Args:
        slide_number: 1-based slide index.

    Returns:
        Dict with slide_number and count of deleted shapes.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    count = slide.Shapes.Count
    for i in range(count, 0, -1):
        slide.Shapes(i).Delete()
    return {"slide_number": slide_number, "deleted": count}


# ============================================================
# Shape Advanced
# ============================================================

def set_shape_opacity(slide_number: int, shape_index: int, opacity: float) -> dict:
    """Set shape opacity (0-100).

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        opacity: Opacity value 0 (fully transparent) to 100 (fully opaque).

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    # Transparency is 0.0 (opaque) to 1.0 (transparent), inverse of opacity
    shape.Fill.Transparency = 1.0 - (opacity / 100.0)
    return {"slide_number": slide_number, "shape_index": shape_index}


def rotate_shape(slide_number: int, shape_index: int, angle: float) -> dict:
    """Rotate shape to specific angle.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        angle: Rotation angle in degrees.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.Rotation = angle
    return {"slide_number": slide_number, "shape_index": shape_index}


def flip_shape(slide_number: int, shape_index: int, direction: str) -> dict:
    """Flip shape horizontally or vertically.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        direction: "horizontal" or "vertical".

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    # msoFlipHorizontal=0, msoFlipVertical=1
    if direction == "horizontal":
        shape.Flip(0)
    elif direction == "vertical":
        shape.Flip(1)
    return {"slide_number": slide_number, "shape_index": shape_index}


def set_shape_name(slide_number: int, shape_index: int, name: str) -> dict:
    """Rename a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        name: New name for the shape.

    Returns:
        Dict with slide_number, shape_index, and name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.Name = name
    return {"slide_number": slide_number, "shape_index": shape_index, "name": name}


def find_shape_by_name(slide_number: int, name: str) -> dict:
    """Find shape index by name.

    Args:
        slide_number: 1-based slide index.
        name: Shape name to search for.

    Returns:
        Dict with slide_number, name, and shape_index (or -1 if not found).
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    for i in range(1, slide.Shapes.Count + 1):
        if slide.Shapes(i).Name == name:
            return {"slide_number": slide_number, "name": name, "shape_index": i}
    return {"slide_number": slide_number, "name": name, "shape_index": -1}


def list_shapes(slide_number: int) -> dict:
    """List all shapes with their index, name, type, position, size.

    Args:
        slide_number: 1-based slide index.

    Returns:
        Dict with slide_number and list of shape info dicts.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shapes = []
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        shapes.append({
            "index": i,
            "name": shape.Name,
            "type": shape.Type,
            "left": shape.Left,
            "top": shape.Top,
            "width": shape.Width,
            "height": shape.Height,
        })
    return {"slide_number": slide_number, "shapes": shapes}


# ============================================================
# Text Advanced
# ============================================================

def add_superscript(slide_number: int, shape_index: int, text: str,
                    base_text: str | None = None) -> dict:
    """Add superscript text to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        text: Superscript text to add.
        base_text: Optional base text to prepend before superscript.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if not shape.HasTextFrame:
        return {"error": f"Shape {shape_index} has no text frame"}
    tf = shape.TextFrame.TextRange
    if base_text is not None:
        tf.Text = base_text
        # Append superscript after base text
        run = tf.InsertAfter(text)
    else:
        run = tf.InsertAfter(text)
    run.Font.Superscript = True
    return {"slide_number": slide_number, "shape_index": shape_index}


def add_subscript(slide_number: int, shape_index: int, text: str,
                  base_text: str | None = None) -> dict:
    """Add subscript text to a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        text: Subscript text to add.
        base_text: Optional base text to prepend before subscript.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if not shape.HasTextFrame:
        return {"error": f"Shape {shape_index} has no text frame"}
    tf = shape.TextFrame.TextRange
    if base_text is not None:
        tf.Text = base_text
        run = tf.InsertAfter(text)
    else:
        run = tf.InsertAfter(text)
    run.Font.Subscript = True
    return {"slide_number": slide_number, "shape_index": shape_index}


def set_paragraph_spacing(
    slide_number: int,
    shape_index: int,
    paragraph_index: int,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
) -> dict:
    """Set paragraph spacing control.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        paragraph_index: 1-based paragraph index.
        space_before: Space before paragraph in points.
        space_after: Space after paragraph in points.
        line_spacing: Line spacing in points.

    Returns:
        Dict with slide_number, shape_index, and paragraph_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if not shape.HasTextFrame:
        return {"error": f"Shape {shape_index} has no text frame"}
    para = shape.TextFrame.TextRange.Paragraphs(paragraph_index)
    pf = para.ParagraphFormat
    if space_before is not None:
        pf.SpaceBefore = space_before
    if space_after is not None:
        pf.SpaceAfter = space_after
    if line_spacing is not None:
        pf.LineRuleWithin = True
        pf.SpaceWithin = line_spacing
    return {
        "slide_number": slide_number,
        "shape_index": shape_index,
        "paragraph_index": paragraph_index,
    }


def set_text_columns(slide_number: int, shape_index: int,
                     num_columns: int, spacing: float = 18) -> dict:
    """Set text columns in a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        num_columns: Number of text columns.
        spacing: Spacing between columns in points.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if not shape.HasTextFrame:
        return {"error": f"Shape {shape_index} has no text frame"}
    shape.TextFrame2.Column.Number = num_columns
    shape.TextFrame2.Column.Spacing = spacing
    return {"slide_number": slide_number, "shape_index": shape_index}


def clear_shape_text(slide_number: int, shape_index: int) -> dict:
    """Clear all text from a shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    if not shape.HasTextFrame:
        return {"error": f"Shape {shape_index} has no text frame"}
    shape.TextFrame.TextRange.Text = ""
    return {"slide_number": slide_number, "shape_index": shape_index}


# ============================================================
# Advanced Chart Operations
# ============================================================

def format_chart_title(
    slide_number: int,
    shape_index: int,
    title: str,
    font_size: float | None = None,
    bold: bool | None = None,
    font_color: tuple[int, int, int] | None = None,
) -> dict:
    """Format chart title.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the chart.
        title: Chart title text.
        font_size: Font size in points.
        bold: Whether title should be bold.
        font_color: Font color as (R, G, B).

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    chart = shape.Chart
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    font = chart.ChartTitle.Format.TextFrame2.TextRange.Font
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = bold
    if font_color is not None:
        font.Fill.ForeColor.RGB = rgb(*font_color)
    return {"slide_number": slide_number, "shape_index": shape_index}


def format_chart_axis(
    slide_number: int,
    shape_index: int,
    axis_type: str,
    title: str | None = None,
    min_value: float | None = None,
    max_value: float | None = None,
    number_format: str | None = None,
) -> dict:
    """Format chart axis.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the chart.
        axis_type: "x" for category axis, "y" for value axis.
        title: Axis title text.
        min_value: Minimum axis value.
        max_value: Maximum axis value.
        number_format: Number format string (e.g. "#,##0").

    Returns:
        Dict with slide_number, shape_index, and axis_type.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    chart = shape.Chart
    # xlCategory=1, xlValue=2
    axis_id = 1 if axis_type == "x" else 2
    axis = chart.Axes(axis_id)
    if title is not None:
        axis.HasTitle = True
        axis.AxisTitle.Text = title
    if min_value is not None:
        axis.MinimumScale = min_value
    if max_value is not None:
        axis.MaximumScale = max_value
    if number_format is not None:
        axis.TickLabels.NumberFormat = number_format
    return {"slide_number": slide_number, "shape_index": shape_index, "axis_type": axis_type}


def format_chart_legend(
    slide_number: int,
    shape_index: int,
    position: str | None = None,
    font_size: float | None = None,
    visible: bool = True,
) -> dict:
    """Format chart legend.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the chart.
        position: Legend position: "bottom", "top", "left", "right".
        font_size: Font size in points.
        visible: Whether legend is visible.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    chart = shape.Chart
    chart.HasLegend = visible
    if visible and chart.HasLegend:
        legend = chart.Legend
        # xlLegendPositionBottom=-4107, xlLegendPositionTop=-4160,
        # xlLegendPositionLeft=-4131, xlLegendPositionRight=-4152
        pos_map = {
            "bottom": -4107,
            "top": -4160,
            "left": -4131,
            "right": -4152,
        }
        if position and position in pos_map:
            legend.Position = pos_map[position]
        if font_size is not None:
            legend.Format.TextFrame2.TextRange.Font.Size = font_size
    return {"slide_number": slide_number, "shape_index": shape_index}


def set_chart_style(slide_number: int, shape_index: int, style_index: int) -> dict:
    """Apply a built-in chart style (1-48).

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the chart.
        style_index: Chart style index (1-48).

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    shape.Chart.ChartStyle = style_index
    return {"slide_number": slide_number, "shape_index": shape_index}


def format_chart_data_labels(
    slide_number: int,
    shape_index: int,
    show_value: bool = True,
    show_percentage: bool = False,
    show_category: bool = False,
    font_size: float | None = None,
    font_color: tuple[int, int, int] | None = None,
) -> dict:
    """Format chart data labels.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the chart.
        show_value: Show data values.
        show_percentage: Show percentages.
        show_category: Show category names.
        font_size: Font size in points.
        font_color: Font color as (R, G, B).

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    chart = shape.Chart
    for i in range(1, chart.SeriesCollection().Count + 1):
        series = chart.SeriesCollection(i)
        series.HasDataLabels = True
        labels = series.DataLabels()
        labels.ShowValue = show_value
        labels.ShowPercentage = show_percentage
        labels.ShowCategoryName = show_category
        if font_size is not None:
            labels.Format.TextFrame2.TextRange.Font.Size = font_size
        if font_color is not None:
            labels.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb(*font_color)
    return {"slide_number": slide_number, "shape_index": shape_index}


def add_chart_data_table(
    slide_number: int,
    shape_index: int,
    show_legend_keys: bool = True,
    show_outline: bool = True,
) -> dict:
    """Add a data table below a chart.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index of the chart.
        show_legend_keys: Whether to show legend keys in the data table.
        show_outline: Whether to show data table outline.

    Returns:
        Dict with slide_number and shape_index.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)
    chart = shape.Chart
    chart.HasDataTable = True
    dt = chart.DataTable
    dt.ShowLegendKey = show_legend_keys
    dt.HasBorderOutline = show_outline
    return {"slide_number": slide_number, "shape_index": shape_index}


# ============================================================
# SmartArt-like Diagrams (built with shapes)
# ============================================================

def _create_org_box(slide, name, title, left, top, width, height,
                    fill_rgb=(41, 65, 122), font_color_rgb=(255, 255, 255)):
    """Helper: create a single org chart box with name and title."""
    shape = slide.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE, left, top, width, height)
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = rgb(*fill_rgb)
    shape.Line.Visible = False
    if shape.HasTextFrame:
        tf = shape.TextFrame
        tf.WordWrap = True
        tf.MarginLeft = 4
        tf.MarginRight = 4
        tf.MarginTop = 4
        tf.MarginBottom = 4
        tr = tf.TextRange
        tr.Text = name
        if title:
            tr.Text = name + "\n" + title
            # Format name (first line) bold
            tr.Paragraphs(1).Font.Bold = True
            tr.Paragraphs(1).Font.Size = 10
            tr.Paragraphs(1).Font.Color.RGB = rgb(*font_color_rgb)
            if tr.Paragraphs().Count > 1:
                tr.Paragraphs(2).Font.Size = 8
                tr.Paragraphs(2).Font.Color.RGB = rgb(*font_color_rgb)
        else:
            tr.Font.Bold = True
            tr.Font.Size = 10
            tr.Font.Color.RGB = rgb(*font_color_rgb)
        tr.ParagraphFormat.Alignment = PP_ALIGN_CENTER
    return shape


def _draw_org_tree(slide, node, cx, top, width, height, h_gap, v_gap, level=0):
    """Recursively draw org chart nodes and connectors."""
    box = _create_org_box(
        slide, node.get("name", ""), node.get("title", ""),
        cx - width / 2, top, width, height,
    )
    children = node.get("children", [])
    if not children:
        return
    child_top = top + height + v_gap
    total_width = len(children) * width + (len(children) - 1) * h_gap
    start_x = cx - total_width / 2 + width / 2
    for i, child in enumerate(children):
        child_cx = start_x + i * (width + h_gap)
        _draw_org_tree(slide, child, child_cx, child_top, width, height, h_gap, v_gap, level + 1)
        # Draw connector from parent bottom center to child top center
        parent_bx = cx
        parent_by = top + height
        child_tx = child_cx
        child_ty = child_top
        connector = slide.Shapes.AddLine(parent_bx, parent_by, child_tx, child_ty)
        connector.Line.ForeColor.RGB = rgb(100, 100, 100)
        connector.Line.Weight = 1.5


def create_org_chart(
    slide_number: int,
    data: dict,
    left: float = 50,
    top: float = 80,
    width: float = 860,
    height: float = 400,
) -> dict:
    """Create an org chart from hierarchy data.

    Args:
        slide_number: 1-based slide index.
        data: Hierarchy dict: {"name": "CEO", "title": "Chief Executive",
              "children": [{"name": "VP Sales", ...}, ...]}.
        left: Left boundary in points.
        top: Top boundary in points.
        width: Total width in points.
        height: Total height in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    # Calculate depth for sizing
    def _depth(node):
        children = node.get("children", [])
        if not children:
            return 1
        return 1 + max(_depth(c) for c in children)

    depth = _depth(data)
    box_w = 130
    box_h = 45
    v_gap = max(20, (height - depth * box_h) / max(depth, 1))
    h_gap = 20
    cx = left + width / 2
    _draw_org_tree(slide, data, cx, top, box_w, box_h, h_gap, v_gap)
    return {"slide_number": slide_number}


def create_pyramid_diagram(
    slide_number: int,
    items: list[str],
    left: float = 200,
    top: float = 60,
    width: float = 560,
    height: float = 420,
    colors: list[tuple[int, int, int]] | None = None,
) -> dict:
    """Create a pyramid/triangle diagram.

    Args:
        slide_number: 1-based slide index.
        items: List of strings from top (smallest) to bottom (largest).
        left: Left position in points.
        top: Top position in points.
        width: Total width in points.
        height: Total height in points.
        colors: Optional list of (R,G,B) tuples for each layer.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    n = len(items)
    if not colors:
        # Generate gradient from dark to light blue
        colors = []
        for i in range(n):
            ratio = i / max(n - 1, 1)
            r = int(41 + ratio * 100)
            g = int(65 + ratio * 100)
            b = int(122 + ratio * 80)
            colors.append((r, g, b))

    layer_h = height / n
    for i, item in enumerate(items):
        # Each layer is a trapezoid approximated by a rectangle
        # Width grows from top to bottom
        ratio_top = (i) / n
        ratio_bottom = (i + 1) / n
        layer_w = width * (0.2 + 0.8 * (ratio_top + ratio_bottom) / 2)
        layer_left = left + (width - layer_w) / 2
        layer_top = top + i * layer_h

        shape = slide.Shapes.AddShape(
            MSO_SHAPE_ISOSCELES_TRIANGLE if i == 0 and n > 1
            else MSO_SHAPE_RECTANGLE,
            layer_left, layer_top, layer_w, layer_h - 2,
        )
        color = colors[i] if i < len(colors) else (100, 100, 200)
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = rgb(*color)
        shape.Line.Visible = False

        # Add text label next to or on the shape
        txt = slide.Shapes.AddTextbox(
            MSO_TEXT_ORIENTATION_HORIZONTAL,
            layer_left, layer_top, layer_w, layer_h - 2,
        )
        txt.TextFrame.TextRange.Text = item
        txt.TextFrame.TextRange.Font.Size = 11
        txt.TextFrame.TextRange.Font.Color.RGB = rgb(255, 255, 255)
        txt.TextFrame.TextRange.Font.Bold = True
        txt.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        txt.TextFrame.WordWrap = True
        txt.Fill.Background()
        txt.Line.Visible = False

    return {"slide_number": slide_number}


def create_circular_diagram(
    slide_number: int,
    items: list[str],
    center_text: str | None = None,
    left: float = 180,
    top: float = 40,
    width: float = 600,
    height: float = 460,
    colors: list[tuple[int, int, int]] | None = None,
) -> dict:
    """Create a circular/radial diagram with items around a center.

    Args:
        slide_number: 1-based slide index.
        items: List of item labels arranged in a circle.
        center_text: Optional text for the center circle.
        left: Left position in points.
        top: Top position in points.
        width: Total width in points.
        height: Total height in points.
        colors: Optional list of (R,G,B) tuples.

    Returns:
        Dict with slide_number.
    """
    import math

    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    n = len(items)

    if not colors:
        colors = []
        for i in range(n):
            hue = i / max(n, 1)
            r = int(80 + 120 * abs(math.sin(hue * math.pi * 2)))
            g = int(80 + 120 * abs(math.sin(hue * math.pi * 2 + 2)))
            b = int(80 + 120 * abs(math.sin(hue * math.pi * 2 + 4)))
            colors.append((r, g, b))

    cx = left + width / 2
    cy = top + height / 2
    radius = min(width, height) / 2 - 50
    item_w = 100
    item_h = 50

    # Center circle
    center_r = 45
    center_shape = slide.Shapes.AddShape(
        MSO_SHAPE_OVAL,
        cx - center_r, cy - center_r, center_r * 2, center_r * 2,
    )
    center_shape.Fill.Solid()
    center_shape.Fill.ForeColor.RGB = rgb(41, 65, 122)
    center_shape.Line.Visible = False
    if center_text and center_shape.HasTextFrame:
        center_shape.TextFrame.TextRange.Text = center_text
        center_shape.TextFrame.TextRange.Font.Size = 11
        center_shape.TextFrame.TextRange.Font.Bold = True
        center_shape.TextFrame.TextRange.Font.Color.RGB = rgb(255, 255, 255)
        center_shape.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        center_shape.TextFrame.WordWrap = True

    for i, item in enumerate(items):
        angle = 2 * math.pi * i / n - math.pi / 2  # Start from top
        ix = cx + radius * math.cos(angle) - item_w / 2
        iy = cy + radius * math.sin(angle) - item_h / 2

        color = colors[i] if i < len(colors) else (100, 100, 200)
        shape = slide.Shapes.AddShape(
            MSO_SHAPE_ROUNDED_RECTANGLE, ix, iy, item_w, item_h,
        )
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = rgb(*color)
        shape.Line.Visible = False
        if shape.HasTextFrame:
            shape.TextFrame.TextRange.Text = item
            shape.TextFrame.TextRange.Font.Size = 9
            shape.TextFrame.TextRange.Font.Bold = True
            shape.TextFrame.TextRange.Font.Color.RGB = rgb(255, 255, 255)
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_CENTER
            shape.TextFrame.WordWrap = True

        # Draw connector from center to item
        line = slide.Shapes.AddLine(cx, cy, ix + item_w / 2, iy + item_h / 2)
        line.Line.ForeColor.RGB = rgb(150, 150, 150)
        line.Line.Weight = 1.0
        # Send line behind the shapes
        line.ZOrder(MSO_SEND_TO_BACK)

    return {"slide_number": slide_number}


def create_matrix_diagram(
    slide_number: int,
    quadrants: list[dict],
    title: str | None = None,
    x_label: str | None = None,
    y_label: str | None = None,
    left: float = 100,
    top: float = 60,
    width: float = 760,
    height: float = 420,
) -> dict:
    """Create a 2x2 matrix diagram.

    Args:
        slide_number: 1-based slide index.
        quadrants: List of 4 dicts: [{"title": "Q1", "items": ["item1"]}, ...].
                   Order: top-left, top-right, bottom-left, bottom-right.
        title: Optional matrix title.
        x_label: Optional X-axis label.
        y_label: Optional Y-axis label.
        left: Left position in points.
        top: Top position in points.
        width: Total width in points.
        height: Total height in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)

    colors = [
        (52, 109, 179),   # top-left blue
        (76, 153, 96),    # top-right green
        (204, 153, 0),    # bottom-left yellow
        (192, 80, 77),    # bottom-right red
    ]

    # Title
    title_h = 30 if title else 0
    if title:
        txt = slide.Shapes.AddTextbox(
            MSO_TEXT_ORIENTATION_HORIZONTAL,
            left, top, width, title_h,
        )
        txt.TextFrame.TextRange.Text = title
        txt.TextFrame.TextRange.Font.Size = 16
        txt.TextFrame.TextRange.Font.Bold = True
        txt.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        txt.Fill.Background()
        txt.Line.Visible = False

    # Axis labels offset
    label_margin = 25 if y_label else 0
    matrix_left = left + label_margin
    matrix_top = top + title_h + 5
    matrix_w = width - label_margin
    matrix_h = height - title_h - (25 if x_label else 0)
    cell_w = matrix_w / 2 - 3
    cell_h = matrix_h / 2 - 3

    positions = [
        (matrix_left, matrix_top),                      # top-left
        (matrix_left + cell_w + 6, matrix_top),         # top-right
        (matrix_left, matrix_top + cell_h + 6),         # bottom-left
        (matrix_left + cell_w + 6, matrix_top + cell_h + 6),  # bottom-right
    ]

    for i, (px, py) in enumerate(positions):
        if i >= len(quadrants):
            break
        q = quadrants[i]
        color = colors[i]
        shape = slide.Shapes.AddShape(
            MSO_SHAPE_ROUNDED_RECTANGLE, px, py, cell_w, cell_h,
        )
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = rgb(*color)
        shape.Fill.Transparency = 0.15
        shape.Line.Visible = False

        # Quadrant title and items
        content = q.get("title", "")
        items = q.get("items", [])
        if items:
            content += "\n" + "\n".join("• " + it for it in items)
        if shape.HasTextFrame:
            tf = shape.TextFrame
            tf.WordWrap = True
            tf.MarginLeft = 8
            tf.MarginTop = 8
            tr = tf.TextRange
            tr.Text = content
            tr.Font.Size = 9
            tr.Font.Color.RGB = rgb(255, 255, 255)
            # Bold the title line
            if tr.Paragraphs().Count > 0:
                tr.Paragraphs(1).Font.Bold = True
                tr.Paragraphs(1).Font.Size = 12

    # Y-axis label
    if y_label:
        yl = slide.Shapes.AddTextbox(
            MSO_TEXT_ORIENTATION_HORIZONTAL,
            left, matrix_top, label_margin, matrix_h,
        )
        yl.TextFrame.TextRange.Text = y_label
        yl.TextFrame.TextRange.Font.Size = 10
        yl.TextFrame.TextRange.Font.Bold = True
        yl.TextFrame.Orientation = 4  # msoTextOrientationUpward
        yl.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        yl.Fill.Background()
        yl.Line.Visible = False

    # X-axis label
    if x_label:
        xl = slide.Shapes.AddTextbox(
            MSO_TEXT_ORIENTATION_HORIZONTAL,
            matrix_left, matrix_top + matrix_h + 5, matrix_w, 20,
        )
        xl.TextFrame.TextRange.Text = x_label
        xl.TextFrame.TextRange.Font.Size = 10
        xl.TextFrame.TextRange.Font.Bold = True
        xl.TextFrame.TextRange.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        xl.Fill.Background()
        xl.Line.Visible = False

    return {"slide_number": slide_number}


# ============================================================
# Presentation-wide Operations
# ============================================================

def set_all_slides_background(
    color: tuple[int, int, int] | None = None,
    image_path: str | None = None,
) -> dict:
    """Set background for all slides.

    Args:
        color: Solid background color as (R, G, B).
        image_path: Path to background image.

    Returns:
        Dict with slide_count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    count = prs.Slides.Count
    for i in range(1, count + 1):
        slide = prs.Slides(i)
        slide.FollowMasterBackground = False
        if image_path:
            abs_path = ensure_absolute_path(image_path)
            slide.Background.Fill.UserPicture(abs_path)
        elif color:
            slide.Background.Fill.Solid()
            slide.Background.Fill.ForeColor.RGB = rgb(*color)
    return {"slide_count": count}


def apply_font_to_all(font_name: str, target: str = "all") -> dict:
    """Apply font to all text in the presentation.

    Args:
        font_name: Font family name to apply.
        target: "all" for all text, "titles" for title shapes only,
                "body" for body shapes only.

    Returns:
        Dict with slide_count and shapes_modified count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    modified = 0
    for s in range(1, prs.Slides.Count + 1):
        slide = prs.Slides(s)
        for sh in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(sh)
            if not shape.HasTextFrame:
                continue
            is_title = False
            try:
                if slide.Shapes.HasTitle and shape.Name == slide.Shapes.Title.Name:
                    is_title = True
            except Exception:
                pass

            if target == "titles" and not is_title:
                continue
            if target == "body" and is_title:
                continue

            shape.TextFrame.TextRange.Font.Name = font_name
            modified += 1
    return {"slide_count": prs.Slides.Count, "shapes_modified": modified}


def get_presentation_summary() -> dict:
    """Return presentation summary: slide count, total shapes, total text chars, fonts used.

    Returns:
        Dict with slide_count, total_shapes, total_text_chars, and fonts_used.
    """
    app = _get_app()
    prs = app.ActivePresentation
    total_shapes = 0
    total_chars = 0
    fonts_used = set()
    for s in range(1, prs.Slides.Count + 1):
        slide = prs.Slides(s)
        total_shapes += slide.Shapes.Count
        for sh in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(sh)
            if shape.HasTextFrame:
                text = shape.TextFrame.TextRange.Text
                total_chars += len(text)
                try:
                    for p in range(1, shape.TextFrame.TextRange.Paragraphs().Count + 1):
                        para = shape.TextFrame.TextRange.Paragraphs(p)
                        fname = para.Font.Name
                        if fname:
                            fonts_used.add(fname)
                except Exception:
                    pass
    return {
        "slide_count": prs.Slides.Count,
        "total_shapes": total_shapes,
        "total_text_chars": total_chars,
        "fonts_used": sorted(fonts_used),
    }


# ============================================================
# Remaining Pro Features
# ============================================================

def set_shape_gradient(
    slide_number: int,
    shape_index: int,
    color1: tuple[int, int, int],
    color2: tuple[int, int, int],
    angle: float = 0,
    gradient_type: str = "linear",
) -> dict:
    """Apply gradient fill to an existing shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index on the slide.
        color1: Start color as (R, G, B).
        color2: End color as (R, G, B).
        angle: Gradient angle in degrees (for linear).
        gradient_type: "linear", "radial", "rectangular", or "path".

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)

    gradient_map = {
        "linear": MSO_GRADIENT_HORIZONTAL,
        "radial": MSO_GRADIENT_FROM_CENTER,
        "rectangular": MSO_GRADIENT_FROM_CORNER,
        "path": MSO_GRADIENT_FROM_CENTER,
    }
    direction = gradient_map.get(gradient_type, MSO_GRADIENT_HORIZONTAL)

    fill = shape.Fill
    fill.TwoColorGradient(direction, 1)
    fill.ForeColor.RGB = rgb(*color1)
    fill.BackColor.RGB = rgb(*color2)

    # For linear gradient, set angle via GradientAngle if supported
    if gradient_type == "linear" and angle != 0:
        try:
            fill.GradientAngle = angle
        except Exception:
            pass

    return {"slide_number": slide_number, "shape_name": shape.Name}


def add_qr_code_shape(
    slide_number: int,
    text: str,
    left: float,
    top: float,
    size: float = 100,
    color: tuple[int, int, int] | None = None,
) -> dict:
    """Create a QR code-like pattern from text using small shapes (grid pattern).

    Uses a simple hash-based pattern to generate a visual grid.

    Args:
        slide_number: 1-based slide index.
        text: Text to encode into the pattern.
        left: Left position in points.
        top: Top position in points.
        size: Overall size in points.
        color: Fill color as (R, G, B). Defaults to black.

    Returns:
        Dict with slide_number and shapes_added count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    fill_color = color or (0, 0, 0)
    grid_size = 10
    cell_size = size / grid_size

    # Generate a deterministic pattern from text hash
    import hashlib
    hash_bytes = hashlib.md5(text.encode("utf-8")).digest()
    shapes_added = 0

    for row in range(grid_size):
        for col in range(grid_size):
            # Use hash bytes to decide if cell is filled
            idx = (row * grid_size + col) % len(hash_bytes)
            if hash_bytes[idx] & (1 << ((row * grid_size + col) % 8)):
                cell_left = left + col * cell_size
                cell_top = top + row * cell_size
                shape = slide.Shapes.AddShape(
                    MSO_SHAPE_RECTANGLE,
                    cell_left, cell_top, cell_size, cell_size,
                )
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = rgb(*fill_color)
                shape.Line.Visible = False
                shapes_added += 1

    # Always add corner markers (3 corners)
    marker_size = cell_size * 3
    for (mx, my) in [(left, top), (left + size - marker_size, top), (left, top + size - marker_size)]:
        border = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, mx, my, marker_size, marker_size)
        border.Fill.Solid()
        border.Fill.ForeColor.RGB = rgb(*fill_color)
        border.Line.Visible = False
        shapes_added += 1
        inner = slide.Shapes.AddShape(
            MSO_SHAPE_RECTANGLE,
            mx + cell_size * 0.5, my + cell_size * 0.5,
            marker_size - cell_size, marker_size - cell_size,
        )
        inner.Fill.Solid()
        inner.Fill.ForeColor.RGB = rgb(255, 255, 255)
        inner.Line.Visible = False
        shapes_added += 1
        center = slide.Shapes.AddShape(
            MSO_SHAPE_RECTANGLE,
            mx + cell_size, my + cell_size,
            cell_size, cell_size,
        )
        center.Fill.Solid()
        center.Fill.ForeColor.RGB = rgb(*fill_color)
        center.Line.Visible = False
        shapes_added += 1

    return {"slide_number": slide_number, "shapes_added": shapes_added}


def set_slide_notes_format(
    slide_number: int,
    font_name: str | None = None,
    font_size: float | None = None,
) -> dict:
    """Format speaker notes font.

    Args:
        slide_number: 1-based slide index.
        font_name: Font family name.
        font_size: Font size in points.

    Returns:
        Dict with slide_number.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    notes_text_range = slide.NotesPage.Shapes(2).TextFrame.TextRange
    if font_name:
        notes_text_range.Font.Name = font_name
    if font_size is not None:
        notes_text_range.Font.Size = font_size
    return {"slide_number": slide_number}


def replace_text(
    slide_number: int,
    old_text: str,
    new_text: str,
    match_case: bool = False,
) -> dict:
    """Find and replace text across all shapes in a slide.

    Args:
        slide_number: 1-based slide index.
        old_text: Text to find.
        new_text: Replacement text.
        match_case: Whether to match case.

    Returns:
        Dict with slide_number and replacements count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    count = 0
    for i in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(i)
        if not shape.HasTextFrame:
            continue
        tf = shape.TextFrame.TextRange
        text = tf.Text
        if not match_case:
            if old_text.lower() not in text.lower():
                continue
            # Case-insensitive replace
            import re
            new_val = re.sub(re.escape(old_text), new_text, text, flags=re.IGNORECASE)
            occurrences = len(re.findall(re.escape(old_text), text, flags=re.IGNORECASE))
        else:
            if old_text not in text:
                continue
            new_val = text.replace(old_text, new_text)
            occurrences = text.count(old_text)
        tf.Text = new_val
        count += occurrences
    return {"slide_number": slide_number, "replacements": count}


def replace_text_all_slides(
    old_text: str,
    new_text: str,
    match_case: bool = False,
) -> dict:
    """Find and replace text across ALL slides.

    Args:
        old_text: Text to find.
        new_text: Replacement text.
        match_case: Whether to match case.

    Returns:
        Dict with total_replacements and slides_modified count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    total = 0
    slides_modified = 0
    for s in range(1, prs.Slides.Count + 1):
        result = replace_text(s, old_text, new_text, match_case)
        if result["replacements"] > 0:
            total += result["replacements"]
            slides_modified += 1
    return {"total_replacements": total, "slides_modified": slides_modified}


def get_all_text(slide_number: int | None = None) -> dict:
    """Extract all text from one slide or all slides.

    Args:
        slide_number: 1-based slide index. None for all slides.

    Returns:
        Dict with list of {slide, shape, text} items.
    """
    app = _get_app()
    prs = app.ActivePresentation
    results = []
    if slide_number:
        slides = [prs.Slides(slide_number)]
    else:
        slides = [prs.Slides(i) for i in range(1, prs.Slides.Count + 1)]
    for slide in slides:
        s_num = slide.SlideNumber
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if shape.HasTextFrame:
                text = shape.TextFrame.TextRange.Text
                if text.strip():
                    results.append({
                        "slide": s_num,
                        "shape": shape.Name,
                        "text": text,
                    })
    return {"texts": results, "count": len(results)}


def set_shape_hyperlink(
    slide_number: int,
    shape_index: int,
    url: str | None = None,
    slide_target: int | None = None,
    tooltip: str | None = None,
) -> dict:
    """Add hyperlink to shape.

    Args:
        slide_number: 1-based slide index.
        shape_index: 1-based shape index.
        url: Web URL for external link.
        slide_target: Slide number for internal link.
        tooltip: Tooltip text on hover.

    Returns:
        Dict with slide_number and shape_name.
    """
    app = _get_app()
    prs = app.ActivePresentation
    slide = prs.Slides(slide_number)
    shape = slide.Shapes(shape_index)

    if url:
        hl = shape.ActionSettings(1).Hyperlink  # ppMouseClick = 1
        hl.Address = url
        if tooltip:
            hl.ScreenTip = tooltip
    elif slide_target:
        action = shape.ActionSettings(1)
        action.Action = 101  # ppActionHyperlink
        action.Hyperlink.SubAddress = str(prs.Slides(slide_target).SlideID) + "," + str(slide_target) + ","
        if tooltip:
            action.Hyperlink.ScreenTip = tooltip

    return {"slide_number": slide_number, "shape_name": shape.Name}


def add_header_footer(
    show_date: bool = True,
    show_slide_number: bool = True,
    show_footer: bool = True,
    footer_text: str = "",
    date_format: str = "auto",
) -> dict:
    """Configure header/footer settings for the presentation.

    Args:
        show_date: Whether to show date.
        show_slide_number: Whether to show slide numbers.
        show_footer: Whether to show footer text.
        footer_text: Footer text content.
        date_format: "auto" for automatic date, "fixed" for fixed text.

    Returns:
        Dict with settings applied.
    """
    app = _get_app()
    prs = app.ActivePresentation
    hf = prs.Slides.Range().HeadersFooters

    hf.DateAndTime.Visible = show_date
    if show_date:
        if date_format == "auto":
            hf.DateAndTime.UseFormat = True
        else:
            hf.DateAndTime.UseFormat = False

    hf.SlideNumber.Visible = show_slide_number
    hf.Footer.Visible = show_footer
    if show_footer and footer_text:
        hf.Footer.Text = footer_text

    return {
        "show_date": show_date,
        "show_slide_number": show_slide_number,
        "show_footer": show_footer,
        "footer_text": footer_text,
    }


def duplicate_presentation(file_path: str) -> dict:
    """Save a copy of the current presentation to a new path.

    Args:
        file_path: Destination file path.

    Returns:
        Dict with file_path.
    """
    app = _get_app()
    prs = app.ActivePresentation
    abs_path = ensure_absolute_path(file_path)
    prs.SaveCopyAs(abs_path)
    return {"file_path": abs_path}


def insert_slides_from(
    file_path: str,
    slide_numbers: list[int] | None = None,
    insert_at: int | None = None,
) -> dict:
    """Insert slides from another presentation.

    Args:
        file_path: Path to the source presentation.
        slide_numbers: List of 1-based slide numbers to insert. None for all.
        insert_at: Position to insert at (1-based). None for end.

    Returns:
        Dict with slides_inserted count and insert_position.
    """
    app = _get_app()
    prs = app.ActivePresentation
    abs_path = ensure_absolute_path(file_path)
    pos = insert_at if insert_at else prs.Slides.Count

    if slide_numbers:
        inserted = 0
        for idx, sn in enumerate(slide_numbers):
            prs.Slides.InsertFromFile(abs_path, pos + idx, sn, sn)
            inserted += 1
    else:
        # Open source to count slides
        source = app.Presentations.Open(abs_path, WithWindow=False)
        total = source.Slides.Count
        source.Close()
        prs.Slides.InsertFromFile(abs_path, pos, 1, total)
        inserted = total

    return {"slides_inserted": inserted, "insert_position": pos}
