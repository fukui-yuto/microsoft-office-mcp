"""PowerPoint Advanced Design - High-level design presets and theme management.

Provides professional slide design presets, theme colors, watermarks,
and decorative elements for creating presentation-quality slides
with a single function call.
"""

from microsoft_office.com_utils import get_or_create_app, rgb

# ============================================================
# Constants
# ============================================================

# Slide dimensions (standard 16:9 in points)
SLIDE_WIDTH = 960  # 13.333 inches * 72
SLIDE_HEIGHT = 540  # 7.5 inches * 72

# Shape types
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_OVAL = 9
MSO_SHAPE_ISOSCELES_TRIANGLE = 7
MSO_SHAPE_DIAMOND = 4
MSO_SHAPE_CHEVRON = 52
MSO_SHAPE_RIGHT_ARROW = 33
MSO_SHAPE_PENTAGON = 56

# Text alignment
PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
PP_ALIGN_RIGHT = 3

# Gradient direction
MSO_GRADIENT_HORIZONTAL = 1
MSO_GRADIENT_VERTICAL = 2
MSO_GRADIENT_DIAGONAL_UP = 3
MSO_GRADIENT_DIAGONAL_DOWN = 4
MSO_GRADIENT_FROM_CORNER = 5
MSO_GRADIENT_FROM_CENTER = 7

# Text orientation
MSO_TEXT_ORIENTATION_HORIZONTAL = 1

# Z-order
MSO_BRING_TO_FRONT = 0
MSO_SEND_TO_BACK = 1

# Shadow type
MSO_SHADOW_21 = 1

# Vertical anchor
MSO_ANCHOR_TOP = 1
MSO_ANCHOR_MIDDLE = 3
MSO_ANCHOR_BOTTOM = 4

# Line dash
MSO_LINE_SOLID = 1
MSO_LINE_DASH = 3

# ============================================================
# Color Schemes
# ============================================================

COLOR_SCHEMES = {
    "corporate_blue": {
        "primary": (0, 82, 136),
        "secondary": (0, 120, 190),
        "accent1": (0, 166, 214),
        "accent2": (242, 169, 0),
        "accent3": (0, 150, 136),
        "accent4": (96, 125, 139),
        "background": (248, 250, 252),
        "text": (33, 37, 41),
    },
    "modern_dark": {
        "primary": (30, 30, 30),
        "secondary": (55, 55, 55),
        "accent1": (0, 200, 255),
        "accent2": (255, 82, 82),
        "accent3": (105, 240, 174),
        "accent4": (255, 196, 0),
        "background": (18, 18, 18),
        "text": (245, 245, 245),
    },
    "nature_green": {
        "primary": (27, 94, 32),
        "secondary": (56, 142, 60),
        "accent1": (129, 199, 132),
        "accent2": (255, 183, 77),
        "accent3": (100, 181, 246),
        "accent4": (161, 136, 127),
        "background": (245, 250, 245),
        "text": (33, 33, 33),
    },
    "sunset_warm": {
        "primary": (183, 28, 28),
        "secondary": (230, 74, 25),
        "accent1": (255, 143, 0),
        "accent2": (255, 202, 40),
        "accent3": (121, 85, 72),
        "accent4": (78, 52, 46),
        "background": (255, 253, 248),
        "text": (62, 39, 35),
    },
    "ocean_breeze": {
        "primary": (0, 105, 148),
        "secondary": (0, 151, 167),
        "accent1": (77, 208, 225),
        "accent2": (255, 171, 64),
        "accent3": (129, 212, 250),
        "accent4": (0, 77, 64),
        "background": (240, 248, 255),
        "text": (13, 71, 161),
    },
    "monochrome": {
        "primary": (33, 33, 33),
        "secondary": (66, 66, 66),
        "accent1": (117, 117, 117),
        "accent2": (158, 158, 158),
        "accent3": (189, 189, 189),
        "accent4": (224, 224, 224),
        "background": (250, 250, 250),
        "text": (33, 33, 33),
    },
    "royal_purple": {
        "primary": (74, 20, 140),
        "secondary": (106, 27, 154),
        "accent1": (171, 71, 188),
        "accent2": (255, 109, 0),
        "accent3": (206, 147, 216),
        "accent4": (69, 39, 160),
        "background": (248, 244, 252),
        "text": (38, 18, 60),
    },
    "tech_neon": {
        "primary": (18, 18, 30),
        "secondary": (30, 30, 50),
        "accent1": (0, 255, 136),
        "accent2": (0, 176, 255),
        "accent3": (255, 0, 128),
        "accent4": (168, 0, 255),
        "background": (10, 10, 20),
        "text": (230, 230, 250),
    },
    "pastel_soft": {
        "primary": (149, 117, 205),
        "secondary": (121, 134, 203),
        "accent1": (240, 152, 152),
        "accent2": (129, 212, 172),
        "accent3": (255, 213, 145),
        "accent4": (144, 202, 249),
        "background": (254, 252, 254),
        "text": (69, 69, 69),
    },
    "bold_contrast": {
        "primary": (0, 0, 0),
        "secondary": (255, 255, 255),
        "accent1": (255, 23, 68),
        "accent2": (0, 230, 118),
        "accent3": (41, 121, 255),
        "accent4": (255, 214, 0),
        "background": (255, 255, 255),
        "text": (0, 0, 0),
    },
}

# ============================================================
# Internal helpers
# ============================================================


def _get_app():
    """Get or create a PowerPoint COM application instance."""
    return get_or_create_app("PowerPoint.Application")


def _get_slide(slide_number):
    """Get a slide by number from the active presentation."""
    app = _get_app()
    prs = app.ActivePresentation
    return prs, prs.Slides(slide_number)


def _get_slide_dimensions():
    """Get the current slide dimensions."""
    app = _get_app()
    prs = app.ActivePresentation
    return prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight


def _add_shape(slide, shape_type, left, top, width, height,
               fill_rgb=None, fill_transparency=0.0,
               line_visible=False, line_rgb=None, line_weight=None,
               rotation=0.0, z_order_back=False):
    """Add a shape with common styling."""
    shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
    if fill_rgb:
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = rgb(*fill_rgb)
        if fill_transparency > 0:
            shape.Fill.Transparency = fill_transparency
    else:
        shape.Fill.Background()
    if line_visible and line_rgb:
        shape.Line.Visible = True
        shape.Line.ForeColor.RGB = rgb(*line_rgb)
        if line_weight:
            shape.Line.Weight = line_weight
    else:
        shape.Line.Visible = False
    if rotation:
        shape.Rotation = rotation
    if z_order_back:
        shape.ZOrder(MSO_SEND_TO_BACK)
    return shape


def _add_gradient_shape(slide, shape_type, left, top, width, height,
                        color1, color2, direction=MSO_GRADIENT_HORIZONTAL,
                        transparency=0.0, z_order_back=False):
    """Add a shape with gradient fill."""
    shape = slide.Shapes.AddShape(shape_type, left, top, width, height)
    fill = shape.Fill
    fill.TwoColorGradient(direction, 1)
    fill.ForeColor.RGB = rgb(*color1)
    fill.BackColor.RGB = rgb(*color2)
    if transparency > 0:
        fill.Transparency = transparency
    shape.Line.Visible = False
    if z_order_back:
        shape.ZOrder(MSO_SEND_TO_BACK)
    return shape


def _add_textbox(slide, left, top, width, height, text,
                 font_name=None, font_size=None, font_color=None,
                 bold=False, italic=False, alignment=PP_ALIGN_LEFT,
                 fill_rgb=None, fill_transparency=0.0,
                 vertical_anchor=None):
    """Add a text box with styling."""
    shape = slide.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL,
                                    left, top, width, height)
    tf = shape.TextFrame
    tf.WordWrap = True
    if vertical_anchor is not None:
        tf.VerticalAnchor = vertical_anchor
    tr = tf.TextRange
    tr.Text = text
    tr.ParagraphFormat.Alignment = alignment
    font = tr.Font
    if font_name:
        font.Name = font_name
    if font_size:
        font.Size = font_size
    if font_color:
        font.Color.RGB = rgb(*font_color)
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
    return shape


def _set_solid_bg(slide, color_rgb):
    """Set solid background on a slide."""
    slide.FollowMasterBackground = False
    slide.Background.Fill.Solid()
    slide.Background.Fill.ForeColor.RGB = rgb(*color_rgb)


def _set_gradient_bg(slide, color1, color2, direction=MSO_GRADIENT_HORIZONTAL):
    """Set gradient background on a slide."""
    slide.FollowMasterBackground = False
    fill = slide.Background.Fill
    fill.TwoColorGradient(direction, 1)
    fill.ForeColor.RGB = rgb(*color1)
    fill.BackColor.RGB = rgb(*color2)


def _add_shadow(shape, blur=8, offset_x=3, offset_y=3, transparency=0.6):
    """Add a drop shadow to a shape."""
    shape.Shadow.Visible = True
    shape.Shadow.Type = MSO_SHADOW_21
    shape.Shadow.Blur = blur
    shape.Shadow.OffsetX = offset_x
    shape.Shadow.OffsetY = offset_y
    shape.Shadow.Transparency = transparency
    shape.Shadow.ForeColor.RGB = rgb(0, 0, 0)


def _position_map(sw, sh):
    """Return position coordinates for common positions."""
    return {
        "top_left": (20, 20),
        "top_right": (sw - 200, 20),
        "bottom_left": (20, sh - 200),
        "bottom_right": (sw - 200, sh - 200),
        "center": (sw / 2 - 100, sh / 2 - 100),
        "background": (0, 0),
    }


def _size_map():
    """Return size multipliers."""
    return {"small": 0.6, "medium": 1.0, "large": 1.5}


def _lerp_color(c1, c2, t):
    """Linearly interpolate between two RGB colors."""
    return (
        int(c1[0] + (c2[0] - c1[0]) * t),
        int(c1[1] + (c2[1] - c1[1]) * t),
        int(c1[2] + (c2[2] - c1[2]) * t),
    )


# ============================================================
# 1. apply_theme_colors
# ============================================================

def apply_theme_colors(color_scheme: str) -> dict:
    """Apply a named color scheme to the active presentation's slide master.

    Args:
        color_scheme: Name of the color scheme. Options: "corporate_blue",
            "modern_dark", "nature_green", "sunset_warm", "ocean_breeze",
            "monochrome", "royal_purple", "tech_neon", "pastel_soft",
            "bold_contrast".

    Returns:
        Dict with applied scheme name and colors.
    """
    if color_scheme not in COLOR_SCHEMES:
        raise ValueError(
            f"Unknown color scheme '{color_scheme}'. "
            f"Available: {', '.join(COLOR_SCHEMES.keys())}"
        )

    scheme = COLOR_SCHEMES[color_scheme]
    app = _get_app()
    prs = app.ActivePresentation
    master = prs.SlideMaster

    # Apply background color to the slide master
    master.Background.Fill.Solid()
    master.Background.Fill.ForeColor.RGB = rgb(*scheme["background"])

    # Apply text color to master title and body styles
    try:
        title_style = master.TextStyles(1)  # Title style
        title_style.TextFrame.TextRange.Font.Color.RGB = rgb(*scheme["text"])
    except Exception:
        pass

    try:
        body_style = master.TextStyles(2)  # Body style
        body_style.TextFrame.TextRange.Font.Color.RGB = rgb(*scheme["text"])
    except Exception:
        pass

    # Apply theme colors via the color scheme object
    try:
        theme = prs.SlideMaster.Theme
        color_obj = theme.ThemeColorScheme
        # Map our scheme to Office theme color indices
        # 1=Background1, 2=Text1, 3=Background2, 4=Text2, 5=Accent1..10=Accent6
        color_obj(1).RGB = rgb(*scheme["background"])
        color_obj(2).RGB = rgb(*scheme["text"])
        color_obj(3).RGB = rgb(*scheme["secondary"])
        color_obj(4).RGB = rgb(*scheme["primary"])
        color_obj(5).RGB = rgb(*scheme["accent1"])
        color_obj(6).RGB = rgb(*scheme["accent2"])
        color_obj(7).RGB = rgb(*scheme["accent3"])
        color_obj(8).RGB = rgb(*scheme["accent4"])
    except Exception:
        pass

    return {"scheme": color_scheme, "colors": scheme}


# ============================================================
# 2. set_master_font
# ============================================================

def set_master_font(title_font: str, body_font: str) -> dict:
    """Set the master slide title and body fonts.

    Args:
        title_font: Font name for titles (e.g., "Segoe UI", "Arial").
        body_font: Font name for body text.

    Returns:
        Dict with applied font names.
    """
    app = _get_app()
    prs = app.ActivePresentation
    master = prs.SlideMaster

    # Set title font via TextStyles
    try:
        title_style = master.TextStyles(1)  # Title
        for level in range(1, 10):
            try:
                title_style.Levels(level).Font.Name = title_font
            except Exception:
                break
    except Exception:
        pass

    # Set body font via TextStyles
    try:
        body_style = master.TextStyles(2)  # Body
        for level in range(1, 10):
            try:
                body_style.Levels(level).Font.Name = body_font
            except Exception:
                break
    except Exception:
        pass

    # Also set via Theme fonts if available
    try:
        theme = prs.SlideMaster.Theme
        theme.ThemeFontScheme.MajorFont(1).Name = title_font
        theme.ThemeFontScheme.MinorFont(1).Name = body_font
    except Exception:
        pass

    return {"title_font": title_font, "body_font": body_font}


# ============================================================
# 3. add_watermark
# ============================================================

def add_watermark(
    text: str,
    slide_number: int | None = None,
    font_size: float = 54,
    color: tuple[int, int, int] = (200, 200, 200),
    rotation: float = -45,
    transparency: float = 75,
) -> dict:
    """Add a diagonal watermark text to one or all slides.

    Args:
        text: Watermark text.
        slide_number: Specific slide (1-based), or None for all slides.
        font_size: Font size in points.
        color: RGB color tuple.
        rotation: Rotation in degrees (negative = counter-clockwise).
        transparency: Transparency percentage 0-100.

    Returns:
        Dict with affected slide count.
    """
    app = _get_app()
    prs = app.ActivePresentation
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if slide_number:
        slides = [prs.Slides(slide_number)]
    else:
        slides = [prs.Slides(i) for i in range(1, prs.Slides.Count + 1)]

    trans_fraction = transparency / 100.0

    for slide in slides:
        # Center the watermark on the slide
        box_w = sw * 0.8
        box_h = font_size * 2
        left = (sw - box_w) / 2
        top = (sh - box_h) / 2

        shape = slide.Shapes.AddTextbox(
            MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, box_w, box_h
        )
        tf = shape.TextFrame
        tf.WordWrap = False
        tr = tf.TextRange
        tr.Text = text
        tr.ParagraphFormat.Alignment = PP_ALIGN_CENTER
        tr.Font.Size = font_size
        tr.Font.Color.RGB = rgb(*color)
        tr.Font.Bold = False

        shape.Fill.Background()
        shape.Line.Visible = False
        shape.Rotation = rotation

        # Set transparency on the font color
        try:
            tr.Font.Color.Brightness = trans_fraction
        except Exception:
            pass

        # Send to back so it doesn't obscure content
        shape.ZOrder(MSO_SEND_TO_BACK)

    return {"slides_affected": len(slides), "text": text}


# ============================================================
# 4. create_title_slide_design
# ============================================================

def create_title_slide_design(
    slide_number: int,
    title: str,
    subtitle: str | None = None,
    style: str = "modern_gradient",
) -> dict:
    """Create a professionally designed title slide.

    Args:
        slide_number: 1-based slide index.
        title: Main title text.
        subtitle: Optional subtitle text.
        style: Design style - "modern_gradient", "minimal_white",
            "bold_split", "dark_premium", "geometric".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "modern_gradient":
        _design_title_modern_gradient(slide, sw, sh, title, subtitle)
    elif style == "minimal_white":
        _design_title_minimal_white(slide, sw, sh, title, subtitle)
    elif style == "bold_split":
        _design_title_bold_split(slide, sw, sh, title, subtitle)
    elif style == "dark_premium":
        _design_title_dark_premium(slide, sw, sh, title, subtitle)
    elif style == "geometric":
        _design_title_geometric(slide, sw, sh, title, subtitle)
    else:
        raise ValueError(f"Unknown title style '{style}'")

    return {"slide_number": slide_number, "style": style}


def _design_title_modern_gradient(slide, sw, sh, title, subtitle):
    """Modern gradient background with large title and subtle subtitle."""
    _set_gradient_bg(slide, (0, 82, 136), (0, 150, 200), MSO_GRADIENT_DIAGONAL_DOWN)

    # Subtle decorative circle in top right
    _add_shape(slide, MSO_SHAPE_OVAL, sw - 250, -80, 400, 400,
               fill_rgb=(255, 255, 255), fill_transparency=0.9)

    # Another smaller circle
    _add_shape(slide, MSO_SHAPE_OVAL, sw - 180, sh - 200, 250, 250,
               fill_rgb=(255, 255, 255), fill_transparency=0.92)

    # Title
    _add_textbox(slide, sw * 0.08, sh * 0.3, sw * 0.84, sh * 0.25, title,
                 font_name="Segoe UI Light", font_size=48, font_color=(255, 255, 255),
                 bold=False, alignment=PP_ALIGN_LEFT)

    # Thin accent line
    line = slide.Shapes.AddLine(sw * 0.08, sh * 0.58, sw * 0.25, sh * 0.58)
    line.Line.ForeColor.RGB = rgb(255, 200, 50)
    line.Line.Weight = 3

    # Subtitle
    if subtitle:
        _add_textbox(slide, sw * 0.08, sh * 0.62, sw * 0.6, sh * 0.12, subtitle,
                     font_name="Segoe UI", font_size=20, font_color=(220, 235, 250),
                     alignment=PP_ALIGN_LEFT)


def _design_title_minimal_white(slide, sw, sh, title, subtitle):
    """White background with thin accent line and clean typography."""
    _set_solid_bg(slide, (255, 255, 255))

    # Left accent bar
    _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, 0, 8, sh,
               fill_rgb=(0, 82, 136))

    # Thin horizontal line
    line = slide.Shapes.AddLine(sw * 0.06, sh * 0.55, sw * 0.35, sh * 0.55)
    line.Line.ForeColor.RGB = rgb(0, 82, 136)
    line.Line.Weight = 2

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.25, sw * 0.88, sh * 0.25, title,
                 font_name="Segoe UI Semibold", font_size=44, font_color=(33, 33, 33),
                 bold=False, alignment=PP_ALIGN_LEFT)

    # Subtitle
    if subtitle:
        _add_textbox(slide, sw * 0.06, sh * 0.6, sw * 0.65, sh * 0.1, subtitle,
                     font_name="Segoe UI", font_size=18, font_color=(120, 120, 120),
                     alignment=PP_ALIGN_LEFT)

    # Bottom-right decorative square
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - 120, sh - 120, 120, 120,
               fill_rgb=(0, 82, 136), fill_transparency=0.08)


def _design_title_bold_split(slide, sw, sh, title, subtitle):
    """Half color / half image area with bold title."""
    _set_solid_bg(slide, (245, 245, 245))

    # Left colored half
    _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE, 0, 0, sw * 0.55, sh,
                        (0, 82, 136), (0, 55, 100), MSO_GRADIENT_VERTICAL)

    # Right side remains light
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.55, 0, sw * 0.45, sh,
               fill_rgb=(240, 242, 245))

    # Title on left side
    _add_textbox(slide, sw * 0.06, sh * 0.3, sw * 0.44, sh * 0.3, title,
                 font_name="Segoe UI", font_size=42, font_color=(255, 255, 255),
                 bold=True, alignment=PP_ALIGN_LEFT)

    # Subtitle below title
    if subtitle:
        _add_textbox(slide, sw * 0.06, sh * 0.63, sw * 0.44, sh * 0.12, subtitle,
                     font_name="Segoe UI Light", font_size=18,
                     font_color=(200, 220, 240), alignment=PP_ALIGN_LEFT)

    # Decorative accent on the dividing line
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.545, sh * 0.2, 6, sh * 0.6,
               fill_rgb=(255, 200, 50))


def _design_title_dark_premium(slide, sw, sh, title, subtitle):
    """Dark background with gold/white accents."""
    _set_gradient_bg(slide, (20, 20, 30), (40, 40, 55), MSO_GRADIENT_DIAGONAL_DOWN)

    # Gold accent line top
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.1, sh * 0.25, sw * 0.15, 3,
               fill_rgb=(212, 175, 55))

    # Title
    _add_textbox(slide, sw * 0.1, sh * 0.3, sw * 0.8, sh * 0.25, title,
                 font_name="Georgia", font_size=48, font_color=(255, 255, 255),
                 bold=False, alignment=PP_ALIGN_LEFT)

    # Gold accent line below title
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.1, sh * 0.58, sw * 0.15, 3,
               fill_rgb=(212, 175, 55))

    # Subtitle
    if subtitle:
        _add_textbox(slide, sw * 0.1, sh * 0.63, sw * 0.7, sh * 0.1, subtitle,
                     font_name="Georgia", font_size=20, font_color=(180, 180, 190),
                     italic=True, alignment=PP_ALIGN_LEFT)

    # Corner accent - top right
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - 80, 0, 80, 4,
               fill_rgb=(212, 175, 55))
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - 4, 0, 4, 80,
               fill_rgb=(212, 175, 55))

    # Corner accent - bottom left
    _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, sh - 4, 80, 4,
               fill_rgb=(212, 175, 55))
    _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, sh - 80, 4, 80,
               fill_rgb=(212, 175, 55))


def _design_title_geometric(slide, sw, sh, title, subtitle):
    """Abstract geometric shapes with title overlay."""
    _set_solid_bg(slide, (245, 247, 250))

    # Large triangle - bottom right
    _add_shape(slide, MSO_SHAPE_ISOSCELES_TRIANGLE, sw * 0.5, sh * 0.3,
               sw * 0.6, sh * 0.8,
               fill_rgb=(0, 82, 136), fill_transparency=0.12, rotation=15)

    # Medium circle - top left
    _add_shape(slide, MSO_SHAPE_OVAL, -60, -60, 300, 300,
               fill_rgb=(0, 166, 214), fill_transparency=0.15)

    # Small diamond
    _add_shape(slide, MSO_SHAPE_DIAMOND, sw * 0.7, sh * 0.05, 120, 120,
               fill_rgb=(242, 169, 0), fill_transparency=0.2)

    # Small circle
    _add_shape(slide, MSO_SHAPE_OVAL, sw * 0.15, sh * 0.7, 80, 80,
               fill_rgb=(0, 150, 136), fill_transparency=0.2)

    # Rectangle accent
    _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.85, sh * 0.6, 60, 60,
               fill_rgb=(0, 82, 136), fill_transparency=0.1, rotation=30)

    # Title with semi-transparent background strip
    strip = _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, sh * 0.32, sw, sh * 0.2,
                       fill_rgb=(255, 255, 255), fill_transparency=0.15)

    _add_textbox(slide, sw * 0.08, sh * 0.33, sw * 0.84, sh * 0.18, title,
                 font_name="Segoe UI", font_size=44, font_color=(33, 37, 41),
                 bold=True, alignment=PP_ALIGN_LEFT)

    if subtitle:
        _add_textbox(slide, sw * 0.08, sh * 0.56, sw * 0.7, sh * 0.1, subtitle,
                     font_name="Segoe UI Light", font_size=20,
                     font_color=(80, 80, 100), alignment=PP_ALIGN_LEFT)


# ============================================================
# 5. create_content_slide_design
# ============================================================

def create_content_slide_design(
    slide_number: int,
    title: str,
    items: list[str],
    style: str = "cards",
) -> dict:
    """Create a professionally designed content slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        items: List of content items/bullet points.
        style: Design style - "cards", "timeline", "comparison",
            "stats", "icon_grid".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "cards":
        _design_content_cards(slide, sw, sh, title, items)
    elif style == "timeline":
        _design_content_timeline(slide, sw, sh, title, items)
    elif style == "comparison":
        _design_content_comparison(slide, sw, sh, title, items)
    elif style == "stats":
        _design_content_stats(slide, sw, sh, title, items)
    elif style == "icon_grid":
        _design_content_icon_grid(slide, sw, sh, title, items)
    else:
        raise ValueError(f"Unknown content style '{style}'")

    return {"slide_number": slide_number, "style": style}


def _design_content_cards(slide, sw, sh, title, items):
    """Card layout with colored accent bars."""
    _set_solid_bg(slide, (245, 247, 250))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=32,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

    # Title underline
    line = slide.Shapes.AddLine(sw * 0.06, sh * 0.17, sw * 0.2, sh * 0.17)
    line.Line.ForeColor.RGB = rgb(0, 120, 190)
    line.Line.Weight = 3

    # Cards
    n = len(items)
    if n == 0:
        return

    cols = min(n, 3)
    rows = (n + cols - 1) // cols
    card_margin = 18
    content_left = sw * 0.06
    content_width = sw * 0.88
    content_top = sh * 0.22
    content_height = sh * 0.72

    card_w = (content_width - card_margin * (cols - 1)) / cols
    card_h = (content_height - card_margin * (rows - 1)) / rows
    card_h = min(card_h, 160)

    accent_colors = [
        (0, 120, 190), (0, 166, 214), (242, 169, 0),
        (0, 150, 136), (96, 125, 139), (183, 28, 28),
    ]

    for i, item in enumerate(items):
        col = i % cols
        row = i // cols
        x = content_left + col * (card_w + card_margin)
        y = content_top + row * (card_h + card_margin)
        accent = accent_colors[i % len(accent_colors)]

        # Card background
        card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE, x, y, card_w, card_h,
                          fill_rgb=(255, 255, 255))
        _add_shadow(card, blur=6, offset_x=2, offset_y=2, transparency=0.7)

        # Top accent bar
        _add_shape(slide, MSO_SHAPE_RECTANGLE, x, y, card_w, 5, fill_rgb=accent)

        # Card text
        _add_textbox(slide, x + 16, y + 20, card_w - 32, card_h - 30, item,
                     font_name="Segoe UI", font_size=14, font_color=(50, 50, 50),
                     alignment=PP_ALIGN_LEFT, vertical_anchor=MSO_ANCHOR_TOP)


def _design_content_timeline(slide, sw, sh, title, items):
    """Horizontal timeline layout."""
    _set_solid_bg(slide, (250, 250, 252))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=32,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

    n = len(items)
    if n == 0:
        return

    line_y = sh * 0.5
    margin_x = sw * 0.1
    usable_width = sw * 0.8

    # Main timeline line
    line = slide.Shapes.AddLine(margin_x, line_y, margin_x + usable_width, line_y)
    line.Line.ForeColor.RGB = rgb(200, 200, 210)
    line.Line.Weight = 3

    step = usable_width / max(n - 1, 1) if n > 1 else 0
    colors = [
        (0, 120, 190), (0, 166, 214), (0, 150, 136),
        (242, 169, 0), (183, 28, 28), (106, 27, 154),
    ]

    for i, item in enumerate(items):
        cx = margin_x + i * step if n > 1 else margin_x + usable_width / 2
        c = colors[i % len(colors)]

        # Circle node
        node_r = 14
        _add_shape(slide, MSO_SHAPE_OVAL,
                   cx - node_r, line_y - node_r, node_r * 2, node_r * 2,
                   fill_rgb=c)

        # Alternating above/below
        if i % 2 == 0:
            text_y = line_y - 130
            # Connector line
            conn = slide.Shapes.AddLine(cx, line_y - node_r - 2, cx, text_y + 70)
            conn.Line.ForeColor.RGB = rgb(*c)
            conn.Line.Weight = 1.5
        else:
            text_y = line_y + 40
            conn = slide.Shapes.AddLine(cx, line_y + node_r + 2, cx, text_y)
            conn.Line.ForeColor.RGB = rgb(*c)
            conn.Line.Weight = 1.5

        # Text
        text_w = min(step * 0.85, 180) if n > 1 else 200
        _add_textbox(slide, cx - text_w / 2, text_y, text_w, 70, item,
                     font_name="Segoe UI", font_size=12, font_color=(50, 50, 60),
                     alignment=PP_ALIGN_CENTER)


def _design_content_comparison(slide, sw, sh, title, items):
    """Side-by-side comparison with items split in half."""
    _set_solid_bg(slide, (248, 249, 252))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=32,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

    mid = len(items) // 2
    left_items = items[:mid] if mid > 0 else items[:1]
    right_items = items[mid:] if mid > 0 else items[1:]

    col_w = sw * 0.4
    left_x = sw * 0.06
    right_x = sw * 0.54

    # Left column background
    _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
               left_x, sh * 0.2, col_w, sh * 0.72,
               fill_rgb=(0, 82, 136), fill_transparency=0.05)

    # Right column background
    _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
               right_x, sh * 0.2, col_w, sh * 0.72,
               fill_rgb=(0, 150, 136), fill_transparency=0.05)

    # Left items
    for i, item in enumerate(left_items):
        y = sh * 0.25 + i * 55
        _add_shape(slide, MSO_SHAPE_OVAL, left_x + 15, y + 5, 10, 10,
                   fill_rgb=(0, 82, 136))
        _add_textbox(slide, left_x + 35, y, col_w - 50, 45, item,
                     font_name="Segoe UI", font_size=14, font_color=(40, 40, 50),
                     alignment=PP_ALIGN_LEFT)

    # Right items
    for i, item in enumerate(right_items):
        y = sh * 0.25 + i * 55
        _add_shape(slide, MSO_SHAPE_OVAL, right_x + 15, y + 5, 10, 10,
                   fill_rgb=(0, 150, 136))
        _add_textbox(slide, right_x + 35, y, col_w - 50, 45, item,
                     font_name="Segoe UI", font_size=14, font_color=(40, 40, 50),
                     alignment=PP_ALIGN_LEFT)


def _design_content_stats(slide, sw, sh, title, items):
    """Big number statistics layout. Items format: 'number|label' or just text."""
    _set_gradient_bg(slide, (0, 60, 110), (0, 100, 160), MSO_GRADIENT_DIAGONAL_DOWN)

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Light", font_size=30,
                 font_color=(255, 255, 255), alignment=PP_ALIGN_LEFT)

    n = len(items)
    if n == 0:
        return

    cols = min(n, 4)
    margin = sw * 0.06
    usable = sw - 2 * margin
    card_w = (usable - 20 * (cols - 1)) / cols
    card_h = sh * 0.55

    for i, item in enumerate(items):
        if i >= cols:
            break
        x = margin + i * (card_w + 20)
        y = sh * 0.25

        # Parse number|label format
        if "|" in item:
            number, label = item.split("|", 1)
        else:
            number = item
            label = ""

        # Semi-transparent card
        card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE, x, y, card_w, card_h,
                          fill_rgb=(255, 255, 255), fill_transparency=0.85)

        # Big number
        _add_textbox(slide, x + 10, y + card_h * 0.15, card_w - 20, card_h * 0.4,
                     number.strip(),
                     font_name="Segoe UI Light", font_size=52,
                     font_color=(255, 255, 255), bold=False,
                     alignment=PP_ALIGN_CENTER)

        # Label
        if label:
            _add_textbox(slide, x + 10, y + card_h * 0.6, card_w - 20, card_h * 0.3,
                         label.strip(),
                         font_name="Segoe UI", font_size=14,
                         font_color=(200, 220, 240),
                         alignment=PP_ALIGN_CENTER)


def _design_content_icon_grid(slide, sw, sh, title, items):
    """2x2 or 3x2 grid with colored icon circles."""
    _set_solid_bg(slide, (250, 250, 252))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=32,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

    n = len(items)
    if n == 0:
        return

    cols = 3 if n > 4 else 2
    rows = (n + cols - 1) // cols
    margin_x = sw * 0.08
    margin_y = sh * 0.22
    usable_w = sw * 0.84
    usable_h = sh * 0.7
    cell_w = usable_w / cols
    cell_h = min(usable_h / rows, 200)

    accent_colors = [
        (0, 120, 190), (0, 166, 214), (242, 169, 0),
        (0, 150, 136), (183, 28, 28), (106, 27, 154),
    ]

    for i, item in enumerate(items):
        col = i % cols
        row = i // cols
        cx = margin_x + col * cell_w + cell_w / 2
        cy = margin_y + row * cell_h
        c = accent_colors[i % len(accent_colors)]

        # Icon circle
        circle_r = 28
        _add_shape(slide, MSO_SHAPE_OVAL,
                   cx - circle_r, cy, circle_r * 2, circle_r * 2,
                   fill_rgb=c)

        # Number in circle
        _add_textbox(slide, cx - circle_r, cy, circle_r * 2, circle_r * 2,
                     str(i + 1),
                     font_name="Segoe UI", font_size=18,
                     font_color=(255, 255, 255), bold=True,
                     alignment=PP_ALIGN_CENTER, vertical_anchor=MSO_ANCHOR_MIDDLE)

        # Item text below circle
        _add_textbox(slide, cx - cell_w * 0.4, cy + circle_r * 2 + 12,
                     cell_w * 0.8, cell_h - circle_r * 2 - 20, item,
                     font_name="Segoe UI", font_size=13, font_color=(60, 60, 70),
                     alignment=PP_ALIGN_CENTER)


# ============================================================
# 6. create_section_divider
# ============================================================

def create_section_divider(
    slide_number: int,
    title: str,
    subtitle: str | None = None,
    style: str = "gradient_wave",
) -> dict:
    """Create a section divider slide.

    Args:
        slide_number: 1-based slide index.
        title: Section title.
        subtitle: Optional subtitle.
        style: Design style - "gradient_wave", "bold_number",
            "minimal_line", "full_bleed_color".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "gradient_wave":
        _set_gradient_bg(slide, (0, 82, 136), (0, 140, 190), MSO_GRADIENT_DIAGONAL_UP)

        # Wave-like curved shapes
        _add_shape(slide, MSO_SHAPE_OVAL, -sw * 0.3, sh * 0.6,
                   sw * 1.6, sh * 0.8,
                   fill_rgb=(255, 255, 255), fill_transparency=0.92)
        _add_shape(slide, MSO_SHAPE_OVAL, -sw * 0.2, sh * 0.7,
                   sw * 1.4, sh * 0.7,
                   fill_rgb=(255, 255, 255), fill_transparency=0.95)

        _add_textbox(slide, sw * 0.1, sh * 0.3, sw * 0.8, sh * 0.2, title,
                     font_name="Segoe UI Light", font_size=44,
                     font_color=(255, 255, 255), bold=False,
                     alignment=PP_ALIGN_CENTER)
        if subtitle:
            _add_textbox(slide, sw * 0.15, sh * 0.52, sw * 0.7, sh * 0.1, subtitle,
                         font_name="Segoe UI", font_size=18,
                         font_color=(220, 235, 250), alignment=PP_ALIGN_CENTER)

    elif style == "bold_number":
        _set_solid_bg(slide, (248, 249, 252))

        # Large faded section number background
        # Extract number from title if possible, else use "01"
        import re
        num_match = re.search(r'\d+', title)
        num_text = num_match.group() if num_match else "01"
        num_text = num_text.zfill(2)

        _add_textbox(slide, sw * 0.55, sh * 0.05, sw * 0.45, sh * 0.9, num_text,
                     font_name="Segoe UI Black", font_size=200,
                     font_color=(0, 82, 136), bold=True,
                     alignment=PP_ALIGN_RIGHT)
        # Make the number semi-transparent by overlaying
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.55, 0, sw * 0.45, sh,
                   fill_rgb=(248, 249, 252), fill_transparency=0.15)

        # Vertical accent bar
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.08, sh * 0.25, 6, sh * 0.5,
                   fill_rgb=(0, 120, 190))

        _add_textbox(slide, sw * 0.12, sh * 0.35, sw * 0.5, sh * 0.15, title,
                     font_name="Segoe UI", font_size=36,
                     font_color=(33, 37, 41), bold=True,
                     alignment=PP_ALIGN_LEFT)
        if subtitle:
            _add_textbox(slide, sw * 0.12, sh * 0.52, sw * 0.5, sh * 0.1, subtitle,
                         font_name="Segoe UI", font_size=16,
                         font_color=(100, 100, 110), alignment=PP_ALIGN_LEFT)

    elif style == "minimal_line":
        _set_solid_bg(slide, (255, 255, 255))

        # Center line
        line = slide.Shapes.AddLine(sw * 0.35, sh * 0.47, sw * 0.65, sh * 0.47)
        line.Line.ForeColor.RGB = rgb(0, 82, 136)
        line.Line.Weight = 2

        _add_textbox(slide, sw * 0.1, sh * 0.5, sw * 0.8, sh * 0.15, title,
                     font_name="Segoe UI Light", font_size=38,
                     font_color=(50, 50, 55), alignment=PP_ALIGN_CENTER)
        if subtitle:
            _add_textbox(slide, sw * 0.2, sh * 0.66, sw * 0.6, sh * 0.08, subtitle,
                         font_name="Segoe UI", font_size=16,
                         font_color=(140, 140, 150), alignment=PP_ALIGN_CENTER)

    elif style == "full_bleed_color":
        _set_solid_bg(slide, (0, 82, 136))

        # Subtle overlay shape
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, 0, sw, sh,
                   fill_rgb=(0, 0, 0), fill_transparency=0.85, z_order_back=True)

        _add_textbox(slide, sw * 0.1, sh * 0.35, sw * 0.8, sh * 0.2, title,
                     font_name="Segoe UI", font_size=46,
                     font_color=(255, 255, 255), bold=True,
                     alignment=PP_ALIGN_CENTER)
        if subtitle:
            _add_textbox(slide, sw * 0.15, sh * 0.58, sw * 0.7, sh * 0.1, subtitle,
                         font_name="Segoe UI Light", font_size=20,
                         font_color=(200, 220, 240), alignment=PP_ALIGN_CENTER)
    else:
        raise ValueError(f"Unknown section divider style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 7. create_closing_slide
# ============================================================

def create_closing_slide(
    slide_number: int,
    title: str = "Thank You",
    subtitle: str | None = None,
    contact_info: str | None = None,
    style: str = "elegant",
) -> dict:
    """Create a closing/thank you slide.

    Args:
        slide_number: 1-based slide index.
        title: Main closing text.
        subtitle: Optional subtitle.
        contact_info: Optional contact information.
        style: Design style - "elegant", "minimal", "bold".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "elegant":
        _set_gradient_bg(slide, (20, 25, 40), (40, 50, 75), MSO_GRADIENT_FROM_CORNER)

        # Decorative circles
        _add_shape(slide, MSO_SHAPE_OVAL, sw * 0.7, -100, 350, 350,
                   fill_rgb=(212, 175, 55), fill_transparency=0.9)
        _add_shape(slide, MSO_SHAPE_OVAL, -120, sh * 0.6, 300, 300,
                   fill_rgb=(212, 175, 55), fill_transparency=0.92)

        # Gold line
        _add_shape(slide, MSO_SHAPE_RECTANGLE,
                   sw * 0.35, sh * 0.32, sw * 0.3, 2,
                   fill_rgb=(212, 175, 55))

        _add_textbox(slide, sw * 0.1, sh * 0.36, sw * 0.8, sh * 0.18, title,
                     font_name="Georgia", font_size=48,
                     font_color=(255, 255, 255), alignment=PP_ALIGN_CENTER)

        _add_shape(slide, MSO_SHAPE_RECTANGLE,
                   sw * 0.35, sh * 0.56, sw * 0.3, 2,
                   fill_rgb=(212, 175, 55))

        if subtitle:
            _add_textbox(slide, sw * 0.15, sh * 0.6, sw * 0.7, sh * 0.08, subtitle,
                         font_name="Georgia", font_size=18,
                         font_color=(180, 180, 195), italic=True,
                         alignment=PP_ALIGN_CENTER)

        if contact_info:
            _add_textbox(slide, sw * 0.2, sh * 0.78, sw * 0.6, sh * 0.1,
                         contact_info,
                         font_name="Segoe UI", font_size=13,
                         font_color=(150, 155, 170), alignment=PP_ALIGN_CENTER)

    elif style == "minimal":
        _set_solid_bg(slide, (255, 255, 255))

        _add_textbox(slide, sw * 0.1, sh * 0.35, sw * 0.8, sh * 0.18, title,
                     font_name="Segoe UI Light", font_size=48,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        # Thin line
        line = slide.Shapes.AddLine(sw * 0.4, sh * 0.55, sw * 0.6, sh * 0.55)
        line.Line.ForeColor.RGB = rgb(0, 82, 136)
        line.Line.Weight = 2

        if subtitle:
            _add_textbox(slide, sw * 0.2, sh * 0.58, sw * 0.6, sh * 0.08, subtitle,
                         font_name="Segoe UI", font_size=16,
                         font_color=(120, 120, 130), alignment=PP_ALIGN_CENTER)

        if contact_info:
            _add_textbox(slide, sw * 0.2, sh * 0.75, sw * 0.6, sh * 0.1,
                         contact_info,
                         font_name="Segoe UI", font_size=12,
                         font_color=(140, 140, 150), alignment=PP_ALIGN_CENTER)

        # Bottom accent bar
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, sh - 6, sw, 6,
                   fill_rgb=(0, 82, 136))

    elif style == "bold":
        _set_solid_bg(slide, (0, 82, 136))

        _add_textbox(slide, sw * 0.08, sh * 0.28, sw * 0.84, sh * 0.25, title,
                     font_name="Segoe UI", font_size=56,
                     font_color=(255, 255, 255), bold=True,
                     alignment=PP_ALIGN_CENTER)

        if subtitle:
            _add_textbox(slide, sw * 0.15, sh * 0.56, sw * 0.7, sh * 0.08, subtitle,
                         font_name="Segoe UI Light", font_size=20,
                         font_color=(200, 220, 240), alignment=PP_ALIGN_CENTER)

        if contact_info:
            # Contact info in a white rounded rect
            info_h = 50
            info_w = sw * 0.5
            info_x = (sw - info_w) / 2
            info_y = sh * 0.72
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              info_x, info_y, info_w, info_h,
                              fill_rgb=(255, 255, 255), fill_transparency=0.15)
            _add_textbox(slide, info_x + 15, info_y + 5, info_w - 30, info_h - 10,
                         contact_info,
                         font_name="Segoe UI", font_size=13,
                         font_color=(220, 230, 245), alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

        # Corner accents
        accent_size = 60
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, 0, accent_size, 4,
                   fill_rgb=(255, 255, 255))
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, 0, 4, accent_size,
                   fill_rgb=(255, 255, 255))
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - accent_size, sh - 4,
                   accent_size, 4, fill_rgb=(255, 255, 255))
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - 4, sh - accent_size,
                   4, accent_size, fill_rgb=(255, 255, 255))
    else:
        raise ValueError(f"Unknown closing slide style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 8. add_decorative_element
# ============================================================

def add_decorative_element(
    slide_number: int,
    element_type: str,
    position: str = "top_right",
    color: tuple[int, int, int] | None = None,
    size: str = "medium",
) -> dict:
    """Add decorative elements to a slide.

    Args:
        slide_number: 1-based slide index.
        element_type: Type of decoration - "circle_cluster", "diagonal_lines",
            "dot_pattern", "corner_accent", "gradient_bar", "wave_shape".
        position: Position - "top_left", "top_right", "bottom_left",
            "bottom_right", "background".
        color: Optional RGB tuple. Defaults to blue.
        size: "small", "medium", "large".

    Returns:
        Dict with slide_number and element_type.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    c = color or (0, 120, 190)
    sizes = _size_map()
    s = sizes.get(size, 1.0)

    pos = _position_map(sw, sh)
    px, py = pos.get(position, pos["top_right"])

    if element_type == "circle_cluster":
        _deco_circle_cluster(slide, px, py, c, s)
    elif element_type == "diagonal_lines":
        _deco_diagonal_lines(slide, sw, sh, px, py, c, s)
    elif element_type == "dot_pattern":
        _deco_dot_pattern(slide, px, py, c, s)
    elif element_type == "corner_accent":
        _deco_corner_accent(slide, sw, sh, position, c, s)
    elif element_type == "gradient_bar":
        _deco_gradient_bar(slide, sw, sh, position, c, s)
    elif element_type == "wave_shape":
        _deco_wave_shape(slide, sw, sh, position, c, s)
    else:
        raise ValueError(f"Unknown element type '{element_type}'")

    return {"slide_number": slide_number, "element_type": element_type}


def _deco_circle_cluster(slide, px, py, color, scale):
    """Cluster of overlapping semi-transparent circles."""
    base = 80 * scale
    _add_shape(slide, MSO_SHAPE_OVAL, px, py, base, base,
               fill_rgb=color, fill_transparency=0.7)
    _add_shape(slide, MSO_SHAPE_OVAL, px + base * 0.5, py - base * 0.3,
               base * 0.7, base * 0.7,
               fill_rgb=color, fill_transparency=0.8)
    _add_shape(slide, MSO_SHAPE_OVAL, px + base * 0.2, py + base * 0.6,
               base * 0.5, base * 0.5,
               fill_rgb=color, fill_transparency=0.75)
    _add_shape(slide, MSO_SHAPE_OVAL, px - base * 0.3, py + base * 0.3,
               base * 0.6, base * 0.6,
               fill_rgb=color, fill_transparency=0.85)


def _deco_diagonal_lines(slide, sw, sh, px, py, color, scale):
    """Set of diagonal lines."""
    length = 120 * scale
    spacing = 18 * scale
    for i in range(5):
        x = px + i * spacing
        line = slide.Shapes.AddLine(x, py, x + length * 0.6, py + length)
        line.Line.ForeColor.RGB = rgb(*color)
        line.Line.Weight = 1.5
        try:
            line.Line.Transparency = 0.5 + i * 0.08
        except Exception:
            pass


def _deco_dot_pattern(slide, px, py, color, scale):
    """Grid of small dots."""
    dot_r = 4 * scale
    spacing = 20 * scale
    rows, cols = 4, 5
    for r in range(rows):
        for c in range(cols):
            x = px + c * spacing
            y = py + r * spacing
            _add_shape(slide, MSO_SHAPE_OVAL, x, y, dot_r, dot_r,
                       fill_rgb=color, fill_transparency=0.5)


def _deco_corner_accent(slide, sw, sh, position, color, scale):
    """L-shaped corner accent."""
    length = 80 * scale
    weight = 5 * scale
    if position == "top_left":
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, 0, length, weight, fill_rgb=color)
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, 0, weight, length, fill_rgb=color)
    elif position == "top_right":
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - length, 0, length, weight,
                   fill_rgb=color)
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - weight, 0, weight, length,
                   fill_rgb=color)
    elif position == "bottom_left":
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, sh - weight, length, weight,
                   fill_rgb=color)
        _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, sh - length, weight, length,
                   fill_rgb=color)
    elif position == "bottom_right":
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - length, sh - weight,
                   length, weight, fill_rgb=color)
        _add_shape(slide, MSO_SHAPE_RECTANGLE, sw - weight, sh - length,
                   weight, length, fill_rgb=color)


def _deco_gradient_bar(slide, sw, sh, position, color, scale):
    """Gradient accent bar."""
    bar_h = 8 * scale
    lighter = _lerp_color(color, (255, 255, 255), 0.5)
    if position in ("top_left", "top_right"):
        _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE,
                            0, 0, sw, bar_h,
                            color, lighter, MSO_GRADIENT_HORIZONTAL)
    elif position in ("bottom_left", "bottom_right"):
        _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE,
                            0, sh - bar_h, sw, bar_h,
                            color, lighter, MSO_GRADIENT_HORIZONTAL)
    else:
        _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE,
                            0, 0, sw, bar_h,
                            color, lighter, MSO_GRADIENT_HORIZONTAL)


def _deco_wave_shape(slide, sw, sh, position, color, scale):
    """Wave-like decorative oval shapes."""
    w = sw * 1.2 * scale
    h = 200 * scale
    if position in ("top_left", "top_right"):
        _add_shape(slide, MSO_SHAPE_OVAL, -sw * 0.1, -h * 0.6, w, h,
                   fill_rgb=color, fill_transparency=0.85)
    elif position in ("bottom_left", "bottom_right"):
        _add_shape(slide, MSO_SHAPE_OVAL, -sw * 0.1, sh - h * 0.4, w, h,
                   fill_rgb=color, fill_transparency=0.85)
    else:
        _add_shape(slide, MSO_SHAPE_OVAL, -sw * 0.1, sh * 0.35, w, h,
                   fill_rgb=color, fill_transparency=0.9)


# ============================================================
# 9. create_agenda_slide
# ============================================================

def create_agenda_slide(
    slide_number: int,
    title: str,
    items: list[str],
    highlight_index: int | None = None,
    style: str = "numbered",
) -> dict:
    """Create an agenda/table of contents slide.

    Args:
        slide_number: 1-based slide index.
        title: Agenda title.
        items: List of agenda items.
        highlight_index: 0-based index of item to highlight (optional).
        style: Design style - "numbered", "cards", "steps".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "numbered":
        _set_solid_bg(slide, (248, 249, 252))

        # Title
        _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=32,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

        # Title underline
        line = slide.Shapes.AddLine(sw * 0.06, sh * 0.17, sw * 0.18, sh * 0.17)
        line.Line.ForeColor.RGB = rgb(0, 120, 190)
        line.Line.Weight = 3

        item_h = min(55, (sh * 0.72) / max(len(items), 1))
        for i, item in enumerate(items):
            y = sh * 0.22 + i * item_h
            is_highlighted = (highlight_index is not None and i == highlight_index)

            # Number circle
            circle_color = (0, 120, 190) if is_highlighted else (180, 190, 200)
            _add_shape(slide, MSO_SHAPE_OVAL,
                       sw * 0.08, y + 5, 32, 32,
                       fill_rgb=circle_color)
            _add_textbox(slide, sw * 0.08, y + 5, 32, 32, str(i + 1),
                         font_name="Segoe UI", font_size=14,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER, vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Item text
            text_color = (33, 37, 41) if is_highlighted else (100, 105, 115)
            font_weight = is_highlighted
            _add_textbox(slide, sw * 0.15, y + 5, sw * 0.75, 32, item,
                         font_name="Segoe UI", font_size=18,
                         font_color=text_color, bold=font_weight,
                         alignment=PP_ALIGN_LEFT, vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Highlight bar
            if is_highlighted:
                _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                           sw * 0.06, y, sw * 0.88, item_h - 5,
                           fill_rgb=(0, 120, 190), fill_transparency=0.92,
                           z_order_back=True)

    elif style == "cards":
        _set_solid_bg(slide, (245, 247, 250))

        _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=32,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        n = len(items)
        if n == 0:
            return {"slide_number": slide_number, "style": style}

        margin = sw * 0.06
        gap = 16
        card_w = (sw - 2 * margin - gap * (n - 1)) / n
        card_h = sh * 0.55
        y = sh * 0.25

        accent_colors = [
            (0, 120, 190), (0, 166, 214), (242, 169, 0),
            (0, 150, 136), (183, 28, 28), (106, 27, 154),
        ]

        for i, item in enumerate(items):
            x = margin + i * (card_w + gap)
            is_hl = (highlight_index is not None and i == highlight_index)
            c = accent_colors[i % len(accent_colors)]

            if is_hl:
                card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                  x, y - 8, card_w, card_h + 16,
                                  fill_rgb=c)
                _add_shadow(card, blur=10, offset_x=3, offset_y=3, transparency=0.5)
                text_color = (255, 255, 255)
            else:
                card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                  x, y, card_w, card_h,
                                  fill_rgb=(255, 255, 255))
                _add_shadow(card, blur=5, offset_x=2, offset_y=2, transparency=0.75)
                text_color = (50, 55, 65)

            # Number
            _add_textbox(slide, x + 10, y + (0 if is_hl else 15), card_w - 20, 50,
                         str(i + 1),
                         font_name="Segoe UI Light", font_size=36,
                         font_color=text_color if is_hl else c,
                         bold=False, alignment=PP_ALIGN_CENTER)

            # Item text
            _add_textbox(slide, x + 12, y + 70, card_w - 24, card_h - 90, item,
                         font_name="Segoe UI", font_size=13,
                         font_color=text_color, alignment=PP_ALIGN_CENTER)

    elif style == "steps":
        _set_solid_bg(slide, (250, 250, 252))

        _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=32,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

        n = len(items)
        if n == 0:
            return {"slide_number": slide_number, "style": style}

        step_w = (sw * 0.84) / n
        y_center = sh * 0.45
        circle_r = 24

        for i, item in enumerate(items):
            cx = sw * 0.08 + i * step_w + step_w / 2
            is_hl = (highlight_index is not None and i == highlight_index)
            c = (0, 120, 190) if is_hl else (190, 195, 205)

            # Connector line (except last)
            if i < n - 1:
                next_cx = sw * 0.08 + (i + 1) * step_w + step_w / 2
                conn = slide.Shapes.AddLine(
                    cx + circle_r, y_center,
                    next_cx - circle_r, y_center)
                conn.Line.ForeColor.RGB = rgb(210, 210, 220)
                conn.Line.Weight = 2

            # Step circle
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx - circle_r, y_center - circle_r,
                       circle_r * 2, circle_r * 2, fill_rgb=c)

            # Step number
            _add_textbox(slide, cx - circle_r, y_center - circle_r,
                         circle_r * 2, circle_r * 2, str(i + 1),
                         font_name="Segoe UI", font_size=16,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Label below
            _add_textbox(slide, cx - step_w * 0.4, y_center + circle_r + 15,
                         step_w * 0.8, 60, item,
                         font_name="Segoe UI", font_size=12,
                         font_color=(60, 60, 70) if is_hl else (130, 130, 140),
                         bold=is_hl, alignment=PP_ALIGN_CENTER)
    else:
        raise ValueError(f"Unknown agenda style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 10. create_comparison_slide
# ============================================================

def create_comparison_slide(
    slide_number: int,
    title: str,
    left_title: str,
    left_items: list[str],
    right_title: str,
    right_items: list[str],
    style: str = "versus",
) -> dict:
    """Create a comparison slide for two things.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        left_title: Left column title.
        left_items: Left column items.
        right_title: Right column title.
        right_items: Right column items.
        style: Design style - "versus", "columns", "pros_cons".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "versus":
        _set_solid_bg(slide, (248, 249, 252))

        # Title
        _add_textbox(slide, sw * 0.06, sh * 0.05, sw * 0.88, sh * 0.08, title,
                     font_name="Segoe UI Semibold", font_size=28,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        # VS circle in center
        vs_r = 30
        _add_shape(slide, MSO_SHAPE_OVAL,
                   sw / 2 - vs_r, sh * 0.45 - vs_r, vs_r * 2, vs_r * 2,
                   fill_rgb=(50, 50, 60))
        _add_textbox(slide, sw / 2 - vs_r, sh * 0.45 - vs_r,
                     vs_r * 2, vs_r * 2, "VS",
                     font_name="Segoe UI", font_size=16,
                     font_color=(255, 255, 255), bold=True,
                     alignment=PP_ALIGN_CENTER, vertical_anchor=MSO_ANCHOR_MIDDLE)

        col_w = sw * 0.4
        # Left column
        _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                   sw * 0.04, sh * 0.18, col_w, sh * 0.75,
                   fill_rgb=(0, 82, 136), fill_transparency=0.04)
        _add_textbox(slide, sw * 0.06, sh * 0.2, col_w - 20, sh * 0.06,
                     left_title,
                     font_name="Segoe UI Semibold", font_size=20,
                     font_color=(0, 82, 136), alignment=PP_ALIGN_CENTER)

        for i, item in enumerate(left_items):
            y = sh * 0.3 + i * 48
            _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.08, y + 15, 4, 16,
                       fill_rgb=(0, 120, 190))
            _add_textbox(slide, sw * 0.1, y + 8, col_w - 60, 30, item,
                         font_name="Segoe UI", font_size=13,
                         font_color=(50, 50, 60), alignment=PP_ALIGN_LEFT)

        # Right column
        _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                   sw * 0.56, sh * 0.18, col_w, sh * 0.75,
                   fill_rgb=(0, 150, 136), fill_transparency=0.04)
        _add_textbox(slide, sw * 0.58, sh * 0.2, col_w - 20, sh * 0.06,
                     right_title,
                     font_name="Segoe UI Semibold", font_size=20,
                     font_color=(0, 130, 116), alignment=PP_ALIGN_CENTER)

        for i, item in enumerate(right_items):
            y = sh * 0.3 + i * 48
            _add_shape(slide, MSO_SHAPE_RECTANGLE, sw * 0.6, y + 15, 4, 16,
                       fill_rgb=(0, 150, 136))
            _add_textbox(slide, sw * 0.62, y + 8, col_w - 60, 30, item,
                         font_name="Segoe UI", font_size=13,
                         font_color=(50, 50, 60), alignment=PP_ALIGN_LEFT)

    elif style == "columns":
        _set_solid_bg(slide, (255, 255, 255))

        _add_textbox(slide, sw * 0.06, sh * 0.05, sw * 0.88, sh * 0.08, title,
                     font_name="Segoe UI Semibold", font_size=28,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        # Divider line
        div = slide.Shapes.AddLine(sw / 2, sh * 0.2, sw / 2, sh * 0.9)
        div.Line.ForeColor.RGB = rgb(220, 220, 225)
        div.Line.Weight = 1.5

        col_w = sw * 0.42

        # Left
        _add_textbox(slide, sw * 0.06, sh * 0.16, col_w, sh * 0.06, left_title,
                     font_name="Segoe UI Semibold", font_size=22,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)
        line_l = slide.Shapes.AddLine(sw * 0.06, sh * 0.24, sw * 0.2, sh * 0.24)
        line_l.Line.ForeColor.RGB = rgb(0, 120, 190)
        line_l.Line.Weight = 2.5

        for i, item in enumerate(left_items):
            y = sh * 0.28 + i * 45
            _add_textbox(slide, sw * 0.08, y, col_w - 20, 35, item,
                         font_name="Segoe UI", font_size=14,
                         font_color=(60, 60, 70), alignment=PP_ALIGN_LEFT)

        # Right
        _add_textbox(slide, sw * 0.54, sh * 0.16, col_w, sh * 0.06, right_title,
                     font_name="Segoe UI Semibold", font_size=22,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)
        line_r = slide.Shapes.AddLine(sw * 0.54, sh * 0.24, sw * 0.68, sh * 0.24)
        line_r.Line.ForeColor.RGB = rgb(0, 150, 136)
        line_r.Line.Weight = 2.5

        for i, item in enumerate(right_items):
            y = sh * 0.28 + i * 45
            _add_textbox(slide, sw * 0.56, y, col_w - 20, 35, item,
                         font_name="Segoe UI", font_size=14,
                         font_color=(60, 60, 70), alignment=PP_ALIGN_LEFT)

    elif style == "pros_cons":
        _set_solid_bg(slide, (250, 250, 252))

        _add_textbox(slide, sw * 0.06, sh * 0.05, sw * 0.88, sh * 0.08, title,
                     font_name="Segoe UI Semibold", font_size=28,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        col_w = sw * 0.42
        card_h = sh * 0.75

        # Left (green / pros)
        left_card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                               sw * 0.04, sh * 0.17, col_w, card_h,
                               fill_rgb=(255, 255, 255))
        _add_shadow(left_card, blur=5, offset_x=2, offset_y=2, transparency=0.8)
        _add_shape(slide, MSO_SHAPE_RECTANGLE,
                   sw * 0.04, sh * 0.17, col_w, 6, fill_rgb=(46, 125, 50))
        _add_textbox(slide, sw * 0.06, sh * 0.2, col_w - 20, sh * 0.06,
                     left_title,
                     font_name="Segoe UI Semibold", font_size=20,
                     font_color=(46, 125, 50), alignment=PP_ALIGN_LEFT)

        for i, item in enumerate(left_items):
            y = sh * 0.3 + i * 45
            _add_textbox(slide, sw * 0.08, y, 20, 25, "+",
                         font_name="Segoe UI", font_size=16,
                         font_color=(46, 125, 50), bold=True,
                         alignment=PP_ALIGN_CENTER)
            _add_textbox(slide, sw * 0.11, y, col_w - 60, 30, item,
                         font_name="Segoe UI", font_size=13,
                         font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)

        # Right (red / cons)
        right_card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                sw * 0.54, sh * 0.17, col_w, card_h,
                                fill_rgb=(255, 255, 255))
        _add_shadow(right_card, blur=5, offset_x=2, offset_y=2, transparency=0.8)
        _add_shape(slide, MSO_SHAPE_RECTANGLE,
                   sw * 0.54, sh * 0.17, col_w, 6, fill_rgb=(198, 40, 40))
        _add_textbox(slide, sw * 0.56, sh * 0.2, col_w - 20, sh * 0.06,
                     right_title,
                     font_name="Segoe UI Semibold", font_size=20,
                     font_color=(198, 40, 40), alignment=PP_ALIGN_LEFT)

        for i, item in enumerate(right_items):
            y = sh * 0.3 + i * 45
            _add_textbox(slide, sw * 0.58, y, 20, 25, "-",
                         font_name="Segoe UI", font_size=16,
                         font_color=(198, 40, 40), bold=True,
                         alignment=PP_ALIGN_CENTER)
            _add_textbox(slide, sw * 0.61, y, col_w - 60, 30, item,
                         font_name="Segoe UI", font_size=13,
                         font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)
    else:
        raise ValueError(f"Unknown comparison style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 11. create_process_flow
# ============================================================

def create_process_flow(
    slide_number: int,
    title: str,
    steps: list[str],
    style: str = "arrows",
) -> dict:
    """Create a process/flow diagram slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        steps: List of process steps.
        style: Design style - "arrows", "circles", "chevrons".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "arrows":
        _set_solid_bg(slide, (248, 249, 252))

        _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=32,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

        n = len(steps)
        if n == 0:
            return {"slide_number": slide_number, "style": style}

        margin = sw * 0.06
        gap = 12
        arrow_w = (sw - 2 * margin - gap * (n - 1)) / n
        arrow_h = 70
        y = sh * 0.4

        colors = [
            (0, 82, 136), (0, 120, 190), (0, 166, 214),
            (0, 150, 136), (242, 169, 0), (183, 28, 28),
        ]

        for i, step in enumerate(steps):
            x = margin + i * (arrow_w + gap)
            c = colors[i % len(colors)]

            # Arrow/pentagon shape
            shape = slide.Shapes.AddShape(
                MSO_SHAPE_CHEVRON if i < n - 1 else MSO_SHAPE_ROUNDED_RECTANGLE,
                x, y, arrow_w, arrow_h
            )
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = rgb(*c)
            shape.Line.Visible = False
            _add_shadow(shape, blur=4, offset_x=2, offset_y=2, transparency=0.7)

            # Step number above
            _add_textbox(slide, x, y - 35, arrow_w, 30, f"Step {i + 1}",
                         font_name="Segoe UI", font_size=11,
                         font_color=c, bold=True, alignment=PP_ALIGN_CENTER)

            # Step text inside
            _add_textbox(slide, x + 10, y + 10, arrow_w - 20, arrow_h - 20, step,
                         font_name="Segoe UI", font_size=12,
                         font_color=(255, 255, 255), bold=False,
                         alignment=PP_ALIGN_CENTER, vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Description area below
            _add_textbox(slide, x, y + arrow_h + 15, arrow_w, 80, "",
                         font_name="Segoe UI", font_size=11,
                         font_color=(100, 100, 110), alignment=PP_ALIGN_CENTER)

    elif style == "circles":
        _set_solid_bg(slide, (250, 250, 252))

        _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=32,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

        n = len(steps)
        if n == 0:
            return {"slide_number": slide_number, "style": style}

        margin = sw * 0.1
        usable = sw - 2 * margin
        circle_r = min(40, usable / (n * 2.5))
        step_spacing = usable / max(n - 1, 1) if n > 1 else 0
        y_center = sh * 0.45

        for i, step in enumerate(steps):
            cx = margin + i * step_spacing if n > 1 else sw / 2
            c = _lerp_color((0, 82, 136), (0, 180, 200), i / max(n - 1, 1))

            # Connector line
            if i < n - 1:
                next_cx = margin + (i + 1) * step_spacing
                conn = slide.Shapes.AddLine(
                    cx + circle_r + 2, y_center,
                    next_cx - circle_r - 2, y_center
                )
                conn.Line.ForeColor.RGB = rgb(200, 205, 215)
                conn.Line.Weight = 2.5
                # Arrow head
                try:
                    conn.Line.EndArrowheadStyle = 2  # msoArrowheadTriangle
                except Exception:
                    pass

            # Circle
            circle = _add_shape(slide, MSO_SHAPE_OVAL,
                                cx - circle_r, y_center - circle_r,
                                circle_r * 2, circle_r * 2,
                                fill_rgb=c)
            _add_shadow(circle, blur=5, offset_x=2, offset_y=2, transparency=0.7)

            # Number
            _add_textbox(slide, cx - circle_r, y_center - circle_r,
                         circle_r * 2, circle_r * 2, str(i + 1),
                         font_name="Segoe UI", font_size=20,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Label below
            label_w = step_spacing * 0.8 if n > 1 else 200
            _add_textbox(slide, cx - label_w / 2, y_center + circle_r + 15,
                         label_w, 70, step,
                         font_name="Segoe UI", font_size=12,
                         font_color=(60, 60, 70), alignment=PP_ALIGN_CENTER)

    elif style == "chevrons":
        _set_solid_bg(slide, (245, 247, 250))

        _add_textbox(slide, sw * 0.06, sh * 0.06, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=32,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

        n = len(steps)
        if n == 0:
            return {"slide_number": slide_number, "style": style}

        margin = sw * 0.06
        total_w = sw - 2 * margin
        chev_w = total_w / n
        chev_h = 80
        y = sh * 0.38

        for i, step in enumerate(steps):
            x = margin + i * chev_w
            # Color gradient across steps
            t = i / max(n - 1, 1)
            c = _lerp_color((0, 82, 136), (0, 166, 214), t)

            shape = slide.Shapes.AddShape(MSO_SHAPE_CHEVRON,
                                          x, y, chev_w, chev_h)
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = rgb(*c)
            shape.Line.Visible = False

            # Step text below
            _add_textbox(slide, x, y + chev_h + 20, chev_w, 70, step,
                         font_name="Segoe UI", font_size=12,
                         font_color=(50, 55, 65), alignment=PP_ALIGN_CENTER)

            # Step number inside chevron
            _add_textbox(slide, x + chev_w * 0.15, y + 15,
                         chev_w * 0.7, chev_h - 30,
                         str(i + 1),
                         font_name="Segoe UI Light", font_size=26,
                         font_color=(255, 255, 255), alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)
    else:
        raise ValueError(f"Unknown process flow style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 12. add_progress_bar
# ============================================================

def add_progress_bar(
    slide_number: int,
    progress: float,
    position: str = "bottom",
    color: tuple[int, int, int] | None = None,
    height: float = 8,
) -> dict:
    """Add a progress bar to a slide.

    Args:
        slide_number: 1-based slide index.
        progress: Progress value from 0.0 to 1.0.
        position: "top" or "bottom".
        color: RGB color tuple. Defaults to blue.
        height: Bar height in points.

    Returns:
        Dict with slide_number and progress percentage.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    c = color or (0, 120, 190)
    progress = max(0.0, min(1.0, progress))
    y = 0 if position == "top" else sh - height

    # Background track
    _add_shape(slide, MSO_SHAPE_RECTANGLE, 0, y, sw, height,
               fill_rgb=(230, 230, 235))

    # Progress fill
    if progress > 0:
        fill_w = sw * progress
        lighter = _lerp_color(c, (255, 255, 255), 0.3)
        _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE,
                            0, y, fill_w, height,
                            c, lighter, MSO_GRADIENT_HORIZONTAL)

    return {"slide_number": slide_number, "progress": f"{progress * 100:.0f}%"}


# ============================================================
# 13. create_quote_slide
# ============================================================

def create_quote_slide(
    slide_number: int,
    quote: str,
    author: str | None = None,
    style: str = "large_quote",
) -> dict:
    """Create a quote slide.

    Args:
        slide_number: 1-based slide index.
        quote: The quote text.
        author: Optional attribution.
        style: Design style - "large_quote", "minimal", "highlighted".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "large_quote":
        _set_solid_bg(slide, (248, 249, 252))

        # Large opening quotation mark
        _add_textbox(slide, sw * 0.06, sh * 0.12, 120, 120,
                     "\u201C",
                     font_name="Georgia", font_size=120,
                     font_color=(0, 120, 190), bold=False,
                     alignment=PP_ALIGN_LEFT)

        # Quote text
        _add_textbox(slide, sw * 0.12, sh * 0.3, sw * 0.76, sh * 0.35, quote,
                     font_name="Georgia", font_size=24,
                     font_color=(50, 55, 65), italic=True,
                     alignment=PP_ALIGN_LEFT)

        # Closing quotation mark
        _add_textbox(slide, sw * 0.8, sh * 0.55, 80, 80,
                     "\u201D",
                     font_name="Georgia", font_size=80,
                     font_color=(0, 120, 190),
                     alignment=PP_ALIGN_RIGHT)

        # Author
        if author:
            # Thin line before author
            line = slide.Shapes.AddLine(sw * 0.12, sh * 0.72, sw * 0.22, sh * 0.72)
            line.Line.ForeColor.RGB = rgb(0, 120, 190)
            line.Line.Weight = 2

            _add_textbox(slide, sw * 0.12, sh * 0.75, sw * 0.6, sh * 0.06,
                         f"- {author}",
                         font_name="Segoe UI", font_size=16,
                         font_color=(100, 105, 115), alignment=PP_ALIGN_LEFT)

    elif style == "minimal":
        _set_solid_bg(slide, (255, 255, 255))

        # Left accent line
        _add_shape(slide, MSO_SHAPE_RECTANGLE,
                   sw * 0.1, sh * 0.3, 4, sh * 0.3,
                   fill_rgb=(0, 120, 190))

        # Quote
        _add_textbox(slide, sw * 0.14, sh * 0.3, sw * 0.72, sh * 0.3, quote,
                     font_name="Segoe UI Light", font_size=26,
                     font_color=(50, 50, 55), italic=True,
                     alignment=PP_ALIGN_LEFT)

        if author:
            _add_textbox(slide, sw * 0.14, sh * 0.65, sw * 0.6, sh * 0.06,
                         author,
                         font_name="Segoe UI", font_size=14,
                         font_color=(130, 130, 140), alignment=PP_ALIGN_LEFT)

    elif style == "highlighted":
        _set_solid_bg(slide, (250, 250, 252))

        # Colored background strip
        _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE,
                            0, sh * 0.25, sw, sh * 0.45,
                            (0, 82, 136), (0, 120, 180),
                            MSO_GRADIENT_HORIZONTAL)

        # Quote on the colored strip
        _add_textbox(slide, sw * 0.1, sh * 0.3, sw * 0.8, sh * 0.3,
                     f"\u201C{quote}\u201D",
                     font_name="Georgia", font_size=24,
                     font_color=(255, 255, 255), italic=True,
                     alignment=PP_ALIGN_CENTER)

        if author:
            _add_textbox(slide, sw * 0.2, sh * 0.62, sw * 0.6, sh * 0.06,
                         f"- {author}",
                         font_name="Segoe UI", font_size=15,
                         font_color=(200, 220, 240), alignment=PP_ALIGN_CENTER)
    else:
        raise ValueError(f"Unknown quote style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 14. create_team_slide
# ============================================================

def create_team_slide(
    slide_number: int,
    title: str,
    members: list[dict],
    style: str = "grid",
) -> dict:
    """Create a team/people slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        members: List of dicts with "name", "role", optional "description".
        style: Design style - "grid", "horizontal".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "grid":
        _set_solid_bg(slide, (248, 249, 252))

        _add_textbox(slide, sw * 0.06, sh * 0.05, sw * 0.88, sh * 0.08, title,
                     font_name="Segoe UI Semibold", font_size=30,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        n = len(members)
        if n == 0:
            return {"slide_number": slide_number, "style": style}

        cols = min(n, 4)
        rows = (n + cols - 1) // cols
        margin = sw * 0.06
        gap = 20
        card_w = (sw - 2 * margin - gap * (cols - 1)) / cols
        card_h = min((sh * 0.78 - gap * (rows - 1)) / rows, 250)
        start_y = sh * 0.16

        accent_colors = [
            (0, 120, 190), (0, 150, 136), (242, 169, 0),
            (183, 28, 28), (106, 27, 154), (0, 166, 214),
        ]

        for i, member in enumerate(members):
            col = i % cols
            row = i // cols
            x = margin + col * (card_w + gap)
            y = start_y + row * (card_h + gap)
            c = accent_colors[i % len(accent_colors)]

            # Card
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              x, y, card_w, card_h,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=6, offset_x=2, offset_y=2, transparency=0.75)

            # Avatar circle placeholder
            avatar_r = 30
            avatar_cx = x + card_w / 2
            avatar_cy = y + 25 + avatar_r
            _add_shape(slide, MSO_SHAPE_OVAL,
                       avatar_cx - avatar_r, avatar_cy - avatar_r,
                       avatar_r * 2, avatar_r * 2,
                       fill_rgb=c)
            # Initials in circle
            initials = "".join(w[0].upper() for w in member.get("name", "?").split()[:2])
            _add_textbox(slide, avatar_cx - avatar_r, avatar_cy - avatar_r,
                         avatar_r * 2, avatar_r * 2, initials,
                         font_name="Segoe UI", font_size=18,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Name
            name_y = avatar_cy + avatar_r + 12
            _add_textbox(slide, x + 8, name_y, card_w - 16, 28,
                         member.get("name", ""),
                         font_name="Segoe UI Semibold", font_size=14,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

            # Role
            _add_textbox(slide, x + 8, name_y + 26, card_w - 16, 22,
                         member.get("role", ""),
                         font_name="Segoe UI", font_size=11,
                         font_color=c, alignment=PP_ALIGN_CENTER)

            # Description
            desc = member.get("description", "")
            if desc:
                _add_textbox(slide, x + 12, name_y + 52, card_w - 24, 50, desc,
                             font_name="Segoe UI", font_size=10,
                             font_color=(110, 110, 120),
                             alignment=PP_ALIGN_CENTER)

    elif style == "horizontal":
        _set_solid_bg(slide, (250, 250, 252))

        _add_textbox(slide, sw * 0.06, sh * 0.05, sw * 0.88, sh * 0.08, title,
                     font_name="Segoe UI Semibold", font_size=30,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        n = len(members)
        if n == 0:
            return {"slide_number": slide_number, "style": style}

        margin = sw * 0.06
        gap = 24
        card_w = (sw - 2 * margin - gap * (n - 1)) / n
        card_h = sh * 0.7
        y = sh * 0.18

        accent_colors = [
            (0, 120, 190), (0, 150, 136), (242, 169, 0),
            (183, 28, 28), (106, 27, 154), (0, 166, 214),
        ]

        for i, member in enumerate(members):
            x = margin + i * (card_w + gap)
            c = accent_colors[i % len(accent_colors)]

            # Top color bar
            _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE,
                                x, y, card_w, card_h * 0.35,
                                c, _lerp_color(c, (255, 255, 255), 0.3),
                                MSO_GRADIENT_VERTICAL)

            # White card below
            card = _add_shape(slide, MSO_SHAPE_RECTANGLE,
                              x, y + card_h * 0.35, card_w, card_h * 0.65,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=5, offset_x=2, offset_y=2, transparency=0.8)

            # Avatar circle overlapping the bar/card boundary
            avatar_r = 35
            avatar_cx = x + card_w / 2
            avatar_cy = y + card_h * 0.35
            _add_shape(slide, MSO_SHAPE_OVAL,
                       avatar_cx - avatar_r, avatar_cy - avatar_r,
                       avatar_r * 2, avatar_r * 2,
                       fill_rgb=(255, 255, 255))
            inner_r = avatar_r - 4
            _add_shape(slide, MSO_SHAPE_OVAL,
                       avatar_cx - inner_r, avatar_cy - inner_r,
                       inner_r * 2, inner_r * 2,
                       fill_rgb=c)
            initials = "".join(w[0].upper() for w in member.get("name", "?").split()[:2])
            _add_textbox(slide, avatar_cx - inner_r, avatar_cy - inner_r,
                         inner_r * 2, inner_r * 2, initials,
                         font_name="Segoe UI", font_size=20,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Name
            info_y = avatar_cy + avatar_r + 15
            _add_textbox(slide, x + 8, info_y, card_w - 16, 28,
                         member.get("name", ""),
                         font_name="Segoe UI Semibold", font_size=16,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

            # Role
            _add_textbox(slide, x + 8, info_y + 30, card_w - 16, 22,
                         member.get("role", ""),
                         font_name="Segoe UI", font_size=12,
                         font_color=c, alignment=PP_ALIGN_CENTER)

            # Description
            desc = member.get("description", "")
            if desc:
                _add_textbox(slide, x + 12, info_y + 58, card_w - 24, 70, desc,
                             font_name="Segoe UI", font_size=10,
                             font_color=(100, 105, 115),
                             alignment=PP_ALIGN_CENTER)
    else:
        raise ValueError(f"Unknown team slide style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 15. add_slide_number_footer
# ============================================================

def add_slide_number_footer(
    start_slide: int = 1,
    end_slide: int | None = None,
    format_type: str = "number",
    position: str = "bottom_right",
    font_size: float = 10,
    font_color: tuple[int, int, int] | None = None,
) -> dict:
    """Add formatted slide numbers to slides.

    Args:
        start_slide: First slide (1-based).
        end_slide: Last slide (1-based), or None for last slide.
        format_type: "number" (just number), "of_total" (1/10),
            "dash_total" (1 - 10).
        position: "bottom_left", "bottom_center", "bottom_right".
        font_size: Font size in points.
        font_color: RGB color tuple.

    Returns:
        Dict with number of slides affected.
    """
    app = _get_app()
    prs = app.ActivePresentation
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight
    total = prs.Slides.Count

    end = end_slide or total
    fc = font_color or (140, 140, 150)

    # Position coordinates
    box_w = 80
    box_h = 24
    margin = 20
    if position == "bottom_left":
        bx = margin
        alignment = PP_ALIGN_LEFT
    elif position == "bottom_center":
        bx = (sw - box_w) / 2
        alignment = PP_ALIGN_CENTER
    else:  # bottom_right
        bx = sw - box_w - margin
        alignment = PP_ALIGN_RIGHT

    by = sh - box_h - margin

    count = 0
    for i in range(start_slide, end + 1):
        if i > total:
            break
        slide = prs.Slides(i)

        # Format the number
        if format_type == "of_total":
            text = f"{i}/{total}"
        elif format_type == "dash_total":
            text = f"{i} - {total}"
        else:
            text = str(i)

        _add_textbox(slide, bx, by, box_w, box_h, text,
                     font_name="Segoe UI", font_size=font_size,
                     font_color=fc, alignment=alignment)
        count += 1

    return {"slides_affected": count, "format": format_type}
