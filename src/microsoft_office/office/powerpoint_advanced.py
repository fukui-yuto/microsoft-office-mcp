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


# ============================================================
# 16. create_chart_slide
# ============================================================

def create_chart_slide(
    slide_number: int,
    title: str,
    chart_type: str,
    data: dict,
    style: str = "modern",
) -> dict:
    """Create a slide with a professionally styled chart.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        chart_type: "bar", "line", "pie", "doughnut".
        data: {"categories": [...], "series": [{"name": "...", "values": [...]}]}.
        style: "modern" (gradient fills), "minimal" (thin lines), "bold" (vivid colors).

    Returns:
        Dict with slide_number and style.
    """
    import math

    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    categories = data.get("categories", [])
    series_list = data.get("series", [])

    # Style palettes
    if style == "modern":
        _set_solid_bg(slide, (248, 250, 252))
        title_color = (33, 37, 41)
        palette = [(0, 120, 215), (0, 180, 160), (255, 140, 0), (200, 50, 80), (100, 80, 200)]
        bg_card_color = (255, 255, 255)
        axis_color = (180, 185, 195)
        label_color = (100, 105, 115)
    elif style == "minimal":
        _set_solid_bg(slide, (255, 255, 255))
        title_color = (60, 60, 65)
        palette = [(100, 100, 110), (160, 160, 170), (60, 60, 70), (200, 200, 210), (130, 130, 140)]
        bg_card_color = None
        axis_color = (220, 220, 225)
        label_color = (130, 130, 140)
    elif style == "bold":
        _set_solid_bg(slide, (245, 245, 248))
        title_color = (20, 20, 25)
        palette = [(0, 100, 210), (230, 50, 50), (20, 180, 80), (255, 170, 0), (150, 50, 200)]
        bg_card_color = (255, 255, 255)
        axis_color = (190, 190, 200)
        label_color = (80, 80, 90)
    else:
        raise ValueError(f"Unknown chart style '{style}'")

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=28,
                 font_color=title_color, alignment=PP_ALIGN_LEFT)

    # Title underline
    line = slide.Shapes.AddLine(sw * 0.06, sh * 0.14, sw * 0.18, sh * 0.14)
    line.Line.ForeColor.RGB = rgb(*palette[0])
    line.Line.Weight = 3

    # Chart area
    chart_left = sw * 0.08
    chart_top = sh * 0.2
    chart_width = sw * 0.84
    chart_height = sh * 0.7

    # Background card for chart area
    if bg_card_color:
        card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                          chart_left - 10, chart_top - 10,
                          chart_width + 20, chart_height + 20,
                          fill_rgb=bg_card_color)
        _add_shadow(card, blur=8, offset_x=2, offset_y=2, transparency=0.8)

    if chart_type in ("bar", "line"):
        n_cats = len(categories)
        n_series = len(series_list)
        if n_cats == 0 or n_series == 0:
            return {"slide_number": slide_number, "style": style}

        # Find max value for scaling
        all_vals = [v for s in series_list for v in s.get("values", [])]
        max_val = max(all_vals) if all_vals else 1

        # Draw axes
        axis_left = chart_left + 50
        axis_bottom = chart_top + chart_height - 40
        axis_right = chart_left + chart_width - 20
        axis_top = chart_top + 20
        plot_h = axis_bottom - axis_top
        plot_w = axis_right - axis_left

        # Y-axis
        y_axis = slide.Shapes.AddLine(axis_left, axis_top, axis_left, axis_bottom)
        y_axis.Line.ForeColor.RGB = rgb(*axis_color)
        y_axis.Line.Weight = 1.5

        # X-axis
        x_axis = slide.Shapes.AddLine(axis_left, axis_bottom, axis_right, axis_bottom)
        x_axis.Line.ForeColor.RGB = rgb(*axis_color)
        x_axis.Line.Weight = 1.5

        # Grid lines
        for g in range(1, 5):
            gy = axis_bottom - (plot_h * g / 4)
            grid = slide.Shapes.AddLine(axis_left, gy, axis_right, gy)
            grid.Line.ForeColor.RGB = rgb(235, 237, 240)
            grid.Line.Weight = 0.75
            # Y-axis labels
            val_label = f"{max_val * g / 4:.0f}"
            _add_textbox(slide, chart_left, gy - 10, 48, 20, val_label,
                         font_name="Segoe UI", font_size=9,
                         font_color=label_color, alignment=PP_ALIGN_RIGHT)

        cat_width = plot_w / n_cats

        if chart_type == "bar":
            bar_gap = 4
            total_bar_width = (cat_width - 20) / n_series
            for si, series in enumerate(series_list):
                c = palette[si % len(palette)]
                values = series.get("values", [])
                for ci, val in enumerate(values):
                    if ci >= n_cats:
                        break
                    bar_h = (val / max_val) * plot_h if max_val > 0 else 0
                    bx = axis_left + ci * cat_width + 10 + si * total_bar_width + bar_gap / 2
                    by = axis_bottom - bar_h
                    bw = total_bar_width - bar_gap
                    if style == "modern":
                        bar = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                         bx, by, bw, bar_h, fill_rgb=c)
                        _add_shadow(bar, blur=3, offset_x=1, offset_y=1, transparency=0.8)
                    elif style == "bold":
                        bar = _add_shape(slide, MSO_SHAPE_RECTANGLE,
                                         bx, by, bw, bar_h, fill_rgb=c)
                    else:  # minimal
                        lighter = _lerp_color(c, (255, 255, 255), 0.4)
                        bar = _add_shape(slide, MSO_SHAPE_RECTANGLE,
                                         bx, by, bw, bar_h, fill_rgb=lighter)
                        bar.Line.Visible = True
                        bar.Line.ForeColor.RGB = rgb(*c)
                        bar.Line.Weight = 1
        else:  # line
            for si, series in enumerate(series_list):
                c = palette[si % len(palette)]
                values = series.get("values", [])
                points = []
                for ci, val in enumerate(values):
                    if ci >= n_cats:
                        break
                    px = axis_left + ci * cat_width + cat_width / 2
                    py = axis_bottom - (val / max_val) * plot_h if max_val > 0 else axis_bottom
                    points.append((px, py))
                # Draw line segments
                for pi in range(len(points) - 1):
                    seg = slide.Shapes.AddLine(
                        points[pi][0], points[pi][1],
                        points[pi + 1][0], points[pi + 1][1])
                    seg.Line.ForeColor.RGB = rgb(*c)
                    seg.Line.Weight = 3 if style == "bold" else 2
                # Draw data points
                for px, py in points:
                    dot_r = 5 if style == "bold" else 4
                    _add_shape(slide, MSO_SHAPE_OVAL,
                               px - dot_r, py - dot_r, dot_r * 2, dot_r * 2,
                               fill_rgb=c)

        # Category labels along X-axis
        for ci, cat in enumerate(categories):
            cx = axis_left + ci * cat_width + cat_width / 2
            _add_textbox(slide, cx - cat_width / 2, axis_bottom + 5,
                         cat_width, 25, cat,
                         font_name="Segoe UI", font_size=9,
                         font_color=label_color, alignment=PP_ALIGN_CENTER)

        # Legend
        legend_x = axis_left
        legend_y = chart_top + chart_height - 20
        for si, series in enumerate(series_list):
            c = palette[si % len(palette)]
            lx = legend_x + si * 120
            _add_shape(slide, MSO_SHAPE_RECTANGLE, lx, legend_y, 12, 12, fill_rgb=c)
            _add_textbox(slide, lx + 16, legend_y - 2, 100, 16,
                         series.get("name", ""),
                         font_name="Segoe UI", font_size=9,
                         font_color=label_color, alignment=PP_ALIGN_LEFT)

    elif chart_type in ("pie", "doughnut"):
        # Draw pie/doughnut using wedge-like segments approximated with shapes
        if not series_list or not categories:
            return {"slide_number": slide_number, "style": style}

        values = series_list[0].get("values", [])
        total = sum(values) if values else 1
        n = len(values)

        cx = chart_left + chart_width * 0.45
        cy = chart_top + chart_height * 0.45
        radius = min(chart_width, chart_height) * 0.35

        # Draw colored segments as oval shapes positioned around center
        # Since COM doesn't have native pie drawing, we create visual representation
        # using colored rectangles in a legend + a large circle with segment indicators

        # Main circle
        _add_shape(slide, MSO_SHAPE_OVAL,
                   cx - radius, cy - radius, radius * 2, radius * 2,
                   fill_rgb=palette[0])

        # For doughnut, add inner white circle
        if chart_type == "doughnut":
            inner_r = radius * 0.55
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx - inner_r, cy - inner_r, inner_r * 2, inner_r * 2,
                       fill_rgb=bg_card_color or (255, 255, 255))

        # Draw colored segment indicators as small shapes around the circle
        # and show the data via a clean legend
        cumulative = 0
        for i, val in enumerate(values):
            if i >= len(categories):
                break
            pct = val / total if total > 0 else 0
            angle = cumulative + pct / 2
            # Place a colored indicator at the segment position
            ix = cx + radius * 0.7 * math.cos(2 * math.pi * angle - math.pi / 2)
            iy = cy + radius * 0.7 * math.sin(2 * math.pi * angle - math.pi / 2)
            c = palette[i % len(palette)]

            # Segment label with percentage
            if i > 0:
                # Add colored oval overlay to show segments
                seg_size = max(radius * pct * 2, 20)
                sx = cx + radius * 0.65 * math.cos(2 * math.pi * angle - math.pi / 2)
                sy = cy + radius * 0.65 * math.sin(2 * math.pi * angle - math.pi / 2)
                _add_shape(slide, MSO_SHAPE_OVAL,
                           sx - seg_size / 2, sy - seg_size / 2,
                           seg_size, seg_size, fill_rgb=c)

            cumulative += pct

        # Legend on the right side
        legend_x = cx + radius + 30
        legend_y = cy - (n * 28) / 2
        for i, val in enumerate(values):
            if i >= len(categories):
                break
            c = palette[i % len(palette)]
            pct = val / total * 100 if total > 0 else 0
            ly = legend_y + i * 32

            _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                       legend_x, ly + 2, 14, 14, fill_rgb=c)
            _add_textbox(slide, legend_x + 20, ly, 160, 20,
                         f"{categories[i]}  ({pct:.0f}%)",
                         font_name="Segoe UI", font_size=11,
                         font_color=label_color, alignment=PP_ALIGN_LEFT)

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 17. create_dashboard_slide
# ============================================================

def create_dashboard_slide(
    slide_number: int,
    title: str,
    kpis: list[dict],
    chart_data: dict | None = None,
    style: str = "executive",
) -> dict:
    """Create a dashboard slide with KPI cards and optional chart.

    Args:
        slide_number: 1-based slide index.
        title: Dashboard title.
        kpis: List of {"label": "...", "value": "...", "change": "...", "trend": "up"/"down"}.
        chart_data: Optional chart data in the same format as create_chart_slide.
        style: "executive" (dark bg), "corporate" (blue/white), "startup" (vibrant).

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    if style == "executive":
        _set_gradient_bg(slide, (25, 28, 38), (35, 40, 55), MSO_GRADIENT_DIAGONAL_DOWN)
        title_color = (255, 255, 255)
        card_bg = (45, 50, 65)
        value_color = (255, 255, 255)
        label_color = (160, 165, 180)
        up_color = (0, 210, 140)
        down_color = (255, 75, 75)
        accent = (0, 150, 255)
    elif style == "corporate":
        _set_solid_bg(slide, (240, 244, 250))
        title_color = (30, 55, 100)
        card_bg = (255, 255, 255)
        value_color = (30, 55, 100)
        label_color = (100, 115, 140)
        up_color = (0, 160, 80)
        down_color = (210, 50, 50)
        accent = (0, 100, 200)
    elif style == "startup":
        _set_solid_bg(slide, (250, 248, 255))
        title_color = (50, 20, 80)
        card_bg = (255, 255, 255)
        value_color = (50, 20, 80)
        label_color = (120, 100, 150)
        up_color = (0, 200, 120)
        down_color = (255, 60, 100)
        accent = (120, 60, 220)
    else:
        raise ValueError(f"Unknown dashboard style '{style}'")

    # Title
    _add_textbox(slide, sw * 0.05, sh * 0.03, sw * 0.9, sh * 0.08, title,
                 font_name="Segoe UI Semibold", font_size=26,
                 font_color=title_color, alignment=PP_ALIGN_LEFT)

    # Accent line under title
    line = slide.Shapes.AddLine(sw * 0.05, sh * 0.11, sw * 0.15, sh * 0.11)
    line.Line.ForeColor.RGB = rgb(*accent)
    line.Line.Weight = 3

    # KPI cards
    n_kpis = len(kpis)
    if n_kpis > 0:
        kpi_margin = sw * 0.05
        kpi_gap = 14
        kpi_w = (sw - 2 * kpi_margin - kpi_gap * (n_kpis - 1)) / n_kpis
        kpi_h = sh * 0.2
        kpi_y = sh * 0.14

        for i, kpi in enumerate(kpis):
            kx = kpi_margin + i * (kpi_w + kpi_gap)

            # Card background
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              kx, kpi_y, kpi_w, kpi_h, fill_rgb=card_bg)
            if style != "executive":
                _add_shadow(card, blur=6, offset_x=2, offset_y=2, transparency=0.8)

            # Top accent stripe
            _add_shape(slide, MSO_SHAPE_RECTANGLE, kx, kpi_y, kpi_w, 4,
                       fill_rgb=accent)

            # Value (large)
            _add_textbox(slide, kx + 12, kpi_y + 14, kpi_w - 24, kpi_h * 0.45,
                         kpi.get("value", ""),
                         font_name="Segoe UI Light", font_size=28,
                         font_color=value_color, bold=False,
                         alignment=PP_ALIGN_LEFT)

            # Label
            _add_textbox(slide, kx + 12, kpi_y + kpi_h * 0.5, kpi_w * 0.6, 22,
                         kpi.get("label", ""),
                         font_name="Segoe UI", font_size=10,
                         font_color=label_color, alignment=PP_ALIGN_LEFT)

            # Change indicator
            change = kpi.get("change", "")
            trend = kpi.get("trend", "")
            if change:
                arrow = "\u25B2 " if trend == "up" else "\u25BC " if trend == "down" else ""
                change_color = up_color if trend == "up" else down_color if trend == "down" else label_color
                _add_textbox(slide, kx + kpi_w * 0.6, kpi_y + kpi_h * 0.5,
                             kpi_w * 0.35, 22,
                             f"{arrow}{change}",
                             font_name="Segoe UI Semibold", font_size=11,
                             font_color=change_color, alignment=PP_ALIGN_RIGHT)

    # Optional chart area below KPIs
    if chart_data:
        chart_top = sh * 0.38
        chart_height = sh * 0.58
        chart_left = sw * 0.05
        chart_w = sw * 0.9

        # Chart background card
        chart_card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                chart_left, chart_top, chart_w, chart_height,
                                fill_rgb=card_bg)
        if style != "executive":
            _add_shadow(chart_card, blur=6, offset_x=2, offset_y=2, transparency=0.8)

        # Draw a simple bar chart within the card
        categories = chart_data.get("categories", [])
        series_list = chart_data.get("series", [])
        if categories and series_list:
            n_cats = len(categories)
            values = series_list[0].get("values", [])
            max_val = max(values) if values else 1

            plot_left = chart_left + 60
            plot_bottom = chart_top + chart_height - 45
            plot_top_y = chart_top + 30
            plot_right = chart_left + chart_w - 30
            plot_h = plot_bottom - plot_top_y
            plot_w = plot_right - plot_left
            cat_w = plot_w / n_cats

            # Grid lines
            for g in range(1, 5):
                gy = plot_bottom - (plot_h * g / 4)
                grid = slide.Shapes.AddLine(plot_left, gy, plot_right, gy)
                grid.Line.ForeColor.RGB = rgb(*(70, 75, 85) if style == "executive" else (235, 237, 240))
                grid.Line.Weight = 0.5

            bar_colors = [accent, _lerp_color(accent, (255, 255, 255), 0.3)]
            for si, series in enumerate(series_list[:2]):
                sc = bar_colors[si % len(bar_colors)]
                svals = series.get("values", [])
                bw = (cat_w - 16) / min(len(series_list), 2)
                for ci, val in enumerate(svals):
                    if ci >= n_cats:
                        break
                    bar_h = (val / max_val) * plot_h if max_val > 0 else 0
                    bx = plot_left + ci * cat_w + 8 + si * bw
                    by = plot_bottom - bar_h
                    _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                               bx, by, bw - 4, bar_h, fill_rgb=sc)

            # Category labels
            for ci, cat in enumerate(categories):
                cx = plot_left + ci * cat_w + cat_w / 2
                _add_textbox(slide, cx - cat_w / 2, plot_bottom + 5, cat_w, 22, cat,
                             font_name="Segoe UI", font_size=9,
                             font_color=label_color, alignment=PP_ALIGN_CENTER)

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 18. create_two_column_slide
# ============================================================

def create_two_column_slide(
    slide_number: int,
    title: str,
    left_content: dict,
    right_content: dict,
    style: str = "balanced",
) -> dict:
    """Create a two-column layout slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        left_content: {"heading": "...", "items": ["..."]}.
        right_content: {"heading": "...", "items": ["..."]}.
        style: "balanced" (equal), "emphasis_left" (wider left), "emphasis_right".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    _set_solid_bg(slide, (248, 250, 252))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.05, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=30,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

    line = slide.Shapes.AddLine(sw * 0.06, sh * 0.16, sw * 0.18, sh * 0.16)
    line.Line.ForeColor.RGB = rgb(0, 120, 190)
    line.Line.Weight = 3

    margin = sw * 0.06
    gap = 24
    content_top = sh * 0.2
    content_h = sh * 0.74

    if style == "balanced":
        left_w = (sw - 2 * margin - gap) * 0.5
        right_w = left_w
    elif style == "emphasis_left":
        left_w = (sw - 2 * margin - gap) * 0.62
        right_w = (sw - 2 * margin - gap) * 0.38
    elif style == "emphasis_right":
        left_w = (sw - 2 * margin - gap) * 0.38
        right_w = (sw - 2 * margin - gap) * 0.62
    else:
        raise ValueError(f"Unknown two-column style '{style}'")

    left_x = margin
    right_x = margin + left_w + gap

    colors = [(0, 120, 190), (0, 150, 136)]

    for idx, (cx, cw, content, accent) in enumerate([
        (left_x, left_w, left_content, colors[0]),
        (right_x, right_w, right_content, colors[1]),
    ]):
        # Card background
        card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                          cx, content_top, cw, content_h,
                          fill_rgb=(255, 255, 255))
        _add_shadow(card, blur=6, offset_x=2, offset_y=2, transparency=0.8)

        # Top accent bar
        _add_shape(slide, MSO_SHAPE_RECTANGLE, cx, content_top, cw, 5,
                   fill_rgb=accent)

        # Heading
        heading = content.get("heading", "")
        _add_textbox(slide, cx + 18, content_top + 18, cw - 36, 30, heading,
                     font_name="Segoe UI Semibold", font_size=18,
                     font_color=accent, alignment=PP_ALIGN_LEFT)

        # Divider line under heading
        h_line = slide.Shapes.AddLine(cx + 18, content_top + 52, cx + cw * 0.4, content_top + 52)
        h_line.Line.ForeColor.RGB = rgb(*accent)
        h_line.Line.Weight = 1.5

        # Items
        items = content.get("items", [])
        for i, item in enumerate(items):
            iy = content_top + 65 + i * 38
            # Bullet dot
            _add_shape(slide, MSO_SHAPE_OVAL, cx + 20, iy + 8, 8, 8,
                       fill_rgb=accent)
            _add_textbox(slide, cx + 36, iy, cw - 56, 32, item,
                         font_name="Segoe UI", font_size=13,
                         font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 19. create_image_text_slide
# ============================================================

def create_image_text_slide(
    slide_number: int,
    title: str,
    text_items: list[str],
    image_path: str | None = None,
    image_position: str = "right",
    style: str = "modern",
) -> dict:
    """Create a slide with text on one side and image/placeholder on other.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        text_items: List of text bullet points.
        image_path: Optional path to image file.
        image_position: "left" or "right".
        style: "modern", "clean", "overlap".

    Returns:
        Dict with slide_number and style.
    """
    import os

    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    margin = sw * 0.06

    if style == "modern":
        _set_solid_bg(slide, (248, 250, 252))
        accent = (0, 120, 190)
        text_color = (50, 55, 65)
        heading_color = (33, 37, 41)
        placeholder_bg = (230, 235, 242)
    elif style == "clean":
        _set_solid_bg(slide, (255, 255, 255))
        accent = (80, 80, 90)
        text_color = (60, 60, 65)
        heading_color = (30, 30, 35)
        placeholder_bg = (240, 240, 242)
    elif style == "overlap":
        _set_solid_bg(slide, (245, 247, 250))
        accent = (0, 100, 180)
        text_color = (50, 55, 65)
        heading_color = (255, 255, 255)
        placeholder_bg = (220, 230, 242)
    else:
        raise ValueError(f"Unknown image-text style '{style}'")

    text_w = sw * 0.45
    img_w = sw * 0.42
    img_h = sh * 0.7

    if image_position == "right":
        text_x = margin
        img_x = sw - margin - img_w
    else:
        text_x = sw - margin - text_w
        img_x = margin

    img_y = sh * 0.18

    if style == "overlap":
        # Colored background block behind text area
        block_x = text_x - 10 if image_position == "right" else text_x - 10
        _add_gradient_shape(slide, MSO_SHAPE_RECTANGLE,
                            0 if image_position == "right" else sw * 0.48,
                            0, sw * 0.55, sh,
                            (0, 80, 150), (0, 120, 190), MSO_GRADIENT_VERTICAL)

    # Title
    title_color = heading_color
    _add_textbox(slide, text_x, sh * 0.06, text_w, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=26,
                 font_color=title_color, alignment=PP_ALIGN_LEFT)

    # Accent line
    line = slide.Shapes.AddLine(text_x, sh * 0.17, text_x + text_w * 0.3, sh * 0.17)
    line.Line.ForeColor.RGB = rgb(*(accent if style != "overlap" else (255, 200, 50)))
    line.Line.Weight = 3

    # Text items
    item_text_color = (255, 255, 255) if style == "overlap" else text_color
    for i, item in enumerate(text_items):
        iy = sh * 0.22 + i * 42
        dot_c = (255, 200, 50) if style == "overlap" else accent
        _add_shape(slide, MSO_SHAPE_OVAL, text_x + 2, iy + 8, 8, 8,
                   fill_rgb=dot_c)
        _add_textbox(slide, text_x + 18, iy, text_w - 24, 36, item,
                     font_name="Segoe UI", font_size=14,
                     font_color=item_text_color, alignment=PP_ALIGN_LEFT)

    # Image or placeholder
    if image_path and os.path.exists(image_path):
        abs_path = os.path.abspath(image_path)
        try:
            pic = slide.Shapes.AddPicture(
                abs_path, LinkToFile=False, SaveWithDocument=True,
                Left=img_x, Top=img_y, Width=img_w, Height=img_h)
            # Maintain aspect ratio - let PowerPoint handle it
            pic.LockAspectRatio = True
        except Exception:
            # Fallback to placeholder
            _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                       img_x, img_y, img_w, img_h,
                       fill_rgb=placeholder_bg)
            _add_textbox(slide, img_x + 20, img_y + img_h / 2 - 20,
                         img_w - 40, 40, "Image Placeholder",
                         font_name="Segoe UI", font_size=16,
                         font_color=(160, 165, 175), alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)
    else:
        # Placeholder rectangle
        ph = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                        img_x, img_y, img_w, img_h,
                        fill_rgb=placeholder_bg)
        if style in ("modern", "clean"):
            _add_shadow(ph, blur=6, offset_x=2, offset_y=2, transparency=0.8)
        # Placeholder icon/text
        _add_textbox(slide, img_x + 20, img_y + img_h / 2 - 20,
                     img_w - 40, 40, "Image Placeholder",
                     font_name="Segoe UI Light", font_size=16,
                     font_color=(160, 165, 175), alignment=PP_ALIGN_CENTER,
                     vertical_anchor=MSO_ANCHOR_MIDDLE)

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 20. create_three_column_slide
# ============================================================

def create_three_column_slide(
    slide_number: int,
    title: str,
    columns: list[dict],
    style: str = "cards",
) -> dict:
    """Create a three-column layout slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        columns: List of 3 dicts with "heading" and "items".
        style: "cards" (card boxes), "clean" (divider lines), "icons" (with icon circles).

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    accent_colors = [(0, 120, 190), (0, 166, 140), (242, 150, 0)]

    # Title
    _set_solid_bg(slide, (248, 250, 252) if style != "clean" else (255, 255, 255))

    _add_textbox(slide, sw * 0.06, sh * 0.05, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=30,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

    margin = sw * 0.05
    gap = 18
    col_w = (sw - 2 * margin - gap * 2) / 3
    content_top = sh * 0.2
    content_h = sh * 0.72

    for i, col_data in enumerate(columns[:3]):
        cx = margin + i * (col_w + gap)
        c = accent_colors[i % len(accent_colors)]
        heading = col_data.get("heading", "")
        items = col_data.get("items", [])

        if style == "cards":
            # Card with shadow
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              cx, content_top, col_w, content_h,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=6, offset_x=2, offset_y=2, transparency=0.78)

            # Top color accent bar
            _add_shape(slide, MSO_SHAPE_RECTANGLE, cx, content_top, col_w, 5,
                       fill_rgb=c)

            # Heading
            _add_textbox(slide, cx + 16, content_top + 18, col_w - 32, 28, heading,
                         font_name="Segoe UI Semibold", font_size=17,
                         font_color=c, alignment=PP_ALIGN_CENTER)

            # Heading underline
            h_line = slide.Shapes.AddLine(cx + col_w * 0.2, content_top + 52,
                                          cx + col_w * 0.8, content_top + 52)
            h_line.Line.ForeColor.RGB = rgb(225, 228, 235)
            h_line.Line.Weight = 1

            # Items
            for j, item in enumerate(items):
                iy = content_top + 65 + j * 36
                _add_shape(slide, MSO_SHAPE_RECTANGLE, cx + 16, iy + 9, 6, 6,
                           fill_rgb=c)
                _add_textbox(slide, cx + 28, iy, col_w - 44, 30, item,
                             font_name="Segoe UI", font_size=12,
                             font_color=(55, 60, 70), alignment=PP_ALIGN_LEFT)

        elif style == "clean":
            # Vertical divider (except first column)
            if i > 0:
                div = slide.Shapes.AddLine(cx - gap / 2, content_top,
                                           cx - gap / 2, content_top + content_h)
                div.Line.ForeColor.RGB = rgb(220, 222, 228)
                div.Line.Weight = 1

            # Heading with accent color underline
            _add_textbox(slide, cx + 8, content_top, col_w - 16, 28, heading,
                         font_name="Segoe UI Semibold", font_size=17,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

            h_line = slide.Shapes.AddLine(cx + 8, content_top + 32,
                                          cx + col_w * 0.3, content_top + 32)
            h_line.Line.ForeColor.RGB = rgb(*c)
            h_line.Line.Weight = 2.5

            for j, item in enumerate(items):
                iy = content_top + 48 + j * 36
                _add_textbox(slide, cx + 8, iy, col_w - 16, 30, item,
                             font_name="Segoe UI", font_size=12,
                             font_color=(70, 70, 80), alignment=PP_ALIGN_LEFT)

        elif style == "icons":
            # Icon circle at top
            icon_r = 30
            icon_cx = cx + col_w / 2
            icon_cy = content_top + icon_r + 10

            _add_shape(slide, MSO_SHAPE_OVAL,
                       icon_cx - icon_r, icon_cy - icon_r,
                       icon_r * 2, icon_r * 2, fill_rgb=c)
            # Number in circle
            _add_textbox(slide, icon_cx - icon_r, icon_cy - icon_r,
                         icon_r * 2, icon_r * 2, str(i + 1),
                         font_name="Segoe UI", font_size=22,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Heading below icon
            _add_textbox(slide, cx + 8, icon_cy + icon_r + 12, col_w - 16, 28, heading,
                         font_name="Segoe UI Semibold", font_size=16,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

            # Items below heading
            for j, item in enumerate(items):
                iy = icon_cy + icon_r + 50 + j * 34
                _add_textbox(slide, cx + 12, iy, col_w - 24, 28, item,
                             font_name="Segoe UI", font_size=12,
                             font_color=(70, 75, 85), alignment=PP_ALIGN_CENTER)
        else:
            raise ValueError(f"Unknown three-column style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 21. create_stat_highlight
# ============================================================

def create_stat_highlight(
    slide_number: int,
    stats: list[dict],
    style: str = "big_numbers",
) -> dict:
    """Create a statistics highlight slide.

    Args:
        slide_number: 1-based slide index.
        stats: List of {"value": "95%", "label": "...", "color": [r,g,b] (optional)}.
        style: "big_numbers", "circles", "bars".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    n = len(stats)
    if n == 0:
        return {"slide_number": slide_number, "style": style}

    default_colors = [
        (0, 120, 215), (0, 180, 140), (255, 140, 0),
        (200, 50, 80), (100, 80, 200), (0, 166, 214),
    ]

    if style == "big_numbers":
        _set_gradient_bg(slide, (20, 25, 40), (35, 42, 65), MSO_GRADIENT_DIAGONAL_DOWN)

        margin = sw * 0.06
        gap = 20
        card_w = (sw - 2 * margin - gap * (n - 1)) / n
        card_h = sh * 0.55
        card_y = (sh - card_h) / 2

        for i, stat in enumerate(stats):
            cx = margin + i * (card_w + gap)
            c = tuple(stat.get("color", default_colors[i % len(default_colors)]))

            # Semi-transparent card
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              cx, card_y, card_w, card_h,
                              fill_rgb=(255, 255, 255), fill_transparency=0.9)

            # Top accent line
            _add_shape(slide, MSO_SHAPE_RECTANGLE, cx, card_y, card_w, 4,
                       fill_rgb=c)

            # Big value
            _add_textbox(slide, cx + 10, card_y + card_h * 0.15,
                         card_w - 20, card_h * 0.45,
                         stat.get("value", ""),
                         font_name="Segoe UI Light", font_size=56,
                         font_color=c, bold=False,
                         alignment=PP_ALIGN_CENTER)

            # Label below
            _add_textbox(slide, cx + 10, card_y + card_h * 0.6,
                         card_w - 20, card_h * 0.3,
                         stat.get("label", ""),
                         font_name="Segoe UI", font_size=14,
                         font_color=(200, 205, 220),
                         alignment=PP_ALIGN_CENTER)

    elif style == "circles":
        _set_solid_bg(slide, (248, 250, 252))

        margin = sw * 0.08
        gap = 20
        cell_w = (sw - 2 * margin) / n

        for i, stat in enumerate(stats):
            cx = margin + i * cell_w + cell_w / 2
            cy = sh * 0.4
            c = tuple(stat.get("color", default_colors[i % len(default_colors)]))

            # Outer circle (track)
            outer_r = min(cell_w * 0.35, sh * 0.25)
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx - outer_r, cy - outer_r,
                       outer_r * 2, outer_r * 2,
                       fill_rgb=_lerp_color(c, (255, 255, 255), 0.85))

            # Inner colored circle (progress indicator)
            inner_r = outer_r * 0.78
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx - inner_r, cy - inner_r,
                       inner_r * 2, inner_r * 2,
                       fill_rgb=c)

            # White center circle
            center_r = outer_r * 0.6
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx - center_r, cy - center_r,
                       center_r * 2, center_r * 2,
                       fill_rgb=(255, 255, 255))

            # Value in center
            _add_textbox(slide, cx - center_r, cy - center_r,
                         center_r * 2, center_r * 2,
                         stat.get("value", ""),
                         font_name="Segoe UI Semibold", font_size=24,
                         font_color=c, bold=False,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Label below circle
            _add_textbox(slide, cx - cell_w * 0.4, cy + outer_r + 15,
                         cell_w * 0.8, 40,
                         stat.get("label", ""),
                         font_name="Segoe UI", font_size=13,
                         font_color=(70, 75, 85),
                         alignment=PP_ALIGN_CENTER)

    elif style == "bars":
        _set_solid_bg(slide, (250, 250, 252))

        margin_x = sw * 0.1
        bar_area_w = sw * 0.65
        bar_h = 22
        gap = max(20, (sh * 0.7) / n - bar_h)
        start_y = sh * 0.12

        for i, stat in enumerate(stats):
            c = tuple(stat.get("color", default_colors[i % len(default_colors)]))
            y = start_y + i * (bar_h + gap)

            # Label on the left
            _add_textbox(slide, margin_x, y - 22, bar_area_w, 20,
                         stat.get("label", ""),
                         font_name="Segoe UI Semibold", font_size=13,
                         font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)

            # Background track
            _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                       margin_x, y, bar_area_w, bar_h,
                       fill_rgb=_lerp_color(c, (255, 255, 255), 0.85))

            # Filled bar - try to parse percentage from value
            val_str = stat.get("value", "0")
            try:
                pct = float(val_str.replace("%", "").replace(",", "").strip()) / 100
            except (ValueError, AttributeError):
                pct = 0.7
            pct = max(0.05, min(1.0, pct))

            fill_w = bar_area_w * pct
            _add_gradient_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                margin_x, y, fill_w, bar_h,
                                c, _lerp_color(c, (255, 255, 255), 0.25),
                                MSO_GRADIENT_HORIZONTAL)

            # Value at end of bar
            _add_textbox(slide, margin_x + fill_w + 10, y - 2,
                         80, bar_h + 4, stat.get("value", ""),
                         font_name="Segoe UI Semibold", font_size=15,
                         font_color=c, bold=True,
                         alignment=PP_ALIGN_LEFT)
    else:
        raise ValueError(f"Unknown stat highlight style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 22. create_timeline_slide
# ============================================================

def create_timeline_slide(
    slide_number: int,
    title: str,
    events: list[dict],
    style: str = "horizontal",
) -> dict:
    """Create a timeline slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        events: List of {"date": "...", "title": "...", "description": "..."}.
        style: "horizontal", "vertical", "alternating".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    n = len(events)
    if n == 0:
        return {"slide_number": slide_number, "style": style}

    colors = [
        (0, 120, 190), (0, 166, 140), (242, 150, 0),
        (200, 50, 80), (100, 80, 200), (0, 166, 214),
    ]

    if style == "horizontal":
        _set_solid_bg(slide, (248, 250, 252))

        # Title
        _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=28,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

        line_y = sh * 0.52
        margin_x = sw * 0.08
        usable_w = sw * 0.84

        # Main timeline line
        main_line = slide.Shapes.AddLine(margin_x, line_y,
                                          margin_x + usable_w, line_y)
        main_line.Line.ForeColor.RGB = rgb(200, 205, 215)
        main_line.Line.Weight = 3

        step = usable_w / max(n - 1, 1) if n > 1 else 0

        for i, event in enumerate(events):
            cx = margin_x + i * step if n > 1 else margin_x + usable_w / 2
            c = colors[i % len(colors)]

            # Node circle
            node_r = 10
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx - node_r, line_y - node_r,
                       node_r * 2, node_r * 2, fill_rgb=c)

            # Connector
            if i % 2 == 0:
                # Above timeline
                card_y = line_y - 155
                conn = slide.Shapes.AddLine(cx, line_y - node_r - 2, cx, card_y + 110)
                conn.Line.ForeColor.RGB = rgb(*c)
                conn.Line.Weight = 1.5
            else:
                # Below timeline
                card_y = line_y + 30
                conn = slide.Shapes.AddLine(cx, line_y + node_r + 2, cx, card_y)
                conn.Line.ForeColor.RGB = rgb(*c)
                conn.Line.Weight = 1.5

            card_w = min(step * 0.85, 160) if n > 1 else 180

            # Date label
            _add_textbox(slide, cx - card_w / 2, card_y, card_w, 20,
                         event.get("date", ""),
                         font_name="Segoe UI Semibold", font_size=10,
                         font_color=c, bold=True, alignment=PP_ALIGN_CENTER)

            # Event title
            _add_textbox(slide, cx - card_w / 2, card_y + 22, card_w, 22,
                         event.get("title", ""),
                         font_name="Segoe UI Semibold", font_size=12,
                         font_color=(40, 42, 50), alignment=PP_ALIGN_CENTER)

            # Description
            desc = event.get("description", "")
            if desc:
                _add_textbox(slide, cx - card_w / 2, card_y + 46, card_w, 60, desc,
                             font_name="Segoe UI", font_size=10,
                             font_color=(100, 105, 115), alignment=PP_ALIGN_CENTER)

    elif style == "vertical":
        _set_solid_bg(slide, (250, 250, 252))

        # Title
        _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=28,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

        line_x = sw * 0.2
        start_y = sh * 0.18
        usable_h = sh * 0.76
        step_h = usable_h / max(n, 1)

        # Vertical timeline line
        v_line = slide.Shapes.AddLine(line_x, start_y, line_x, start_y + usable_h)
        v_line.Line.ForeColor.RGB = rgb(200, 205, 215)
        v_line.Line.Weight = 3

        for i, event in enumerate(events):
            cy = start_y + i * step_h + step_h / 2
            c = colors[i % len(colors)]

            # Node
            node_r = 8
            _add_shape(slide, MSO_SHAPE_OVAL,
                       line_x - node_r, cy - node_r,
                       node_r * 2, node_r * 2, fill_rgb=c)

            # Date on left
            _add_textbox(slide, sw * 0.03, cy - 12, line_x - sw * 0.05, 24,
                         event.get("date", ""),
                         font_name="Segoe UI Semibold", font_size=11,
                         font_color=c, alignment=PP_ALIGN_RIGHT)

            # Content card on right
            card_x = line_x + 25
            card_w = sw * 0.68
            card_h = step_h - 12

            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              card_x, cy - card_h / 2, card_w, card_h,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=4, offset_x=1, offset_y=1, transparency=0.85)

            # Left accent bar on card
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       card_x, cy - card_h / 2, 4, card_h, fill_rgb=c)

            # Event title
            _add_textbox(slide, card_x + 14, cy - card_h / 2 + 6,
                         card_w - 24, 22,
                         event.get("title", ""),
                         font_name="Segoe UI Semibold", font_size=13,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

            # Description
            desc = event.get("description", "")
            if desc:
                _add_textbox(slide, card_x + 14, cy - card_h / 2 + 28,
                             card_w - 24, card_h - 34, desc,
                             font_name="Segoe UI", font_size=10,
                             font_color=(100, 105, 115), alignment=PP_ALIGN_LEFT)

            # Connector line
            conn = slide.Shapes.AddLine(line_x + node_r + 2, cy, card_x, cy)
            conn.Line.ForeColor.RGB = rgb(*c)
            conn.Line.Weight = 1.5

    elif style == "alternating":
        _set_solid_bg(slide, (245, 247, 250))

        # Title
        _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.1, title,
                     font_name="Segoe UI Semibold", font_size=28,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        center_x = sw * 0.5
        start_y = sh * 0.18
        usable_h = sh * 0.76
        step_h = usable_h / max(n, 1)

        # Center vertical line
        v_line = slide.Shapes.AddLine(center_x, start_y, center_x, start_y + usable_h)
        v_line.Line.ForeColor.RGB = rgb(200, 205, 215)
        v_line.Line.Weight = 3

        for i, event in enumerate(events):
            cy = start_y + i * step_h + step_h / 2
            c = colors[i % len(colors)]
            is_left = (i % 2 == 0)

            # Node
            node_r = 10
            node = _add_shape(slide, MSO_SHAPE_OVAL,
                              center_x - node_r, cy - node_r,
                              node_r * 2, node_r * 2, fill_rgb=c)

            card_w = sw * 0.36
            card_h = step_h - 16

            if is_left:
                card_x = center_x - card_w - 30
                conn = slide.Shapes.AddLine(center_x - node_r - 2, cy,
                                            card_x + card_w, cy)
            else:
                card_x = center_x + 30
                conn = slide.Shapes.AddLine(center_x + node_r + 2, cy,
                                            card_x, cy)
            conn.Line.ForeColor.RGB = rgb(*c)
            conn.Line.Weight = 1.5

            # Card
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              card_x, cy - card_h / 2, card_w, card_h,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=5, offset_x=2, offset_y=2, transparency=0.8)

            # Top accent
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       card_x, cy - card_h / 2, card_w, 4, fill_rgb=c)

            # Date
            _add_textbox(slide, card_x + 10, cy - card_h / 2 + 8,
                         card_w - 20, 18,
                         event.get("date", ""),
                         font_name="Segoe UI Semibold", font_size=10,
                         font_color=c, alignment=PP_ALIGN_LEFT)

            # Title
            _add_textbox(slide, card_x + 10, cy - card_h / 2 + 26,
                         card_w - 20, 20,
                         event.get("title", ""),
                         font_name="Segoe UI Semibold", font_size=12,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

            # Description
            desc = event.get("description", "")
            if desc:
                _add_textbox(slide, card_x + 10, cy - card_h / 2 + 48,
                             card_w - 20, card_h - 54, desc,
                             font_name="Segoe UI", font_size=10,
                             font_color=(100, 105, 115), alignment=PP_ALIGN_LEFT)
    else:
        raise ValueError(f"Unknown timeline style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 23. create_funnel_diagram
# ============================================================

def create_funnel_diagram(
    slide_number: int,
    title: str,
    stages: list[dict],
    style: str = "gradient",
) -> dict:
    """Create a funnel diagram slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        stages: List of {"label": "...", "value": "..."} from widest to narrowest.
        style: "gradient", "flat", "3d".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    n = len(stages)
    if n == 0:
        return {"slide_number": slide_number, "style": style}

    # Title
    _set_solid_bg(slide, (248, 250, 252) if style != "3d" else (240, 242, 248))

    _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=28,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

    funnel_left = sw * 0.15
    funnel_right = sw * 0.6
    funnel_top = sh * 0.2
    funnel_bottom = sh * 0.9
    max_w = funnel_right - funnel_left
    step_h = (funnel_bottom - funnel_top) / n
    gap = 4

    base_colors = [
        (0, 100, 200), (0, 140, 180), (0, 170, 140),
        (80, 180, 80), (200, 180, 0), (230, 130, 0),
    ]

    for i, stage in enumerate(stages):
        # Progressively narrower
        t = i / max(n - 1, 1)
        current_w = max_w * (1 - t * 0.6)
        next_w = max_w * (1 - (i + 1) / max(n - 1, 1) * 0.6) if i < n - 1 else current_w * 0.7
        cx = funnel_left + (max_w - current_w) / 2
        y = funnel_top + i * step_h
        c = base_colors[i % len(base_colors)]

        if style == "gradient":
            lighter = _lerp_color(c, (255, 255, 255), 0.25)
            shape = _add_gradient_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                        cx, y, current_w, step_h - gap,
                                        c, lighter, MSO_GRADIENT_HORIZONTAL)
            _add_shadow(shape, blur=4, offset_x=2, offset_y=2, transparency=0.75)
        elif style == "flat":
            shape = _add_shape(slide, MSO_SHAPE_RECTANGLE,
                               cx, y, current_w, step_h - gap, fill_rgb=c)
        elif style == "3d":
            # Main shape
            shape = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                               cx, y, current_w, step_h - gap, fill_rgb=c)
            # Side shadow for 3D effect
            darker = _lerp_color(c, (0, 0, 0), 0.3)
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       cx + current_w, y + 4, 8, step_h - gap - 4,
                       fill_rgb=darker)
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       cx + 4, y + step_h - gap, current_w, 6,
                       fill_rgb=darker)
        else:
            raise ValueError(f"Unknown funnel style '{style}'")

        # Label text inside
        _add_textbox(slide, cx + 12, y + 4, current_w - 24, step_h - gap - 8,
                     stage.get("label", ""),
                     font_name="Segoe UI Semibold", font_size=14,
                     font_color=(255, 255, 255),
                     alignment=PP_ALIGN_CENTER,
                     vertical_anchor=MSO_ANCHOR_MIDDLE)

        # Value on the right side
        val_x = funnel_left + max_w + 30
        _add_textbox(slide, val_x, y + (step_h - gap) / 2 - 12,
                     sw * 0.25, 24,
                     stage.get("value", ""),
                     font_name="Segoe UI Light", font_size=20,
                     font_color=c, bold=False,
                     alignment=PP_ALIGN_LEFT)

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 24. create_swot_slide
# ============================================================

def create_swot_slide(
    slide_number: int,
    strengths: list[str],
    weaknesses: list[str],
    opportunities: list[str],
    threats: list[str],
    style: str = "colored",
) -> dict:
    """Create a SWOT analysis slide.

    Args:
        slide_number: 1-based slide index.
        strengths: List of strength items.
        weaknesses: List of weakness items.
        opportunities: List of opportunity items.
        threats: List of threat items.
        style: "colored" (4 bg colors), "minimal" (borders only), "icons" (with S/W/O/T circles).

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    _set_solid_bg(slide, (248, 250, 252) if style != "minimal" else (255, 255, 255))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.03, sw * 0.88, sh * 0.08, "SWOT Analysis",
                 font_name="Segoe UI Semibold", font_size=28,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

    margin = sw * 0.05
    gap = 14
    quad_w = (sw - 2 * margin - gap) / 2
    quad_h = (sh * 0.82 - gap) / 2
    top_y = sh * 0.13

    quadrants = [
        ("S", "Strengths", strengths, (0, 150, 80), (230, 248, 238)),
        ("W", "Weaknesses", weaknesses, (220, 60, 60), (252, 235, 235)),
        ("O", "Opportunities", opportunities, (0, 120, 200), (230, 242, 255)),
        ("T", "Threats", threats, (230, 150, 0), (255, 245, 225)),
    ]

    positions = [
        (margin, top_y),
        (margin + quad_w + gap, top_y),
        (margin, top_y + quad_h + gap),
        (margin + quad_w + gap, top_y + quad_h + gap),
    ]

    for idx, (letter, label, items, accent, bg_color) in enumerate(quadrants):
        qx, qy = positions[idx]

        if style == "colored":
            # Colored background card
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              qx, qy, quad_w, quad_h,
                              fill_rgb=bg_color)
            # Top accent bar
            _add_shape(slide, MSO_SHAPE_RECTANGLE, qx, qy, quad_w, 5,
                       fill_rgb=accent)

            # Label
            _add_textbox(slide, qx + 14, qy + 12, quad_w - 28, 24, label,
                         font_name="Segoe UI Semibold", font_size=16,
                         font_color=accent, alignment=PP_ALIGN_LEFT)

            # Items
            for j, item in enumerate(items):
                iy = qy + 44 + j * 30
                if iy + 24 > qy + quad_h:
                    break
                _add_shape(slide, MSO_SHAPE_OVAL, qx + 16, iy + 6, 7, 7,
                           fill_rgb=accent)
                _add_textbox(slide, qx + 30, iy, quad_w - 44, 24, item,
                             font_name="Segoe UI", font_size=11,
                             font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)

        elif style == "minimal":
            # White card with border
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              qx, qy, quad_w, quad_h,
                              fill_rgb=(255, 255, 255),
                              line_visible=True, line_rgb=(210, 212, 220),
                              line_weight=1.5)

            # Label with accent color
            _add_textbox(slide, qx + 14, qy + 12, quad_w - 28, 24, label,
                         font_name="Segoe UI Semibold", font_size=16,
                         font_color=accent, alignment=PP_ALIGN_LEFT)

            # Underline
            h_line = slide.Shapes.AddLine(qx + 14, qy + 40,
                                          qx + quad_w * 0.35, qy + 40)
            h_line.Line.ForeColor.RGB = rgb(*accent)
            h_line.Line.Weight = 2

            # Items
            for j, item in enumerate(items):
                iy = qy + 50 + j * 28
                if iy + 22 > qy + quad_h:
                    break
                _add_textbox(slide, qx + 18, iy, quad_w - 36, 22, item,
                             font_name="Segoe UI", font_size=11,
                             font_color=(60, 65, 75), alignment=PP_ALIGN_LEFT)

        elif style == "icons":
            # White card with shadow
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              qx, qy, quad_w, quad_h,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=5, offset_x=2, offset_y=2, transparency=0.8)

            # Icon circle with letter
            icon_r = 20
            _add_shape(slide, MSO_SHAPE_OVAL,
                       qx + 14, qy + 12, icon_r * 2, icon_r * 2,
                       fill_rgb=accent)
            _add_textbox(slide, qx + 14, qy + 12, icon_r * 2, icon_r * 2, letter,
                         font_name="Segoe UI", font_size=18,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Label next to icon
            _add_textbox(slide, qx + icon_r * 2 + 22, qy + 16, quad_w - 80, 28,
                         label,
                         font_name="Segoe UI Semibold", font_size=15,
                         font_color=accent, alignment=PP_ALIGN_LEFT)

            # Items
            for j, item in enumerate(items):
                iy = qy + icon_r * 2 + 22 + j * 28
                if iy + 22 > qy + quad_h:
                    break
                _add_shape(slide, MSO_SHAPE_RECTANGLE, qx + 16, iy + 7, 5, 5,
                           fill_rgb=accent)
                _add_textbox(slide, qx + 28, iy, quad_w - 44, 22, item,
                             font_name="Segoe UI", font_size=11,
                             font_color=(55, 60, 70), alignment=PP_ALIGN_LEFT)
        else:
            raise ValueError(f"Unknown SWOT style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 25. create_roadmap_slide
# ============================================================

def create_roadmap_slide(
    slide_number: int,
    title: str,
    phases: list[dict],
    style: str = "arrow",
) -> dict:
    """Create a roadmap slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        phases: List of {"name": "...", "period": "...", "items": [...]}.
        style: "arrow" (arrow shapes), "lane" (swim lanes), "milestone" (milestone dots).

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    n = len(phases)
    if n == 0:
        return {"slide_number": slide_number, "style": style}

    _set_solid_bg(slide, (248, 250, 252))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=28,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

    colors = [
        (0, 100, 200), (0, 160, 140), (230, 140, 0),
        (200, 50, 80), (100, 80, 200), (0, 166, 214),
    ]

    if style == "arrow":
        margin = sw * 0.05
        gap = 6
        arrow_w = (sw - 2 * margin - gap * (n - 1)) / n
        arrow_h = 65
        arrow_y = sh * 0.2

        for i, phase in enumerate(phases):
            ax = margin + i * (arrow_w + gap)
            c = colors[i % len(colors)]

            # Arrow/chevron shape
            shape_type = MSO_SHAPE_CHEVRON if i < n - 1 else MSO_SHAPE_ROUNDED_RECTANGLE
            shape = slide.Shapes.AddShape(shape_type, ax, arrow_y, arrow_w, arrow_h)
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = rgb(*c)
            shape.Line.Visible = False
            _add_shadow(shape, blur=4, offset_x=2, offset_y=2, transparency=0.75)

            # Phase name in arrow
            _add_textbox(slide, ax + arrow_w * 0.1, arrow_y + 8,
                         arrow_w * 0.8, 28,
                         phase.get("name", ""),
                         font_name="Segoe UI Semibold", font_size=12,
                         font_color=(255, 255, 255),
                         alignment=PP_ALIGN_CENTER)

            # Period below arrow
            _add_textbox(slide, ax + arrow_w * 0.1, arrow_y + 36,
                         arrow_w * 0.8, 22,
                         phase.get("period", ""),
                         font_name="Segoe UI", font_size=9,
                         font_color=(230, 235, 245),
                         alignment=PP_ALIGN_CENTER)

            # Items below arrow in a card
            items = phase.get("items", [])
            if items:
                card_y = arrow_y + arrow_h + 15
                card_h = sh * 0.5
                card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                  ax, card_y, arrow_w, card_h,
                                  fill_rgb=(255, 255, 255))
                _add_shadow(card, blur=4, offset_x=1, offset_y=1, transparency=0.85)

                # Top accent line on card
                _add_shape(slide, MSO_SHAPE_RECTANGLE,
                           ax, card_y, arrow_w, 3, fill_rgb=c)

                for j, item in enumerate(items):
                    iy = card_y + 14 + j * 30
                    if iy + 24 > card_y + card_h:
                        break
                    _add_shape(slide, MSO_SHAPE_RECTANGLE,
                               ax + 12, iy + 7, 5, 5, fill_rgb=c)
                    _add_textbox(slide, ax + 22, iy, arrow_w - 34, 24, item,
                                 font_name="Segoe UI", font_size=10,
                                 font_color=(55, 60, 70), alignment=PP_ALIGN_LEFT)

    elif style == "lane":
        margin = sw * 0.05
        gap = 4
        lane_w = (sw - 2 * margin - gap * (n - 1)) / n
        lane_top = sh * 0.18
        lane_h = sh * 0.76

        for i, phase in enumerate(phases):
            lx = margin + i * (lane_w + gap)
            c = colors[i % len(colors)]

            # Lane header
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       lx, lane_top, lane_w, 50, fill_rgb=c)
            _add_textbox(slide, lx + 6, lane_top + 4, lane_w - 12, 24,
                         phase.get("name", ""),
                         font_name="Segoe UI Semibold", font_size=12,
                         font_color=(255, 255, 255), alignment=PP_ALIGN_CENTER)
            _add_textbox(slide, lx + 6, lane_top + 26, lane_w - 12, 18,
                         phase.get("period", ""),
                         font_name="Segoe UI", font_size=9,
                         font_color=(220, 225, 240), alignment=PP_ALIGN_CENTER)

            # Lane body
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       lx, lane_top + 50, lane_w, lane_h - 50,
                       fill_rgb=_lerp_color(c, (255, 255, 255), 0.92))

            # Items in lane
            items = phase.get("items", [])
            for j, item in enumerate(items):
                iy = lane_top + 60 + j * 36
                if iy + 28 > lane_top + lane_h:
                    break
                item_card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                       lx + 6, iy, lane_w - 12, 30,
                                       fill_rgb=(255, 255, 255))
                _add_shadow(item_card, blur=2, offset_x=1, offset_y=1, transparency=0.85)
                _add_shape(slide, MSO_SHAPE_RECTANGLE,
                           lx + 6, iy, 4, 30, fill_rgb=c)
                _add_textbox(slide, lx + 16, iy + 3, lane_w - 28, 24, item,
                             font_name="Segoe UI", font_size=10,
                             font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)

    elif style == "milestone":
        margin = sw * 0.08
        usable_w = sw * 0.84
        line_y = sh * 0.32

        # Main horizontal line
        main_line = slide.Shapes.AddLine(margin, line_y,
                                          margin + usable_w, line_y)
        main_line.Line.ForeColor.RGB = rgb(190, 195, 205)
        main_line.Line.Weight = 4

        step = usable_w / max(n - 1, 1) if n > 1 else 0

        for i, phase in enumerate(phases):
            cx = margin + i * step if n > 1 else margin + usable_w / 2
            c = colors[i % len(colors)]

            # Milestone diamond
            diamond_size = 20
            diamond = _add_shape(slide, MSO_SHAPE_DIAMOND,
                                 cx - diamond_size, line_y - diamond_size,
                                 diamond_size * 2, diamond_size * 2,
                                 fill_rgb=c)
            _add_shadow(diamond, blur=3, offset_x=1, offset_y=1, transparency=0.75)

            # Phase name and period above
            _add_textbox(slide, cx - 80, line_y - 65, 160, 22,
                         phase.get("name", ""),
                         font_name="Segoe UI Semibold", font_size=13,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)
            _add_textbox(slide, cx - 80, line_y - 45, 160, 18,
                         phase.get("period", ""),
                         font_name="Segoe UI", font_size=10,
                         font_color=c, alignment=PP_ALIGN_CENTER)

            # Items below milestone
            items = phase.get("items", [])
            item_w = min(step * 0.85, 170) if n > 1 else 180
            conn = slide.Shapes.AddLine(cx, line_y + diamond_size + 2,
                                        cx, line_y + diamond_size + 20)
            conn.Line.ForeColor.RGB = rgb(*c)
            conn.Line.Weight = 1.5

            for j, item in enumerate(items):
                iy = line_y + diamond_size + 25 + j * 28
                _add_textbox(slide, cx - item_w / 2, iy, item_w, 22, item,
                             font_name="Segoe UI", font_size=10,
                             font_color=(70, 75, 85), alignment=PP_ALIGN_CENTER)
    else:
        raise ValueError(f"Unknown roadmap style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 26. create_pricing_table
# ============================================================

def create_pricing_table(
    slide_number: int,
    title: str,
    plans: list[dict],
    style: str = "cards",
) -> dict:
    """Create a pricing comparison slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        plans: List of {"name": "...", "price": "...", "features": [...], "highlighted": bool}.
        style: "cards" (vertical cards), "table" (horizontal table), "minimal".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    n = len(plans)
    if n == 0:
        return {"slide_number": slide_number, "style": style}

    _set_solid_bg(slide, (245, 247, 252))

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.03, sw * 0.88, sh * 0.08, title,
                 font_name="Segoe UI Semibold", font_size=28,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

    accent = (0, 100, 200)
    highlight_color = (0, 120, 215)

    if style == "cards":
        margin = sw * 0.06
        gap = 16
        card_w = (sw - 2 * margin - gap * (n - 1)) / n
        card_h = sh * 0.8
        card_y = sh * 0.14

        for i, plan in enumerate(plans):
            px = margin + i * (card_w + gap)
            is_hl = plan.get("highlighted", False)

            if is_hl:
                # Highlighted card - taller, colored header
                card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                  px, card_y - 10, card_w, card_h + 20,
                                  fill_rgb=(255, 255, 255))
                _add_shadow(card, blur=10, offset_x=3, offset_y=3, transparency=0.6)

                # Colored header
                _add_shape(slide, MSO_SHAPE_RECTANGLE,
                           px, card_y - 10, card_w, 80,
                           fill_rgb=highlight_color)

                # Plan name
                _add_textbox(slide, px + 8, card_y - 4, card_w - 16, 28,
                             plan.get("name", ""),
                             font_name="Segoe UI Semibold", font_size=16,
                             font_color=(255, 255, 255), alignment=PP_ALIGN_CENTER)
                # Price
                _add_textbox(slide, px + 8, card_y + 24, card_w - 16, 36,
                             plan.get("price", ""),
                             font_name="Segoe UI Light", font_size=28,
                             font_color=(255, 255, 255), alignment=PP_ALIGN_CENTER)

                # "POPULAR" badge
                badge_w = 80
                _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                           px + (card_w - badge_w) / 2, card_y + 70,
                           badge_w, 22, fill_rgb=(255, 200, 0))
                _add_textbox(slide, px + (card_w - badge_w) / 2, card_y + 70,
                             badge_w, 22, "POPULAR",
                             font_name="Segoe UI", font_size=8,
                             font_color=(40, 40, 50), bold=True,
                             alignment=PP_ALIGN_CENTER,
                             vertical_anchor=MSO_ANCHOR_MIDDLE)

                feat_start = card_y + 100
            else:
                card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                                  px, card_y, card_w, card_h,
                                  fill_rgb=(255, 255, 255))
                _add_shadow(card, blur=6, offset_x=2, offset_y=2, transparency=0.78)

                # Plan name
                _add_textbox(slide, px + 8, card_y + 14, card_w - 16, 28,
                             plan.get("name", ""),
                             font_name="Segoe UI Semibold", font_size=16,
                             font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)
                # Price
                _add_textbox(slide, px + 8, card_y + 42, card_w - 16, 36,
                             plan.get("price", ""),
                             font_name="Segoe UI Light", font_size=28,
                             font_color=highlight_color, alignment=PP_ALIGN_CENTER)

                feat_start = card_y + 85

            # Divider line
            div = slide.Shapes.AddLine(px + 16, feat_start,
                                        px + card_w - 16, feat_start)
            div.Line.ForeColor.RGB = rgb(225, 228, 235)
            div.Line.Weight = 1

            # Features
            features = plan.get("features", [])
            for j, feat in enumerate(features):
                fy = feat_start + 12 + j * 28
                # Checkmark
                _add_textbox(slide, px + 12, fy, 18, 20, "\u2713",
                             font_name="Segoe UI", font_size=12,
                             font_color=(0, 160, 80), bold=True,
                             alignment=PP_ALIGN_CENTER)
                _add_textbox(slide, px + 32, fy, card_w - 48, 22, feat,
                             font_name="Segoe UI", font_size=11,
                             font_color=(60, 65, 75), alignment=PP_ALIGN_LEFT)

    elif style == "table":
        margin = sw * 0.05
        table_w = sw - 2 * margin
        header_h = 70
        row_h = 32
        table_y = sh * 0.14

        # Determine max features
        max_feats = max(len(p.get("features", [])) for p in plans) if plans else 0
        col_w = table_w / (n + 1)  # +1 for feature label column
        feat_col_w = col_w * 1.2
        plan_col_w = (table_w - feat_col_w) / n

        # Header row
        for i, plan in enumerate(plans):
            hx = margin + feat_col_w + i * plan_col_w
            is_hl = plan.get("highlighted", False)
            h_bg = highlight_color if is_hl else (60, 65, 80)

            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       hx, table_y, plan_col_w, header_h, fill_rgb=h_bg)
            _add_textbox(slide, hx + 4, table_y + 8, plan_col_w - 8, 24,
                         plan.get("name", ""),
                         font_name="Segoe UI Semibold", font_size=14,
                         font_color=(255, 255, 255), alignment=PP_ALIGN_CENTER)
            _add_textbox(slide, hx + 4, table_y + 32, plan_col_w - 8, 30,
                         plan.get("price", ""),
                         font_name="Segoe UI Light", font_size=20,
                         font_color=(255, 255, 255), alignment=PP_ALIGN_CENTER)

        # Feature header
        _add_shape(slide, MSO_SHAPE_RECTANGLE,
                   margin, table_y, feat_col_w, header_h,
                   fill_rgb=(45, 50, 60))
        _add_textbox(slide, margin + 10, table_y + 20, feat_col_w - 20, 30,
                     "Features",
                     font_name="Segoe UI Semibold", font_size=15,
                     font_color=(255, 255, 255), alignment=PP_ALIGN_LEFT)

        # Feature rows - collect all unique features
        all_features = []
        for plan in plans:
            for f in plan.get("features", []):
                if f not in all_features:
                    all_features.append(f)

        for j, feat in enumerate(all_features):
            ry = table_y + header_h + j * row_h
            row_bg = (248, 250, 252) if j % 2 == 0 else (255, 255, 255)

            # Feature name
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       margin, ry, feat_col_w, row_h, fill_rgb=row_bg)
            _add_textbox(slide, margin + 12, ry + 4, feat_col_w - 16, row_h - 8, feat,
                         font_name="Segoe UI", font_size=11,
                         font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)

            # Check/cross for each plan
            for i, plan in enumerate(plans):
                cx = margin + feat_col_w + i * plan_col_w
                _add_shape(slide, MSO_SHAPE_RECTANGLE,
                           cx, ry, plan_col_w, row_h, fill_rgb=row_bg)
                has_feat = feat in plan.get("features", [])
                check = "\u2713" if has_feat else "\u2014"
                check_color = (0, 160, 80) if has_feat else (190, 190, 200)
                _add_textbox(slide, cx + 4, ry + 4, plan_col_w - 8, row_h - 8,
                             check,
                             font_name="Segoe UI", font_size=13,
                             font_color=check_color, bold=True,
                             alignment=PP_ALIGN_CENTER)

    elif style == "minimal":
        margin = sw * 0.06
        gap = 20
        card_w = (sw - 2 * margin - gap * (n - 1)) / n
        card_h = sh * 0.78
        card_y = sh * 0.15

        for i, plan in enumerate(plans):
            px = margin + i * (card_w + gap)
            is_hl = plan.get("highlighted", False)

            # Thin border card
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              px, card_y, card_w, card_h,
                              fill_rgb=(255, 255, 255),
                              line_visible=True,
                              line_rgb=highlight_color if is_hl else (210, 212, 220),
                              line_weight=2 if is_hl else 1)

            # Plan name
            _add_textbox(slide, px + 12, card_y + 16, card_w - 24, 24,
                         plan.get("name", ""),
                         font_name="Segoe UI", font_size=14,
                         font_color=(100, 105, 115), alignment=PP_ALIGN_CENTER)

            # Price
            _add_textbox(slide, px + 12, card_y + 44, card_w - 24, 40,
                         plan.get("price", ""),
                         font_name="Segoe UI Light", font_size=32,
                         font_color=(33, 37, 41) if is_hl else (80, 85, 95),
                         alignment=PP_ALIGN_CENTER)

            # Divider
            div = slide.Shapes.AddLine(px + 20, card_y + 90,
                                        px + card_w - 20, card_y + 90)
            div.Line.ForeColor.RGB = rgb(230, 232, 238)
            div.Line.Weight = 1

            # Features
            features = plan.get("features", [])
            for j, feat in enumerate(features):
                fy = card_y + 100 + j * 28
                _add_textbox(slide, px + 16, fy, card_w - 32, 22, feat,
                             font_name="Segoe UI", font_size=11,
                             font_color=(70, 75, 85), alignment=PP_ALIGN_CENTER)
    else:
        raise ValueError(f"Unknown pricing table style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 27. create_icon_list_slide
# ============================================================

def create_icon_list_slide(
    slide_number: int,
    title: str,
    items: list[dict],
    style: str = "horizontal",
) -> dict:
    """Create an icon + text list layout slide.

    Args:
        slide_number: 1-based slide index.
        title: Slide title.
        items: List of {"icon_text": "01", "title": "...", "description": "..."}.
        style: "horizontal" (side by side), "vertical" (stacked), "grid" (2x2/2x3).

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    n = len(items)
    if n == 0:
        return {"slide_number": slide_number, "style": style}

    _set_solid_bg(slide, (248, 250, 252))

    colors = [
        (0, 120, 215), (0, 170, 140), (242, 150, 0),
        (200, 50, 80), (100, 80, 200), (0, 166, 214),
    ]

    # Title
    _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.1, title,
                 font_name="Segoe UI Semibold", font_size=28,
                 font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

    line = slide.Shapes.AddLine(sw * 0.06, sh * 0.14, sw * 0.18, sh * 0.14)
    line.Line.ForeColor.RGB = rgb(*colors[0])
    line.Line.Weight = 3

    if style == "horizontal":
        margin = sw * 0.06
        gap = 20
        item_w = (sw - 2 * margin - gap * (n - 1)) / n
        item_y = sh * 0.22

        for i, item in enumerate(items):
            ix = margin + i * (item_w + gap)
            c = colors[i % len(colors)]

            # Icon circle
            icon_r = 28
            icon_cx = ix + item_w / 2
            _add_shape(slide, MSO_SHAPE_OVAL,
                       icon_cx - icon_r, item_y, icon_r * 2, icon_r * 2,
                       fill_rgb=c)
            _add_textbox(slide, icon_cx - icon_r, item_y,
                         icon_r * 2, icon_r * 2,
                         item.get("icon_text", ""),
                         font_name="Segoe UI", font_size=16,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Title
            _add_textbox(slide, ix, item_y + icon_r * 2 + 14, item_w, 24,
                         item.get("title", ""),
                         font_name="Segoe UI Semibold", font_size=14,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

            # Description
            desc = item.get("description", "")
            if desc:
                _add_textbox(slide, ix + 4, item_y + icon_r * 2 + 42,
                             item_w - 8, sh * 0.45, desc,
                             font_name="Segoe UI", font_size=12,
                             font_color=(80, 85, 95), alignment=PP_ALIGN_CENTER)

    elif style == "vertical":
        margin = sw * 0.08
        item_h = min(70, (sh * 0.72) / n)
        start_y = sh * 0.2

        for i, item in enumerate(items):
            y = start_y + i * item_h
            c = colors[i % len(colors)]

            # Icon circle on the left
            icon_r = 22
            _add_shape(slide, MSO_SHAPE_OVAL,
                       margin, y + 5, icon_r * 2, icon_r * 2,
                       fill_rgb=c)
            _add_textbox(slide, margin, y + 5,
                         icon_r * 2, icon_r * 2,
                         item.get("icon_text", ""),
                         font_name="Segoe UI", font_size=14,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Title next to icon
            _add_textbox(slide, margin + icon_r * 2 + 16, y + 2,
                         sw * 0.3, 24,
                         item.get("title", ""),
                         font_name="Segoe UI Semibold", font_size=15,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

            # Description
            desc = item.get("description", "")
            if desc:
                _add_textbox(slide, margin + icon_r * 2 + 16, y + 26,
                             sw * 0.6, 36, desc,
                             font_name="Segoe UI", font_size=12,
                             font_color=(90, 95, 105), alignment=PP_ALIGN_LEFT)

    elif style == "grid":
        cols = 3 if n > 4 else 2
        rows = (n + cols - 1) // cols
        margin = sw * 0.06
        gap_x = 24
        gap_y = 18
        cell_w = (sw - 2 * margin - gap_x * (cols - 1)) / cols
        cell_h = min((sh * 0.72) / rows - gap_y, 180)
        start_y = sh * 0.2

        for i, item in enumerate(items):
            col = i % cols
            row = i // cols
            cx = margin + col * (cell_w + gap_x)
            cy = start_y + row * (cell_h + gap_y)
            c = colors[i % len(colors)]

            # Card background
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              cx, cy, cell_w, cell_h,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=5, offset_x=2, offset_y=2, transparency=0.8)

            # Icon circle
            icon_r = 22
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx + 16, cy + 16, icon_r * 2, icon_r * 2,
                       fill_rgb=c)
            _add_textbox(slide, cx + 16, cy + 16,
                         icon_r * 2, icon_r * 2,
                         item.get("icon_text", ""),
                         font_name="Segoe UI", font_size=14,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Title
            _add_textbox(slide, cx + icon_r * 2 + 24, cy + 16,
                         cell_w - icon_r * 2 - 40, 24,
                         item.get("title", ""),
                         font_name="Segoe UI Semibold", font_size=13,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)

            # Description
            desc = item.get("description", "")
            if desc:
                _add_textbox(slide, cx + 16, cy + icon_r * 2 + 24,
                             cell_w - 32, cell_h - icon_r * 2 - 36, desc,
                             font_name="Segoe UI", font_size=11,
                             font_color=(80, 85, 95), alignment=PP_ALIGN_LEFT)
    else:
        raise ValueError(f"Unknown icon list style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 28. create_testimonial_slide
# ============================================================

def create_testimonial_slide(
    slide_number: int,
    testimonials: list[dict],
    style: str = "cards",
) -> dict:
    """Create a testimonial/review slide.

    Args:
        slide_number: 1-based slide index.
        testimonials: List of {"quote": "...", "author": "...", "company": "..."}.
        style: "cards", "single_large", "minimal".

    Returns:
        Dict with slide_number and style.
    """
    prs, slide = _get_slide(slide_number)
    sw, sh = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight

    n = len(testimonials)
    if n == 0:
        return {"slide_number": slide_number, "style": style}

    accent_colors = [
        (0, 120, 215), (0, 170, 140), (242, 150, 0),
        (200, 50, 80), (100, 80, 200),
    ]

    if style == "cards":
        _set_solid_bg(slide, (245, 247, 252))

        _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.08,
                     "What Our Clients Say",
                     font_name="Segoe UI Semibold", font_size=28,
                     font_color=(33, 37, 41), alignment=PP_ALIGN_CENTER)

        margin = sw * 0.06
        gap = 18
        card_w = (sw - 2 * margin - gap * (n - 1)) / n
        card_h = sh * 0.7
        card_y = sh * 0.16

        for i, test in enumerate(testimonials):
            cx = margin + i * (card_w + gap)
            c = accent_colors[i % len(accent_colors)]

            # Card
            card = _add_shape(slide, MSO_SHAPE_ROUNDED_RECTANGLE,
                              cx, card_y, card_w, card_h,
                              fill_rgb=(255, 255, 255))
            _add_shadow(card, blur=6, offset_x=2, offset_y=2, transparency=0.78)

            # Top accent
            _add_shape(slide, MSO_SHAPE_RECTANGLE, cx, card_y, card_w, 5,
                       fill_rgb=c)

            # Large quote mark
            _add_textbox(slide, cx + 14, card_y + 12, 50, 50, "\u201C",
                         font_name="Georgia", font_size=60,
                         font_color=c, alignment=PP_ALIGN_LEFT)

            # Quote text
            _add_textbox(slide, cx + 16, card_y + 65, card_w - 32, card_h * 0.45,
                         test.get("quote", ""),
                         font_name="Segoe UI", font_size=12,
                         font_color=(60, 65, 75), italic=True,
                         alignment=PP_ALIGN_LEFT)

            # Divider
            div_y = card_y + card_h * 0.7
            div = slide.Shapes.AddLine(cx + 16, div_y, cx + card_w * 0.4, div_y)
            div.Line.ForeColor.RGB = rgb(220, 225, 235)
            div.Line.Weight = 1

            # Author avatar circle
            av_r = 18
            _add_shape(slide, MSO_SHAPE_OVAL,
                       cx + 16, div_y + 12, av_r * 2, av_r * 2,
                       fill_rgb=c)
            initials = "".join(w[0].upper() for w in test.get("author", "?").split()[:2])
            _add_textbox(slide, cx + 16, div_y + 12, av_r * 2, av_r * 2, initials,
                         font_name="Segoe UI", font_size=12,
                         font_color=(255, 255, 255), bold=True,
                         alignment=PP_ALIGN_CENTER,
                         vertical_anchor=MSO_ANCHOR_MIDDLE)

            # Author name & company
            _add_textbox(slide, cx + av_r * 2 + 24, div_y + 12,
                         card_w - av_r * 2 - 40, 18,
                         test.get("author", ""),
                         font_name="Segoe UI Semibold", font_size=11,
                         font_color=(33, 37, 41), alignment=PP_ALIGN_LEFT)
            _add_textbox(slide, cx + av_r * 2 + 24, div_y + 30,
                         card_w - av_r * 2 - 40, 16,
                         test.get("company", ""),
                         font_name="Segoe UI", font_size=10,
                         font_color=c, alignment=PP_ALIGN_LEFT)

    elif style == "single_large":
        _set_gradient_bg(slide, (25, 30, 45), (40, 48, 70), MSO_GRADIENT_DIAGONAL_DOWN)

        test = testimonials[0]
        c = accent_colors[0]

        # Large opening quote mark
        _add_textbox(slide, sw * 0.08, sh * 0.1, 100, 100, "\u201C",
                     font_name="Georgia", font_size=120,
                     font_color=(255, 255, 255), bold=False,
                     alignment=PP_ALIGN_LEFT)

        # Quote
        _add_textbox(slide, sw * 0.12, sh * 0.3, sw * 0.76, sh * 0.35,
                     test.get("quote", ""),
                     font_name="Georgia", font_size=24,
                     font_color=(240, 242, 248), italic=True,
                     alignment=PP_ALIGN_LEFT)

        # Accent line
        line = slide.Shapes.AddLine(sw * 0.12, sh * 0.72, sw * 0.25, sh * 0.72)
        line.Line.ForeColor.RGB = rgb(*c)
        line.Line.Weight = 3

        # Author
        _add_textbox(slide, sw * 0.12, sh * 0.75, sw * 0.5, 24,
                     test.get("author", ""),
                     font_name="Segoe UI Semibold", font_size=18,
                     font_color=(255, 255, 255), alignment=PP_ALIGN_LEFT)

        # Company
        _add_textbox(slide, sw * 0.12, sh * 0.8, sw * 0.5, 20,
                     test.get("company", ""),
                     font_name="Segoe UI", font_size=14,
                     font_color=c, alignment=PP_ALIGN_LEFT)

        # Decorative circles
        _add_shape(slide, MSO_SHAPE_OVAL, sw * 0.8, sh * 0.6, 200, 200,
                   fill_rgb=c, fill_transparency=0.9)
        _add_shape(slide, MSO_SHAPE_OVAL, sw * 0.85, -80, 250, 250,
                   fill_rgb=(255, 255, 255), fill_transparency=0.95)

    elif style == "minimal":
        _set_solid_bg(slide, (255, 255, 255))

        _add_textbox(slide, sw * 0.06, sh * 0.04, sw * 0.88, sh * 0.08,
                     "Testimonials",
                     font_name="Segoe UI Light", font_size=28,
                     font_color=(50, 55, 65), alignment=PP_ALIGN_LEFT)

        item_h = min(sh * 0.22, (sh * 0.78) / n)
        start_y = sh * 0.15

        for i, test in enumerate(testimonials):
            y = start_y + i * item_h
            c = accent_colors[i % len(accent_colors)]

            # Left accent bar
            _add_shape(slide, MSO_SHAPE_RECTANGLE,
                       sw * 0.08, y, 4, item_h - 10, fill_rgb=c)

            # Quote
            _add_textbox(slide, sw * 0.11, y + 4, sw * 0.7, item_h * 0.55,
                         f"\u201C{test.get('quote', '')}\u201D",
                         font_name="Segoe UI", font_size=13,
                         font_color=(50, 55, 65), italic=True,
                         alignment=PP_ALIGN_LEFT)

            # Author - Company
            author_text = test.get("author", "")
            company = test.get("company", "")
            if company:
                author_text = f"{author_text}, {company}"
            _add_textbox(slide, sw * 0.11, y + item_h * 0.6, sw * 0.6, 20,
                         author_text,
                         font_name="Segoe UI Semibold", font_size=11,
                         font_color=c, alignment=PP_ALIGN_LEFT)
    else:
        raise ValueError(f"Unknown testimonial style '{style}'")

    return {"slide_number": slide_number, "style": style}


# ============================================================
# 29. apply_consistent_branding
# ============================================================

def apply_consistent_branding(
    primary_color: tuple[int, int, int],
    secondary_color: tuple[int, int, int],
    accent_color: tuple[int, int, int],
    font_title: str = "Segoe UI",
    font_body: str = "Segoe UI",
) -> dict:
    """Apply consistent branding across ALL slides in the presentation.

    Args:
        primary_color: RGB tuple for primary brand color.
        secondary_color: RGB tuple for secondary brand color.
        accent_color: RGB tuple for accent color.
        font_title: Font name for titles.
        font_body: Font name for body text.

    Returns:
        Dict with number of slides affected and branding details.
    """
    app = _get_app()
    prs = app.ActivePresentation
    total = prs.Slides.Count

    # Apply to slide master
    master = prs.SlideMaster

    # Set master fonts
    try:
        title_style = master.TextStyles(1)
        for level in range(1, 10):
            try:
                title_style.Levels(level).Font.Name = font_title
                title_style.Levels(level).Font.Color.RGB = rgb(*primary_color)
            except Exception:
                break
    except Exception:
        pass

    try:
        body_style = master.TextStyles(2)
        for level in range(1, 10):
            try:
                body_style.Levels(level).Font.Name = font_body
            except Exception:
                break
    except Exception:
        pass

    # Apply theme colors
    try:
        theme = prs.SlideMaster.Theme
        color_obj = theme.ThemeColorScheme
        color_obj(4).RGB = rgb(*primary_color)    # Text2 / dark
        color_obj(5).RGB = rgb(*primary_color)    # Accent1
        color_obj(6).RGB = rgb(*secondary_color)  # Accent2
        color_obj(7).RGB = rgb(*accent_color)     # Accent3
    except Exception:
        pass

    # Apply branding to each slide
    for si in range(1, total + 1):
        slide = prs.Slides(si)

        # Add a thin accent bar at the bottom of each slide
        sw = prs.PageSetup.SlideWidth
        sh = prs.PageSetup.SlideHeight

        # Bottom accent bar
        bar = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 0, sh - 5, sw, 5)
        bar.Fill.Solid()
        bar.Fill.ForeColor.RGB = rgb(*primary_color)
        bar.Line.Visible = False

        # Small accent dot in bottom-right corner
        dot = slide.Shapes.AddShape(MSO_SHAPE_OVAL, sw - 22, sh - 22, 12, 12)
        dot.Fill.Solid()
        dot.Fill.ForeColor.RGB = rgb(*accent_color)
        dot.Line.Visible = False

        # Apply font to all text boxes on the slide
        for shape_idx in range(1, slide.Shapes.Count + 1):
            try:
                shape = slide.Shapes(shape_idx)
                if shape.HasTextFrame:
                    tf = shape.TextFrame
                    tr = tf.TextRange
                    # Check if it looks like a title (large font size)
                    try:
                        if tr.Font.Size and tr.Font.Size >= 24:
                            tr.Font.Name = font_title
                        else:
                            tr.Font.Name = font_body
                    except Exception:
                        tr.Font.Name = font_body
            except Exception:
                continue

    return {
        "slides_affected": total,
        "primary_color": list(primary_color),
        "secondary_color": list(secondary_color),
        "accent_color": list(accent_color),
        "font_title": font_title,
        "font_body": font_body,
    }
