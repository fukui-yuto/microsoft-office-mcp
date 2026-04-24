"""Safe PowerPoint COM operations — design-token constrained.

All visual decisions (color, font, size) are resolved from design tokens.
No gradients, shadows, 3D effects, or arbitrary RGB values.
"""

from __future__ import annotations

from microsoft_office.com_utils import get_or_create_app, rgb
from microsoft_office.design_tokens import (
    FontToken,
    SizeToken,
    ThemeName,
    TokenRole,
    get_color,
    get_font,
    get_size,
)
from microsoft_office.templates.layouts import (
    SLIDE_HEIGHT,
    SLIDE_WIDTH,
    SlideLayout,
    get_regions,
    split_items_area,
)
from microsoft_office.validators import (
    check_contrast,
    validate_bullets,
    validate_comparison_items,
)

# COM constants (subset — no gradient/shadow/3D)
MSO_SHAPE_RECTANGLE = 1
PP_ALIGN_LEFT = 1
PP_ALIGN_CENTER = 2
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
MSO_ANCHOR_TOP = 1
MSO_ANCHOR_MIDDLE = 3


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_app():
    return get_or_create_app("PowerPoint.Application")


def _get_prs():
    return _get_app().ActivePresentation


def _add_blank_slide() -> tuple:
    """Append a blank slide and return (presentation, slide, slide_number)."""
    prs = _get_prs()
    # Layout index 7 = blank in most Office versions
    slide = prs.Slides.Add(prs.Slides.Count + 1, 7)
    return prs, slide, slide.SlideIndex


def _set_bg(slide, color: tuple[int, int, int]):
    """Set solid background — no gradients."""
    slide.FollowMasterBackground = False
    slide.Background.Fill.Solid()
    slide.Background.Fill.ForeColor.RGB = rgb(*color)


def _place_textbox(
    slide,
    region: tuple[float, float, float, float],
    text: str,
    font_token: FontToken,
    size_token: SizeToken,
    color_role: TokenRole,
    theme: ThemeName,
    bold: bool = False,
    alignment: int = PP_ALIGN_LEFT,
    vertical_anchor: int | None = None,
) -> object:
    """Place a text box in a region using only design tokens."""
    left, top, width, height = region
    shape = slide.Shapes.AddTextbox(
        MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height,
    )
    tf = shape.TextFrame
    tf.WordWrap = True
    if vertical_anchor is not None:
        tf.VerticalAnchor = vertical_anchor

    tr = tf.TextRange
    tr.Text = text
    tr.ParagraphFormat.Alignment = alignment

    font = tr.Font
    font.Name = get_font(font_token)
    font.Size = get_size(size_token)
    font.Bold = bold

    text_color = get_color(theme, color_role)
    bg_color = get_color(theme, TokenRole.BG_PRIMARY)
    check_contrast(text_color, bg_color)
    font.Color.RGB = rgb(*text_color)

    shape.Fill.Background()
    shape.Line.Visible = False
    return shape


def _place_rect(
    slide,
    region: tuple[float, float, float, float],
    fill_role: TokenRole,
    theme: ThemeName,
) -> object:
    """Place a solid rectangle — no gradient, no shadow."""
    left, top, width, height = region
    shape = slide.Shapes.AddShape(MSO_SHAPE_RECTANGLE, left, top, width, height)
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = rgb(*get_color(theme, fill_role))
    shape.Line.Visible = False
    return shape


def _place_divider(slide, region, theme: ThemeName):
    """Place a thin horizontal line as a divider."""
    _place_rect(slide, region, TokenRole.BORDER, theme)


def _place_bullets_text(
    slide,
    region: tuple[float, float, float, float],
    bullets: list[str],
    theme: ThemeName,
) -> object:
    """Place bullet list as a single text box with paragraph breaks."""
    left, top, width, height = region
    shape = slide.Shapes.AddTextbox(
        MSO_TEXT_ORIENTATION_HORIZONTAL, left, top, width, height,
    )
    tf = shape.TextFrame
    tf.WordWrap = True

    text_color = get_color(theme, TokenRole.TEXT_PRIMARY)
    bg_color = get_color(theme, TokenRole.BG_PRIMARY)
    check_contrast(text_color, bg_color)

    for i, bullet in enumerate(bullets):
        if i == 0:
            para = tf.TextRange
        else:
            para = tf.TextRange.InsertAfter("\r")
        para.InsertAfter(bullet)

    # Format all text
    tr = tf.TextRange
    tr.Font.Name = get_font(FontToken.BODY)
    tr.Font.Size = get_size(SizeToken.MD)
    tr.Font.Color.RGB = rgb(*text_color)
    tr.ParagraphFormat.Alignment = PP_ALIGN_LEFT

    # Set bullet style
    for i in range(1, tf.TextRange.Paragraphs().Count + 1):
        para = tf.TextRange.Paragraphs(i)
        para.ParagraphFormat.Bullet.Type = 1  # ppBulletUnnumbered
        para.ParagraphFormat.SpaceAfter = 8

    shape.Fill.Background()
    shape.Line.Visible = False
    return shape


# ---------------------------------------------------------------------------
# Public API — one function per layout
# ---------------------------------------------------------------------------

def add_title_slide(
    title: str,
    subtitle: str | None = None,
    theme: ThemeName = ThemeName.BUSINESS,
) -> dict:
    """Create a title slide using design tokens only."""
    prs, slide, idx = _add_blank_slide()
    regions = get_regions(SlideLayout.TITLE)

    _set_bg(slide, get_color(theme, TokenRole.BG_PRIMARY))

    # Accent bar
    _place_rect(slide, regions["accent_bar"], TokenRole.ACCENT, theme)

    # Title
    _place_textbox(
        slide, regions["title"], title,
        FontToken.TITLE, SizeToken.XL, TokenRole.TEXT_PRIMARY, theme,
        bold=True, alignment=PP_ALIGN_LEFT,
    )

    # Subtitle
    if subtitle:
        _place_textbox(
            slide, regions["subtitle"], subtitle,
            FontToken.BODY, SizeToken.MD, TokenRole.TEXT_SECONDARY, theme,
            alignment=PP_ALIGN_LEFT,
        )

    return {"slide_number": idx, "layout": "title", "theme": theme.value}


def add_section_header(
    title: str,
    accent: bool = False,
    theme: ThemeName = ThemeName.BUSINESS,
) -> dict:
    """Create a section header slide."""
    prs, slide, idx = _add_blank_slide()
    regions = get_regions(SlideLayout.SECTION_HEADER)

    bg_role = TokenRole.BG_SECONDARY if accent else TokenRole.BG_PRIMARY
    _set_bg(slide, get_color(theme, bg_role))

    # Accent bar
    _place_rect(slide, regions["accent_bar"], TokenRole.ACCENT, theme)

    # Title
    _place_textbox(
        slide, regions["title"], title,
        FontToken.TITLE, SizeToken.LG, TokenRole.TEXT_PRIMARY, theme,
        bold=True, alignment=PP_ALIGN_LEFT,
    )

    return {"slide_number": idx, "layout": "section_header", "theme": theme.value}


def add_bullet_slide(
    title: str,
    bullets: list[str],
    theme: ThemeName = ThemeName.BUSINESS,
) -> dict:
    """Create a bullet point slide (max 5 bullets)."""
    validate_bullets(bullets)

    prs, slide, idx = _add_blank_slide()
    regions = get_regions(SlideLayout.BULLET)

    _set_bg(slide, get_color(theme, TokenRole.BG_PRIMARY))

    # Title
    _place_textbox(
        slide, regions["title"], title,
        FontToken.TITLE, SizeToken.LG, TokenRole.TEXT_PRIMARY, theme,
        bold=True,
    )

    # Divider
    _place_divider(slide, regions["divider"], theme)

    # Bullets
    _place_bullets_text(slide, regions["content"], bullets, theme)

    return {
        "slide_number": idx,
        "layout": "bullet",
        "theme": theme.value,
        "bullet_count": len(bullets),
    }


def add_two_column_slide(
    title: str,
    left_items: list[str],
    right_items: list[str],
    theme: ThemeName = ThemeName.BUSINESS,
) -> dict:
    """Create a two-column slide (max 5 items per column)."""
    validate_bullets(left_items)
    validate_bullets(right_items)

    prs, slide, idx = _add_blank_slide()
    regions = get_regions(SlideLayout.TWO_COLUMN)

    _set_bg(slide, get_color(theme, TokenRole.BG_PRIMARY))

    _place_textbox(
        slide, regions["title"], title,
        FontToken.TITLE, SizeToken.LG, TokenRole.TEXT_PRIMARY, theme,
        bold=True,
    )
    _place_divider(slide, regions["divider"], theme)
    _place_bullets_text(slide, regions["left"], left_items, theme)
    _place_bullets_text(slide, regions["right"], right_items, theme)

    return {
        "slide_number": idx,
        "layout": "two_column",
        "theme": theme.value,
    }


def add_image_slide(
    title: str,
    image_path: str,
    caption: str | None = None,
    theme: ThemeName = ThemeName.BUSINESS,
) -> dict:
    """Create an image slide with optional caption."""
    import os
    image_path = os.path.abspath(image_path)
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"Image not found: {image_path}")

    prs, slide, idx = _add_blank_slide()
    regions = get_regions(SlideLayout.IMAGE)

    _set_bg(slide, get_color(theme, TokenRole.BG_PRIMARY))

    _place_textbox(
        slide, regions["title"], title,
        FontToken.TITLE, SizeToken.LG, TokenRole.TEXT_PRIMARY, theme,
        bold=True,
    )

    # Image
    left, top, width, height = regions["image"]
    slide.Shapes.AddPicture(
        image_path, LinkToFile=False, SaveWithDocument=True,
        Left=left, Top=top, Width=width, Height=height,
    )

    if caption:
        _place_textbox(
            slide, regions["caption"], caption,
            FontToken.BODY, SizeToken.SM, TokenRole.TEXT_SECONDARY, theme,
            alignment=PP_ALIGN_CENTER,
        )

    return {"slide_number": idx, "layout": "image", "theme": theme.value}


def add_comparison_slide(
    title: str,
    items: list[dict],
    theme: ThemeName = ThemeName.BUSINESS,
) -> dict:
    """Create a comparison slide with 2-4 item cards.

    Each item dict: {"heading": str, "points": list[str]}
    """
    validate_comparison_items(items)

    prs, slide, idx = _add_blank_slide()
    regions = get_regions(SlideLayout.COMPARISON)

    _set_bg(slide, get_color(theme, TokenRole.BG_PRIMARY))

    _place_textbox(
        slide, regions["title"], title,
        FontToken.TITLE, SizeToken.LG, TokenRole.TEXT_PRIMARY, theme,
        bold=True,
    )
    _place_divider(slide, regions["divider"], theme)

    # Split items area into columns
    columns = split_items_area(regions["items_area"], len(items))

    for col_region, item in zip(columns, items):
        left, top, width, height = col_region
        heading = item.get("heading", "")
        points = item.get("points", [])
        validate_bullets(points)

        # Card background
        _place_rect(
            slide, (left, top, width, height), TokenRole.SURFACE, theme,
        )

        # Card heading
        heading_region = (left + 12, top + 12, width - 24, 30)
        _place_textbox(
            slide, heading_region, heading,
            FontToken.TITLE, SizeToken.MD, TokenRole.TEXT_PRIMARY, theme,
            bold=True,
        )

        # Card bullet points
        if points:
            points_region = (left + 12, top + 50, width - 24, height - 62)
            _place_bullets_text(slide, points_region, points, theme)

    return {
        "slide_number": idx,
        "layout": "comparison",
        "theme": theme.value,
        "item_count": len(items),
    }
