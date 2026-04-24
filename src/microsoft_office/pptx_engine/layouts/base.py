"""Base helpers shared by all layout modules.

Every layout module must expose:
    render(headline: str, content: dict) -> SlidePlacement
"""

from __future__ import annotations

from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
    GAP_PT,
    MARGIN_H_PT,
    MARGIN_V_PT,
)
from microsoft_office.pptx_engine.schemas import (
    ColorToken,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def headline_element(headline: str) -> TextElement:
    """Standard headline textbox at the top of the slide."""
    return TextElement(
        region=Region(
            left=MARGIN_H_PT,
            top=MARGIN_V_PT,
            width=CONTENT_WIDTH_PT,
            height=48.0,
        ),
        text=headline,
        style=TextStyle.HEADLINE,
    )


def divider_element(top: float) -> RectElement:
    """Thin horizontal divider line."""
    return RectElement(
        region=Region(
            left=MARGIN_H_PT,
            top=top,
            width=CONTENT_WIDTH_PT,
            height=1.5,
        ),
        color=ColorToken.MONO_200,
    )


def body_top() -> float:
    """Y position where body content starts (below headline + gap)."""
    return MARGIN_V_PT + 48.0 + GAP_PT


def body_area_height() -> float:
    """Available height for body content."""
    from microsoft_office.pptx_engine.design_system import SLIDE_HEIGHT_PT
    return SLIDE_HEIGHT_PT - body_top() - MARGIN_V_PT


def accent_bar(top: float, width: float = 50.0) -> RectElement:
    """Small accent color bar."""
    return RectElement(
        region=Region(
            left=MARGIN_H_PT,
            top=top,
            width=width,
            height=3.0,
        ),
        color=ColorToken.ACCENT_PRIMARY,
    )
