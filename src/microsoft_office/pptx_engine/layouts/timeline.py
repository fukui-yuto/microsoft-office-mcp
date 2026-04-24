"""timeline layout — horizontal timeline with date markers."""

from __future__ import annotations

from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
    GAP_PT,
    MARGIN_H_PT,
)
from microsoft_office.pptx_engine.layouts.base import (
    accent_bar,
    body_area_height,
    body_top,
    headline_element,
)
from microsoft_office.pptx_engine.schemas import (
    ColorToken,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
    TimelineContent,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = TimelineContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))

    y = body_top()
    h = body_area_height()
    n = len(c.entries)

    # Horizontal track line
    track_y = y + h * 0.4
    elements.append(RectElement(
        region=Region(left=MARGIN_H_PT, top=track_y, width=CONTENT_WIDTH_PT, height=2),
        color=ColorToken.MONO_400,
    ))

    # Entries
    step_w = CONTENT_WIDTH_PT / n
    for i, entry in enumerate(c.entries):
        cx = MARGIN_H_PT + i * step_w + step_w / 2

        # Dot on track
        dot_size = 10.0
        elements.append(RectElement(
            region=Region(left=cx - dot_size / 2, top=track_y - dot_size / 2 + 1, width=dot_size, height=dot_size),
            color=ColorToken.ACCENT_PRIMARY,
        ))

        # Date above
        elements.append(TextElement(
            region=Region(left=cx - step_w / 2 + 4, top=track_y - 35, width=step_w - 8, height=24),
            text=entry.date,
            style=TextStyle.CAPTION,
        ))

        # Event below
        elements.append(TextElement(
            region=Region(left=cx - step_w / 2 + 4, top=track_y + 18, width=step_w - 8, height=60),
            text=entry.event,
            style=TextStyle.BODY,
        ))

    return SlidePlacement(elements=elements)
