"""three_column layout — three equal columns with heading + body."""

from __future__ import annotations

from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
    GAP_PT,
    MARGIN_H_PT,
)
from microsoft_office.pptx_engine.layouts.base import (
    body_area_height,
    body_top,
    divider_element,
    headline_element,
)
from microsoft_office.pptx_engine.schemas import (
    ColorToken,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
    ThreeColumnContent,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = ThreeColumnContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))
    elements.append(divider_element(body_top() - GAP_PT / 2))

    y = body_top()
    h = body_area_height()
    col_w = (CONTENT_WIDTH_PT - GAP_PT * 2) / 3

    for i, col in enumerate(c.columns):
        x = MARGIN_H_PT + i * (col_w + GAP_PT)

        # Card background
        elements.append(RectElement(
            region=Region(left=x, top=y, width=col_w, height=h),
            color=ColorToken.MONO_200,
        ))

        # Heading
        elements.append(TextElement(
            region=Region(left=x + 12, top=y + 15, width=col_w - 24, height=28),
            text=col.heading,
            style=TextStyle.SUBHEAD,
        ))

        # Body
        elements.append(TextElement(
            region=Region(left=x + 12, top=y + 55, width=col_w - 24, height=h - 75),
            text=col.body,
            style=TextStyle.BODY,
        ))

    return SlidePlacement(elements=elements)
