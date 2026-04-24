"""big_number layout — single dominant metric with context."""

from __future__ import annotations

from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
    MARGIN_H_PT,
    SLIDE_HEIGHT_PT,
)
from microsoft_office.pptx_engine.layouts.base import (
    accent_bar,
    body_top,
    headline_element,
)
from microsoft_office.pptx_engine.schemas import (
    BigNumberContent,
    ColorToken,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = BigNumberContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))

    # Big number — centered, dominant
    center_y = body_top() + 30.0
    value_text = f"{c.value}{c.unit}" if c.unit else c.value
    elements.append(TextElement(
        region=Region(
            left=MARGIN_H_PT,
            top=center_y,
            width=CONTENT_WIDTH_PT,
            height=80.0,
        ),
        text=value_text,
        style=TextStyle.HEADLINE,
    ))

    # Accent bar under the number
    elements.append(accent_bar(center_y + 85.0, width=80.0))

    # Delta
    if c.delta:
        elements.append(TextElement(
            region=Region(
                left=MARGIN_H_PT,
                top=center_y + 100.0,
                width=CONTENT_WIDTH_PT,
                height=30.0,
            ),
            text=c.delta,
            style=TextStyle.SUBHEAD,
        ))

    # Context
    if c.context:
        elements.append(TextElement(
            region=Region(
                left=MARGIN_H_PT,
                top=center_y + 140.0,
                width=CONTENT_WIDTH_PT,
                height=24.0,
            ),
            text=c.context,
            style=TextStyle.CAPTION,
        ))

    return SlidePlacement(elements=elements)
