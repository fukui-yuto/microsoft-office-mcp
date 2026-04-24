"""key_message_with_evidence layout — prominent message + supporting evidence."""

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
    KeyMessageContent,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = KeyMessageContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))

    y = body_top()

    # Key message — large
    elements.append(accent_bar(y, width=60.0))
    elements.append(TextElement(
        region=Region(
            left=MARGIN_H_PT,
            top=y + 10,
            width=CONTENT_WIDTH_PT,
            height=50,
        ),
        text=c.message,
        style=TextStyle.SUBHEAD,
    ))

    # Evidence items
    evidence_y = y + 80.0
    n = len(c.evidence)
    ev_w = (CONTENT_WIDTH_PT - GAP_PT * (n - 1)) / n

    for i, ev in enumerate(c.evidence):
        ex = MARGIN_H_PT + i * (ev_w + GAP_PT)
        elements.append(RectElement(
            region=Region(left=ex, top=evidence_y, width=ev_w, height=body_area_height() - 100),
            color=ColorToken.MONO_200,
        ))
        elements.append(TextElement(
            region=Region(left=ex + 10, top=evidence_y + 15, width=ev_w - 20, height=body_area_height() - 130),
            text=ev,
            style=TextStyle.BODY,
        ))

    return SlidePlacement(elements=elements)
