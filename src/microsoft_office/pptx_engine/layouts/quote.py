"""quote layout — centered quotation with attribution."""

from __future__ import annotations

from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
    MARGIN_H_PT,
    SLIDE_HEIGHT_PT,
)
from microsoft_office.pptx_engine.layouts.base import accent_bar
from microsoft_office.pptx_engine.schemas import (
    QuoteContent,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = QuoteContent.model_validate(content)
    elements = []

    # No standard headline for quote — the quote IS the content
    # But we still include the headline as a small label
    elements.append(TextElement(
        region=Region(
            left=MARGIN_H_PT,
            top=SLIDE_HEIGHT_PT * 0.15,
            width=CONTENT_WIDTH_PT,
            height=24,
        ),
        text=headline,
        style=TextStyle.CAPTION,
    ))

    # Accent bar
    elements.append(accent_bar(SLIDE_HEIGHT_PT * 0.25, width=40.0))

    # Quote text — large, centered
    quote_inset = CONTENT_WIDTH_PT * 0.1
    elements.append(TextElement(
        region=Region(
            left=MARGIN_H_PT + quote_inset,
            top=SLIDE_HEIGHT_PT * 0.30,
            width=CONTENT_WIDTH_PT - 2 * quote_inset,
            height=120,
        ),
        text=f"\u201C{c.quote}\u201D",
        style=TextStyle.SUBHEAD,
    ))

    # Attribution
    if c.attribution:
        elements.append(TextElement(
            region=Region(
                left=MARGIN_H_PT + quote_inset,
                top=SLIDE_HEIGHT_PT * 0.30 + 130,
                width=CONTENT_WIDTH_PT - 2 * quote_inset,
                height=24,
            ),
            text=f"— {c.attribution}",
            style=TextStyle.CAPTION,
        ))

    return SlidePlacement(elements=elements)
