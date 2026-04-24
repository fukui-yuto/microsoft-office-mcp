"""before_after layout — side-by-side comparison."""

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
    BeforeAfterContent,
    ColorToken,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = BeforeAfterContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))
    elements.append(divider_element(body_top() - GAP_PT / 2))

    col_width = (CONTENT_WIDTH_PT - GAP_PT) / 2
    y = body_top()
    h = body_area_height()

    for i, side in enumerate([c.before, c.after]):
        x = MARGIN_H_PT + i * (col_width + GAP_PT)

        # Column background
        elements.append(RectElement(
            region=Region(left=x, top=y, width=col_width, height=h),
            color=ColorToken.MONO_200 if i == 0 else ColorToken.MONO_000,
        ))

        # Label
        elements.append(TextElement(
            region=Region(left=x + 12, top=y + 10, width=col_width - 24, height=30),
            text=side.label,
            style=TextStyle.SUBHEAD,
        ))

        # Points
        point_y = y + 50.0
        for pt_text in side.points:
            elements.append(TextElement(
                region=Region(left=x + 12, top=point_y, width=col_width - 24, height=24),
                text=f"• {pt_text}",
                style=TextStyle.BODY,
            ))
            point_y += 30.0

    return SlidePlacement(elements=elements)
