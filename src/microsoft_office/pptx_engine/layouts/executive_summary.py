"""executive_summary layout — headline + 1-3 evidence-backed points."""

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
    ExecutiveSummaryContent,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = ExecutiveSummaryContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))
    elements.append(divider_element(body_top() - GAP_PT / 2))

    n = len(c.points)
    row_height = body_area_height() / n
    y = body_top()

    for point in c.points:
        # Point text
        elements.append(TextElement(
            region=Region(
                left=MARGIN_H_PT,
                top=y,
                width=CONTENT_WIDTH_PT,
                height=row_height * 0.5,
            ),
            text=point.text,
            style=TextStyle.BODY,
        ))
        # Evidence (if present)
        if point.evidence:
            elements.append(TextElement(
                region=Region(
                    left=MARGIN_H_PT,
                    top=y + row_height * 0.5,
                    width=CONTENT_WIDTH_PT,
                    height=row_height * 0.35,
                ),
                text=point.evidence,
                style=TextStyle.CAPTION,
            ))
        y += row_height

    return SlidePlacement(elements=elements)
