"""process_flow layout — horizontal step sequence."""

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
    divider_element,
    headline_element,
)
from microsoft_office.pptx_engine.schemas import (
    ColorToken,
    ProcessFlowContent,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = ProcessFlowContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))
    elements.append(divider_element(body_top() - GAP_PT / 2))

    n = len(c.steps)
    y = body_top() + 20.0
    step_w = (CONTENT_WIDTH_PT - GAP_PT * (n - 1)) / n
    h = body_area_height() - 20.0

    for i, step in enumerate(c.steps):
        x = MARGIN_H_PT + i * (step_w + GAP_PT)

        # Step card background
        elements.append(RectElement(
            region=Region(left=x, top=y, width=step_w, height=h),
            color=ColorToken.MONO_200,
        ))

        # Step number accent bar
        elements.append(accent_bar(y + 8.0, width=step_w * 0.3))

        # Step label
        elements.append(TextElement(
            region=Region(left=x + 10, top=y + 20, width=step_w - 20, height=28),
            text=f"{i + 1}. {step.label}",
            style=TextStyle.SUBHEAD,
        ))

        # Step description
        if step.description:
            elements.append(TextElement(
                region=Region(left=x + 10, top=y + 55, width=step_w - 20, height=h - 70),
                text=step.description,
                style=TextStyle.BODY,
            ))

        # Arrow between steps
        if i < n - 1:
            arrow_x = x + step_w + GAP_PT * 0.2
            elements.append(TextElement(
                region=Region(left=arrow_x, top=y + h / 2 - 12, width=GAP_PT * 0.6, height=24),
                text="→",
                style=TextStyle.SUBHEAD,
            ))

    return SlidePlacement(elements=elements)
