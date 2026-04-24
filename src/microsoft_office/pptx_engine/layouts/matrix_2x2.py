"""matrix_2x2 layout — 2x2 quadrant diagram with axis labels."""

from __future__ import annotations

from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
    GAP_PT,
    MARGIN_H_PT,
)
from microsoft_office.pptx_engine.layouts.base import (
    body_area_height,
    body_top,
    headline_element,
)
from microsoft_office.pptx_engine.schemas import (
    ColorToken,
    Matrix2x2Content,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = Matrix2x2Content.model_validate(content)
    elements = []

    elements.append(headline_element(headline))

    y = body_top()
    h = body_area_height()
    # Reserve left strip for Y-axis label
    axis_w = 30.0
    grid_x = MARGIN_H_PT + axis_w + 8
    grid_w = CONTENT_WIDTH_PT - axis_w - 8
    cell_w = (grid_w - GAP_PT) / 2
    cell_h = (h - GAP_PT - 24) / 2  # 24 for x-axis label

    # Y-axis label (rotated text not supported well, just place vertically)
    elements.append(TextElement(
        region=Region(left=MARGIN_H_PT, top=y + h / 2 - 12, width=axis_w, height=24),
        text=c.y_label,
        style=TextStyle.CAPTION,
    ))

    # Quadrants: [top-left, top-right, bottom-left, bottom-right]
    positions = [
        (grid_x, y),
        (grid_x + cell_w + GAP_PT, y),
        (grid_x, y + cell_h + GAP_PT),
        (grid_x + cell_w + GAP_PT, y + cell_h + GAP_PT),
    ]
    bg_colors = [
        ColorToken.MONO_200,
        ColorToken.MONO_200,
        ColorToken.MONO_200,
        ColorToken.MONO_200,
    ]

    for i, ((cx, cy), q_text) in enumerate(zip(positions, c.quadrants)):
        elements.append(RectElement(
            region=Region(left=cx, top=cy, width=cell_w, height=cell_h),
            color=bg_colors[i],
        ))
        elements.append(TextElement(
            region=Region(left=cx + 10, top=cy + cell_h / 2 - 12, width=cell_w - 20, height=24),
            text=q_text,
            style=TextStyle.BODY,
        ))

    # X-axis label
    elements.append(TextElement(
        region=Region(left=grid_x, top=y + h - 20, width=grid_w, height=20),
        text=c.x_label,
        style=TextStyle.CAPTION,
    ))

    return SlidePlacement(elements=elements)
