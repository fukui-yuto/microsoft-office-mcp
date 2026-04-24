"""comparison_table layout — structured table with headers."""

from __future__ import annotations

from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
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
    ComparisonTableContent,
    RectElement,
    Region,
    SlidePlacement,
    TextElement,
    TextStyle,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = ComparisonTableContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))

    y = body_top()
    h = body_area_height()
    n_cols = len(c.headers)
    n_rows = len(c.rows) + 1  # +1 for header row
    col_w = CONTENT_WIDTH_PT / n_cols
    row_h = min(h / n_rows, 50.0)

    # Header row
    for j, header in enumerate(c.headers):
        x = MARGIN_H_PT + j * col_w
        elements.append(RectElement(
            region=Region(left=x, top=y, width=col_w, height=row_h),
            color=ColorToken.MONO_200,
        ))
        elements.append(TextElement(
            region=Region(left=x + 8, top=y + 4, width=col_w - 16, height=row_h - 8),
            text=header,
            style=TextStyle.SUBHEAD,
        ))

    # Data rows
    for i, row in enumerate(c.rows):
        ry = y + row_h * (i + 1)
        # Label column
        elements.append(TextElement(
            region=Region(left=MARGIN_H_PT + 8, top=ry + 4, width=col_w - 16, height=row_h - 8),
            text=row.label,
            style=TextStyle.BODY,
        ))
        # Value columns
        for j, val in enumerate(row.values):
            x = MARGIN_H_PT + (j + 1) * col_w if len(row.values) < n_cols else MARGIN_H_PT + j * col_w
            # Adjust: if label is separate, values start at col 1
            vx = MARGIN_H_PT + (j + 1) * col_w
            if j + 1 < n_cols:
                elements.append(TextElement(
                    region=Region(left=vx + 8, top=ry + 4, width=col_w - 16, height=row_h - 8),
                    text=val,
                    style=TextStyle.BODY,
                ))

        # Row divider
        elements.append(RectElement(
            region=Region(left=MARGIN_H_PT, top=ry + row_h - 1, width=CONTENT_WIDTH_PT, height=1),
            color=ColorToken.MONO_200,
        ))

    return SlidePlacement(elements=elements)
