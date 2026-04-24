"""waterfall layout — waterfall chart generated via matplotlib."""

from __future__ import annotations

from microsoft_office.pptx_engine.charts import render_waterfall_chart
from microsoft_office.pptx_engine.design_system import (
    CONTENT_WIDTH_PT,
    MARGIN_H_PT,
)
from microsoft_office.pptx_engine.layouts.base import (
    body_area_height,
    body_top,
    headline_element,
)
from microsoft_office.pptx_engine.schemas import (
    ImageElement,
    Region,
    SlidePlacement,
    WaterfallContent,
)


def render(headline: str, content: dict) -> SlidePlacement:
    c = WaterfallContent.model_validate(content)
    elements = []

    elements.append(headline_element(headline))

    labels = [item.label for item in c.items]
    values = [item.value for item in c.items]
    is_totals = [item.is_total for item in c.items]

    chart_path = render_waterfall_chart(labels, values, is_totals, unit=c.unit)

    y = body_top()
    h = body_area_height()
    elements.append(ImageElement(
        region=Region(left=MARGIN_H_PT, top=y, width=CONTENT_WIDTH_PT, height=h),
        path=chart_path,
    ))

    return SlidePlacement(elements=elements)
