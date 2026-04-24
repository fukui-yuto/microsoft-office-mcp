"""Layer 2 — Designer: convert a PresentationPlan into a Placement.

Reads each slide's layout + content, delegates to the corresponding
layout module, and assembles the full Placement.
"""

from __future__ import annotations

from microsoft_office.pptx_engine.layouts import get_layout_module
from microsoft_office.pptx_engine.schemas import (
    Placement,
    PresentationPlan,
    SlidePlacement,
)
from microsoft_office.pptx_engine.validator import validate_plan


def design(plan: PresentationPlan) -> Placement:
    """Convert a validated plan into positioned elements.

    Args:
        plan: A PresentationPlan (Layer 1 output).

    Returns:
        A Placement ready for Layer 3 rendering.

    Raises:
        ValueError: If plan validation fails.
    """
    errors = validate_plan(plan)
    if errors:
        raise ValueError(
            "Plan validation failed:\n" + "\n".join(errors)
        )

    slide_placements: list[SlidePlacement] = []

    for slide_plan in plan.slides:
        layout_mod = get_layout_module(slide_plan.layout)
        placement = layout_mod.render(slide_plan.headline, slide_plan.content)
        slide_placements.append(placement)

    return Placement(slides=slide_placements)
