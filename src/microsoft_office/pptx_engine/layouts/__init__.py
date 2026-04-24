"""Layout registry — maps LayoutName to render functions."""

from __future__ import annotations

from typing import TYPE_CHECKING

from microsoft_office.pptx_engine.schemas import LayoutName, SlidePlacement

if TYPE_CHECKING:
    from microsoft_office.pptx_engine.schemas import SlidePlan

from microsoft_office.pptx_engine.layouts import (
    big_number,
    executive_summary,
    before_after,
    matrix_2x2,
    waterfall,
    process_flow,
    quote,
    comparison_table,
    key_message,
    three_column,
    timeline,
)

_REGISTRY: dict[LayoutName, object] = {
    LayoutName.BIG_NUMBER: big_number,
    LayoutName.EXECUTIVE_SUMMARY: executive_summary,
    LayoutName.BEFORE_AFTER: before_after,
    LayoutName.MATRIX_2X2: matrix_2x2,
    LayoutName.WATERFALL: waterfall,
    LayoutName.PROCESS_FLOW: process_flow,
    LayoutName.QUOTE: quote,
    LayoutName.COMPARISON_TABLE: comparison_table,
    LayoutName.KEY_MESSAGE: key_message,
    LayoutName.THREE_COLUMN: three_column,
    LayoutName.TIMELINE: timeline,
}


def get_layout_module(name: LayoutName):
    """Return the layout module for a given layout name."""
    if name not in _REGISTRY:
        raise ValueError(f"Unknown layout: {name.value}")
    return _REGISTRY[name]
