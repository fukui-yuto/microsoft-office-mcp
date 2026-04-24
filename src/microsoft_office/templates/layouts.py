"""Slide layout definitions.

Each layout specifies named regions (bounding boxes in points)
that the safe COM layer places content into.  No design logic here —
only spatial geometry.

Coordinates assume standard 16:9 slide (960 x 540 pt).
"""

from __future__ import annotations

from enum import Enum


SLIDE_WIDTH = 960.0
SLIDE_HEIGHT = 540.0

# Margins
MARGIN_H = 60.0   # horizontal margin (left/right)
MARGIN_V = 40.0    # vertical margin (top/bottom)

CONTENT_WIDTH = SLIDE_WIDTH - 2 * MARGIN_H   # 840
CONTENT_HEIGHT = SLIDE_HEIGHT - 2 * MARGIN_V  # 460


class SlideLayout(str, Enum):
    """Available constrained slide layouts."""
    TITLE = "title"
    SECTION_HEADER = "section_header"
    BULLET = "bullet"
    TWO_COLUMN = "two_column"
    IMAGE = "image"
    COMPARISON = "comparison"


# Region = (left, top, width, height) in points
Region = tuple[float, float, float, float]

LAYOUT_REGIONS: dict[SlideLayout, dict[str, Region]] = {
    SlideLayout.TITLE: {
        "title": (MARGIN_H, 170.0, CONTENT_WIDTH, 80.0),
        "subtitle": (MARGIN_H, 260.0, CONTENT_WIDTH, 36.0),
        "accent_bar": (MARGIN_H, 155.0, 80.0, 4.0),
    },
    SlideLayout.SECTION_HEADER: {
        "title": (MARGIN_H, 200.0, CONTENT_WIDTH, 60.0),
        "accent_bar": (MARGIN_H, 190.0, 60.0, 3.0),
    },
    SlideLayout.BULLET: {
        "title": (MARGIN_H, MARGIN_V, CONTENT_WIDTH, 50.0),
        "divider": (MARGIN_H, 95.0, CONTENT_WIDTH, 1.0),
        "content": (MARGIN_H, 110.0, CONTENT_WIDTH, 390.0),
    },
    SlideLayout.TWO_COLUMN: {
        "title": (MARGIN_H, MARGIN_V, CONTENT_WIDTH, 50.0),
        "divider": (MARGIN_H, 95.0, CONTENT_WIDTH, 1.0),
        "left": (MARGIN_H, 110.0, 395.0, 390.0),
        "right": (MARGIN_H + 445.0, 110.0, 395.0, 390.0),
    },
    SlideLayout.IMAGE: {
        "title": (MARGIN_H, MARGIN_V, CONTENT_WIDTH, 50.0),
        "image": (MARGIN_H, 100.0, CONTENT_WIDTH, 350.0),
        "caption": (MARGIN_H, 460.0, CONTENT_WIDTH, 30.0),
    },
    SlideLayout.COMPARISON: {
        "title": (MARGIN_H, MARGIN_V, CONTENT_WIDTH, 50.0),
        "divider": (MARGIN_H, 95.0, CONTENT_WIDTH, 1.0),
        # Items laid out dynamically (2-4 equal-width columns)
        "items_area": (MARGIN_H, 110.0, CONTENT_WIDTH, 390.0),
    },
}


def get_regions(layout: SlideLayout) -> dict[str, Region]:
    """Return region definitions for a layout."""
    return LAYOUT_REGIONS[layout]


def split_items_area(
    area: Region,
    n_items: int,
    gap: float = 20.0,
) -> list[Region]:
    """Split an area into n equal-width columns with gaps.

    Used by comparison layout to position 2-4 item cards.
    """
    left, top, total_width, height = area
    col_width = (total_width - gap * (n_items - 1)) / n_items
    return [
        (left + i * (col_width + gap), top, col_width, height)
        for i in range(n_items)
    ]
