"""Tests for templates/layouts module."""

from microsoft_office.templates.layouts import (
    CONTENT_HEIGHT,
    CONTENT_WIDTH,
    LAYOUT_REGIONS,
    MARGIN_H,
    MARGIN_V,
    SLIDE_HEIGHT,
    SLIDE_WIDTH,
    SlideLayout,
    get_regions,
    split_items_area,
)


class TestSlideConstants:
    def test_dimensions(self):
        assert SLIDE_WIDTH == 960.0
        assert SLIDE_HEIGHT == 540.0

    def test_content_area(self):
        assert CONTENT_WIDTH == SLIDE_WIDTH - 2 * MARGIN_H
        assert CONTENT_HEIGHT == SLIDE_HEIGHT - 2 * MARGIN_V


class TestLayoutRegions:
    def test_all_layouts_defined(self):
        for layout in SlideLayout:
            assert layout in LAYOUT_REGIONS

    def test_regions_within_slide(self):
        for layout in SlideLayout:
            for name, (left, top, width, height) in LAYOUT_REGIONS[layout].items():
                assert left >= 0, f"{layout.value}.{name} left={left}"
                assert top >= 0, f"{layout.value}.{name} top={top}"
                assert left + width <= SLIDE_WIDTH + 1, (
                    f"{layout.value}.{name} right edge {left + width} > {SLIDE_WIDTH}"
                )
                assert top + height <= SLIDE_HEIGHT + 1, (
                    f"{layout.value}.{name} bottom edge {top + height} > {SLIDE_HEIGHT}"
                )

    def test_title_layout_has_required_regions(self):
        regions = get_regions(SlideLayout.TITLE)
        assert "title" in regions
        assert "subtitle" in regions
        assert "accent_bar" in regions

    def test_bullet_layout_has_required_regions(self):
        regions = get_regions(SlideLayout.BULLET)
        assert "title" in regions
        assert "content" in regions

    def test_two_column_layout_has_required_regions(self):
        regions = get_regions(SlideLayout.TWO_COLUMN)
        assert "left" in regions
        assert "right" in regions

    def test_comparison_layout_has_items_area(self):
        regions = get_regions(SlideLayout.COMPARISON)
        assert "items_area" in regions


class TestSplitItemsArea:
    def test_two_columns(self):
        area = (60.0, 110.0, 840.0, 390.0)
        cols = split_items_area(area, 2, gap=20.0)
        assert len(cols) == 2
        assert cols[0][0] == 60.0  # first col starts at left edge
        # Widths should be equal
        assert cols[0][2] == cols[1][2]

    def test_four_columns(self):
        area = (60.0, 110.0, 840.0, 390.0)
        cols = split_items_area(area, 4, gap=20.0)
        assert len(cols) == 4
        # No overlap: each col ends before next starts
        for i in range(len(cols) - 1):
            right_edge = cols[i][0] + cols[i][2]
            next_left = cols[i + 1][0]
            assert right_edge <= next_left

    def test_columns_fill_area(self):
        area = (60.0, 110.0, 840.0, 390.0)
        cols = split_items_area(area, 3, gap=20.0)
        last_right = cols[-1][0] + cols[-1][2]
        area_right = area[0] + area[2]
        assert abs(last_right - area_right) < 1.0  # within 1pt tolerance
