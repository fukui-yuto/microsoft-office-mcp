"""Tests for validators module."""

import pytest

from microsoft_office.validators import (
    check_contrast,
    contrast_ratio,
    validate_bullets,
    validate_color_usage,
    validate_comparison_items,
    validate_font_size,
)


class TestContrastRatio:
    def test_black_on_white(self):
        ratio = contrast_ratio((0, 0, 0), (255, 255, 255))
        assert ratio == pytest.approx(21.0, abs=0.1)

    def test_white_on_white(self):
        ratio = contrast_ratio((255, 255, 255), (255, 255, 255))
        assert ratio == pytest.approx(1.0, abs=0.01)

    def test_symmetric(self):
        r1 = contrast_ratio((37, 99, 235), (255, 255, 255))
        r2 = contrast_ratio((255, 255, 255), (37, 99, 235))
        assert r1 == pytest.approx(r2, abs=0.01)

    def test_business_theme_text_passes(self):
        """BUSINESS theme text_primary on bg_primary must pass WCAG AA."""
        ratio = contrast_ratio((23, 23, 23), (255, 255, 255))
        assert ratio >= 4.5

    def test_minimal_theme_text_passes(self):
        ratio = contrast_ratio((28, 28, 28), (255, 255, 255))
        assert ratio >= 4.5

    def test_warm_theme_text_passes(self):
        ratio = contrast_ratio((41, 37, 36), (255, 252, 248))
        assert ratio >= 4.5


class TestCheckContrast:
    def test_passes(self):
        check_contrast((0, 0, 0), (255, 255, 255))

    def test_fails(self):
        with pytest.raises(ValueError, match="below WCAG AA"):
            check_contrast((200, 200, 200), (255, 255, 255))


class TestValidateBullets:
    def test_within_limit(self):
        validate_bullets(["a", "b", "c", "d", "e"])

    def test_exceeds_limit(self):
        with pytest.raises(ValueError, match="Too many bullets"):
            validate_bullets(["a"] * 6)

    def test_empty(self):
        validate_bullets([])


class TestValidateComparisonItems:
    def test_within_limit(self):
        validate_comparison_items([{}, {}, {}, {}])

    def test_exceeds_limit(self):
        with pytest.raises(ValueError, match="Too many comparison"):
            validate_comparison_items([{}] * 5)


class TestValidateFontSize:
    def test_within_range(self):
        validate_font_size(18.0)

    def test_too_small(self):
        with pytest.raises(ValueError, match="outside allowed range"):
            validate_font_size(8.0)

    def test_too_large(self):
        with pytest.raises(ValueError, match="outside allowed range"):
            validate_font_size(48.0)

    def test_boundaries(self):
        validate_font_size(11.0)
        validate_font_size(36.0)


class TestValidateColorUsage:
    def test_within_limits(self):
        colors = [(0, 0, 0), (255, 255, 255), (37, 99, 235)]
        errors = validate_color_usage(colors, accent_count=1)
        assert errors == []

    def test_too_many_colors(self):
        colors = [(i, i, i) for i in range(7)]
        errors = validate_color_usage(colors, accent_count=1)
        assert any("Too many colors" in e for e in errors)

    def test_too_many_accents(self):
        colors = [(0, 0, 0)]
        errors = validate_color_usage(colors, accent_count=2)
        assert any("Accent color" in e for e in errors)
