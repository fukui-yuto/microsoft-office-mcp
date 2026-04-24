"""Validation utilities for safe PowerPoint generation.

Enforces WCAG AA contrast, color count limits, accent restrictions,
font size ranges, and bullet count caps.
"""

from __future__ import annotations

from microsoft_office.design_tokens import (
    FONT_SIZE_MAX,
    FONT_SIZE_MIN,
    MAX_ACCENT_PER_SLIDE,
    MAX_BULLETS,
    MAX_COLORS_PER_SLIDE,
    MAX_COMPARISON_ITEMS,
)


# ---------------------------------------------------------------------------
# WCAG 2.0 contrast ratio
# ---------------------------------------------------------------------------

def _relative_luminance(rgb: tuple[int, int, int]) -> float:
    """Compute WCAG 2.0 relative luminance from sRGB (0-255)."""
    channels = []
    for c in rgb:
        s = c / 255.0
        channels.append(s / 12.92 if s <= 0.04045 else ((s + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def contrast_ratio(
    fg: tuple[int, int, int],
    bg: tuple[int, int, int],
) -> float:
    """Return WCAG 2.0 contrast ratio between two sRGB colors.

    Result is always >= 1.0.  WCAG AA requires >= 4.5 for normal text.
    """
    l1 = _relative_luminance(fg)
    l2 = _relative_luminance(bg)
    lighter = max(l1, l2)
    darker = min(l1, l2)
    return (lighter + 0.05) / (darker + 0.05)


def check_contrast(
    fg: tuple[int, int, int],
    bg: tuple[int, int, int],
    min_ratio: float = 4.5,
) -> None:
    """Raise ValueError if contrast ratio is below min_ratio (WCAG AA)."""
    ratio = contrast_ratio(fg, bg)
    if ratio < min_ratio:
        raise ValueError(
            f"Contrast ratio {ratio:.2f} between {fg} and {bg} "
            f"is below WCAG AA minimum {min_ratio}"
        )


# ---------------------------------------------------------------------------
# Bullet / item count
# ---------------------------------------------------------------------------

def validate_bullets(bullets: list[str]) -> None:
    """Raise ValueError if bullet count exceeds MAX_BULLETS."""
    if len(bullets) > MAX_BULLETS:
        raise ValueError(
            f"Too many bullets: {len(bullets)} (max {MAX_BULLETS}). "
            f"Split across multiple slides instead."
        )


def validate_comparison_items(items: list[dict]) -> None:
    """Raise ValueError if comparison item count exceeds MAX_COMPARISON_ITEMS."""
    if len(items) > MAX_COMPARISON_ITEMS:
        raise ValueError(
            f"Too many comparison items: {len(items)} (max {MAX_COMPARISON_ITEMS}). "
            f"Simplify or split across slides."
        )


# ---------------------------------------------------------------------------
# Font size range
# ---------------------------------------------------------------------------

def validate_font_size(size_pt: float) -> None:
    """Raise ValueError if font size is outside the allowed range."""
    if not (FONT_SIZE_MIN <= size_pt <= FONT_SIZE_MAX):
        raise ValueError(
            f"Font size {size_pt}pt is outside allowed range "
            f"[{FONT_SIZE_MIN}, {FONT_SIZE_MAX}]"
        )


# ---------------------------------------------------------------------------
# Color usage on a single slide
# ---------------------------------------------------------------------------

def validate_color_usage(
    colors_used: list[tuple[int, int, int]],
    accent_count: int,
) -> list[str]:
    """Return list of warning/error messages for color rule violations.

    Rules:
      - Unique colors on one slide must be <= MAX_COLORS_PER_SLIDE
      - Accent color usage must be <= MAX_ACCENT_PER_SLIDE
    """
    errors: list[str] = []
    unique = set(colors_used)
    if len(unique) > MAX_COLORS_PER_SLIDE:
        errors.append(
            f"Too many colors on slide: {len(unique)} "
            f"(max {MAX_COLORS_PER_SLIDE})"
        )
    if accent_count > MAX_ACCENT_PER_SLIDE:
        errors.append(
            f"Accent color used {accent_count} times "
            f"(max {MAX_ACCENT_PER_SLIDE} per slide)"
        )
    return errors
