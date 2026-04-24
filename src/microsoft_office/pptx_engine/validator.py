"""Validation rules enforced across the pipeline.

Every prohibition is a hard error — no warnings, no overrides.
"""

from __future__ import annotations

import re

from microsoft_office.pptx_engine.design_system import (
    FONT_SIZE_MAX,
    FONT_SIZE_MIN,
    MAX_COLORS_PER_SLIDE,
)
from microsoft_office.pptx_engine.schemas import (
    CONTENT_MODELS,
    ColorToken,
    LayoutName,
    Placement,
    PresentationPlan,
    SlidePlan,
    SlidePlacement,
)


# ============================================================
# Headline validation
# ============================================================

# Simple heuristic: a headline ending in a noun without a verb-like ending
# is likely taigen-dome (体言止め).  Japanese verb endings include:
# る/た/だ/す/く/ぐ/む/ぶ/ぬ/う + ます/です/ある/いる/なる/する/できる etc.
_VERB_PATTERN = re.compile(
    r"(る|た|だ|す|く|ぐ|む|ぶ|ぬ|う|ます|です|ある|いる|なる|する|できる|"
    r"した|された|している|される|させる|ない|べきだ|見込む|目指す|"
    r"\b(?:needed|achieved|increased|decreased|improved|enabled|"
    r"is|are|was|were|will|should|must|can|have|has|had|"
    r"reach|expect|grow|decline|drive|support|deliver)\b)",
    re.IGNORECASE,
)


def validate_headline(headline: str) -> None:
    """Ensure headline is a sentence with a verb, not a noun phrase."""
    if len(headline) > 40:
        raise ValueError(
            f"Headline exceeds 40 chars ({len(headline)}): '{headline}'"
        )
    if not headline.strip():
        raise ValueError("Headline must not be empty")

    # Check last meaningful token for verb ending
    text = headline.rstrip("。．.、")
    if not _VERB_PATTERN.search(text):
        raise ValueError(
            f"Headline appears to be a noun phrase (体言止め): '{headline}'. "
            f"Rewrite to include a verb or judgement."
        )


# ============================================================
# Slide-level content validation
# ============================================================

def validate_slide_content(slide: SlidePlan) -> None:
    """Validate layout-specific content against its schema."""
    model_cls = CONTENT_MODELS.get(slide.layout)
    if model_cls is None:
        raise ValueError(f"No content schema for layout '{slide.layout.value}'")
    # This will raise ValidationError if content doesn't match
    model_cls.model_validate(slide.content)


# ============================================================
# Plan validation (Layer 1)
# ============================================================

def validate_plan(plan: PresentationPlan) -> list[str]:
    """Validate a full presentation plan. Returns list of errors."""
    errors: list[str] = []
    for i, slide in enumerate(plan.slides):
        try:
            validate_headline(slide.headline)
        except ValueError as e:
            errors.append(f"Slide {i + 1}: {e}")
        try:
            validate_slide_content(slide)
        except Exception as e:
            errors.append(f"Slide {i + 1} content: {e}")
    return errors


# ============================================================
# Placement validation (Layer 2 output)
# ============================================================

def validate_placement(placement: Placement) -> list[str]:
    """Validate placement for design rule violations."""
    errors: list[str] = []
    for i, slide in enumerate(placement.slides):
        slide_errors = _validate_slide_placement(slide)
        for err in slide_errors:
            errors.append(f"Slide {i + 1}: {err}")
    return errors


def _validate_slide_placement(slide: SlidePlacement) -> list[str]:
    """Check a single slide's placement for violations."""
    errors: list[str] = []

    # Collect color tokens used
    accent_count = 0
    unique_colors: set[str] = set()
    for el in slide.elements:
        if el.type == "rect":
            unique_colors.add(el.color.value)
            if el.color in (ColorToken.ACCENT_PRIMARY, ColorToken.ACCENT_SUB):
                accent_count += 1
        if el.type == "textbox":
            # Text colors are derived from style, count mono + accent
            if el.style.value == "headline":
                unique_colors.add("mono_900")
            elif el.style.value == "subhead":
                unique_colors.add("mono_900")
            elif el.style.value == "body":
                unique_colors.add("mono_600")
            elif el.style.value == "caption":
                unique_colors.add("mono_400")

    if len(unique_colors) > MAX_COLORS_PER_SLIDE:
        errors.append(
            f"Too many colors: {len(unique_colors)} (max {MAX_COLORS_PER_SLIDE})"
        )

    if accent_count > 1:
        errors.append(
            f"Accent color used {accent_count} times (max 1 per slide)"
        )

    return errors


# ============================================================
# Prohibited features (checked at render time)
# ============================================================

PROHIBITED_FEATURES = [
    "gradient",
    "shadow",
    "3d",
    "wordart",
    "animation",
    "bold",
]


def assert_no_bold(is_bold: bool, context: str = "") -> None:
    """Bold is prohibited — Medium weight is used instead."""
    if is_bold:
        raise ValueError(
            f"Bold is prohibited. Use Medium weight. Context: {context}"
        )


def validate_font_size(size_pt: float, context: str = "") -> None:
    """Font size must be within design system range."""
    if not (FONT_SIZE_MIN <= size_pt <= FONT_SIZE_MAX):
        raise ValueError(
            f"Font size {size_pt}pt outside [{FONT_SIZE_MIN}, {FONT_SIZE_MAX}]. "
            f"Context: {context}"
        )


def validate_bullet_count(count: int, max_count: int = 5) -> None:
    """No more than 5 bullets per list."""
    if count > max_count:
        raise ValueError(
            f"Too many bullets: {count} (max {max_count}). Split across slides."
        )
