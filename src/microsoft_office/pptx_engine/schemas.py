"""Pydantic models defining the JSON contracts between pipeline layers.

Layer 1 output → PresentationPlan
Layer 2 output → Placement
Layer 3 input  → Placement
"""

from __future__ import annotations

from enum import Enum
from typing import Literal

from pydantic import BaseModel, Field, field_validator


# ============================================================
# Enums
# ============================================================

class NarrativeArc(str, Enum):
    SCQA = "SCQA"         # Situation-Complication-Question-Answer
    PREP = "PREP"         # Point-Reason-Example-Point
    PYRAMID = "Pyramid"   # Minto Pyramid Principle


class LayoutName(str, Enum):
    EXECUTIVE_SUMMARY = "executive_summary"
    BIG_NUMBER = "big_number"
    BEFORE_AFTER = "before_after"
    MATRIX_2X2 = "matrix_2x2"
    WATERFALL = "waterfall"
    PROCESS_FLOW = "process_flow"
    QUOTE = "quote"
    COMPARISON_TABLE = "comparison_table"
    KEY_MESSAGE = "key_message_with_evidence"
    THREE_COLUMN = "three_column"
    TIMELINE = "timeline"


class TextStyle(str, Enum):
    HEADLINE = "headline"
    SUBHEAD = "subhead"
    BODY = "body"
    CAPTION = "caption"


class ColorToken(str, Enum):
    MONO_900 = "mono_900"
    MONO_600 = "mono_600"
    MONO_400 = "mono_400"
    MONO_200 = "mono_200"
    MONO_000 = "mono_000"
    ACCENT_PRIMARY = "accent_primary"
    ACCENT_SUB = "accent_sub"


# ============================================================
# Layer 1 — Plan Schema
# ============================================================

class SlidePlan(BaseModel):
    """Single slide in the presentation plan."""
    layout: LayoutName
    headline: str = Field(
        ...,
        min_length=1,
        max_length=40,
        description="Conclusion sentence (not a noun phrase). Must contain a verb or judgement.",
    )
    content: dict = Field(
        default_factory=dict,
        description="Layout-specific content (validated per layout).",
    )


class PresentationPlan(BaseModel):
    """Layer 1 output — the full presentation plan."""
    narrative_arc: NarrativeArc
    slides: list[SlidePlan] = Field(..., min_length=1, max_length=20)


# ============================================================
# Layer 2 — Placement Schema
# ============================================================

class Region(BaseModel):
    """Bounding box in points: [left, top, width, height]."""
    left: float
    top: float
    width: float
    height: float


class TextElement(BaseModel):
    type: Literal["textbox"] = "textbox"
    region: Region
    text: str
    style: TextStyle


class RectElement(BaseModel):
    type: Literal["rect"] = "rect"
    region: Region
    color: ColorToken


class ImageElement(BaseModel):
    type: Literal["image"] = "image"
    region: Region
    path: str


PlacementElement = TextElement | RectElement | ImageElement


class SlidePlacement(BaseModel):
    """All elements for a single slide."""
    elements: list[PlacementElement]


class Placement(BaseModel):
    """Layer 2 output — positioned elements for every slide."""
    slides: list[SlidePlacement] = Field(..., min_length=1)


# ============================================================
# Layout-specific content schemas
# ============================================================

class BigNumberContent(BaseModel):
    value: str = Field(..., description="The main number, e.g. '3.2億'")
    unit: str = Field(default="", description="Unit label, e.g. '円'")
    delta: str = Field(default="", description="Change indicator, e.g. '+24%'")
    context: str = Field(default="", description="Context line, e.g. '前年同期比'")


class EvidencePoint(BaseModel):
    text: str
    evidence: str = ""


class ExecutiveSummaryContent(BaseModel):
    points: list[EvidencePoint] = Field(..., min_length=1, max_length=3)


class BeforeAfterSide(BaseModel):
    label: str
    points: list[str] = Field(..., min_length=1, max_length=5)


class BeforeAfterContent(BaseModel):
    before: BeforeAfterSide
    after: BeforeAfterSide


class Matrix2x2Content(BaseModel):
    x_label: str
    y_label: str
    quadrants: list[str] = Field(..., min_length=4, max_length=4)


class WaterfallItem(BaseModel):
    label: str
    value: float
    is_total: bool = False


class WaterfallContent(BaseModel):
    items: list[WaterfallItem] = Field(..., min_length=2, max_length=8)
    unit: str = ""


class ProcessStep(BaseModel):
    label: str
    description: str = ""


class ProcessFlowContent(BaseModel):
    steps: list[ProcessStep] = Field(..., min_length=2, max_length=5)


class QuoteContent(BaseModel):
    quote: str
    attribution: str = ""


class ComparisonRow(BaseModel):
    label: str
    values: list[str] = Field(..., min_length=2, max_length=4)


class ComparisonTableContent(BaseModel):
    headers: list[str] = Field(..., min_length=2, max_length=4)
    rows: list[ComparisonRow] = Field(..., min_length=1, max_length=5)


class KeyMessageContent(BaseModel):
    message: str
    evidence: list[str] = Field(..., min_length=1, max_length=3)


class ColumnItem(BaseModel):
    heading: str
    body: str


class ThreeColumnContent(BaseModel):
    columns: list[ColumnItem] = Field(..., min_length=3, max_length=3)


class TimelineEntry(BaseModel):
    date: str
    event: str


class TimelineContent(BaseModel):
    entries: list[TimelineEntry] = Field(..., min_length=2, max_length=5)


# Registry: layout name → content model
CONTENT_MODELS: dict[LayoutName, type[BaseModel]] = {
    LayoutName.BIG_NUMBER: BigNumberContent,
    LayoutName.EXECUTIVE_SUMMARY: ExecutiveSummaryContent,
    LayoutName.BEFORE_AFTER: BeforeAfterContent,
    LayoutName.MATRIX_2X2: Matrix2x2Content,
    LayoutName.WATERFALL: WaterfallContent,
    LayoutName.PROCESS_FLOW: ProcessFlowContent,
    LayoutName.QUOTE: QuoteContent,
    LayoutName.COMPARISON_TABLE: ComparisonTableContent,
    LayoutName.KEY_MESSAGE: KeyMessageContent,
    LayoutName.THREE_COLUMN: ThreeColumnContent,
    LayoutName.TIMELINE: TimelineContent,
}
