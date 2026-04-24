"""Tests for the pptx_engine 3-layer pipeline."""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

import pytest

from microsoft_office.pptx_engine.design_system import (
    ACCENT_PRIMARY,
    COLOR_TOKENS,
    FONT_SIZE_MAX,
    FONT_SIZE_MIN,
    GAP_PT,
    MARGIN_H_PT,
    MARGIN_V_PT,
    MAX_COLORS_PER_SLIDE,
    MONO,
    SLIDE_HEIGHT_PT,
    SLIDE_WIDTH_PT,
    TYPO,
    resolve_color,
)
from microsoft_office.pptx_engine.schemas import (
    BigNumberContent,
    ColorToken,
    ExecutiveSummaryContent,
    LayoutName,
    NarrativeArc,
    Placement,
    PresentationPlan,
    Region,
    SlidePlan,
    SlidePlacement,
    TextElement,
    TextStyle,
)
from microsoft_office.pptx_engine.validator import (
    assert_no_bold,
    validate_bullet_count,
    validate_font_size,
    validate_headline,
    validate_plan,
    validate_placement,
)
from microsoft_office.pptx_engine.planner import parse_plan, get_plan_schema
from microsoft_office.pptx_engine.designer import design
from microsoft_office.pptx_engine.renderer import render


# ============================================================
# Design System tests
# ============================================================

class TestDesignSystem:
    def test_mono_has_5_steps(self):
        assert len(MONO) == 5

    def test_color_tokens_count(self):
        assert len(COLOR_TOKENS) == 7  # 5 mono + 2 accent

    def test_typo_styles(self):
        for style_name in ("headline", "subhead", "body", "caption"):
            s = TYPO[style_name]
            assert s["bold"] is False, f"{style_name} must not be bold"
            assert FONT_SIZE_MIN <= s["size_pt"] <= FONT_SIZE_MAX

    def test_resolve_color_valid(self):
        assert resolve_color("mono_900") == MONO["900"]
        assert resolve_color("accent_primary") == ACCENT_PRIMARY

    def test_resolve_color_invalid(self):
        with pytest.raises(ValueError, match="Unknown color token"):
            resolve_color("neon_pink")

    def test_grid_margins(self):
        assert MARGIN_H_PT > 0
        assert MARGIN_V_PT > 0
        assert 2 * MARGIN_H_PT < SLIDE_WIDTH_PT
        assert 2 * MARGIN_V_PT < SLIDE_HEIGHT_PT

    def test_gap(self):
        assert GAP_PT == 18.0 * 1.5


# ============================================================
# Validator tests
# ============================================================

class TestHeadlineValidation:
    def test_good_headline(self):
        validate_headline("売上は前年比20%成長した")

    def test_good_headline_english(self):
        validate_headline("Revenue increased by 20%")

    def test_noun_phrase_rejected(self):
        with pytest.raises(ValueError, match="noun phrase"):
            validate_headline("売上推移")

    def test_too_long(self):
        with pytest.raises(ValueError, match="exceeds 40"):
            validate_headline("あ" * 41)

    def test_empty(self):
        with pytest.raises(ValueError, match="empty"):
            validate_headline("")

    def test_ends_with_desu(self):
        validate_headline("売上は好調です")

    def test_ends_with_naru(self):
        validate_headline("成長率は20%になる")


class TestBoldProhibition:
    def test_bold_raises(self):
        with pytest.raises(ValueError, match="Bold is prohibited"):
            assert_no_bold(True, context="test")

    def test_not_bold_ok(self):
        assert_no_bold(False)


class TestFontSizeValidation:
    def test_in_range(self):
        validate_font_size(18.0)

    def test_below_min(self):
        with pytest.raises(ValueError):
            validate_font_size(8.0)

    def test_above_max(self):
        with pytest.raises(ValueError):
            validate_font_size(48.0)


class TestBulletCount:
    def test_within_limit(self):
        validate_bullet_count(5)

    def test_exceeds(self):
        with pytest.raises(ValueError, match="Too many bullets"):
            validate_bullet_count(6)


class TestPlacementValidation:
    def test_valid_placement(self):
        p = Placement(slides=[SlidePlacement(elements=[
            TextElement(
                region=Region(left=0, top=0, width=100, height=50),
                text="test",
                style=TextStyle.HEADLINE,
            ),
        ])])
        errors = validate_placement(p)
        assert errors == []


# ============================================================
# Schema tests
# ============================================================

class TestSchemas:
    def test_big_number_content(self):
        c = BigNumberContent(value="3.2億", unit="円", delta="+24%", context="前年同期比")
        assert c.value == "3.2億"

    def test_executive_summary_max_3(self):
        with pytest.raises(Exception):
            ExecutiveSummaryContent.model_validate({
                "points": [
                    {"text": "a"}, {"text": "b"}, {"text": "c"}, {"text": "d"},
                ]
            })

    def test_plan_schema_json(self):
        schema = get_plan_schema()
        assert "properties" in schema
        assert "slides" in schema["properties"]


# ============================================================
# Planner tests (Layer 1)
# ============================================================

class TestPlanner:
    def test_parse_valid_plan(self):
        raw = {
            "narrative_arc": "SCQA",
            "slides": [
                {
                    "layout": "big_number",
                    "headline": "売上は前年比20%成長した",
                    "content": {
                        "value": "3.2億",
                        "unit": "円",
                        "delta": "+24%",
                        "context": "前年同期比",
                    },
                },
            ],
        }
        plan = parse_plan(raw)
        assert plan.narrative_arc == NarrativeArc.SCQA
        assert len(plan.slides) == 1

    def test_parse_json_string(self):
        raw = json.dumps({
            "narrative_arc": "PREP",
            "slides": [
                {
                    "layout": "big_number",
                    "headline": "利益率が大幅に改善した",
                    "content": {"value": "18.5", "unit": "%"},
                },
            ],
        })
        plan = parse_plan(raw)
        assert plan.narrative_arc == NarrativeArc.PREP

    def test_parse_rejects_bad_headline(self):
        raw = {
            "narrative_arc": "SCQA",
            "slides": [
                {
                    "layout": "big_number",
                    "headline": "売上推移",
                    "content": {"value": "100"},
                },
            ],
        }
        with pytest.raises(ValueError, match="noun phrase"):
            parse_plan(raw)


# ============================================================
# Designer tests (Layer 2)
# ============================================================

class TestDesigner:
    def _make_plan(self, layout: str, headline: str, content: dict) -> PresentationPlan:
        return PresentationPlan(
            narrative_arc=NarrativeArc.SCQA,
            slides=[SlidePlan(layout=layout, headline=headline, content=content)],
        )

    def test_big_number_design(self):
        plan = self._make_plan(
            "big_number",
            "売上は過去最高を更新した",
            {"value": "3.2億", "unit": "円", "delta": "+24%", "context": "前年同期比"},
        )
        placement = design(plan)
        assert len(placement.slides) == 1
        assert len(placement.slides[0].elements) > 0

    def test_executive_summary_design(self):
        plan = self._make_plan(
            "executive_summary",
            "3つの施策が成長を牽引している",
            {
                "points": [
                    {"text": "DX推進", "evidence": "売上20%増"},
                    {"text": "海外展開", "evidence": "3カ国進出済"},
                ],
            },
        )
        placement = design(plan)
        assert len(placement.slides) == 1


# ============================================================
# Renderer tests (Layer 3)
# ============================================================

class TestRenderer:
    def test_render_big_number(self):
        plan = PresentationPlan(
            narrative_arc=NarrativeArc.SCQA,
            slides=[SlidePlan(
                layout="big_number",
                headline="売上は過去最高を更新した",
                content={"value": "3.2億", "unit": "円", "delta": "+24%", "context": "前年同期比"},
            )],
        )
        placement = design(plan)
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output = render(placement, f.name)
        path = Path(output)
        assert path.exists()
        assert path.stat().st_size > 0

    def test_render_two_slide_deck(self):
        plan = PresentationPlan(
            narrative_arc=NarrativeArc.PYRAMID,
            slides=[
                SlidePlan(
                    layout="big_number",
                    headline="売上は過去最高を更新した",
                    content={"value": "120億", "delta": "+15%"},
                ),
                SlidePlan(
                    layout="executive_summary",
                    headline="3つの要因が成長を支えている",
                    content={
                        "points": [
                            {"text": "新規顧客獲得が加速した", "evidence": "YoY +40%"},
                            {"text": "既存顧客の単価が上昇した", "evidence": "ARPU +12%"},
                        ],
                    },
                ),
            ],
        )
        placement = design(plan)
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output = render(placement, f.name)
        path = Path(output)
        assert path.exists()


# ============================================================
# End-to-end test (all 3 layers)
# ============================================================

class TestEndToEnd:
    def test_full_pipeline(self):
        """Plan → Design → Render for a multi-layout deck."""
        raw = {
            "narrative_arc": "SCQA",
            "slides": [
                {
                    "layout": "big_number",
                    "headline": "売上は120億円に到達した",
                    "content": {"value": "120", "unit": "億円", "delta": "+15%", "context": "前年度比"},
                },
                {
                    "layout": "executive_summary",
                    "headline": "3つの施策が成長をドライブする",
                    "content": {
                        "points": [
                            {"text": "プロダクト拡充が奏功した", "evidence": "新規ARR 30億円"},
                            {"text": "APAC展開が本格化した", "evidence": "3拠点を開設"},
                        ],
                    },
                },
                {
                    "layout": "three_column",
                    "headline": "注力領域は3つに絞られる",
                    "content": {
                        "columns": [
                            {"heading": "プロダクト", "body": "AI機能統合とAPI連携"},
                            {"heading": "海外", "body": "シンガポール・台湾・タイ"},
                            {"heading": "組織", "body": "エンジニア30名採用"},
                        ],
                    },
                },
            ],
        }

        plan = parse_plan(raw)
        placement = design(plan)
        with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as f:
            output = render(placement, f.name)
        assert Path(output).exists()
