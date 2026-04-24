"""Layer 1 — Planner: structure and validate a presentation plan.

In an MCP context, the LLM is the planner.  This module provides:
  1. The JSON schema (via Pydantic) that constrains LLM output
  2. Validation that enforces headline quality, content schemas, etc.
  3. A convenience function to parse raw JSON into a validated plan

The MCP instruction tells the LLM to produce output conforming to
PresentationPlan schema.  This module validates it server-side.
"""

from __future__ import annotations

import json
from typing import Any

from microsoft_office.pptx_engine.schemas import PresentationPlan
from microsoft_office.pptx_engine.validator import validate_plan


# System prompt fragment for LLM guidance (used in MCP instructions)
SYSTEM_PROMPT_FRAGMENT = """
You are creating a presentation plan in the style of McKinsey / BCG executive summaries.

RULES:
- Every slide MUST have a "headline" that is a CONCLUSION SENTENCE (not a title).
  The headline must contain a verb or judgement. Noun phrases (体言止め) are prohibited.
  Maximum 40 characters.

BAD headlines (prohibited):
  ✗ "売上推移"  (noun phrase)
  ✗ "市場概況"  (noun phrase)
  ✗ "次期戦略の概要"  (noun phrase)

GOOD headlines (required style):
  ✓ "売上は前年比20%成長し過去最高を更新した"
  ✓ "DX投資が営業利益率を3pt改善させている"
  ✓ "APAC進出により2027年に売上倍増を見込む"

- narrative_arc must be one of: SCQA, PREP, Pyramid
- Each slide's layout must be one of the 11 available layouts
- Content must match the layout's schema exactly
"""


def parse_plan(raw: str | dict[str, Any]) -> PresentationPlan:
    """Parse and validate a presentation plan from JSON string or dict.

    Args:
        raw: JSON string or dict conforming to PresentationPlan schema.

    Returns:
        Validated PresentationPlan.

    Raises:
        ValueError: If validation fails with detailed error messages.
    """
    if isinstance(raw, str):
        data = json.loads(raw)
    else:
        data = raw

    plan = PresentationPlan.model_validate(data)

    errors = validate_plan(plan)
    if errors:
        raise ValueError(
            "Plan validation failed:\n" + "\n".join(errors)
        )

    return plan


def get_plan_schema() -> dict:
    """Return the JSON Schema for PresentationPlan.

    Useful for including in MCP tool descriptions or
    system prompts to guide LLM output.
    """
    return PresentationPlan.model_json_schema()
