"""Design system constants — the single source of truth for all visual decisions.

References:
  - Gene Zelazny "Say It With Charts"
  - IBM Carbon Design System typography scale
  - Apple Keynote event slides (2020+)

Nothing in this file is configurable at runtime.
"""

from __future__ import annotations

from pptx.util import Pt, Emu

# ============================================================
# Slide dimensions (16:9)
# ============================================================
SLIDE_WIDTH_PT = 960.0
SLIDE_HEIGHT_PT = 540.0

SLIDE_WIDTH = Emu(int(SLIDE_WIDTH_PT * 12700))
SLIDE_HEIGHT = Emu(int(SLIDE_HEIGHT_PT * 12700))

# ============================================================
# GRID — 12-column with 7% margin
# ============================================================
MARGIN_RATIO = 0.07
MARGIN_H_PT = SLIDE_WIDTH_PT * MARGIN_RATIO   # ~67.2 pt
MARGIN_V_PT = SLIDE_HEIGHT_PT * MARGIN_RATIO   # ~37.8 pt
CONTENT_WIDTH_PT = SLIDE_WIDTH_PT - 2 * MARGIN_H_PT   # ~825.6 pt
CONTENT_HEIGHT_PT = SLIDE_HEIGHT_PT - 2 * MARGIN_V_PT  # ~464.4 pt

GRID_COLUMNS = 12
COLUMN_WIDTH_PT = CONTENT_WIDTH_PT / GRID_COLUMNS  # ~68.8 pt

# Element gap = body size * 1.5
GAP_PT = 18.0 * 1.5  # 27 pt

# ============================================================
# TYPOGRAPHY
# ============================================================
# Font families — Japanese-friendly, available on Windows
FONT_HEADLINE = "BIZ UDPGothic"
FONT_SUBHEAD = "BIZ UDPGothic"
FONT_BODY = "BIZ UDPGothic"
FONT_CAPTION = "BIZ UDPGothic"

# Sizes and weights
#   Weight: "Medium" for headline, "Regular" for everything else
#   Bold is PROHIBITED — use Medium weight instead
TYPO = {
    "headline": {"font": FONT_HEADLINE, "size": Pt(40), "size_pt": 40.0, "bold": False, "line_spacing": 1.4},
    "subhead":  {"font": FONT_SUBHEAD,  "size": Pt(24), "size_pt": 24.0, "bold": False, "line_spacing": 1.4},
    "body":     {"font": FONT_BODY,     "size": Pt(18), "size_pt": 18.0, "bold": False, "line_spacing": 1.4},
    "caption":  {"font": FONT_CAPTION,  "size": Pt(12), "size_pt": 12.0, "bold": False, "line_spacing": 1.4},
}

FONT_SIZE_MIN = 12.0
FONT_SIZE_MAX = 40.0

# ============================================================
# COLOR PALETTE
# ============================================================
# Mono scale (5 steps): near-black → white
MONO = {
    "900": (10, 10, 10),      # #0A0A0A  — primary text
    "600": (74, 74, 74),      # #4A4A4A  — secondary text
    "400": (156, 163, 175),   # #9CA3AF  — tertiary / border
    "200": (229, 231, 235),   # #E5E7EB  — divider / surface
    "000": (255, 255, 255),   # #FFFFFF  — background
}

# Accent (max 1 per slide)
ACCENT_PRIMARY = (30, 58, 138)    # #1E3A8A  — deep blue
ACCENT_SUB = (15, 118, 110)       # #0F766E  — teal (charts only)

# Named color tokens for schema reference
COLOR_TOKENS: dict[str, tuple[int, int, int]] = {
    "mono_900": MONO["900"],
    "mono_600": MONO["600"],
    "mono_400": MONO["400"],
    "mono_200": MONO["200"],
    "mono_000": MONO["000"],
    "accent_primary": ACCENT_PRIMARY,
    "accent_sub": ACCENT_SUB,
}

MAX_COLORS_PER_SLIDE = 4   # mono_900 + mono_600 + mono_400 + 1 accent (or mono_200)

# ============================================================
# CHART defaults (matplotlib)
# ============================================================
CHART_DPI = 200
CHART_BG_COLOR = "#FFFFFF"
CHART_FONT_FAMILY = "BIZ UDPGothic"
CHART_FONT_SIZE = 10
CHART_ACCENT_HEX = "#1E3A8A"
CHART_SUB_HEX = "#0F766E"
CHART_GREY_HEX = "#9CA3AF"

# ============================================================
# Helpers
# ============================================================

def resolve_color(token: str) -> tuple[int, int, int]:
    """Resolve a color token name to RGB tuple."""
    if token not in COLOR_TOKENS:
        valid = ", ".join(COLOR_TOKENS.keys())
        raise ValueError(f"Unknown color token '{token}'. Valid: {valid}")
    return COLOR_TOKENS[token]


def pt_to_emu(pt: float) -> int:
    """Convert points to EMU."""
    return int(pt * 12700)
