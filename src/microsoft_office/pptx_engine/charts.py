"""Chart generation via matplotlib → PNG.

All charts follow the design system:
  - No axis lines / spines
  - No legend box (labels inline or data-label)
  - Data labels directly on bars/points
  - White background
  - Design system fonts and colors
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import matplotlib
matplotlib.use("Agg")  # non-interactive backend

import matplotlib.pyplot as plt
import matplotlib.font_manager as fm

from microsoft_office.pptx_engine.design_system import (
    CHART_ACCENT_HEX,
    CHART_BG_COLOR,
    CHART_DPI,
    CHART_FONT_FAMILY,
    CHART_FONT_SIZE,
    CHART_GREY_HEX,
    CHART_SUB_HEX,
)


def _setup_style():
    """Apply design system defaults to matplotlib."""
    plt.rcParams.update({
        "figure.facecolor": CHART_BG_COLOR,
        "axes.facecolor": CHART_BG_COLOR,
        "font.size": CHART_FONT_SIZE,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "axes.spines.left": False,
        "axes.spines.bottom": False,
        "axes.grid": False,
        "xtick.bottom": False,
        "ytick.left": False,
    })
    # Try to use the design system font
    available = {f.name for f in fm.fontManager.ttflist}
    if CHART_FONT_FAMILY in available:
        plt.rcParams["font.family"] = CHART_FONT_FAMILY


def render_bar_chart(
    labels: list[str],
    values: list[float],
    unit: str = "",
    highlight_index: int | None = None,
) -> str:
    """Render a horizontal bar chart and return the PNG path."""
    _setup_style()
    fig, ax = plt.subplots(figsize=(8, 0.6 * len(labels) + 1))

    colors = []
    for i in range(len(values)):
        if highlight_index is not None and i == highlight_index:
            colors.append(CHART_ACCENT_HEX)
        else:
            colors.append(CHART_GREY_HEX)

    bars = ax.barh(labels, values, color=colors, height=0.5)

    # Data labels
    for bar, val in zip(bars, values):
        label = f"{val:,.0f}{unit}" if unit else f"{val:,.0f}"
        ax.text(
            bar.get_width() + max(values) * 0.02,
            bar.get_y() + bar.get_height() / 2,
            label,
            va="center",
            fontsize=CHART_FONT_SIZE,
            color="#4A4A4A",
        )

    ax.set_xlim(0, max(values) * 1.2)
    ax.invert_yaxis()
    ax.set_xticks([])
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=CHART_FONT_SIZE, color="#4A4A4A")

    fig.tight_layout()
    path = Path(tempfile.mktemp(suffix=".png"))
    fig.savefig(path, dpi=CHART_DPI, bbox_inches="tight", pad_inches=0.1)
    plt.close(fig)
    return str(path)


def render_waterfall_chart(
    labels: list[str],
    values: list[float],
    is_totals: list[bool],
    unit: str = "",
) -> str:
    """Render a waterfall chart and return the PNG path."""
    _setup_style()
    fig, ax = plt.subplots(figsize=(max(6, len(labels) * 1.2), 4))

    cumulative = 0.0
    bottoms = []
    heights = []
    colors = []

    for val, is_total in zip(values, is_totals):
        if is_total:
            bottoms.append(0)
            heights.append(val)
            colors.append(CHART_ACCENT_HEX)
            cumulative = val
        else:
            if val >= 0:
                bottoms.append(cumulative)
                heights.append(val)
                colors.append(CHART_SUB_HEX)
            else:
                bottoms.append(cumulative + val)
                heights.append(abs(val))
                colors.append(CHART_GREY_HEX)
            cumulative += val

    x = range(len(labels))
    ax.bar(x, heights, bottom=bottoms, color=colors, width=0.5)

    # Data labels
    for i, (b, h, v) in enumerate(zip(bottoms, heights, values)):
        label = f"{v:+,.0f}{unit}" if not is_totals[i] else f"{v:,.0f}{unit}"
        ax.text(i, b + h + max(heights) * 0.02, label,
                ha="center", fontsize=CHART_FONT_SIZE, color="#4A4A4A")

    ax.set_xticks(list(x))
    ax.set_xticklabels(labels, fontsize=CHART_FONT_SIZE, color="#4A4A4A")
    ax.set_yticks([])

    fig.tight_layout()
    path = Path(tempfile.mktemp(suffix=".png"))
    fig.savefig(path, dpi=CHART_DPI, bbox_inches="tight", pad_inches=0.1)
    plt.close(fig)
    return str(path)
