"""Safe PowerPoint MCP tool definitions — design-token constrained.

These tools enforce design tokens: no arbitrary RGB, no gradients,
no shadows, no 3D.  Colors, fonts, and sizes are resolved from
ThemeName enum only.
"""

import json

from microsoft_office.design_tokens import ThemeName
from microsoft_office.office import powerpoint_safe as ppt_safe
from microsoft_office.server import mcp


def _resolve_theme(theme: str) -> ThemeName:
    """Resolve theme string to ThemeName enum, raising clear errors."""
    try:
        return ThemeName(theme)
    except ValueError:
        valid = ", ".join(t.value for t in ThemeName)
        raise ValueError(f"Unknown theme '{theme}'. Choose from: {valid}")


@mcp.tool()
def pptx_add_title_slide(
    title: str,
    subtitle: str | None = None,
    theme: str = "business",
) -> str:
    """タイトルスライドを追加します。theme: business, minimal, warm"""
    try:
        result = ppt_safe.add_title_slide(
            title, subtitle=subtitle, theme=_resolve_theme(theme),
        )
        return (
            f"スライド {result['slide_number']} にタイトルスライドを追加しました "
            f"(テーマ: {result['theme']})"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def pptx_add_section_header(
    title: str,
    accent: bool = False,
    theme: str = "business",
) -> str:
    """セクション見出しスライドを追加します。accent=Trueで背景にアクセント色を使用。theme: business, minimal, warm"""
    try:
        result = ppt_safe.add_section_header(
            title, accent=accent, theme=_resolve_theme(theme),
        )
        return (
            f"スライド {result['slide_number']} にセクション見出しを追加しました "
            f"(テーマ: {result['theme']})"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def pptx_add_bullet_slide(
    title: str,
    bullets: list[str],
    theme: str = "business",
) -> str:
    """箇条書きスライドを追加します（最大5項目）。theme: business, minimal, warm"""
    try:
        result = ppt_safe.add_bullet_slide(
            title, bullets, theme=_resolve_theme(theme),
        )
        return (
            f"スライド {result['slide_number']} に箇条書きスライドを追加しました "
            f"({result['bullet_count']}項目, テーマ: {result['theme']})"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def pptx_add_two_column_slide(
    title: str,
    left: list[str],
    right: list[str],
    theme: str = "business",
) -> str:
    """2カラムスライドを追加します（各列最大5項目）。theme: business, minimal, warm"""
    try:
        result = ppt_safe.add_two_column_slide(
            title, left, right, theme=_resolve_theme(theme),
        )
        return (
            f"スライド {result['slide_number']} に2カラムスライドを追加しました "
            f"(テーマ: {result['theme']})"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def pptx_add_image_slide(
    title: str,
    image_path: str,
    caption: str | None = None,
    theme: str = "business",
) -> str:
    """画像スライドを追加します。theme: business, minimal, warm"""
    try:
        result = ppt_safe.add_image_slide(
            title, image_path, caption=caption, theme=_resolve_theme(theme),
        )
        return (
            f"スライド {result['slide_number']} に画像スライドを追加しました "
            f"(テーマ: {result['theme']})"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def pptx_add_comparison_slide(
    title: str,
    items: list[dict],
    theme: str = "business",
) -> str:
    """比較スライドを追加します（2〜4項目）。items: [{"heading": "見出し", "points": ["項目1", ...]}]。theme: business, minimal, warm"""
    try:
        result = ppt_safe.add_comparison_slide(
            title, items, theme=_resolve_theme(theme),
        )
        return (
            f"スライド {result['slide_number']} に比較スライドを追加しました "
            f"({result['item_count']}項目, テーマ: {result['theme']})"
        )
    except Exception as e:
        return f"エラー: {e}"
