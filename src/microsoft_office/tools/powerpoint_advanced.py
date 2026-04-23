"""PowerPoint Advanced Design MCP tool definitions."""

import json

from microsoft_office.server import mcp
from microsoft_office.office import powerpoint_advanced as ppt_adv


@mcp.tool()
def powerpoint_apply_theme_colors(color_scheme: str) -> str:
    """プレゼンテーションにカラースキームを適用します。選択肢: corporate_blue, modern_dark, nature_green, sunset_warm, ocean_breeze, monochrome, royal_purple, tech_neon, pastel_soft, bold_contrast"""
    try:
        result = ppt_adv.apply_theme_colors(color_scheme)
        return f"カラースキーム '{result['scheme']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_master_font(title_font: str, body_font: str) -> str:
    """マスタースライドのタイトルフォントと本文フォントを設定します。"""
    try:
        result = ppt_adv.set_master_font(title_font, body_font)
        return f"マスターフォントを設定しました（タイトル: {result['title_font']}, 本文: {result['body_font']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_watermark(
    text: str,
    slide_number: int | None = None,
    font_size: float = 54,
    color: list[int] | None = None,
    rotation: float = -45,
    transparency: float = 75,
) -> str:
    """スライドに透かし（ウォーターマーク）を追加します。slide_numberを省略すると全スライドに適用します。transparencyは0-100のパーセント値です。"""
    try:
        color_tuple = tuple(color) if color else (200, 200, 200)
        result = ppt_adv.add_watermark(
            text, slide_number=slide_number, font_size=font_size,
            color=color_tuple, rotation=rotation, transparency=transparency,
        )
        return f"透かし '{result['text']}' を {result['slides_affected']} 枚のスライドに追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_title_slide_design(
    slide_number: int,
    title: str,
    subtitle: str | None = None,
    style: str = "modern_gradient",
) -> str:
    """プロフェッショナルなタイトルスライドをデザインします。style: modern_gradient, minimal_white, bold_split, dark_premium, geometric"""
    try:
        result = ppt_adv.create_title_slide_design(
            slide_number, title, subtitle=subtitle, style=style,
        )
        return f"スライド {result['slide_number']} にタイトルデザイン '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_content_slide_design(
    slide_number: int,
    title: str,
    items: list[str],
    style: str = "cards",
) -> str:
    """プロフェッショナルなコンテンツスライドをデザインします。style: cards, timeline, comparison, stats, icon_grid。statsスタイルでは'数値|ラベル'形式が使えます。"""
    try:
        result = ppt_adv.create_content_slide_design(
            slide_number, title, items, style=style,
        )
        return f"スライド {result['slide_number']} にコンテンツデザイン '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_section_divider(
    slide_number: int,
    title: str,
    subtitle: str | None = None,
    style: str = "gradient_wave",
) -> str:
    """セクション区切りスライドを作成します。style: gradient_wave, bold_number, minimal_line, full_bleed_color"""
    try:
        result = ppt_adv.create_section_divider(
            slide_number, title, subtitle=subtitle, style=style,
        )
        return f"スライド {result['slide_number']} にセクション区切り '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_closing_slide(
    slide_number: int,
    title: str = "Thank You",
    subtitle: str | None = None,
    contact_info: str | None = None,
    style: str = "elegant",
) -> str:
    """クロージング（Thank You）スライドを作成します。style: elegant, minimal, bold"""
    try:
        result = ppt_adv.create_closing_slide(
            slide_number, title=title, subtitle=subtitle,
            contact_info=contact_info, style=style,
        )
        return f"スライド {result['slide_number']} にクロージングデザイン '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_decorative_element(
    slide_number: int,
    element_type: str,
    position: str = "top_right",
    color: list[int] | None = None,
    size: str = "medium",
) -> str:
    """スライドに装飾要素を追加します。element_type: circle_cluster, diagonal_lines, dot_pattern, corner_accent, gradient_bar, wave_shape。position: top_left, top_right, bottom_left, bottom_right, background。size: small, medium, large"""
    try:
        color_tuple = tuple(color) if color else None
        result = ppt_adv.add_decorative_element(
            slide_number, element_type, position=position,
            color=color_tuple, size=size,
        )
        return f"スライド {result['slide_number']} に装飾要素 '{result['element_type']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_agenda_slide(
    slide_number: int,
    title: str,
    items: list[str],
    highlight_index: int | None = None,
    style: str = "numbered",
) -> str:
    """アジェンダ（目次）スライドを作成します。highlight_indexで現在のセクションをハイライトできます（0始まり）。style: numbered, cards, steps"""
    try:
        result = ppt_adv.create_agenda_slide(
            slide_number, title, items,
            highlight_index=highlight_index, style=style,
        )
        return f"スライド {result['slide_number']} にアジェンダデザイン '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_comparison_slide(
    slide_number: int,
    title: str,
    left_title: str,
    left_items: list[str],
    right_title: str,
    right_items: list[str],
    style: str = "versus",
) -> str:
    """比較スライドを作成します。style: versus（VS中央表示）, columns（クリーンな2列）, pros_cons（緑/赤のメリット・デメリット）"""
    try:
        result = ppt_adv.create_comparison_slide(
            slide_number, title, left_title, left_items,
            right_title, right_items, style=style,
        )
        return f"スライド {result['slide_number']} に比較デザイン '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_process_flow(
    slide_number: int,
    title: str,
    steps: list[str],
    style: str = "arrows",
) -> str:
    """プロセスフロー図スライドを作成します。style: arrows（矢印接続）, circles（番号付き円）, chevrons（シェブロン形状）"""
    try:
        result = ppt_adv.create_process_flow(
            slide_number, title, steps, style=style,
        )
        return f"スライド {result['slide_number']} にプロセスフロー '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_progress_bar(
    slide_number: int,
    progress: float,
    position: str = "bottom",
    color: list[int] | None = None,
    height: float = 8,
) -> str:
    """スライドにプログレスバーを追加します。progressは0.0〜1.0の範囲で指定します。position: top, bottom"""
    try:
        color_tuple = tuple(color) if color else None
        result = ppt_adv.add_progress_bar(
            slide_number, progress, position=position,
            color=color_tuple, height=height,
        )
        return f"スライド {result['slide_number']} にプログレスバー ({result['progress']}) を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_quote_slide(
    slide_number: int,
    quote: str,
    author: str | None = None,
    style: str = "large_quote",
) -> str:
    """引用スライドを作成します。style: large_quote（大きな引用符）, minimal（シンプル＋細い線）, highlighted（色付き背景帯）"""
    try:
        result = ppt_adv.create_quote_slide(
            slide_number, quote, author=author, style=style,
        )
        return f"スライド {result['slide_number']} に引用デザイン '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_team_slide(
    slide_number: int,
    title: str,
    members: list[dict],
    style: str = "grid",
) -> str:
    """チーム紹介スライドを作成します。membersは{name, role, description(任意)}のリストです。style: grid（カードグリッド）, horizontal（横並び）"""
    try:
        result = ppt_adv.create_team_slide(
            slide_number, title, members, style=style,
        )
        return f"スライド {result['slide_number']} にチームデザイン '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_slide_number_footer(
    start_slide: int = 1,
    end_slide: int | None = None,
    format_type: str = "number",
    position: str = "bottom_right",
    font_size: float = 10,
    font_color: list[int] | None = None,
) -> str:
    """スライドにページ番号フッターを追加します。format_type: number（番号のみ）, of_total（1/10形式）, dash_total（1 - 10形式）。position: bottom_left, bottom_center, bottom_right"""
    try:
        color_tuple = tuple(font_color) if font_color else None
        result = ppt_adv.add_slide_number_footer(
            start_slide=start_slide, end_slide=end_slide,
            format_type=format_type, position=position,
            font_size=font_size, font_color=color_tuple,
        )
        return f"{result['slides_affected']} 枚のスライドにページ番号（{result['format']}）を追加しました"
    except Exception as e:
        return f"エラー: {e}"
