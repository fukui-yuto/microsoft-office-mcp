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


@mcp.tool()
def powerpoint_create_chart_slide(
    slide_number: int,
    title: str,
    chart_type: str,
    data: str,
    style: str = "modern",
) -> str:
    """チャート付きスライドを作成します。chart_type: bar, line, pie, doughnut。data: JSON文字列 {"categories": ["Q1","Q2"], "series": [{"name": "Revenue", "values": [100,200]}]}。style: modern, minimal, bold"""
    try:
        data_dict = json.loads(data) if isinstance(data, str) else data
        result = ppt_adv.create_chart_slide(
            slide_number, title, chart_type, data_dict, style=style,
        )
        return f"スライド {result['slide_number']} にチャートスライド（{chart_type}, {result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_dashboard_slide(
    slide_number: int,
    title: str,
    kpis: str,
    chart_data: str | None = None,
    style: str = "executive",
) -> str:
    """ダッシュボードスライドを作成します。kpis: JSON文字列 [{"label": "Revenue", "value": "$1.2M", "change": "+15%", "trend": "up"}]。style: executive, corporate, startup"""
    try:
        kpis_list = json.loads(kpis) if isinstance(kpis, str) else kpis
        chart_dict = json.loads(chart_data) if isinstance(chart_data, str) and chart_data else chart_data
        result = ppt_adv.create_dashboard_slide(
            slide_number, title, kpis_list, chart_data=chart_dict, style=style,
        )
        return f"スライド {result['slide_number']} にダッシュボードスライド（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_two_column_slide(
    slide_number: int,
    title: str,
    left_content: str,
    right_content: str,
    style: str = "balanced",
) -> str:
    """2カラムレイアウトスライドを作成します。content: JSON文字列 {"heading": "...", "items": ["..."]}。style: balanced, emphasis_left, emphasis_right"""
    try:
        left = json.loads(left_content) if isinstance(left_content, str) else left_content
        right = json.loads(right_content) if isinstance(right_content, str) else right_content
        result = ppt_adv.create_two_column_slide(
            slide_number, title, left, right, style=style,
        )
        return f"スライド {result['slide_number']} に2カラムレイアウト（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_image_text_slide(
    slide_number: int,
    title: str,
    text_items: list[str],
    image_path: str | None = None,
    image_position: str = "right",
    style: str = "modern",
) -> str:
    """テキストと画像のスライドを作成します。image_position: left, right。style: modern, clean, overlap"""
    try:
        result = ppt_adv.create_image_text_slide(
            slide_number, title, text_items, image_path=image_path,
            image_position=image_position, style=style,
        )
        return f"スライド {result['slide_number']} にテキスト＋画像レイアウト（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_three_column_slide(
    slide_number: int,
    title: str,
    columns: str,
    style: str = "cards",
) -> str:
    """3カラムレイアウトスライドを作成します。columns: JSON文字列 [{"heading": "...", "items": ["..."]}] (3つ)。style: cards, clean, icons"""
    try:
        cols = json.loads(columns) if isinstance(columns, str) else columns
        result = ppt_adv.create_three_column_slide(
            slide_number, title, cols, style=style,
        )
        return f"スライド {result['slide_number']} に3カラムレイアウト（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_stat_highlight(
    slide_number: int,
    stats: str,
    style: str = "big_numbers",
) -> str:
    """統計ハイライトスライドを作成します。stats: JSON文字列 [{"value": "95%", "label": "Customer Satisfaction", "color": [0,150,255]}]。style: big_numbers, circles, bars"""
    try:
        stats_list = json.loads(stats) if isinstance(stats, str) else stats
        result = ppt_adv.create_stat_highlight(
            slide_number, stats_list, style=style,
        )
        return f"スライド {result['slide_number']} に統計ハイライト（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_timeline_slide(
    slide_number: int,
    title: str,
    events: str,
    style: str = "horizontal",
) -> str:
    """タイムラインスライドを作成します。events: JSON文字列 [{"date": "2024 Q1", "title": "Launch", "description": "..."}]。style: horizontal, vertical, alternating"""
    try:
        events_list = json.loads(events) if isinstance(events, str) else events
        result = ppt_adv.create_timeline_slide(
            slide_number, title, events_list, style=style,
        )
        return f"スライド {result['slide_number']} にタイムライン（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_funnel_diagram(
    slide_number: int,
    title: str,
    stages: str,
    style: str = "gradient",
) -> str:
    """ファネル図スライドを作成します。stages: JSON文字列 [{"label": "Visitors", "value": "10,000"}]。style: gradient, flat, 3d"""
    try:
        stages_list = json.loads(stages) if isinstance(stages, str) else stages
        result = ppt_adv.create_funnel_diagram(
            slide_number, title, stages_list, style=style,
        )
        return f"スライド {result['slide_number']} にファネル図（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_swot_slide(
    slide_number: int,
    strengths: list[str],
    weaknesses: list[str],
    opportunities: list[str],
    threats: list[str],
    style: str = "colored",
) -> str:
    """SWOT分析スライドを作成します。各パラメータは文字列のリストです。style: colored, minimal, icons"""
    try:
        result = ppt_adv.create_swot_slide(
            slide_number, strengths, weaknesses, opportunities, threats,
            style=style,
        )
        return f"スライド {result['slide_number']} にSWOT分析（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_roadmap_slide(
    slide_number: int,
    title: str,
    phases: str,
    style: str = "arrow",
) -> str:
    """ロードマップスライドを作成します。phases: JSON文字列 [{"name": "Phase 1", "period": "Q1 2024", "items": ["Task 1"]}]。style: arrow, lane, milestone"""
    try:
        phases_list = json.loads(phases) if isinstance(phases, str) else phases
        result = ppt_adv.create_roadmap_slide(
            slide_number, title, phases_list, style=style,
        )
        return f"スライド {result['slide_number']} にロードマップ（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_pricing_table(
    slide_number: int,
    title: str,
    plans: str,
    style: str = "cards",
) -> str:
    """料金表スライドを作成します。plans: JSON文字列 [{"name": "Basic", "price": "$9/mo", "features": ["Feature 1"], "highlighted": false}]。style: cards, table, minimal"""
    try:
        plans_list = json.loads(plans) if isinstance(plans, str) else plans
        result = ppt_adv.create_pricing_table(
            slide_number, title, plans_list, style=style,
        )
        return f"スライド {result['slide_number']} に料金表（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_icon_list_slide(
    slide_number: int,
    title: str,
    items: str,
    style: str = "horizontal",
) -> str:
    """アイコン付きリストスライドを作成します。items: JSON文字列 [{"icon_text": "01", "title": "Step One", "description": "Details..."}]。style: horizontal, vertical, grid"""
    try:
        items_list = json.loads(items) if isinstance(items, str) else items
        result = ppt_adv.create_icon_list_slide(
            slide_number, title, items_list, style=style,
        )
        return f"スライド {result['slide_number']} にアイコンリスト（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_testimonial_slide(
    slide_number: int,
    testimonials: str,
    style: str = "cards",
) -> str:
    """お客様の声スライドを作成します。testimonials: JSON文字列 [{"quote": "Great!", "author": "John", "company": "Acme Inc"}]。style: cards, single_large, minimal"""
    try:
        test_list = json.loads(testimonials) if isinstance(testimonials, str) else testimonials
        result = ppt_adv.create_testimonial_slide(
            slide_number, test_list, style=style,
        )
        return f"スライド {result['slide_number']} にお客様の声（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_apply_consistent_branding(
    primary_color: list[int],
    secondary_color: list[int],
    accent_color: list[int],
    font_title: str = "Segoe UI",
    font_body: str = "Segoe UI",
) -> str:
    """プレゼンテーション全体に統一ブランディングを適用します。各colorは[R,G,B]のリストです。"""
    try:
        result = ppt_adv.apply_consistent_branding(
            tuple(primary_color), tuple(secondary_color), tuple(accent_color),
            font_title=font_title, font_body=font_body,
        )
        return f"{result['slides_affected']} 枚のスライドにブランディングを適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_before_after_slide(
    slide_number: int,
    title: str,
    before_items: list[str],
    after_items: list[str],
    style: str = "split",
) -> str:
    """Before/After比較スライドを作成します。style: split（左右分割）, overlay（重なるカード）, arrow（矢印で接続）"""
    try:
        result = ppt_adv.create_before_after_slide(
            slide_number, title, before_items, after_items, style=style,
        )
        return f"スライド {result['slide_number']} にBefore/After比較（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_feature_showcase(
    slide_number: int,
    title: str,
    features: str,
    style: str = "grid",
) -> str:
    """機能紹介スライドを作成します。features: JSON文字列 [{"title": "...", "description": "...", "icon_text": "..."}]。style: grid（2x3グリッド）, list（縦リスト）, cards（横並びカード）"""
    try:
        features_list = json.loads(features) if isinstance(features, str) else features
        result = ppt_adv.create_feature_showcase(
            slide_number, title, features_list, style=style,
        )
        return f"スライド {result['slide_number']} に機能紹介（{result['style']}、{result['feature_count']}件）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_data_table_slide(
    slide_number: int,
    title: str,
    headers: list[str],
    rows: str,
    style: str = "professional",
) -> str:
    """データテーブル付きスライドを作成します。rows: JSON文字列 [[値1, 値2, ...], ...]。style: professional（ダークヘッダー、ストライプ行）, minimal（細いボーダー）, colorful（カラフルセル）"""
    try:
        rows_list = json.loads(rows) if isinstance(rows, str) else rows
        result = ppt_adv.create_data_table_slide(
            slide_number, title, headers, rows_list, style=style,
        )
        return f"スライド {result['slide_number']} にデータテーブル（{result['style']}、{result['row_count']}行×{result['col_count']}列）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_thank_you_slide(
    slide_number: int,
    message: str = "Thank You",
    contact_email: str | None = None,
    contact_phone: str | None = None,
    website: str | None = None,
    social_media: str | None = None,
    style: str = "elegant",
) -> str:
    """Thank You / お問い合わせスライドを作成します。style: elegant（ダーク背景、ゴールドアクセント）, minimal（白背景）, corporate（ブランドカラー）"""
    try:
        result = ppt_adv.create_thank_you_slide(
            slide_number, message=message, contact_email=contact_email,
            contact_phone=contact_phone, website=website,
            social_media=social_media, style=style,
        )
        return f"スライド {result['slide_number']} にThank Youスライド（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_problem_solution_slide(
    slide_number: int,
    problem_title: str,
    problem_items: list[str],
    solution_title: str,
    solution_items: list[str],
    style: str = "contrast",
) -> str:
    """課題→解決スライドを作成します。style: contrast（赤/緑の対比）, split（左右分割）, flow（上→下フロー）"""
    try:
        result = ppt_adv.create_problem_solution_slide(
            slide_number, problem_title, problem_items,
            solution_title, solution_items, style=style,
        )
        return f"スライド {result['slide_number']} に課題→解決スライド（{result['style']}）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_mind_map(
    slide_number: int,
    center_topic: str,
    branches: str,
    style: str = "organic",
) -> str:
    """マインドマップ図を作成します。branches: JSON文字列 [{"topic": "ブランチ名", "subtopics": ["サブトピック1"]}]。style: organic（曲線コネクタ）, structured（直線）, colorful（各ブランチ別色）"""
    try:
        branches_list = json.loads(branches) if isinstance(branches, str) else branches
        result = ppt_adv.create_mind_map(
            slide_number, center_topic, branches_list, style=style,
        )
        return f"スライド {result['slide_number']} にマインドマップ（{result['style']}、{result['branch_count']}ブランチ）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_hierarchy_slide(
    slide_number: int,
    title: str,
    levels: str,
    style: str = "pyramid",
) -> str:
    """階層構造スライドを作成します。levels: JSON文字列 [{"label": "レベル名", "items": ["項目1"]}]。style: pyramid（逆ピラミッド）, tree（ツリー構造）, layers（積み重ねレイヤー）"""
    try:
        levels_list = json.loads(levels) if isinstance(levels, str) else levels
        result = ppt_adv.create_hierarchy_slide(
            slide_number, title, levels_list, style=style,
        )
        return f"スライド {result['slide_number']} に階層構造（{result['style']}、{result['level_count']}レベル）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_metrics_grid(
    slide_number: int,
    title: str,
    metrics: str,
    columns: int = 3,
    style: str = "cards",
) -> str:
    """メトリクスグリッドスライドを作成します。metrics: JSON文字列 [{"label": "ラベル", "value": "値", "unit": "単位", "trend": "up/down/flat"}]。style: cards, minimal, dashboard"""
    try:
        metrics_list = json.loads(metrics) if isinstance(metrics, str) else metrics
        result = ppt_adv.create_metrics_grid(
            slide_number, title, metrics_list, columns=columns, style=style,
        )
        return f"スライド {result['slide_number']} にメトリクスグリッド（{result['style']}、{result['metric_count']}件）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_workflow_slide(
    slide_number: int,
    title: str,
    steps: str,
    style: str = "horizontal",
) -> str:
    """ワークフロースライドを作成します。steps: JSON文字列 [{"title": "ステップ名", "description": "説明", "status": "done/active/pending"}]。style: horizontal（横並び）, vertical（縦並び）, circular（円形配置）"""
    try:
        steps_list = json.loads(steps) if isinstance(steps, str) else steps
        result = ppt_adv.create_workflow_slide(
            slide_number, title, steps_list, style=style,
        )
        return f"スライド {result['slide_number']} にワークフロー（{result['style']}、{result['step_count']}ステップ）を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_create_venn_diagram(
    slide_number: int,
    items: str,
    center_text: str | None = None,
    style: str = "classic",
) -> str:
    """ベン図（2〜3円）を作成します。items: JSON文字列 [{"label": "ラベル", "items": ["項目1"]}]。style: classic（半透明重なり）, solid（不透明）, minimal（輪郭のみ）"""
    try:
        items_list = json.loads(items) if isinstance(items, str) else items
        result = ppt_adv.create_venn_diagram(
            slide_number, items_list, center_text=center_text, style=style,
        )
        return f"スライド {result['slide_number']} にベン図（{result['style']}、{result['circle_count']}円）を作成しました"
    except Exception as e:
        return f"エラー: {e}"
