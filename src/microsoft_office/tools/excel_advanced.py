"""Excel Advanced MCP tool definitions - dashboards, templates, and professional formatting."""

import json

from microsoft_office.server import mcp
from microsoft_office.office import excel_advanced


@mcp.tool()
def excel_create_dashboard_header(
    title: str,
    subtitle: str | None = None,
    color_scheme: str = "corporate_blue",
    sheet: str | None = None,
) -> str:
    """ダッシュボードヘッダーを作成します。カラーバンド、タイトル、サブタイトル、日付を含むプロフェッショナルなヘッダー。color_scheme: corporate_blue, modern_dark, forest_green, sunset_orange, royal_purple。"""
    try:
        result = excel_advanced.create_dashboard_header(sheet, title, subtitle=subtitle, color_scheme=color_scheme)
        return f"[{result['sheet']}] ダッシュボードヘッダーを作成しました（タイトル: {result['title']}, スキーム: {result['scheme']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_kpi_cards(
    kpis: str,
    start_row: int = 3,
    style: str = "modern",
    sheet: str | None = None,
) -> str:
    """KPIカードを作成します。kpisはJSON文字列: [{"title":"Revenue","value":"$1.2M","change":"+15%","change_type":"positive"}]。style: modern, minimal, bold。"""
    try:
        kpis_data = json.loads(kpis)
        result = excel_advanced.create_kpi_cards(sheet, kpis_data, start_row=start_row, style=style)
        return f"[{result['sheet']}] {result['kpi_count']}個のKPIカードを作成しました（スタイル: {result['style']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_data_table(
    headers: list[str],
    data: list[list],
    start_cell: str = "A1",
    style: str = "striped_blue",
    sheet: str | None = None,
) -> str:
    """プロフェッショナルなデータテーブルを作成します。自動フォーマット付き。style: striped_blue, striped_gray, bordered, minimal, dark_header, colorful。"""
    try:
        result = excel_advanced.create_data_table(sheet, headers, data, start_cell=start_cell, style=style)
        return f"[{result['sheet']}] データテーブルを作成しました（{result['rows']}行×{result['cols']}列, 範囲: {result['range']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_summary_row(
    range_str: str,
    summary_type: str = "sum",
    style: str = "bold_bordered",
    sheet: str | None = None,
) -> str:
    """データ範囲の下に集計行を追加します。summary_type: sum, average, count, min, max。style: bold_bordered, colored_band, double_line。"""
    try:
        result = excel_advanced.create_summary_row(sheet, range_str, summary_type=summary_type, style=style)
        return f"[{result['sheet']}] 集計行を追加しました（行{result['summary_row']}, タイプ: {result['type']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_apply_table_theme(
    range_str: str,
    theme: str = "professional",
    sheet: str | None = None,
) -> str:
    """データ範囲にテーブルテーマを適用します。最初の行がヘッダーとして扱われます。theme: professional, financial, marketing, executive, minimal。"""
    try:
        result = excel_advanced.apply_table_theme(sheet, range_str, theme=theme)
        return f"[{result['sheet']}] テーマ '{result['theme']}' を {result['range']} に適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_input_form(
    fields: str,
    start_row: int = 1,
    style: str = "bordered",
    sheet: str | None = None,
) -> str:
    """データ入力フォームを作成します。fieldsはJSON文字列: [{"label":"名前","type":"text","required":true,"validation":"string","hint":"フルネームを入力"}]。style: bordered, shaded, minimal。"""
    try:
        fields_data = json.loads(fields)
        result = excel_advanced.create_input_form(sheet, fields_data, start_row=start_row, style=style)
        return f"[{result['sheet']}] 入力フォームを作成しました（{result['fields']}フィールド, スタイル: {result['style']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_calendar(
    year: int,
    month: int,
    start_cell: str = "A1",
    style: str = "modern",
    sheet: str | None = None,
) -> str:
    """月間カレンダーを作成します。style: modern, minimal, colorful。"""
    try:
        result = excel_advanced.create_calendar(sheet, year, month, start_cell=start_cell, style=style)
        return f"[{result['sheet']}] {result['year']}年{result['month']}月のカレンダーを作成しました（スタイル: {result['style']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_gantt_chart(
    tasks: str,
    start_cell: str = "A1",
    sheet: str | None = None,
) -> str:
    """ガントチャートを作成します。tasksはJSON文字列: [{"name":"Task 1","start_date":"2024-01-01","end_date":"2024-01-15","progress":50,"color":[66,133,244]}]。"""
    try:
        tasks_data = json.loads(tasks)
        # Convert color lists to tuples
        for task in tasks_data:
            if "color" in task and isinstance(task["color"], list):
                task["color"] = tuple(task["color"])
        result = excel_advanced.create_gantt_chart(sheet, tasks_data, start_cell=start_cell)
        return f"[{result['sheet']}] ガントチャートを作成しました（{result['tasks']}タスク, {result['days']}日間）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_scorecard(
    title: str,
    metrics: str,
    start_cell: str = "A1",
    style: str = "traffic_light",
    sheet: str | None = None,
) -> str:
    """パフォーマンススコアカードを作成します。metricsはJSON文字列: [{"name":"Customer Satisfaction","target":90,"actual":87,"unit":"%"}]。style: traffic_light, progress_bar, rating。"""
    try:
        metrics_data = json.loads(metrics)
        result = excel_advanced.create_scorecard(sheet, title, metrics_data, start_cell=start_cell, style=style)
        return f"[{result['sheet']}] スコアカード '{result['title']}' を作成しました（{result['metrics']}指標, スタイル: {result['style']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_financial_report(
    title: str,
    categories: str,
    periods: list[str],
    values: list[list],
    style: str = "standard",
    sheet: str | None = None,
) -> str:
    """財務レポートを作成します。categoriesはJSON文字列: [{"name":"売上","level":0,"bold":true}]。level: 0=セクションヘッダー, 1=項目, 2=小計, 3=合計。"""
    try:
        categories_data = json.loads(categories)
        result = excel_advanced.create_financial_report(sheet, title, categories_data, periods, values, style=style)
        return f"[{result['sheet']}] 財務レポート '{result['title']}' を作成しました（{result['categories']}カテゴリ, {result['periods']}期間）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_comparison_table(
    headers: list[str],
    row_labels: list[str],
    data: list[list],
    highlight_best: bool = True,
    start_cell: str = "A1",
    sheet: str | None = None,
) -> str:
    """比較表を作成します。各行の最良値を自動ハイライトします。"""
    try:
        result = excel_advanced.create_comparison_table(sheet, headers, row_labels, data,
                                                        highlight_best=highlight_best, start_cell=start_cell)
        return f"[{result['sheet']}] 比較表を作成しました（{result['rows']}行×{result['cols']}列, ハイライト: {result['highlighted']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_setup_print_ready(
    title: str | None = None,
    orientation: str = "portrait",
    fit_to_pages: bool = True,
    sheet: str | None = None,
) -> str:
    """シートを印刷用に設定します。余白、ヘッダー/フッター、ページ収まり設定を行います。orientation: portrait, landscape。"""
    try:
        result = excel_advanced.setup_print_ready(sheet, title=title, orientation=orientation, fit_to_pages=fit_to_pages)
        return f"[{result['sheet']}] 印刷設定を完了しました（向き: {result['orientation']}, ページ収まり: {result['fit_to_pages']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_heatmap(
    data_range: str,
    color_low: list[int] | None = None,
    color_mid: list[int] | None = None,
    color_high: list[int] | None = None,
    sheet: str | None = None,
) -> str:
    """数値範囲にヒートマップカラーリングを適用します。デフォルト: 緑→黄→赤。color_low/color_mid/color_high: [R,G,B]形式。"""
    try:
        result = excel_advanced.create_heatmap(
            sheet, data_range,
            color_low=tuple(color_low) if color_low else None,
            color_mid=tuple(color_mid) if color_mid else None,
            color_high=tuple(color_high) if color_high else None,
        )
        return f"[{result['sheet']}] ヒートマップを {result['range']} に適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_sheet_navigation(
    sheets_info: str,
    nav_sheet_name: str = "Menu",
) -> str:
    """ナビゲーション/メニューシートを作成します。sheets_infoはJSON文字列: [{"name":"Sheet1","description":"売上データ"}]。"""
    try:
        info_data = json.loads(sheets_info)
        result = excel_advanced.add_sheet_navigation(info_data, nav_sheet_name=nav_sheet_name)
        return f"ナビゲーションシート '{result['nav_sheet']}' を作成しました（{result['entries']}エントリ）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_invoice_template(
    company_name: str,
    company_address: str | None = None,
    logo_path: str | None = None,
    style: str = "modern",
    sheet: str | None = None,
) -> str:
    """プロフェッショナルな請求書テンプレートを作成します。小計、税金、合計の計算式付き。"""
    try:
        result = excel_advanced.create_invoice_template(sheet, company_name,
                                                         company_address=company_address,
                                                         logo_path=logo_path, style=style)
        return f"[{result['sheet']}] 請求書テンプレートを作成しました（会社名: {result['company']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_timesheet(
    employee_name: str | None = None,
    month: int | None = None,
    year: int | None = None,
    style: str = "standard",
    sheet: str | None = None,
) -> str:
    """タイムシートテンプレートを作成します。自動計算式付きの月間タイムシート。"""
    try:
        result = excel_advanced.create_timesheet(sheet, employee_name=employee_name,
                                                   month=month, year=year, style=style)
        return f"[{result['sheet']}] タイムシートを作成しました（期間: {result['month']}, {result['days']}日）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_format_as_currency(
    range_str: str,
    currency_symbol: str = "¥",
    decimal_places: int = 0,
    sheet: str | None = None,
) -> str:
    """範囲を通貨フォーマットに設定します。currency_symbol: 通貨記号（デフォルト: ¥）。"""
    try:
        result = excel_advanced.format_as_currency(sheet, range_str,
                                                     currency_symbol=currency_symbol,
                                                     decimal_places=decimal_places)
        return f"[{result['sheet']}] {result['range']} を通貨フォーマットに設定しました（{result['symbol']}, 小数点{result['decimals']}桁）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_format_as_percentage(
    range_str: str,
    decimal_places: int = 1,
    sheet: str | None = None,
) -> str:
    """範囲をパーセンテージフォーマットに設定します。"""
    try:
        result = excel_advanced.format_as_percentage(sheet, range_str, decimal_places=decimal_places)
        return f"[{result['sheet']}] {result['range']} をパーセンテージフォーマットに設定しました（小数点{result['decimals']}桁）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_data_summary(
    data_range: str,
    summary_cell: str,
    summary_type: str = "dashboard",
    sheet: str | None = None,
) -> str:
    """データ範囲の統計サマリーを追加します。summary_type: dashboard (グリッド表示), simple (主要数値のみ), detailed (詳細統計)。"""
    try:
        result = excel_advanced.add_data_summary(sheet, data_range, summary_cell, summary_type=summary_type)
        return f"[{result['sheet']}] データサマリーを追加しました（データ範囲: {result['data_range']}, タイプ: {result['summary_type']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_checklist(
    title: str,
    items: list[str],
    start_cell: str = "A1",
    sheet: str | None = None,
) -> str:
    """チェックリストを作成します。チェックボックス付きで、チェック時に取り消し線が適用されます。"""
    try:
        result = excel_advanced.create_checklist(sheet, title, items, start_cell=start_cell)
        return f"[{result['sheet']}] チェックリスト '{result['title']}' を作成しました（{result['items']}項目）"
    except Exception as e:
        return f"エラー: {e}"
