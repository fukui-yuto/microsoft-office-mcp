"""Excel MCP tool definitions."""

import json

from microsoft_office.server import mcp
from microsoft_office.office import excel


@mcp.tool()
def excel_create(file_path: str | None = None) -> str:
    """Excelブックを新規作成します。file_pathを指定すると即座に保存します。"""
    try:
        result = excel.create_workbook(file_path)
        msg = f"ブック '{result['name']}' を作成しました（シート数: {result['sheet_count']}）"
        if file_path:
            msg += f" (保存先: {file_path})"
        return msg
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_open(file_path: str) -> str:
    """既存のExcelファイル(.xlsx)を開きます。"""
    try:
        result = excel.open_workbook(file_path)
        return f"'{result['name']}' を開きました（シート: {', '.join(result['sheets'])}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_save(file_path: str | None = None) -> str:
    """アクティブなブックを保存します。file_pathを指定すると別名保存します。"""
    try:
        result = excel.save_workbook(file_path)
        return f"'{result['name']}' を保存しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_close() -> str:
    """アクティブなブックを閉じます。"""
    try:
        result = excel.close_workbook()
        return f"'{result['name']}' を閉じました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_sheets() -> str:
    """アクティブなブックの全シート名を取得します。"""
    try:
        result = excel.get_sheets()
        return f"ブック '{result['name']}' のシート: {', '.join(result['sheets'])}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_sheet(name: str) -> str:
    """新しいワークシートを追加します。name: シート名。"""
    try:
        result = excel.add_sheet(name)
        return f"シート '{result['sheet_name']}' を追加しました（合計: {result['sheet_count']}シート）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_delete_sheet(sheet: str) -> str:
    """指定したワークシートを削除します。sheet: 削除するシート名。"""
    try:
        result = excel.delete_sheet(sheet)
        return f"シート '{result['sheet_name']}' を削除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_rename_sheet(old_name: str, new_name: str) -> str:
    """ワークシートの名前を変更します。old_name: 現在のシート名、new_name: 新しいシート名。"""
    try:
        result = excel.rename_sheet(old_name, new_name)
        return f"シート '{result['old_name']}' を '{result['new_name']}' に変更しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_read_cell(cell: str, sheet: str | None = None) -> str:
    """セルの値を読み取ります。例: cell="A1"。sheet省略時はアクティブシート。"""
    try:
        result = excel.read_cell(cell, sheet)
        return f"[{result['sheet']}] {result['cell']} = {result['value']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_write_cell(cell: str, value: str, sheet: str | None = None) -> str:
    """セルに値を書き込みます。例: cell="A1", value="Hello"。sheet省略時はアクティブシート。"""
    try:
        result = excel.write_cell(cell, value, sheet)
        return f"[{result['sheet']}] {result['cell']} に '{result['value']}' を書き込みました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_read_range(range_address: str, sheet: str | None = None) -> str:
    """範囲のデータを読み取ります。例: range_address="A1:C10"。結果はJSON形式で返します。"""
    try:
        result = excel.read_range(range_address, sheet)
        return f"[{result['sheet']}] {result['range']}:\n{json.dumps(result['data'], ensure_ascii=False, default=str)}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_write_range(start_cell: str, data: list[list], sheet: str | None = None) -> str:
    """範囲にデータを書き込みます。dataは2次元配列。例: start_cell="A1", data=[["名前","年齢"],["太郎","25"]]。"""
    try:
        result = excel.write_range(start_cell, data, sheet)
        return f"[{result['sheet']}] {result['start_cell']} から {result['rows']}行×{result['cols']}列 を書き込みました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_sheet_data(sheet: str | None = None) -> str:
    """シートの使用範囲のデータを全て取得します。結果はJSON形式で返します。"""
    try:
        result = excel.get_sheet_data(sheet)
        return f"[{result['sheet']}] {result['rows']}行×{result['cols']}列:\n{json.dumps(result['data'], ensure_ascii=False, default=str)}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_cell_format(
    cell: str,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_name: str | None = None,
    font_color: list[int] | None = None,
    underline: bool | None = None,
    number_format: str | None = None,
    sheet: str | None = None,
) -> str:
    """セルのフォント書式を設定します。font_color: [R,G,B]形式（例: [255,0,0]で赤）。number_formatで表示形式を指定可能（例: "#,##0"）。"""
    try:
        result = excel.set_cell_format(
            cell,
            bold=bold,
            italic=italic,
            font_size=font_size,
            font_name=font_name,
            font_color=tuple(font_color) if font_color else None,
            underline=underline,
            number_format=number_format,
            sheet=sheet,
        )
        return f"[{result['sheet']}] {result['cell']} の書式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_cell_fill(
    range_address: str,
    color: list[int] | None = None,
    pattern: int | None = None,
    pattern_color: list[int] | None = None,
    sheet: str | None = None,
) -> str:
    """セルの塗りつぶしを設定します。color: [R,G,B]形式（例: [255,255,0]で黄色）。pattern: 1=単色。"""
    try:
        result = excel.set_cell_fill(
            range_address,
            color=tuple(color) if color else None,
            pattern=pattern,
            pattern_color=tuple(pattern_color) if pattern_color else None,
            sheet=sheet,
        )
        return f"[{result['sheet']}] {result['range']} の塗りつぶしを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_cell_borders(
    range_address: str,
    border_style: int = 1,
    border_weight: int = 2,
    color: list[int] | None = None,
    edges: list[int] | None = None,
    sheet: str | None = None,
) -> str:
    """セルの罫線を設定します。edges: 罫線位置のリスト（7=左,8=上,9=下,10=右,11=内縦,12=内横）。border_style: 1=実線。border_weight: 1=細,2=中,3=太。"""
    try:
        result = excel.set_cell_borders(
            range_address,
            border_style=border_style,
            border_weight=border_weight,
            color=tuple(color) if color else None,
            edges=edges,
            sheet=sheet,
        )
        return f"[{result['sheet']}] {result['range']} の罫線を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_cell_alignment(
    range_address: str,
    horizontal: int | None = None,
    vertical: int | None = None,
    wrap_text: bool | None = None,
    text_rotation: int | None = None,
    sheet: str | None = None,
) -> str:
    """セルの配置を設定します。horizontal: -4131=左,-4108=中央,-4152=右。vertical: -4160=上,-4108=中央,-4107=下。wrap_text: 折り返し。text_rotation: 回転角度。"""
    try:
        result = excel.set_cell_alignment(
            range_address,
            horizontal=horizontal,
            vertical=vertical,
            wrap_text=wrap_text,
            text_rotation=text_rotation,
            sheet=sheet,
        )
        return f"[{result['sheet']}] {result['range']} の配置を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_merge_cells(range_address: str, sheet: str | None = None) -> str:
    """指定範囲のセルを結合します。例: range_address="A1:D1"。"""
    try:
        result = excel.merge_cells(range_address, sheet)
        return f"[{result['sheet']}] {result['range']} を結合しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_unmerge_cells(range_address: str, sheet: str | None = None) -> str:
    """指定範囲のセル結合を解除します。例: range_address="A1:D1"。"""
    try:
        result = excel.unmerge_cells(range_address, sheet)
        return f"[{result['sheet']}] {result['range']} の結合を解除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_column_width(
    columns: str,
    width: float | None = None,
    auto_fit: bool = False,
    sheet: str | None = None,
) -> str:
    """列幅を設定します。columns: 列指定（例: "A", "A:C"）。width: 幅の値。auto_fit: Trueで自動調整。"""
    try:
        result = excel.set_column_width(columns, width=width, auto_fit=auto_fit, sheet=sheet)
        return f"[{result['sheet']}] 列 {result['columns']} の幅を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_row_height(rows: str, height: float, sheet: str | None = None) -> str:
    """行の高さを設定します。rows: 行指定（例: "1", "1:5"）。height: 高さの値（ポイント）。"""
    try:
        result = excel.set_row_height(rows, height, sheet)
        return f"[{result['sheet']}] 行 {result['rows']} の高さを {height} に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_formula(cell: str, formula: str, sheet: str | None = None) -> str:
    """セルに数式を設定します。例: cell="A11", formula="=SUM(A1:A10)"。"""
    try:
        result = excel.set_formula(cell, formula, sheet)
        return f"[{result['sheet']}] {result['cell']} に数式 '{result['formula']}' を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_number_format(range_address: str, format_string: str, sheet: str | None = None) -> str:
    """セルの表示形式を設定します。例: format_string="#,##0", "yyyy/mm/dd", "0%", "0.00"。"""
    try:
        result = excel.set_number_format(range_address, format_string, sheet)
        return f"[{result['sheet']}] {result['range']} の表示形式を '{format_string}' に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_named_range(name: str, range_address: str, sheet: str | None = None) -> str:
    """名前付き範囲を追加します。name: 範囲名、range_address: 範囲（例: "A1:C10"）。"""
    try:
        result = excel.add_named_range(name, range_address, sheet)
        return f"名前付き範囲 '{result['name']}' を追加しました（範囲: {result['range']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_delete_named_range(name: str) -> str:
    """名前付き範囲を削除します。name: 削除する範囲名。"""
    try:
        result = excel.delete_named_range(name)
        return f"名前付き範囲 '{result['name']}' を削除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_named_ranges() -> str:
    """ブック内の全ての名前付き範囲を一覧表示します。"""
    try:
        result = excel.get_named_ranges()
        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_sort_range(
    range_address: str,
    sort_field: int,
    order: int = 1,
    has_header: bool = True,
    sheet: str | None = None,
) -> str:
    """範囲をソートします。sort_field: ソートキーの列番号（1始まり）。order: 1=昇順,2=降順。has_header: ヘッダー行の有無。"""
    try:
        result = excel.sort_range(range_address, sort_field, order=order, has_header=has_header, sheet=sheet)
        return f"[{result['sheet']}] {result['range']} をソートしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_auto_filter(
    range_address: str | None = None,
    field: int | None = None,
    criteria1: str | None = None,
    criteria2: str | None = None,
    operator: int | None = None,
    sheet: str | None = None,
) -> str:
    """オートフィルターを設定します。range_address: フィルター範囲。field: フィルター列番号（1始まり）。criteria1/criteria2: フィルター条件。"""
    try:
        result = excel.set_auto_filter(
            range_address=range_address,
            field=field,
            criteria1=criteria1,
            criteria2=criteria2,
            operator=operator,
            sheet=sheet,
        )
        return f"[{result['sheet']}] オートフィルターを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_freeze_panes(cell: str, sheet: str | None = None) -> str:
    """ウィンドウ枠を固定します。例: cell="B2"で1行目と1列目を固定。cell="A2"で1行目のみ固定。"""
    try:
        result = excel.freeze_panes(cell, sheet)
        return f"[{result['sheet']}] セル {result['cell']} でウィンドウ枠を固定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_conditional_formatting(
    range_address: str,
    operator: int,
    formula1: str,
    formula2: str | None = None,
    font_color: list[int] | None = None,
    fill_color: list[int] | None = None,
    sheet: str | None = None,
) -> str:
    """条件付き書式を設定します。operator: 1=間,3=等しい,5=より大きい,6=より小さい。font_color/fill_color: [R,G,B]形式。"""
    try:
        result = excel.set_conditional_formatting(
            range_address,
            operator,
            formula1,
            formula2=formula2,
            font_color=tuple(font_color) if font_color else None,
            fill_color=tuple(fill_color) if fill_color else None,
            sheet=sheet,
        )
        return f"[{result['sheet']}] {result['range']} に条件付き書式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_data_validation(
    range_address: str,
    validation_type: int,
    operator: int | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
    input_title: str | None = None,
    input_message: str | None = None,
    error_title: str | None = None,
    error_message: str | None = None,
    sheet: str | None = None,
) -> str:
    """データの入力規則を設定します。validation_type: 3=リスト,1=整数,2=小数,4=日付。リストの場合formula1にカンマ区切りの値を指定。"""
    try:
        result = excel.add_data_validation(
            range_address,
            validation_type,
            operator=operator,
            formula1=formula1,
            formula2=formula2,
            input_title=input_title,
            input_message=input_message,
            error_title=error_title,
            error_message=error_message,
            sheet=sheet,
        )
        return f"[{result['sheet']}] {result['range']} にデータの入力規則を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_chart(
    chart_type: int,
    data_range: str,
    sheet: str | None = None,
    title: str | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float = 400,
    height: float = 300,
    has_legend: bool = True,
) -> str:
    """グラフを追加します。chart_type: 51=縦棒,57=横棒,4=折れ線,5=円,-4169=散布図。data_range: データ範囲（例: "A1:B10"）。"""
    try:
        result = excel.add_chart(
            chart_type,
            data_range,
            sheet=sheet,
            title=title,
            left=left,
            top=top,
            width=width,
            height=height,
            has_legend=has_legend,
        )
        return f"[{result['sheet']}] グラフを追加しました（タイプ: {chart_type}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_pivot_table(
    source_range: str,
    dest_cell: str,
    table_name: str,
    row_fields: list[str],
    column_fields: list[str] | None = None,
    data_fields: list[dict] | None = None,
    page_fields: list[str] | None = None,
    source_sheet: str | None = None,
    dest_sheet: str | None = None,
) -> str:
    """ピボットテーブルを作成します。row_fields: 行フィールド名のリスト。data_fields: [{"name": "売上", "function": -4157}]形式（function: -4157=合計,-4106=カウント,-4154=平均）。"""
    try:
        result = excel.create_pivot_table(
            source_range,
            dest_cell,
            table_name,
            row_fields,
            column_fields=column_fields,
            data_fields=data_fields,
            page_fields=page_fields,
            source_sheet=source_sheet,
            dest_sheet=dest_sheet,
        )
        return f"ピボットテーブル '{result['table_name']}' を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_comment(cell: str, text: str, sheet: str | None = None) -> str:
    """セルにコメントを追加します。例: cell="A1", text="確認が必要"。"""
    try:
        result = excel.add_comment(cell, text, sheet)
        return f"[{result['sheet']}] {result['cell']} にコメントを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_comments(sheet: str | None = None) -> str:
    """シート内の全コメントを一覧表示します。結果はJSON形式で返します。"""
    try:
        result = excel.get_comments(sheet)
        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_print_setup(
    sheet: str | None = None,
    orientation: int | None = None,
    paper_size: int | None = None,
    fit_to_pages_wide: int | None = None,
    fit_to_pages_tall: int | None = None,
    print_title_rows: str | None = None,
    print_area: str | None = None,
    center_horizontally: bool | None = None,
    center_vertically: bool | None = None,
) -> str:
    """印刷設定を行います。orientation: 1=縦,2=横。paper_size: 1=Letter,9=A4。print_title_rows: 繰り返し行（例: "1:1"）。print_area: 印刷範囲（例: "A1:F20"）。"""
    try:
        result = excel.set_print_setup(
            sheet=sheet,
            orientation=orientation,
            paper_size=paper_size,
            fit_to_pages_wide=fit_to_pages_wide,
            fit_to_pages_tall=fit_to_pages_tall,
            print_title_rows=print_title_rows,
            print_area=print_area,
            center_horizontally=center_horizontally,
            center_vertically=center_vertically,
        )
        return f"[{result['sheet']}] 印刷設定を更新しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_export_to_pdf(output_path: str, sheet: str | None = None) -> str:
    """シートまたはブック全体をPDFファイルに出力します。output_path: 出力先のPDFファイルパス。"""
    try:
        result = excel.export_to_pdf(output_path, sheet)
        return f"PDFを出力しました: {result['output_path']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_protect_sheet(
    password: str | None = None,
    allow_formatting_cells: bool = False,
    allow_sorting: bool = False,
    allow_filtering: bool = False,
    sheet: str | None = None,
) -> str:
    """シートを保護します。password: 保護パスワード（省略可）。allow_*で許可する操作を指定します。"""
    try:
        result = excel.protect_sheet(
            password=password,
            allow_formatting_cells=allow_formatting_cells,
            allow_sorting=allow_sorting,
            allow_filtering=allow_filtering,
            sheet=sheet,
        )
        return f"[{result['sheet']}] シートを保護しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_unprotect_sheet(password: str | None = None, sheet: str | None = None) -> str:
    """シートの保護を解除します。password: 保護時に設定したパスワード（省略可）。"""
    try:
        result = excel.unprotect_sheet(password=password, sheet=sheet)
        return f"[{result['sheet']}] シートの保護を解除しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Sparklines
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_add_sparkline(
    data_range: str,
    location_cell: str,
    sheet: str | None = None,
    sparkline_type: str = "line",
    color: list[int] | None = None,
) -> str:
    """スパークラインを追加します。data_range: データ範囲（例: "A1:A10"）。location_cell: 配置先セル。sparkline_type: "line","column","win_loss"。color: [R,G,B]形式。"""
    try:
        result = excel.add_sparkline(
            sheet=sheet,
            data_range=data_range,
            location_cell=location_cell,
            sparkline_type=sparkline_type,
            color=tuple(color) if color else None,
        )
        return f"[{result['sheet']}] {result['location']} にスパークラインを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_format_sparkline(
    location_cell: str,
    sheet: str | None = None,
    high_point: bool = False,
    low_point: bool = False,
    first_point: bool = False,
    last_point: bool = False,
    negative_points: bool = False,
    markers: bool = False,
    line_weight: float | None = None,
) -> str:
    """スパークラインの表示オプションを設定します。location_cell: スパークラインのセル。high_point/low_point: 最高/最低ポイント表示。markers: マーカー表示。"""
    try:
        result = excel.format_sparkline(
            sheet=sheet,
            location_cell=location_cell,
            high_point=high_point,
            low_point=low_point,
            first_point=first_point,
            last_point=last_point,
            negative_points=negative_points,
            markers=markers,
            line_weight=line_weight,
        )
        return f"[{result['sheet']}] {result['location']} のスパークラインを設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Advanced Conditional Formatting
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_set_conditional_formatting_color_scale(
    range_address: str,
    min_color: list[int],
    mid_color: list[int] | None = None,
    max_color: list[int] | None = None,
    sheet: str | None = None,
) -> str:
    """カラースケール条件付き書式を設定します。2色または3色スケール。min_color/mid_color/max_color: [R,G,B]形式。mid_colorを指定すると3色スケール。"""
    try:
        result = excel.set_conditional_formatting_color_scale(
            sheet=sheet,
            range_str=range_address,
            min_color=tuple(min_color),
            mid_color=tuple(mid_color) if mid_color else None,
            max_color=tuple(max_color) if max_color else None,
        )
        return f"[{result['sheet']}] {result['range']} にカラースケールを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_conditional_formatting_data_bar(
    range_address: str,
    bar_color: list[int],
    show_value: bool = True,
    min_type: int | None = None,
    max_type: int | None = None,
    sheet: str | None = None,
) -> str:
    """データバー条件付き書式を設定します。bar_color: [R,G,B]形式。show_value: セル値を表示するか。"""
    try:
        result = excel.set_conditional_formatting_data_bar(
            sheet=sheet,
            range_str=range_address,
            bar_color=tuple(bar_color),
            show_value=show_value,
            min_type=min_type,
            max_type=max_type,
        )
        return f"[{result['sheet']}] {result['range']} にデータバーを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_conditional_formatting_icon_set(
    range_address: str,
    icon_style: str = "3_arrows",
    reverse: bool = False,
    show_icon_only: bool = False,
    sheet: str | None = None,
) -> str:
    """アイコンセット条件付き書式を設定します。icon_style: "3_arrows","3_traffic_lights","3_stars","4_arrows","5_arrows","3_flags","3_symbols"。"""
    try:
        result = excel.set_conditional_formatting_icon_set(
            sheet=sheet,
            range_str=range_address,
            icon_style=icon_style,
            reverse=reverse,
            show_icon_only=show_icon_only,
        )
        return f"[{result['sheet']}] {result['range']} にアイコンセットを設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Grouping & Outline
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_group_rows(start_row: int, end_row: int, sheet: str | None = None) -> str:
    """行をグループ化します（折りたたみ可能）。start_row: 開始行番号、end_row: 終了行番号。"""
    try:
        result = excel.group_rows(sheet=sheet, start_row=start_row, end_row=end_row)
        return f"[{result['sheet']}] 行 {result['rows']} をグループ化しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_ungroup_rows(start_row: int, end_row: int, sheet: str | None = None) -> str:
    """行のグループ化を解除します。start_row: 開始行番号、end_row: 終了行番号。"""
    try:
        result = excel.ungroup_rows(sheet=sheet, start_row=start_row, end_row=end_row)
        return f"[{result['sheet']}] 行 {result['rows']} のグループ化を解除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_group_columns(start_col: str, end_col: str, sheet: str | None = None) -> str:
    """列をグループ化します（折りたたみ可能）。start_col: 開始列（例: "B"）、end_col: 終了列（例: "D"）。"""
    try:
        result = excel.group_columns(sheet=sheet, start_col=start_col, end_col=end_col)
        return f"[{result['sheet']}] 列 {result['columns']} をグループ化しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_ungroup_columns(start_col: str, end_col: str, sheet: str | None = None) -> str:
    """列のグループ化を解除します。start_col: 開始列、end_col: 終了列。"""
    try:
        result = excel.ungroup_columns(sheet=sheet, start_col=start_col, end_col=end_col)
        return f"[{result['sheet']}] 列 {result['columns']} のグループ化を解除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_outline_level(
    sheet: str | None = None,
    show_detail: bool = True,
    summary_below: bool = True,
    summary_right: bool = True,
) -> str:
    """アウトラインの設定を変更します。summary_below: 集計行を下に配置。summary_right: 集計列を右に配置。"""
    try:
        result = excel.set_outline_level(
            sheet=sheet,
            show_detail=show_detail,
            summary_below=summary_below,
            summary_right=summary_right,
        )
        return f"[{result['sheet']}] アウトライン設定を変更しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_collapse_group(level: int, sheet: str | None = None) -> str:
    """アウトラインを指定レベルまで折りたたみます。level: 表示するレベル（1-8）。"""
    try:
        result = excel.collapse_group(sheet=sheet, level=level)
        return f"[{result['sheet']}] アウトラインをレベル {result['level']} に折りたたみました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Hyperlinks
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_add_hyperlink(
    cell: str,
    url: str,
    display_text: str | None = None,
    tooltip: str | None = None,
    sheet: str | None = None,
) -> str:
    """セルにハイパーリンクを追加します。cell: セルアドレス。url: リンク先URL。display_text: 表示テキスト。tooltip: ツールチップ。"""
    try:
        result = excel.add_hyperlink(
            sheet=sheet,
            cell=cell,
            url=url,
            display_text=display_text,
            tooltip=tooltip,
        )
        return f"[{result['sheet']}] {result['cell']} にハイパーリンクを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_internal_link(
    cell: str,
    target_sheet: str,
    target_cell: str,
    display_text: str | None = None,
    sheet: str | None = None,
) -> str:
    """セルにブック内リンク（別シート/セルへのリンク）を追加します。target_sheet: リンク先シート名。target_cell: リンク先セル。"""
    try:
        result = excel.add_internal_link(
            sheet=sheet,
            cell=cell,
            target_sheet=target_sheet,
            target_cell=target_cell,
            display_text=display_text,
        )
        return f"[{result['sheet']}] {result['cell']} に内部リンクを追加しました（{result['target_sheet']}!{result['target_cell']}）"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Images & Shapes
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_insert_image(
    image_path: str,
    cell: str = "A1",
    width: float | None = None,
    height: float | None = None,
    sheet: str | None = None,
) -> str:
    """ワークシートに画像を挿入します。image_path: 画像ファイルのパス。cell: 配置先セル。width/height: サイズ（ポイント）。"""
    try:
        result = excel.insert_image(
            sheet=sheet,
            image_path=image_path,
            cell=cell,
            width=width,
            height=height,
        )
        return f"[{result['sheet']}] セル {result['cell']} 付近に画像を挿入しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_shape(
    shape_type: int,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: list[int] | None = None,
    line_color: list[int] | None = None,
    text: str | None = None,
    sheet: str | None = None,
) -> str:
    """ワークシートに図形を追加します。shape_type: 1=四角形,5=角丸四角形,9=楕円,4=ひし形。fill_color/line_color: [R,G,B]形式。text: 図形内テキスト。"""
    try:
        result = excel.add_shape(
            sheet=sheet,
            shape_type=shape_type,
            left=left,
            top=top,
            width=width,
            height=height,
            fill_color=tuple(fill_color) if fill_color else None,
            line_color=tuple(line_color) if line_color else None,
            text=text,
        )
        return f"[{result['sheet']}] 図形 '{result['name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_textbox(
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_size: float | None = None,
    font_color: list[int] | None = None,
    bold: bool = False,
    sheet: str | None = None,
) -> str:
    """ワークシートにテキストボックスを追加します。left/top: 位置（ポイント）。width/height: サイズ。font_color: [R,G,B]形式。"""
    try:
        result = excel.add_textbox(
            sheet=sheet,
            left=left,
            top=top,
            width=width,
            height=height,
            text=text,
            font_size=font_size,
            font_color=tuple(font_color) if font_color else None,
            bold=bold,
        )
        return f"[{result['sheet']}] テキストボックス '{result['name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Advanced Cell Formatting
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_set_cell_style(range_address: str, style_name: str, sheet: str | None = None) -> str:
    """セルに組み込みスタイルを適用します。style_name: "Heading 1","Total","Accent1"など。range_address: 対象範囲。"""
    try:
        result = excel.set_cell_style(sheet=sheet, range_str=range_address, style_name=style_name)
        return f"[{result['sheet']}] {result['range']} にスタイル '{result['style']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_auto_fit_columns(range_address: str | None = None, sheet: str | None = None) -> str:
    """列幅をコンテンツに合わせて自動調整します。range_address: 対象範囲（省略時は使用範囲全体）。"""
    try:
        result = excel.auto_fit_columns(sheet=sheet, range_str=range_address)
        return f"[{result['sheet']}] {result['range']} の列幅を自動調整しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_auto_fit_rows(range_address: str | None = None, sheet: str | None = None) -> str:
    """行の高さをコンテンツに合わせて自動調整します。range_address: 対象範囲（省略時は使用範囲全体）。"""
    try:
        result = excel.auto_fit_rows(sheet=sheet, range_str=range_address)
        return f"[{result['sheet']}] {result['range']} の行の高さを自動調整しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_cell_indent(range_address: str, indent_level: int, sheet: str | None = None) -> str:
    """セルのインデントレベルを設定します。indent_level: インデントレベル（0以上）。"""
    try:
        result = excel.set_cell_indent(sheet=sheet, range_str=range_address, indent_level=indent_level)
        return f"[{result['sheet']}] {result['range']} のインデントを {result['indent']} に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_text_rotation(range_address: str, angle: int, sheet: str | None = None) -> str:
    """セルのテキスト回転角度を設定します。angle: -90〜90度。"""
    try:
        result = excel.set_text_rotation(sheet=sheet, range_str=range_address, angle=angle)
        return f"[{result['sheet']}] {result['range']} のテキスト回転を {result['angle']}度 に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_wrap_text(range_address: str, wrap: bool = True, sheet: str | None = None) -> str:
    """セルのテキスト折り返しを設定します。wrap: Trueで折り返し有効。"""
    try:
        result = excel.set_wrap_text(sheet=sheet, range_str=range_address, wrap=wrap)
        return f"[{result['sheet']}] {result['range']} の折り返しを {'有効' if result['wrap'] else '無効'} にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_cell_pattern(
    range_address: str,
    pattern_type: int,
    fore_color: list[int] | None = None,
    back_color: list[int] | None = None,
    sheet: str | None = None,
) -> str:
    """セルのパターン塗りつぶしを設定します。pattern_type: パターン種類（1-18）。fore_color/back_color: [R,G,B]形式。"""
    try:
        result = excel.set_cell_pattern(
            sheet=sheet,
            range_str=range_address,
            pattern_type=pattern_type,
            fore_color=tuple(fore_color) if fore_color else None,
            back_color=tuple(back_color) if back_color else None,
        )
        return f"[{result['sheet']}] {result['range']} にパターンを設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Data Features (Advanced)
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_copy_range(
    source_range: str,
    target_cell: str,
    target_sheet: str | None = None,
    sheet: str | None = None,
) -> str:
    """範囲を別の場所にコピーします。source_range: コピー元範囲。target_cell: コピー先セル。target_sheet: コピー先シート（省略時は同じシート）。"""
    try:
        result = excel.copy_range(
            sheet=sheet,
            source_range=source_range,
            target_cell=target_cell,
            target_sheet=target_sheet,
        )
        return f"[{result['sheet']}] {result['source']} を [{result['target_sheet']}] {result['target']} にコピーしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_clear_range(
    range_address: str,
    clear_type: str = "all",
    sheet: str | None = None,
) -> str:
    """範囲をクリアします。clear_type: "all"=全て,"contents"=値のみ,"formats"=書式のみ,"comments"=コメントのみ。"""
    try:
        result = excel.clear_range(sheet=sheet, range_str=range_address, clear_type=clear_type)
        return f"[{result['sheet']}] {result['range']} をクリアしました（{result['clear_type']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_find_value(
    value: str,
    match_case: bool = False,
    match_entire: bool = False,
    sheet: str | None = None,
) -> str:
    """シート内で値を検索します。value: 検索値。match_case: 大文字小文字を区別。match_entire: セル全体と一致。"""
    try:
        result = excel.find_value(sheet=sheet, value=value, match_case=match_case, match_entire=match_entire)
        if result["found"]:
            return f"[{result['sheet']}] '{result['value']}' が見つかりました: {result['found']}"
        else:
            return f"[{result['sheet']}] '{result['value']}' は見つかりませんでした"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_replace_value(
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_entire: bool = False,
    sheet: str | None = None,
) -> str:
    """シート内で値を検索して置換します。find_text: 検索文字列。replace_text: 置換文字列。"""
    try:
        result = excel.replace_value(
            sheet=sheet,
            find_text=find_text,
            replace_text=replace_text,
            match_case=match_case,
            match_entire=match_entire,
        )
        return f"[{result['sheet']}] '{result['find']}' を '{result['replace']}' に置換しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_remove_duplicates(
    range_address: str,
    columns: list[int] | None = None,
    sheet: str | None = None,
) -> str:
    """範囲から重複行を削除します。columns: チェックする列番号リスト（1始まり、省略時は全列）。"""
    try:
        result = excel.remove_duplicates(sheet=sheet, range_str=range_address, columns=columns)
        return f"[{result['sheet']}] {result['range']} の重複を削除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_text_to_columns(
    range_address: str,
    delimiter: str = ",",
    sheet: str | None = None,
) -> str:
    """テキストを列に分割します。delimiter: 区切り文字（","、"\\t"、";"、" "など）。"""
    try:
        result = excel.text_to_columns(sheet=sheet, range_str=range_address, delimiter=delimiter)
        return f"[{result['sheet']}] {result['range']} を区切り文字 '{result['delimiter']}' で分割しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_transpose_range(
    source_range: str,
    target_cell: str,
    sheet: str | None = None,
) -> str:
    """データを転置（行列入替）します。source_range: 元の範囲。target_cell: 転置先セル。"""
    try:
        result = excel.transpose_range(sheet=sheet, source_range=source_range, target_cell=target_cell)
        return f"[{result['sheet']}] {result['source']} を {result['target']} に転置しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Advanced Chart
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_format_chart(
    chart_index: int = 1,
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    legend_position: str | None = None,
    style: int | None = None,
    sheet: str | None = None,
) -> str:
    """既存のグラフを書式設定します。chart_index: グラフ番号（1始まり）。legend_position: "bottom","top","left","right","none"。style: スタイル番号。"""
    try:
        result = excel.format_chart(
            sheet=sheet,
            chart_index=chart_index,
            title=title,
            x_axis_title=x_axis_title,
            y_axis_title=y_axis_title,
            legend_position=legend_position,
            style=style,
        )
        return f"[{result['sheet']}] グラフ '{result['chart']}' を書式設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_format_chart_series(
    chart_index: int = 1,
    series_index: int = 1,
    color: list[int] | None = None,
    line_weight: float | None = None,
    marker_style: int | None = None,
    marker_size: int | None = None,
    sheet: str | None = None,
) -> str:
    """グラフの個別データ系列を書式設定します。series_index: 系列番号（1始まり）。color: [R,G,B]形式。marker_style: -4142=なし,8=丸,1=四角。"""
    try:
        result = excel.format_chart_series(
            sheet=sheet,
            chart_index=chart_index,
            series_index=series_index,
            color=tuple(color) if color else None,
            line_weight=line_weight,
            marker_style=marker_style,
            marker_size=marker_size,
        )
        return f"[{result['sheet']}] グラフのデータ系列 {result['series_index']} を書式設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_chart_trendline(
    chart_index: int = 1,
    series_index: int = 1,
    trend_type: str = "linear",
    display_equation: bool = False,
    display_r_squared: bool = False,
    sheet: str | None = None,
) -> str:
    """グラフにトレンドラインを追加します。trend_type: "linear","exponential","logarithmic","polynomial","power","moving_average"。"""
    try:
        result = excel.add_chart_trendline(
            sheet=sheet,
            chart_index=chart_index,
            series_index=series_index,
            trend_type=trend_type,
            display_equation=display_equation,
            display_r_squared=display_r_squared,
        )
        return f"[{result['sheet']}] グラフにトレンドライン（{result['trend_type']}）を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_chart_area_format(
    chart_index: int = 1,
    fill_color: list[int] | None = None,
    border_color: list[int] | None = None,
    border_weight: float | None = None,
    sheet: str | None = None,
) -> str:
    """グラフエリアの背景と枠線を書式設定します。fill_color/border_color: [R,G,B]形式。"""
    try:
        result = excel.set_chart_area_format(
            sheet=sheet,
            chart_index=chart_index,
            fill_color=tuple(fill_color) if fill_color else None,
            border_color=tuple(border_color) if border_color else None,
            border_weight=border_weight,
        )
        return f"[{result['sheet']}] グラフエリアを書式設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Pivot Table Advanced
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_format_pivot_table(
    pivot_name: str,
    style: str | None = None,
    show_grand_total_rows: bool | None = None,
    show_grand_total_cols: bool | None = None,
    repeat_item_labels: bool | None = None,
    sheet: str | None = None,
) -> str:
    """ピボットテーブルを書式設定します。style: スタイル名（例: "PivotStyleMedium9"）。show_grand_total_rows/cols: 総計表示。"""
    try:
        result = excel.format_pivot_table(
            sheet=sheet,
            pivot_name=pivot_name,
            style=style,
            show_grand_total_rows=show_grand_total_rows,
            show_grand_total_cols=show_grand_total_cols,
            repeat_item_labels=repeat_item_labels,
        )
        return f"[{result['sheet']}] ピボットテーブル '{result['pivot_name']}' を書式設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_pivot_field(
    pivot_name: str,
    field_name: str,
    area: str = "row",
    position: int | None = None,
    sheet: str | None = None,
) -> str:
    """ピボットテーブルにフィールドを追加します。area: "row","column","data","filter"。position: エリア内の位置。"""
    try:
        result = excel.add_pivot_field(
            sheet=sheet,
            pivot_name=pivot_name,
            field_name=field_name,
            area=area,
            position=position,
        )
        return f"[{result['sheet']}] ピボットテーブル '{result['pivot_name']}' にフィールド '{result['field']}' を追加しました（{result['area']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_refresh_pivot_table(pivot_name: str | None = None, sheet: str | None = None) -> str:
    """ピボットテーブルを更新します。pivot_name: 特定のピボットテーブル名（省略時は全て更新）。"""
    try:
        result = excel.refresh_pivot_table(sheet=sheet, pivot_name=pivot_name)
        if "pivot_name" in result:
            return f"[{result['sheet']}] ピボットテーブル '{result['pivot_name']}' を更新しました"
        else:
            return f"[{result['sheet']}] {result['refreshed_count']} 個のピボットテーブルを更新しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Workbook Features
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_set_workbook_properties(
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
) -> str:
    """ブックのドキュメントプロパティを設定します。title: タイトル。author: 作成者。subject: 件名。keywords: キーワード。"""
    try:
        result = excel.set_workbook_properties(
            title=title,
            author=author,
            subject=subject,
            keywords=keywords,
            comments=comments,
        )
        return f"ブック '{result['name']}' のプロパティを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_workbook_statistics() -> str:
    """ブックの統計情報を取得します（シート数、使用範囲など）。結果はJSON形式で返します。"""
    try:
        result = excel.get_workbook_statistics()
        return json.dumps(result, ensure_ascii=False, default=str)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_tab_color(color: list[int], sheet: str | None = None) -> str:
    """ワークシートのタブの色を設定します。color: [R,G,B]形式（例: [255,0,0]で赤）。"""
    try:
        result = excel.set_tab_color(sheet=sheet, color=tuple(color))
        return f"[{result['sheet']}] タブの色を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_hide_sheet(sheet: str) -> str:
    """ワークシートを非表示にします。sheet: 非表示にするシート名。"""
    try:
        result = excel.hide_sheet(sheet)
        return f"シート '{result['sheet']}' を非表示にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_unhide_sheet(sheet: str) -> str:
    """非表示のワークシートを再表示します。sheet: 再表示するシート名。"""
    try:
        result = excel.unhide_sheet(sheet)
        return f"シート '{result['sheet']}' を再表示しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_copy_sheet(
    source_sheet: str,
    target_name: str | None = None,
    before: str | None = None,
    after: str | None = None,
) -> str:
    """ワークシートをコピーします。source_sheet: コピー元シート名。target_name: 新しいシート名。before/after: 配置位置のシート名。"""
    try:
        result = excel.copy_sheet(
            source_sheet=source_sheet,
            target_name=target_name,
            before=before,
            after=after,
        )
        return f"シート '{result['source']}' を '{result['new_sheet']}' としてコピーしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_move_sheet(
    sheet: str,
    before: str | None = None,
    after: str | None = None,
) -> str:
    """ワークシートを移動します。sheet: 移動するシート名。before/after: 移動先のシート名。"""
    try:
        result = excel.move_sheet(sheet=sheet, before=before, after=after)
        return f"シート '{result['sheet']}' を移動しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Print & Page Setup Advanced
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_set_print_area(range_address: str, sheet: str | None = None) -> str:
    """印刷範囲を設定します。range_address: 印刷範囲（例: "A1:F20"）。"""
    try:
        result = excel.set_print_area(sheet=sheet, range_str=range_address)
        return f"[{result['sheet']}] 印刷範囲を {result['print_area']} に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_print_titles(
    rows: str | None = None,
    columns: str | None = None,
    sheet: str | None = None,
) -> str:
    """各ページに繰り返し印刷する行/列を設定します。rows: 繰り返し行（例: "1:2"）。columns: 繰り返し列（例: "A:B"）。"""
    try:
        result = excel.set_print_titles(sheet=sheet, rows=rows, columns=columns)
        return f"[{result['sheet']}] 印刷タイトルを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_page_break(
    row: int | None = None,
    col: int | None = None,
    sheet: str | None = None,
) -> str:
    """改ページを挿入します。row: 水平改ページの行番号。col: 垂直改ページの列番号。"""
    try:
        result = excel.add_page_break(sheet=sheet, row=row, col=col)
        msg = f"[{result['sheet']}] 改ページを挿入しました"
        if result['row']:
            msg += f"（行: {result['row']}）"
        if result['col']:
            msg += f"（列: {result['col']}）"
        return msg
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Protection Advanced
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_protect_workbook(
    password: str | None = None,
    structure: bool = True,
    windows: bool = False,
) -> str:
    """ブック全体を保護します。structure: 構造の保護（シートの追加/削除を防止）。windows: ウィンドウの保護。"""
    try:
        result = excel.protect_workbook(password=password, structure=structure, windows=windows)
        return f"ブック '{result['name']}' を保護しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_unprotect_workbook(password: str | None = None) -> str:
    """ブックの保護を解除します。password: 保護時に設定したパスワード（省略可）。"""
    try:
        result = excel.unprotect_workbook(password=password)
        return f"ブック '{result['name']}' の保護を解除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_lock_cells(
    range_address: str,
    locked: bool = True,
    sheet: str | None = None,
) -> str:
    """セルのロック/ロック解除を設定します（シート保護と併用）。locked: Trueでロック、Falseでロック解除。"""
    try:
        result = excel.lock_cells(sheet=sheet, range_str=range_address, locked=locked)
        status = "ロック" if result['locked'] else "ロック解除"
        return f"[{result['sheet']}] {result['range']} を{status}しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Advanced Formula Functions
# ---------------------------------------------------------------------------


@mcp.tool()
def excel_set_array_formula(
    range_str: str,
    formula: str,
    sheet: str | None = None,
) -> str:
    """配列数式を設定します。CSE配列数式またはダイナミック配列に対応。"""
    try:
        result = excel.set_array_formula(sheet=sheet, range_str=range_str, formula=formula)
        return f"[{result['sheet']}] {result['range']} に配列数式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_evaluate_formula(
    formula: str,
    sheet: str | None = None,
) -> str:
    """数式を評価して結果を返します（セルに配置せずに計算）。"""
    try:
        result = excel.evaluate_formula(sheet=sheet, formula=formula)
        return f"数式 {result['formula']} の結果: {result['result']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_formula_range(
    start_cell: str,
    formulas: list[list[str]],
    sheet: str | None = None,
) -> str:
    """複数の数式を一括設定します。formulasは2次元リスト。"""
    try:
        result = excel.set_formula_range(sheet=sheet, start_cell=start_cell, formulas=formulas)
        return f"[{result['sheet']}] {result['start_cell']}から{result['rows']}行x{result['cols']}列の数式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_named_formula(
    name: str,
    formula: str,
    sheet: str | None = None,
) -> str:
    """名前付き数式を作成します（名前付き範囲ではなく数式）。"""
    try:
        result = excel.create_named_formula(name=name, formula=formula, sheet=sheet)
        return f"名前付き数式 '{result['name']}' を作成しました（スコープ: {result['scope']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_calculate_statistics(
    range_str: str,
    sheet: str | None = None,
) -> str:
    """範囲の記述統計量を計算します（合計、平均、最小、最大、件数、標準偏差、中央値）。"""
    try:
        result = excel.calculate_statistics(sheet=sheet, range_str=range_str)
        lines = [f"[{result['sheet']}] {result['range']} の統計:"]
        for k, label in [("sum","合計"),("avg","平均"),("min","最小"),("max","最大"),("count","件数"),("stdev","標準偏差"),("median","中央値")]:
            lines.append(f"  {label}: {result[k]}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_frequency_distribution(
    data_range: str,
    bins_range: str,
    output_cell: str,
    sheet: str | None = None,
) -> str:
    """度数分布を作成します。"""
    try:
        result = excel.create_frequency_distribution(sheet=sheet, data_range=data_range, bins_range=bins_range, output_cell=output_cell)
        return f"[{result['sheet']}] {result['output_range']} に度数分布を作成しました（{result['bins_count']}ビン）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_goal_seek(
    target_cell: str,
    target_value: float,
    changing_cell: str,
    sheet: str | None = None,
) -> str:
    """ゴールシーク分析を実行します。目標値に達するように変化セルを調整します。"""
    try:
        result = excel.goal_seek(sheet=sheet, target_cell=target_cell, target_value=target_value, changing_cell=changing_cell)
        return f"[{result['sheet']}] ゴールシーク完了: {result['target_cell']}={result['achieved_value']} (変化セル {result['changing_cell']}={result['changing_value']})"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_data_table_analysis(
    formula_cell: str,
    row_input_cell: str | None = None,
    col_input_cell: str | None = None,
    row_values_range: str | None = None,
    col_values_range: str | None = None,
    sheet: str | None = None,
) -> str:
    """What-Ifデータテーブル（感度分析）を作成します。"""
    try:
        result = excel.create_data_table_analysis(
            sheet=sheet, row_input_cell=row_input_cell, col_input_cell=col_input_cell,
            formula_cell=formula_cell, row_values_range=row_values_range, col_values_range=col_values_range,
        )
        return f"[{result['sheet']}] データテーブルを作成しました: {result['table_address']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_rich_text_cell(
    cell: str,
    runs: list[dict],
    sheet: str | None = None,
) -> str:
    """セル内にリッチテキスト（複数書式）を設定します。runs=[{"text":"Hello","bold":true,"color":[255,0,0]},...]"""
    try:
        result = excel.set_rich_text_cell(sheet=sheet, cell=cell, runs=runs)
        return f"[{result['sheet']}] {result['cell']} にリッチテキストを設定しました（{result['runs_count']}ラン）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_cell_dropdown(
    cell: str,
    items: list[str],
    show_error: bool = True,
    sheet: str | None = None,
) -> str:
    """セルにドロップダウンリストを追加します。"""
    try:
        result = excel.add_cell_dropdown(sheet=sheet, cell=cell, items=items, show_error=show_error)
        return f"[{result['sheet']}] {result['cell']} にドロップダウンを追加しました（{result['items_count']}項目）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_cell_hyperlink_format(
    cell: str,
    display_text: str | None = None,
    color: list[int] | None = None,
    underline: bool = True,
    sheet: str | None = None,
) -> str:
    """ハイパーリンクセルの書式を設定します。color=[R,G,B]。"""
    try:
        result = excel.set_cell_hyperlink_format(
            sheet=sheet, cell=cell, display_text=display_text,
            color=tuple(color) if color else None, underline=underline,
        )
        return f"[{result['sheet']}] {result['cell']} のハイパーリンク書式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_clear_all_formatting(
    range_str: str | None = None,
    sheet: str | None = None,
) -> str:
    """範囲またはシート全体の書式をクリアします。"""
    try:
        result = excel.clear_all_formatting(sheet=sheet, range_str=range_str)
        return f"[{result['sheet']}] {result['range']} の書式をクリアしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_combo_chart(
    data_range: str,
    chart_types: list[str],
    series_on_secondary: list[int] | None = None,
    position_cell: str = "E1",
    width: float = 480,
    height: float = 300,
    sheet: str | None = None,
) -> str:
    """コンボチャート（複合グラフ）を追加します。chart_types: 系列ごとの型("column","line","area")。"""
    try:
        result = excel.add_combo_chart(
            sheet=sheet, data_range=data_range, chart_types=chart_types,
            series_on_secondary=series_on_secondary, position_cell=position_cell, width=width, height=height,
        )
        return f"[{result['sheet']}] コンボチャート '{result['chart_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_pie_chart(
    data_range: str,
    position_cell: str = "E1",
    width: float = 400,
    height: float = 300,
    explode: list[int] | None = None,
    show_percentage: bool = True,
    sheet: str | None = None,
) -> str:
    """円グラフを追加します。explode: 分離するデータポイントのインデックスリスト(1始まり)。"""
    try:
        result = excel.add_pie_chart(
            sheet=sheet, data_range=data_range, position_cell=position_cell,
            width=width, height=height, explode=explode, show_percentage=show_percentage,
        )
        return f"[{result['sheet']}] 円グラフ '{result['chart_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_scatter_chart(
    data_range: str,
    position_cell: str = "E1",
    width: float = 480,
    height: float = 300,
    show_trendline: bool = False,
    bubble: bool = False,
    sheet: str | None = None,
) -> str:
    """散布図またはバブルチャートを追加します。"""
    try:
        result = excel.add_scatter_chart(
            sheet=sheet, data_range=data_range, position_cell=position_cell,
            width=width, height=height, show_trendline=show_trendline, bubble=bubble,
        )
        return f"[{result['sheet']}] 散布図 '{result['chart_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_stock_chart(
    data_range: str,
    position_cell: str = "E1",
    width: float = 480,
    height: float = 300,
    chart_subtype: str = "hlc",
    sheet: str | None = None,
) -> str:
    """株価チャートを追加します。chart_subtype: "hlc"(高値-安値-終値)、"ohlc"(始値-高値-安値-終値)。"""
    try:
        result = excel.add_stock_chart(
            sheet=sheet, data_range=data_range, position_cell=position_cell,
            width=width, height=height, chart_subtype=chart_subtype,
        )
        return f"[{result['sheet']}] 株価チャート '{result['chart_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_radar_chart(
    data_range: str,
    position_cell: str = "E1",
    width: float = 400,
    height: float = 300,
    filled: bool = False,
    sheet: str | None = None,
) -> str:
    """レーダーチャート（スパイダーチャート）を追加します。"""
    try:
        result = excel.add_radar_chart(
            sheet=sheet, data_range=data_range, position_cell=position_cell,
            width=width, height=height, filled=filled,
        )
        return f"[{result['sheet']}] レーダーチャート '{result['chart_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_activate_sheet(sheet: str) -> str:
    """指定シートをアクティブ（表示）にします。"""
    try:
        result = excel.activate_sheet(sheet=sheet)
        return f"シート '{result['sheet']}' をアクティブにしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_active_sheet() -> str:
    """現在アクティブなシート名を取得します。"""
    try:
        result = excel.get_active_sheet()
        return f"アクティブシート: {result['sheet']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_used_range(sheet: str | None = None) -> str:
    """使用範囲のアドレスとサイズを取得します。"""
    try:
        result = excel.get_used_range(sheet=sheet)
        return f"[{result['sheet']}] 使用範囲: {result['address']} ({result['rows']}行 x {result['columns']}列)"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_last_row(column: str = "A", sheet: str | None = None) -> str:
    """指定列の最終使用行番号を取得します。"""
    try:
        result = excel.get_last_row(sheet=sheet, column=column)
        return f"[{result['sheet']}] {result['column']}列の最終行: {result['last_row']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_last_column(row: int = 1, sheet: str | None = None) -> str:
    """指定行の最終使用列番号を取得します。"""
    try:
        result = excel.get_last_column(sheet=sheet, row=row)
        return f"[{result['sheet']}] {result['row']}行目の最終列: {result['column_letter']} (列{result['last_column']})"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_fill_series(
    start_cell: str,
    end_cell: str,
    fill_type: str = "linear",
    step: float = 1,
    sheet: str | None = None,
) -> str:
    """連続データを自動入力します。fill_type: "linear","growth","date","auto"。"""
    try:
        result = excel.fill_series(sheet=sheet, start_cell=start_cell, end_cell=end_cell, fill_type=fill_type, step=step)
        return f"[{result['sheet']}] {result['start']}から{result['end']}に{result['fill_type']}系列を入力しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_concatenate_range(
    range_str: str,
    separator: str = ", ",
    sheet: str | None = None,
) -> str:
    """範囲内の値を連結して1つの文字列にします。"""
    try:
        result = excel.concatenate_range(sheet=sheet, range_str=range_str, separator=separator)
        return f"[{result['sheet']}] 連結結果（{result['count']}セル）: {result['result']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_split_text_by_rows(
    cell: str,
    separator: str | None = None,
    target_cell: str | None = None,
    sheet: str | None = None,
) -> str:
    """セルのテキストを分割して複数行に展開します。"""
    try:
        result = excel.split_text_by_rows(sheet=sheet, cell=cell, separator=separator, target_cell=target_cell)
        return f"[{result['sheet']}] {result['parts']}個に分割しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_apply_formula_to_range(
    range_str: str,
    formula_template: str,
    sheet: str | None = None,
) -> str:
    """範囲内の各セルに数式パターンを適用します。{row}と{col}がプレースホルダー。例: "=A{row}*B{row}"。"""
    try:
        result = excel.apply_formula_to_range(sheet=sheet, range_str=range_str, formula_template=formula_template)
        return f"[{result['sheet']}] {result['range']} に数式パターンを適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_sequence(
    start_cell: str,
    count: int,
    start_value: float = 1,
    step: float = 1,
    direction: str = "down",
    sheet: str | None = None,
) -> str:
    """数列を生成します。direction: "down"(縦)、"right"(横)。"""
    try:
        result = excel.create_sequence(
            sheet=sheet, start_cell=start_cell, count=count,
            start_value=start_value, step=step, direction=direction,
        )
        return f"[{result['sheet']}] {result['start_cell']}から{result['count']}個の数列を生成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_highlight_cells(
    range_str: str,
    condition: str,
    color: list[int],
    sheet: str | None = None,
) -> str:
    """条件に合うセルをハイライトします。condition: ">50","<0","=100","contains:text","empty","not_empty"。color=[R,G,B]。"""
    try:
        result = excel.highlight_cells(sheet=sheet, range_str=range_str, condition=condition, color=color)
        return f"[{result['sheet']}] {result['range']} で{result['highlighted']}セルをハイライトしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_count_if(
    range_str: str,
    criteria: str,
    sheet: str | None = None,
) -> str:
    """条件に合うセル数をカウントします（COUNTIF）。"""
    try:
        result = excel.count_if(sheet=sheet, range_str=range_str, criteria=criteria)
        return f"[{result['sheet']}] 条件'{result['criteria']}'に合うセル: {result['count']}件"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_sum_if(
    range_str: str,
    criteria: str,
    sum_range: str | None = None,
    sheet: str | None = None,
) -> str:
    """条件に合うセルの合計を計算します（SUMIF）。"""
    try:
        result = excel.sum_if(sheet=sheet, range_str=range_str, criteria=criteria, sum_range=sum_range)
        return f"[{result['sheet']}] 条件'{result['criteria']}'の合計: {result['sum']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_gridlines_visible(visible: bool = True, sheet: str | None = None) -> str:
    """グリッド線の表示/非表示を切り替えます。"""
    try:
        result = excel.set_gridlines_visible(sheet=sheet, visible=visible)
        status = "表示" if result['visible'] else "非表示"
        return f"[{result['sheet']}] グリッド線を{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_headings_visible(visible: bool = True, sheet: str | None = None) -> str:
    """行列見出しの表示/非表示を切り替えます。"""
    try:
        result = excel.set_headings_visible(sheet=sheet, visible=visible)
        status = "表示" if result['visible'] else "非表示"
        return f"[{result['sheet']}] 行列見出しを{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_zoom_level(zoom_percent: int, sheet: str | None = None) -> str:
    """ズームレベルを設定します（10-400%）。"""
    try:
        result = excel.set_zoom_level(sheet=sheet, zoom_percent=zoom_percent)
        return f"[{result['sheet']}] ズームを{result['zoom']}%に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_sheet_direction(direction: str = "ltr", sheet: str | None = None) -> str:
    """シートの読み取り方向を設定します。direction: "ltr"(左→右)、"rtl"(右→左)。"""
    try:
        result = excel.set_sheet_direction(sheet=sheet, direction=direction)
        return f"[{result['sheet']}] 方向を{result['direction']}に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_number_validation(
    range_str: str,
    min_value: float | None = None,
    max_value: float | None = None,
    input_message: str | None = None,
    error_message: str | None = None,
    sheet: str | None = None,
) -> str:
    """数値範囲の入力規則を追加します。"""
    try:
        result = excel.add_number_validation(
            sheet=sheet, range_str=range_str, min_value=min_value,
            max_value=max_value, input_message=input_message, error_message=error_message,
        )
        return f"[{result['sheet']}] {result['range']} に数値バリデーションを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_date_validation(
    range_str: str,
    min_date: str | None = None,
    max_date: str | None = None,
    input_message: str | None = None,
    sheet: str | None = None,
) -> str:
    """日付範囲の入力規則を追加します。日付形式: "2024-01-01"。"""
    try:
        result = excel.add_date_validation(
            sheet=sheet, range_str=range_str, min_date=min_date,
            max_date=max_date, input_message=input_message,
        )
        return f"[{result['sheet']}] {result['range']} に日付バリデーションを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_text_length_validation(
    range_str: str,
    min_length: int | None = None,
    max_length: int | None = None,
    input_message: str | None = None,
    sheet: str | None = None,
) -> str:
    """文字数制限の入力規則を追加します。"""
    try:
        result = excel.add_text_length_validation(
            sheet=sheet, range_str=range_str, min_length=min_length,
            max_length=max_length, input_message=input_message,
        )
        return f"[{result['sheet']}] {result['range']} にテキスト長バリデーションを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_trace_precedents(cell: str, sheet: str | None = None) -> str:
    """数式が参照しているセル（参照元）を取得します。"""
    try:
        result = excel.trace_precedents(sheet=sheet, cell=cell)
        if result['precedents']:
            return f"[{result['sheet']}] {result['cell']} の参照元: {', '.join(result['precedents'])}"
        return f"[{result['sheet']}] {result['cell']} には参照元がありません"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_trace_dependents(cell: str, sheet: str | None = None) -> str:
    """このセルを参照しているセル（参照先）を取得します。"""
    try:
        result = excel.trace_dependents(sheet=sheet, cell=cell)
        if result['dependents']:
            return f"[{result['sheet']}] {result['cell']} の参照先: {', '.join(result['dependents'])}"
        return f"[{result['sheet']}] {result['cell']} を参照しているセルはありません"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_check_errors(range_str: str | None = None, sheet: str | None = None) -> str:
    """範囲内のエラーセル（#N/A, #VALUE!等）を検出します。"""
    try:
        result = excel.check_errors(sheet=sheet, range_str=range_str)
        if result['errors']:
            lines = [f"[{result['sheet']}] {result['count']}個のエラーを検出:"]
            for err in result['errors'][:20]:
                lines.append(f"  {err['cell']}: {err['error']}")
            return "\n".join(lines)
        return f"[{result['sheet']}] エラーは検出されませんでした"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_cell_formula(cell: str, sheet: str | None = None) -> str:
    """セルの数式を取得します（計算結果ではなく数式そのもの）。"""
    try:
        result = excel.get_cell_formula(sheet=sheet, cell=cell)
        if result['has_formula']:
            return f"[{result['sheet']}] {result['cell']} の数式: {result['formula']}"
        return f"[{result['sheet']}] {result['cell']} には数式がありません（値: {result['formula']}）"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# Remaining Pro Features
# ---------------------------------------------------------------------------

@mcp.tool()
def excel_insert_subtotals(
    group_column: int,
    sum_columns: list[int],
    subtotal_function: str = "sum",
    sheet: str | None = None,
) -> str:
    """グループ列で自動小計を挿入します。subtotal_function: "sum", "count", "average", "max", "min"。"""
    try:
        result = excel.insert_subtotals(
            sheet=sheet, group_column=group_column,
            sum_columns=sum_columns, subtotal_function=subtotal_function,
        )
        return f"[{result['sheet']}] 小計を挿入しました（関数: {result['subtotal_function']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_dropdown_list(
    cell_range: str,
    source_range: str,
    sheet: str | None = None,
) -> str:
    """セル範囲参照からドロップダウンリストを作成します。"""
    try:
        result = excel.create_dropdown_list(sheet=sheet, cell_range=cell_range, source_range=source_range)
        return f"[{result['sheet']}] {result['range']} にドロップダウンリストを作成しました（ソース: {result['source']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_conditional_icon(
    range_str: str,
    icon_type: str,
    thresholds: list[float] | None = None,
    sheet: str | None = None,
) -> str:
    """アイコンセット条件付き書式を適用します。icon_type: "arrows", "circles", "flags", "stars"。"""
    try:
        result = excel.set_conditional_icon(
            sheet=sheet, range_str=range_str,
            icon_type=icon_type, thresholds=thresholds,
        )
        return f"[{result['sheet']}] {result['range']} にアイコンセット '{result['icon_type']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_add_error_bars(
    chart_index: int,
    series_index: int = 1,
    error_type: str = "percentage",
    amount: float = 5,
    sheet: str | None = None,
) -> str:
    """グラフ系列にエラーバーを追加します。error_type: "percentage", "fixed", "standard_deviation", "standard_error"。"""
    try:
        result = excel.add_error_bars(
            sheet=sheet, chart_index=chart_index,
            series_index=series_index, error_type=error_type, amount=amount,
        )
        return (
            f"[{result['sheet']}] グラフ {result['chart_index']} の系列 {result['series_index']} "
            f"にエラーバーを追加しました（{result['error_type']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_set_chart_gradient(
    chart_index: int,
    series_index: int,
    color1: list[int],
    color2: list[int],
    sheet: str | None = None,
) -> str:
    """グラフ系列にグラデーション塗りつぶしを適用します。"""
    try:
        result = excel.set_chart_gradient(
            sheet=sheet, chart_index=chart_index, series_index=series_index,
            color1=tuple(color1), color2=tuple(color2),
        )
        return (
            f"[{result['sheet']}] グラフ {result['chart_index']} の系列 {result['series_index']} "
            f"にグラデーションを適用しました"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_named_style(
    style_name: str,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: list[int] | None = None,
    fill_color: list[int] | None = None,
    bold: bool = False,
    borders: bool = False,
    sheet: str | None = None,
) -> str:
    """再利用可能なセルスタイルを作成します。"""
    try:
        result = excel.create_named_style(
            sheet=sheet, style_name=style_name,
            font_name=font_name, font_size=font_size,
            font_color=tuple(font_color) if font_color else None,
            fill_color=tuple(fill_color) if fill_color else None,
            bold=bold, borders=borders,
        )
        return f"スタイル '{result['style_name']}' を作成しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_apply_alternating_colors(
    range_str: str,
    color1: list[int],
    color2: list[int],
    sheet: str | None = None,
) -> str:
    """行の交互色（ゼブラストライプ）を適用します。"""
    try:
        result = excel.apply_alternating_colors(
            sheet=sheet, range_str=range_str,
            color1=tuple(color1), color2=tuple(color2),
        )
        return f"[{result['sheet']}] {result['range']} に交互色を適用しました（{result['rows_formatted']}行）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_get_distinct_values(range_str: str, sheet: str | None = None) -> str:
    """範囲からユニークな値を取得します。"""
    try:
        result = excel.get_distinct_values(sheet=sheet, range_str=range_str)
        if result['values']:
            lines = [f"[{result['sheet']}] {result['count']}個のユニーク値:"]
            for v in result['values'][:50]:
                lines.append(f"  {v}")
            return "\n".join(lines)
        return f"[{result['sheet']}] 値が見つかりませんでした"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_vlookup(
    lookup_value: str,
    table_range: str,
    col_index: int,
    exact_match: bool = True,
    sheet: str | None = None,
) -> str:
    """VLOOKUPを実行して結果を返します。"""
    try:
        result = excel.vlookup(
            sheet=sheet, lookup_value=lookup_value,
            table_range=table_range, col_index=col_index,
            exact_match=exact_match,
        )
        if result['found']:
            return f"[{result['sheet']}] VLOOKUP結果: '{result['lookup_value']}' → {result['result']}"
        return f"[{result['sheet']}] '{result['lookup_value']}' は見つかりませんでした"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def excel_create_summary_sheet(
    source_sheets: list[str],
    summary_sheet_name: str = "Summary",
) -> str:
    """複数シートからデータを集約するサマリーシートを作成します。"""
    try:
        result = excel.create_summary_sheet(
            source_sheets=source_sheets,
            summary_sheet_name=summary_sheet_name,
        )
        return (
            f"サマリーシート '{result['summary_sheet']}' を作成しました"
            f"（{len(result['source_sheets'])}シートから{result['total_rows']}行）"
        )
    except Exception as e:
        return f"エラー: {e}"
