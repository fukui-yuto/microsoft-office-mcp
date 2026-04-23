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
