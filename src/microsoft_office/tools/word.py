"""Word MCP tool definitions."""

from microsoft_office.server import mcp
from microsoft_office.office import word


@mcp.tool()
def word_create(file_path: str | None = None) -> str:
    """Word文書を新規作成します。file_pathを指定すると即座に保存します。"""
    try:
        result = word.create_document(file_path)
        msg = f"文書 '{result['name']}' を作成しました"
        if file_path:
            msg += f" (保存先: {file_path})"
        return msg
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_open(file_path: str) -> str:
    """既存のWord文書(.docx)を開きます。"""
    try:
        result = word.open_document(file_path)
        return f"'{result['name']}' を開きました（段落数: {result['paragraph_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_save(file_path: str | None = None) -> str:
    """アクティブな文書を保存します。file_pathを指定すると別名保存します。"""
    try:
        result = word.save_document(file_path)
        return f"'{result['name']}' を保存しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_close() -> str:
    """アクティブな文書を閉じます。"""
    try:
        result = word.close_document()
        return f"'{result['name']}' を閉じました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_content() -> str:
    """アクティブな文書の全段落とスタイルを取得します。段落番号・スタイル名・テキストを一覧表示します。"""
    try:
        result = word.get_content()
        lines = [f"文書: {result['name']} (段落数: {result['paragraph_count']})"]
        for para in result["paragraphs"]:
            text = para["text"]
            style = para["style"]
            lines.append(f"  [{para['number']}] ({style}) {text}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_paragraph(text: str, style: str | None = None) -> str:
    """文書の末尾に段落を追加します。styleでスタイル名を指定可能（例: 'Normal', '標準'）。"""
    try:
        result = word.add_paragraph(text, style)
        return f"段落 {result['paragraph_number']} を追加しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_heading(text: str, level: int = 1) -> str:
    """文書の末尾に見出しを追加します。level: 1=見出し1, 2=見出し2, 3=見出し3。"""
    try:
        result = word.add_heading(text, level)
        return f"見出し{result['level']} '{result['text']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_edit_paragraph(paragraph_number: int, text: str) -> str:
    """指定した段落のテキストを編集します。paragraph_numberは1始まりです。"""
    try:
        result = word.edit_paragraph(paragraph_number, text)
        return f"段落 {result['paragraph_number']} を更新しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_delete_paragraph(paragraph_number: int) -> str:
    """指定した段落を削除します。paragraph_numberは1始まりです。"""
    try:
        result = word.delete_paragraph(paragraph_number)
        return f"段落 {result['deleted']} を削除しました（残り: {result['remaining']}段落）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_paragraph_at(paragraph_number: int, text: str, style: str | None = None) -> str:
    """指定した段落番号の位置に段落を挿入します。既存の段落は後ろにずれます。styleでスタイル名を指定可能。"""
    try:
        result = word.insert_paragraph_at(paragraph_number, text, style)
        return f"段落 {result['paragraph_number']} に '{result['text']}' を挿入しました（総段落数: {result['paragraph_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_table(rows: int, cols: int, data: list[list[str]] | None = None) -> str:
    """文書の末尾に表を挿入します。dataで初期データを指定可能（2次元配列）。例: data=[["名前","年齢"],["太郎","25"]]。"""
    try:
        result = word.insert_table(rows, cols, data)
        return f"{result['rows']}×{result['cols']} の表を挿入しました（表の合計: {result['table_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_format_table(
    table_index: int,
    header_bold: bool = True,
    header_fill_color: list[int] | None = None,
    border_style: int = 1,
    auto_fit: bool = True,
    alternating_row_color: list[int] | None = None,
) -> str:
    """表の書式を設定します。table_indexは1始まり。header_fill_color/alternating_row_colorは[R,G,B]形式（例: [0,112,192]）。border_style: 1=単線。"""
    try:
        result = word.format_table(
            table_index,
            header_bold=header_bold,
            header_fill_rgb=tuple(header_fill_color) if header_fill_color else None,
            border_style=border_style,
            auto_fit=auto_fit,
            alternating_row_color_rgb=tuple(alternating_row_color) if alternating_row_color else None,
        )
        return f"表 {result['table_index']} の書式を設定しました（{result['rows']}行×{result['cols']}列）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_page_setup(
    top_margin: float | None = None,
    bottom_margin: float | None = None,
    left_margin: float | None = None,
    right_margin: float | None = None,
    orientation: int | None = None,
    paper_size: int | None = None,
) -> str:
    """ページ設定を変更します。余白はポイント単位（72pt=1inch）。orientation: 0=縦, 1=横。paper_size: 7=A4, 0=Letter。"""
    try:
        result = word.set_page_setup(
            top_margin=top_margin,
            bottom_margin=bottom_margin,
            left_margin=left_margin,
            right_margin=right_margin,
            orientation=orientation,
            paper_size=paper_size,
        )
        return (
            f"ページ設定を更新しました（余白: 上{result['top_margin']}pt 下{result['bottom_margin']}pt "
            f"左{result['left_margin']}pt 右{result['right_margin']}pt, "
            f"向き: {'横' if result['orientation'] == 1 else '縦'}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_header(
    text: str,
    alignment: int = 1,
    font_size: float | None = None,
    section_index: int | None = None,
) -> str:
    """ヘッダーを追加・設定します。alignment: 0=左, 1=中央, 2=右。section_indexで対象セクションを指定（省略時は最終セクション）。"""
    try:
        result = word.add_header(text, alignment=alignment, font_size=font_size, section_index=section_index)
        return f"セクション {result['section']} のヘッダーを設定しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_footer(
    text: str,
    alignment: int = 1,
    font_size: float | None = None,
    section_index: int | None = None,
) -> str:
    """フッターを追加・設定します。alignment: 0=左, 1=中央, 2=右。section_indexで対象セクションを指定（省略時は最終セクション）。"""
    try:
        result = word.add_footer(text, alignment=alignment, font_size=font_size, section_index=section_index)
        return f"セクション {result['section']} のフッターを設定しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_page_numbers(
    alignment: int = 1,
    in_header: bool = False,
    starting_number: int | None = None,
) -> str:
    """ページ番号を追加します。alignment: 0=左, 1=中央, 2=右。in_header=Trueでヘッダーに配置。starting_numberで開始番号を指定。"""
    try:
        result = word.add_page_numbers(alignment=alignment, in_header=in_header, starting_number=starting_number)
        position = "ヘッダー" if result["in_header"] else "フッター"
        return f"{position}にページ番号を追加しました（配置: {['左','中央','右'][result['alignment']]}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_table_of_contents(upper_level: int = 1, lower_level: int = 3) -> str:
    """文書の先頭に目次を挿入します。upper_level/lower_levelで見出しレベルの範囲を指定（デフォルト: 見出し1〜3）。"""
    try:
        result = word.add_table_of_contents(upper_level=upper_level, lower_level=lower_level)
        return f"目次を挿入しました（見出しレベル {result['upper_level']}〜{result['lower_level']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_section_break(break_type: int = 2, paragraph_number: int | None = None) -> str:
    """セクション区切りまたは改ページを挿入します。break_type: 2=次のページから新セクション, 3=現在の位置で新セクション, 7=改ページ。paragraph_numberで挿入位置を指定。"""
    try:
        result = word.add_section_break(break_type=break_type, paragraph_number=paragraph_number)
        return f"区切りを挿入しました（タイプ: {result['break_type']}, セクション数: {result['section_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_columns(
    num_columns: int = 2,
    spacing: float | None = None,
    section_index: int | None = None,
) -> str:
    """段組みを設定します。num_columnsで列数を指定。spacingで列間の間隔（ポイント）。section_indexで対象セクションを指定。"""
    try:
        result = word.set_columns(num_columns=num_columns, spacing=spacing, section_index=section_index)
        return f"セクション {result['section']} を {result['num_columns']} 段組みに設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_paragraph_format_extended(
    paragraph_number: int,
    font_color: list[int] | None = None,
    underline: int | None = None,
    strikethrough: bool | None = None,
    alignment: int | None = None,
    line_spacing: float | None = None,
    line_spacing_rule: int | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    first_line_indent: float | None = None,
    left_indent: float | None = None,
    right_indent: float | None = None,
) -> str:
    """段落の詳細書式を設定します。font_colorは[R,G,B]形式。underline: 1=一重線, 3=二重線。alignment: 0=左, 1=中央, 2=右, 3=両端揃え。line_spacing_rule: 0=1行, 1=1.5行, 2=2行, 4=固定値。間隔・インデントはポイント単位。"""
    try:
        result = word.set_paragraph_format_extended(
            paragraph_number,
            font_color_rgb=tuple(font_color) if font_color else None,
            underline=underline,
            strikethrough=strikethrough,
            alignment=alignment,
            line_spacing=line_spacing,
            line_spacing_rule=line_spacing_rule,
            space_before=space_before,
            space_after=space_after,
            first_line_indent=first_line_indent,
            left_indent=left_indent,
            right_indent=right_indent,
        )
        return f"段落 {result['paragraph_number']} の詳細書式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_font(
    paragraph_number: int,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_color: list[int] | None = None,
    underline: int | None = None,
) -> str:
    """段落のフォントを設定します。font_colorは[R,G,B]形式（例: [255,0,0]で赤）。underline: 1=一重線, 3=二重線。"""
    try:
        result = word.set_font(
            paragraph_number,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            font_color_rgb=tuple(font_color) if font_color else None,
            underline=underline,
        )
        return f"段落 {result['paragraph_number']} のフォントを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_list(items: list[str], list_type: str = "bullet", level: int = 0) -> str:
    """リスト（箇条書きまたは番号付き）を追加します。list_type: 'bullet'=箇条書き, 'number'=番号付き。levelでインデントレベルを指定（0始まり）。"""
    try:
        result = word.add_list(items, list_type=list_type, level=level)
        type_label = "番号付きリスト" if result["list_type"] == "number" else "箇条書き"
        return f"{type_label}を追加しました（{result['items_added']}項目, レベル: {result['level']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_hyperlink(paragraph_number: int, url: str, display_text: str | None = None) -> str:
    """指定した段落にハイパーリンクを追加します。display_textで表示テキストを指定（省略時はURLをそのまま表示）。"""
    try:
        result = word.add_hyperlink(paragraph_number, url, display_text=display_text)
        return f"段落 {result['paragraph_number']} にリンクを追加しました: {result['url']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_bookmark(name: str, paragraph_number: int) -> str:
    """指定した段落にブックマークを追加します。nameはブックマーク名（英数字・アンダースコアのみ）。"""
    try:
        result = word.add_bookmark(name, paragraph_number)
        return f"ブックマーク '{result['name']}' を段落 {result['paragraph_number']} に追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_bookmarks() -> str:
    """アクティブな文書の全ブックマーク一覧を取得します。"""
    try:
        result = word.get_bookmarks()
        if result["bookmark_count"] == 0:
            return "ブックマークはありません"
        lines = [f"ブックマーク数: {result['bookmark_count']}"]
        for bm in result["bookmarks"]:
            lines.append(f"  - {bm['name']} (位置: {bm['start']}〜{bm['end']})")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_image(
    file_path: str,
    width: float | None = None,
    height: float | None = None,
    paragraph_number: int | None = None,
) -> str:
    """画像を挿入します。width/heightでサイズをポイント単位で指定。paragraph_numberで挿入位置を指定（省略時は末尾）。"""
    try:
        result = word.insert_image(file_path, width=width, height=height, paragraph_number=paragraph_number)
        return f"画像を挿入しました: {result['file_path']}（サイズ: {result['width']:.0f}×{result['height']:.0f}pt）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_text_box(
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_name: str | None = None,
    font_size: float | None = None,
    fill_color: list[int] | None = None,
    border_color: list[int] | None = None,
) -> str:
    """テキストボックスを追加します。left/top/width/heightはポイント単位で位置とサイズを指定。fill_color/border_colorは[R,G,B]形式。"""
    try:
        result = word.add_text_box(
            left, top, width, height, text,
            font_name=font_name,
            font_size=font_size,
            fill_rgb=tuple(fill_color) if fill_color else None,
            border_rgb=tuple(border_color) if border_color else None,
        )
        return f"テキストボックス '{result['shape_name']}' を追加しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_find_and_replace(
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_whole_word: bool = False,
    replace_all: bool = True,
) -> str:
    """文書内のテキストを検索して置換します。match_case=Trueで大文字小文字を区別。match_whole_word=Trueで単語単位で検索。replace_all=Trueで全置換。"""
    try:
        result = word.find_and_replace(
            find_text, replace_text,
            match_case=match_case,
            match_whole_word=match_whole_word,
            replace_all=replace_all,
        )
        status = "置換を実行しました" if result["executed"] else "一致するテキストが見つかりませんでした"
        return f"'{result['find_text']}' → '{result['replace_text']}': {status}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_comment(paragraph_number: int, comment_text: str) -> str:
    """指定した段落にコメントを追加します。"""
    try:
        result = word.add_comment(paragraph_number, comment_text)
        return f"段落 {result['paragraph_number']} にコメントを追加しました: '{result['comment_text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_comments() -> str:
    """アクティブな文書の全コメント一覧を取得します。"""
    try:
        result = word.get_comments()
        if result["comment_count"] == 0:
            return "コメントはありません"
        lines = [f"コメント数: {result['comment_count']}"]
        for c in result["comments"]:
            lines.append(f"  [{c['index']}] {c['author']}: {c['text']} (対象: {c['scope_text']})")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_watermark(
    text: str,
    font_size: float = 72,
    color: list[int] | None = None,
    semitransparent: bool = True,
) -> str:
    """透かし（ウォーターマーク）を追加します。例: text='DRAFT'。colorは[R,G,B]形式（デフォルト: [192,192,192]のグレー）。semitransparent=Trueで半透明。"""
    try:
        color_rgb = tuple(color) if color else (192, 192, 192)
        result = word.add_watermark(text, font_size=font_size, color_rgb=color_rgb, semitransparent=semitransparent)
        return f"透かし '{result['text']}' を追加しました（フォントサイズ: {result['font_size']}pt）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_page_borders(
    border_style: int = 1,
    color: list[int] | None = None,
    width: float | None = None,
) -> str:
    """ページ罫線を設定します。border_style: 1=単線。colorは[R,G,B]形式。widthはポイント単位。"""
    try:
        result = word.set_page_borders(
            border_style=border_style,
            color_rgb=tuple(color) if color else None,
            width=width,
        )
        return f"ページ罫線を設定しました（スタイル: {result['border_style']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_export_to_pdf(output_path: str) -> str:
    """アクティブな文書をPDFとして出力します。output_pathに保存先パスを指定。"""
    try:
        result = word.export_to_pdf(output_path)
        return f"PDFを出力しました: {result['output_path']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_enable_track_changes(enable: bool = True) -> str:
    """変更履歴の記録を有効または無効にします。enable=Trueで有効、Falseで無効。"""
    try:
        result = word.enable_track_changes(enable=enable)
        status = "有効" if result["track_changes"] else "無効"
        return f"変更履歴を{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_accept_all_changes() -> str:
    """文書内の全ての変更履歴を承認します。"""
    try:
        word.accept_all_changes()
        return "全ての変更を承認しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_reject_all_changes() -> str:
    """文書内の全ての変更履歴を拒否します。"""
    try:
        word.reject_all_changes()
        return "全ての変更を拒否しました"
    except Exception as e:
        return f"エラー: {e}"
