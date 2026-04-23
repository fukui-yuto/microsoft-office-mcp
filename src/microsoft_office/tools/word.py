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


# ---------------------------------------------------------------------------
# 1. Style Management
# ---------------------------------------------------------------------------

@mcp.tool()
def word_create_style(
    name: str,
    style_type: str = "paragraph",
    base_style: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: list[int] | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
    alignment: str | None = None,
) -> str:
    """カスタムスタイルを作成します。style_type: 'paragraph','character','table','list'。font_colorは[R,G,B]形式。alignment: 'left','center','right','justify'。"""
    try:
        result = word.create_style(
            name,
            style_type=style_type,
            base_style=base_style,
            font_name=font_name,
            font_size=font_size,
            font_color_rgb=tuple(font_color) if font_color else None,
            bold=bold,
            italic=italic,
            space_before=space_before,
            space_after=space_after,
            line_spacing=line_spacing,
            alignment=alignment,
        )
        return f"スタイル '{result['name']}' を作成しました（種類: {result['style_type']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_apply_style(paragraph_index: int, style_name: str) -> str:
    """段落にスタイルを適用します。paragraph_indexは1始まり。style_nameでスタイル名を指定。"""
    try:
        result = word.apply_style(paragraph_index, style_name)
        return f"段落 {result['paragraph_index']} にスタイル '{result['style_name']}' を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_styles() -> str:
    """アクティブな文書の全スタイル一覧を取得します。"""
    try:
        result = word.get_styles()
        lines = [f"スタイル数: {result['style_count']}"]
        for s in result["styles"]:
            built_in = "組込み" if s["built_in"] else "カスタム"
            lines.append(f"  - {s['name']} (タイプ: {s['type']}, {built_in})")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_modify_style(
    style_name: str,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: list[int] | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
) -> str:
    """既存のスタイルを変更します。font_colorは[R,G,B]形式。"""
    try:
        result = word.modify_style(
            style_name,
            font_name=font_name,
            font_size=font_size,
            font_color_rgb=tuple(font_color) if font_color else None,
            bold=bold,
            italic=italic,
            space_before=space_before,
            space_after=space_after,
            line_spacing=line_spacing,
        )
        return f"スタイル '{result['style_name']}' を変更しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 2. Footnotes & Endnotes
# ---------------------------------------------------------------------------

@mcp.tool()
def word_add_footnote(paragraph_index: int, text: str) -> str:
    """指定した段落の末尾に脚注を追加します。paragraph_indexは1始まり。"""
    try:
        result = word.add_footnote(paragraph_index, text)
        return f"段落 {result['paragraph_index']} に脚注を追加しました（脚注数: {result['footnote_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_endnote(paragraph_index: int, text: str) -> str:
    """指定した段落の末尾に文末脚注を追加します。paragraph_indexは1始まり。"""
    try:
        result = word.add_endnote(paragraph_index, text)
        return f"段落 {result['paragraph_index']} に文末脚注を追加しました（文末脚注数: {result['endnote_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_footnotes() -> str:
    """アクティブな文書の全脚注一覧を取得します。"""
    try:
        result = word.get_footnotes()
        if result["footnote_count"] == 0:
            return "脚注はありません"
        lines = [f"脚注数: {result['footnote_count']}"]
        for fn in result["footnotes"]:
            lines.append(f"  [{fn['index']}] {fn['text']}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_endnotes() -> str:
    """アクティブな文書の全文末脚注一覧を取得します。"""
    try:
        result = word.get_endnotes()
        if result["endnote_count"] == 0:
            return "文末脚注はありません"
        lines = [f"文末脚注数: {result['endnote_count']}"]
        for en in result["endnotes"]:
            lines.append(f"  [{en['index']}] {en['text']}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 3. Advanced Table Operations
# ---------------------------------------------------------------------------

@mcp.tool()
def word_merge_table_cells(
    table_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> str:
    """表のセルを結合します。table_indexは1始まり。行・列も1始まり。"""
    try:
        result = word.merge_table_cells(table_index, start_row, start_col, end_row, end_col)
        return f"表 {result['table_index']} のセル {result['merged']} を結合しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_split_table_cell(
    table_index: int,
    row: int,
    col: int,
    num_rows: int,
    num_cols: int,
) -> str:
    """表のセルを分割します。table_indexは1始まり。行・列も1始まり。"""
    try:
        result = word.split_table_cell(table_index, row, col, num_rows, num_cols)
        return f"表 {result['table_index']} のセル {result['cell']} を {result['split_into']} に分割しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_format_table_cell(
    table_index: int,
    row: int,
    col: int,
    fill_color: list[int] | None = None,
    font_color: list[int] | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    alignment: int | None = None,
    vertical_alignment: str | None = None,
) -> str:
    """表の個別セルを書式設定します。fill_color/font_colorは[R,G,B]形式。alignment: 0=左,1=中央,2=右。vertical_alignment: 'top','center','bottom'。"""
    try:
        result = word.format_table_cell(
            table_index, row, col,
            fill_color_rgb=tuple(fill_color) if fill_color else None,
            font_color_rgb=tuple(font_color) if font_color else None,
            font_size=font_size,
            bold=bold,
            alignment=alignment,
            vertical_alignment=vertical_alignment,
        )
        return f"表 {result['table_index']} のセル {result['cell']} の書式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_table_cell_borders(
    table_index: int,
    row: int,
    col: int,
    border_type: str,
    color: list[int] | None = None,
    weight: float | None = None,
    style: int | None = None,
) -> str:
    """表のセルに罫線を設定します。border_type: 'top','bottom','left','right','all'。colorは[R,G,B]形式。"""
    try:
        result = word.set_table_cell_borders(
            table_index, row, col, border_type,
            color_rgb=tuple(color) if color else None,
            weight=weight,
            style=style,
        )
        return f"表 {result['table_index']} のセル {result['cell']} に {result['border_type']} 罫線を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_table_width(
    table_index: int,
    width_type: str,
    width: float | None = None,
) -> str:
    """表の幅を設定します。width_type: 'auto'=自動, 'fixed'=固定(ポイント), 'percent'=パーセント。"""
    try:
        result = word.set_table_width(table_index, width_type, width=width)
        return f"表 {result['table_index']} の幅を {result['width_type']} に設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 4. Shapes & Drawing
# ---------------------------------------------------------------------------

@mcp.tool()
def word_add_shape(
    shape_type: int,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: list[int] | None = None,
    line_color: list[int] | None = None,
) -> str:
    """図形を追加します。shape_type: 1=四角形, 5=角丸四角形, 9=楕円 等。座標はポイント単位。fill_color/line_colorは[R,G,B]形式。"""
    try:
        result = word.add_shape(
            shape_type, left, top, width, height,
            fill_color_rgb=tuple(fill_color) if fill_color else None,
            line_color_rgb=tuple(line_color) if line_color else None,
        )
        return f"図形 '{result['shape_name']}' を追加しました（タイプ: {result['shape_type']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_line(
    start_x: float,
    start_y: float,
    end_x: float,
    end_y: float,
    color: list[int] | None = None,
    weight: float = 1.0,
    dash_style: int | None = None,
) -> str:
    """直線を描画します。座標はポイント単位。colorは[R,G,B]形式。weightは線の太さ。dash_styleで破線スタイルを指定。"""
    try:
        result = word.add_line(
            start_x, start_y, end_x, end_y,
            color_rgb=tuple(color) if color else None,
            weight=weight,
            dash_style=dash_style,
        )
        return f"直線 '{result['shape_name']}' を描画しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 5. Text Enhancements
# ---------------------------------------------------------------------------

@mcp.tool()
def word_add_drop_cap(
    paragraph_index: int,
    lines_to_drop: int = 3,
    font_name: str | None = None,
) -> str:
    """段落にドロップキャップ（先頭文字の拡大表示）を設定します。lines_to_dropで行数を指定（デフォルト3行）。"""
    try:
        result = word.add_drop_cap(paragraph_index, lines_to_drop=lines_to_drop, font_name=font_name)
        return f"段落 {result['paragraph_index']} にドロップキャップを設定しました（{result['lines_to_drop']}行）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_text_highlight(paragraph_index: int, color_index: int) -> str:
    """段落のテキストに蛍光ペン（ハイライト）を設定します。color_index: 1-16のWdColorIndex値。例: 7=黄色, 6=赤, 4=緑。"""
    try:
        result = word.set_text_highlight(paragraph_index, color_index)
        return f"段落 {result['paragraph_index']} にハイライトを設定しました（色: {result['color_index']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_character_spacing(
    paragraph_index: int,
    spacing: float = 0,
    kerning: float | None = None,
    scale: int = 100,
) -> str:
    """段落の文字間隔を設定します。spacingはポイント単位（正=広げる、負=詰める）。scaleは文字幅の倍率（100=通常）。kerningはカーニングのしきい値。"""
    try:
        result = word.set_character_spacing(paragraph_index, spacing=spacing, kerning=kerning, scale=scale)
        return f"段落 {result['paragraph_index']} の文字間隔を設定しました（間隔: {result['spacing']}pt, 倍率: {result['scale']}%）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_text_effect(paragraph_index: int, effect_type: str) -> str:
    """段落にテキスト効果を追加します。effect_type: 'shadow'=影, 'outline'=アウトライン, 'emboss'=浮き出し, 'engrave'=彫り込み。"""
    try:
        result = word.add_text_effect(paragraph_index, effect_type)
        return f"段落 {result['paragraph_index']} にテキスト効果 '{result['effect_type']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 6. Page Layout
# ---------------------------------------------------------------------------

@mcp.tool()
def word_insert_page_break(paragraph_index: int | None = None) -> str:
    """改ページを挿入します。paragraph_indexで挿入位置の段落を指定（省略時は末尾）。"""
    try:
        result = word.insert_page_break(paragraph_index=paragraph_index)
        pos = f"段落 {result['paragraph_index']} の後" if result["paragraph_index"] else "末尾"
        return f"改ページを{pos}に挿入しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_column_break(paragraph_index: int | None = None) -> str:
    """段区切りを挿入します。paragraph_indexで挿入位置の段落を指定（省略時は末尾）。"""
    try:
        result = word.insert_column_break(paragraph_index=paragraph_index)
        pos = f"段落 {result['paragraph_index']} の後" if result["paragraph_index"] else "末尾"
        return f"段区切りを{pos}に挿入しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_paragraph_borders(
    paragraph_index: int,
    border_type: str,
    color: list[int] | None = None,
    weight: float | None = None,
    style: int | None = None,
) -> str:
    """段落に罫線を設定します。border_type: 'top','bottom','left','right','box'。colorは[R,G,B]形式。"""
    try:
        result = word.set_paragraph_borders(
            paragraph_index, border_type,
            color_rgb=tuple(color) if color else None,
            weight=weight,
            style=style,
        )
        return f"段落 {result['paragraph_index']} に {result['border_type']} 罫線を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_paragraph_shading(paragraph_index: int, color: list[int]) -> str:
    """段落の背景色を設定します。colorは[R,G,B]形式（例: [255,255,200]）。"""
    try:
        result = word.set_paragraph_shading(paragraph_index, color_rgb=tuple(color))
        return f"段落 {result['paragraph_index']} の背景色を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_horizontal_line(paragraph_index: int | None = None) -> str:
    """装飾用の水平線を追加します。paragraph_indexで挿入位置を指定（省略時は末尾）。"""
    try:
        result = word.add_horizontal_line(paragraph_index=paragraph_index)
        return f"水平線を段落 {result['paragraph_index']} に追加しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 7. Document Properties
# ---------------------------------------------------------------------------

@mcp.tool()
def word_set_document_properties(
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    comments: str | None = None,
) -> str:
    """文書のメタデータ（プロパティ）を設定します。タイトル、作成者、件名、キーワード、分類、コメントを指定可能。"""
    try:
        word.set_document_properties(
            title=title,
            author=author,
            subject=subject,
            keywords=keywords,
            category=category,
            comments=comments,
        )
        return "文書プロパティを更新しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_document_properties() -> str:
    """文書のメタデータ（プロパティ）を取得します。"""
    try:
        result = word.get_document_properties()
        lines = ["文書プロパティ:"]
        for key, value in result.items():
            lines.append(f"  {key}: {value}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_document_statistics() -> str:
    """文書の統計情報（文字数、単語数、ページ数など）を取得します。"""
    try:
        result = word.get_document_statistics()
        return (
            f"文書統計: ページ数={result['page_count']}, "
            f"単語数={result['word_count']}, "
            f"文字数={result['character_count']}, "
            f"行数={result['line_count']}, "
            f"段落数={result['paragraph_count']}"
        )
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 8. Advanced Features
# ---------------------------------------------------------------------------

@mcp.tool()
def word_add_cross_reference(
    ref_type: str,
    ref_item: int,
    ref_format: int | None = None,
) -> str:
    """相互参照を追加します。ref_type: 'heading','bookmark','footnote'。ref_itemは1始まりのインデックス。"""
    try:
        result = word.add_cross_reference(ref_type, ref_item, ref_format=ref_format)
        return f"相互参照を追加しました（タイプ: {result['ref_type']}, 項目: {result['ref_item']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_field(field_code: str, paragraph_index: int | None = None) -> str:
    """Wordフィールドを挿入します。field_code例: 'DATE', 'PAGE', 'NUMPAGES', 'AUTHOR', 'TOC'。paragraph_indexで挿入位置を指定。"""
    try:
        result = word.insert_field(field_code, paragraph_index=paragraph_index)
        return f"フィールド '{result['field_code']}' を挿入しました（フィールド数: {result['field_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_tab_stops(
    paragraph_index: int,
    positions: list[float],
    alignments: list[str] | None = None,
    leaders: list[str] | None = None,
) -> str:
    """段落にタブストップを設定します。positionsはポイント単位のリスト。alignments: 'left','center','right','decimal'。leaders: 'none','dots','dashes','line'。"""
    try:
        result = word.set_tab_stops(paragraph_index, positions, alignments=alignments, leaders=leaders)
        return f"段落 {result['paragraph_index']} にタブストップを {result['tab_count']} 個設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_protect_document(
    password: str | None = None,
    protection_type: str = "read_only",
) -> str:
    """文書を保護します。protection_type: 'read_only'=読み取り専用, 'comments'=コメントのみ, 'forms'=フォームのみ, 'tracked_changes'=変更履歴。"""
    try:
        result = word.protect_document(password=password, protection_type=protection_type)
        return f"文書を保護しました（種類: {result['protection_type']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_unprotect_document(password: str | None = None) -> str:
    """文書の保護を解除します。パスワードが設定されている場合はpasswordを指定してください。"""
    try:
        word.unprotect_document(password=password)
        return "文書の保護を解除しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 9. Content Controls & Templates
# ---------------------------------------------------------------------------

@mcp.tool()
def word_add_content_control(
    control_type: str,
    paragraph_index: int | None = None,
    title: str | None = None,
    placeholder_text: str | None = None,
) -> str:
    """コンテンツコントロールを追加します。control_type: 'rich_text','plain_text','combo_box','drop_down','date_picker','checkbox'。"""
    try:
        result = word.add_content_control(
            control_type,
            paragraph_index=paragraph_index,
            title=title,
            placeholder_text=placeholder_text,
        )
        return f"コンテンツコントロール '{result['title']}' を追加しました（種類: {result['control_type']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_building_block(name: str, category: str | None = None) -> str:
    """文書パーツ（クイックパーツ）を挿入します。nameで名前を指定。categoryでカテゴリを指定（省略可）。"""
    try:
        result = word.insert_building_block(name, category=category)
        if result["inserted"]:
            return f"文書パーツ '{result['name']}' を挿入しました"
        else:
            return f"文書パーツ '{result['name']}' が見つかりませんでした"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 10. Lists Advanced
# ---------------------------------------------------------------------------

@mcp.tool()
def word_set_list_level(paragraph_index: int, level: int) -> str:
    """リストのインデントレベルを設定します。paragraph_indexは1始まり。level: 0-8。"""
    try:
        result = word.set_list_level(paragraph_index, level)
        return f"段落 {result['paragraph_index']} のリストレベルを {result['level']} に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_restart_list_numbering(paragraph_index: int) -> str:
    """指定した段落でリストの番号を振り直します。paragraph_indexは1始まり。"""
    try:
        result = word.restart_list_numbering(paragraph_index)
        return f"段落 {result['paragraph_index']} でリスト番号を振り直しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 11. Advanced Text Formatting
# ---------------------------------------------------------------------------

@mcp.tool()
def word_set_text_color_range(
    paragraph_index: int,
    start_char: int,
    end_char: int,
    color: list[int],
) -> str:
    """段落内の指定範囲の文字に色を設定します。paragraph_indexは1始まり。start_char/end_charは0始まりのオフセット。colorは[R,G,B]形式。"""
    try:
        result = word.set_text_color_range(paragraph_index, start_char, end_char, tuple(color))
        return f"段落 {result['paragraph_index']} の文字 {result['start_char']}〜{result['end_char']} に色を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_text_size_range(
    paragraph_index: int,
    start_char: int,
    end_char: int,
    font_size: float,
) -> str:
    """段落内の指定範囲の文字にフォントサイズを設定します。paragraph_indexは1始まり。start_char/end_charは0始まりのオフセット。"""
    try:
        result = word.set_text_size_range(paragraph_index, start_char, end_char, font_size)
        return f"段落 {result['paragraph_index']} の文字 {result['start_char']}〜{result['end_char']} のサイズを {result['font_size']}pt に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_strikethrough(paragraph_index: int, double: bool = False) -> str:
    """段落に取り消し線を設定します。double=Trueで二重取り消し線。paragraph_indexは1始まり。"""
    try:
        result = word.add_strikethrough(paragraph_index, double=double)
        style = "二重取り消し線" if result["double"] else "取り消し線"
        return f"段落 {result['paragraph_index']} に{style}を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_underline_style(paragraph_index: int, style: str) -> str:
    """段落の下線スタイルを設定します。style: 'single'=一重線, 'double'=二重線, 'dotted'=点線, 'dashed'=破線, 'wavy'=波線, 'thick'=太線, 'none'=なし。"""
    try:
        result = word.set_underline_style(paragraph_index, style)
        return f"段落 {result['paragraph_index']} の下線スタイルを '{result['style']}' に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_small_caps(paragraph_index: int) -> str:
    """段落にスモールキャップス（小型英大文字）を設定します。paragraph_indexは1始まり。"""
    try:
        result = word.add_small_caps(paragraph_index)
        return f"段落 {result['paragraph_index']} にスモールキャップスを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_all_caps(paragraph_index: int) -> str:
    """段落にオールキャップス（すべて大文字）を設定します。paragraph_indexは1始まり。"""
    try:
        result = word.add_all_caps(paragraph_index)
        return f"段落 {result['paragraph_index']} にオールキャップスを設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 12. Advanced Paragraph
# ---------------------------------------------------------------------------

@mcp.tool()
def word_set_keep_with_next(paragraph_index: int, keep: bool = True) -> str:
    """段落を次の段落と一緒に保持します（段落間で改ページしない）。paragraph_indexは1始まり。"""
    try:
        result = word.set_keep_with_next(paragraph_index, keep=keep)
        status = "有効" if result["keep"] else "無効"
        return f"段落 {result['paragraph_index']} の「次の段落と分離しない」を{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_keep_together(paragraph_index: int, keep: bool = True) -> str:
    """段落を分割しないように設定します（段落内で改ページしない）。paragraph_indexは1始まり。"""
    try:
        result = word.set_keep_together(paragraph_index, keep=keep)
        status = "有効" if result["keep"] else "無効"
        return f"段落 {result['paragraph_index']} の「段落を分割しない」を{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_page_break_before(paragraph_index: int, break_before: bool = True) -> str:
    """段落の前に改ページを設定します。paragraph_indexは1始まり。"""
    try:
        result = word.set_page_break_before(paragraph_index, break_before=break_before)
        status = "有効" if result["break_before"] else "無効"
        return f"段落 {result['paragraph_index']} の「段落前で改ページ」を{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_widow_orphan_control(paragraph_index: int, control: bool = True) -> str:
    """段落の改行・改ページのオーファンコントロールを設定します。paragraph_indexは1始まり。"""
    try:
        result = word.set_widow_orphan_control(paragraph_index, control=control)
        status = "有効" if result["control"] else "無効"
        return f"段落 {result['paragraph_index']} のオーファンコントロールを{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_outline_level(paragraph_index: int, level: int) -> str:
    """段落のアウトラインレベルを設定します。paragraph_indexは1始まり。level: 0=本文, 1-9=アウトラインレベル。"""
    try:
        result = word.set_outline_level(paragraph_index, level)
        level_str = "本文" if result["level"] == 0 else f"レベル{result['level']}"
        return f"段落 {result['paragraph_index']} のアウトラインレベルを{level_str}に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_paragraph_count() -> str:
    """アクティブな文書の段落数を取得します。"""
    try:
        result = word.get_paragraph_count()
        return f"段落数: {result['paragraph_count']}"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 13. Table Advanced
# ---------------------------------------------------------------------------

@mcp.tool()
def word_set_table_row_height(
    table_index: int,
    row: int,
    height: float,
    rule: str = "exact",
) -> str:
    """表の行の高さを設定します。table_index/rowは1始まり。heightはポイント単位。rule: 'exact'=固定, 'at_least'=最小, 'auto'=自動。"""
    try:
        result = word.set_table_row_height(table_index, row, height, rule=rule)
        return f"表 {result['table_index']} の行 {result['row']} の高さを {result['height']}pt ({result['rule']}) に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_table_cell_width(
    table_index: int,
    row: int,
    col: int,
    width: float,
) -> str:
    """表の個別セルの幅を設定します。table_index/row/colは1始まり。widthはポイント単位。"""
    try:
        result = word.set_table_cell_width(table_index, row, col, width)
        return f"表 {result['table_index']} のセル ({result['row']},{result['col']}) の幅を {result['width']}pt に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_table_alignment(table_index: int, alignment: str) -> str:
    """表の配置を設定します。table_indexは1始まり。alignment: 'left'=左, 'center'=中央, 'right'=右。"""
    try:
        result = word.set_table_alignment(table_index, alignment)
        return f"表 {result['table_index']} の配置を '{result['alignment']}' に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_table_row(
    table_index: int,
    position: int | None = None,
    values: list[str] | None = None,
) -> str:
    """表に行を追加します。table_indexは1始まり。positionで挿入位置を指定（1始まり、省略時は末尾）。valuesで初期値を指定。"""
    try:
        result = word.add_table_row(table_index, position=position, values=values)
        return f"表 {result['table_index']} の行 {result['row']} に行を追加しました（行数: {result['row_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_delete_table_row(table_index: int, row: int) -> str:
    """表の行を削除します。table_index/rowは1始まり。"""
    try:
        result = word.delete_table_row(table_index, row)
        return f"表 {result['table_index']} の行 {result['deleted_row']} を削除しました（残り行数: {result['row_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_table_column(table_index: int, position: int | None = None) -> str:
    """表に列を追加します。table_indexは1始まり。positionで挿入位置を指定（1始まり、省略時は末尾）。"""
    try:
        result = word.add_table_column(table_index, position=position)
        return f"表 {result['table_index']} に列を追加しました（列数: {result['col_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_delete_table_column(table_index: int, col: int) -> str:
    """表の列を削除します。table_index/colは1始まり。"""
    try:
        result = word.delete_table_column(table_index, col)
        return f"表 {result['table_index']} の列 {result['deleted_col']} を削除しました（残り列数: {result['col_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_table_data(table_index: int) -> str:
    """表の全データを取得します。table_indexは1始まり。2次元配列として返します。"""
    try:
        result = word.get_table_data(table_index)
        lines = [f"表 {result['table_index']} ({result['rows']}行×{result['cols']}列):"]
        for ri, row_data in enumerate(result["data"]):
            cells = " | ".join(row_data)
            lines.append(f"  行{ri + 1}: {cells}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_table_repeat_header(table_index: int, repeat: bool = True) -> str:
    """表の先頭行を各ページで繰り返すヘッダー行に設定します。table_indexは1始まり。"""
    try:
        result = word.set_table_repeat_header(table_index, repeat=repeat)
        status = "有効" if result["repeat"] else "無効"
        return f"表 {result['table_index']} のヘッダー行繰り返しを{status}にしました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 14. Section Management
# ---------------------------------------------------------------------------

@mcp.tool()
def word_get_section_count() -> str:
    """アクティブな文書のセクション数を取得します。"""
    try:
        result = word.get_section_count()
        return f"セクション数: {result['section_count']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_section_page_setup(
    section_index: int,
    orientation: int | None = None,
    width: float | None = None,
    height: float | None = None,
    top_margin: float | None = None,
    bottom_margin: float | None = None,
    left_margin: float | None = None,
    right_margin: float | None = None,
) -> str:
    """セクション別のページ設定を変更します。section_indexは1始まり。余白・幅・高さはポイント単位。orientation: 0=縦, 1=横。"""
    try:
        result = word.set_section_page_setup(
            section_index,
            orientation=orientation,
            width=width,
            height=height,
            top_margin=top_margin,
            bottom_margin=bottom_margin,
            left_margin=left_margin,
            right_margin=right_margin,
        )
        return f"セクション {result['section_index']} のページ設定を更新しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_section_header(
    section_index: int,
    text: str,
    alignment: int | None = None,
) -> str:
    """セクション別のヘッダーを設定します。section_indexは1始まり。alignment: 0=左, 1=中央, 2=右。"""
    try:
        result = word.set_section_header(section_index, text, alignment=alignment)
        return f"セクション {result['section_index']} のヘッダーを設定しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_section_footer(
    section_index: int,
    text: str,
    alignment: int | None = None,
) -> str:
    """セクション別のフッターを設定します。section_indexは1始まり。alignment: 0=左, 1=中央, 2=右。"""
    try:
        result = word.set_section_footer(section_index, text, alignment=alignment)
        return f"セクション {result['section_index']} のフッターを設定しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_link_section_header(section_index: int, link_to_previous: bool = True) -> str:
    """セクションのヘッダーを前のセクションとリンク/リンク解除します。section_indexは1始まり。"""
    try:
        result = word.link_section_header(section_index, link_to_previous=link_to_previous)
        status = "リンク" if result["link_to_previous"] else "リンク解除"
        return f"セクション {result['section_index']} のヘッダーを前のセクションと{status}しました"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 15. Document Navigation & Structure
# ---------------------------------------------------------------------------

@mcp.tool()
def word_go_to_page(page_number: int) -> str:
    """指定したページに移動します。"""
    try:
        result = word.go_to_page(page_number)
        return f"ページ {result['page_number']} に移動しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_page_count() -> str:
    """アクティブな文書のページ数を取得します。"""
    try:
        result = word.get_page_count()
        return f"ページ数: {result['page_count']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_text_at_bookmark(bookmark_name: str, text: str) -> str:
    """ブックマーク位置にテキストを挿入します。"""
    try:
        result = word.insert_text_at_bookmark(bookmark_name, text)
        return f"ブックマーク '{result['bookmark_name']}' にテキストを挿入しました: '{result['text']}'"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_paragraph_text(paragraph_index: int) -> str:
    """指定した段落のテキストを取得します。paragraph_indexは1始まり。"""
    try:
        result = word.get_paragraph_text(paragraph_index)
        return f"段落 {result['paragraph_index']}: {result['text']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_get_paragraph_range(start_index: int, end_index: int) -> str:
    """指定範囲の段落のテキストを取得します。start_index/end_indexは1始まり。"""
    try:
        result = word.get_paragraph_range(start_index, end_index)
        lines = [f"段落 {result['start_index']}〜{result['end_index']}:"]
        for p in result["paragraphs"]:
            lines.append(f"  [{p['index']}] {p['text']}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 16. Advanced Find
# ---------------------------------------------------------------------------

@mcp.tool()
def word_find_text(
    text: str,
    match_case: bool = False,
    match_whole_word: bool = False,
) -> str:
    """文書内のテキストを検索します。最初に見つかった位置の段落番号と文字位置を返します。"""
    try:
        result = word.find_text(text, match_case=match_case, match_whole_word=match_whole_word)
        if result["found"]:
            return f"'{result['text']}' が見つかりました（段落: {result['paragraph_index']}, 文字位置: {result['char_position']}）"
        return f"'{result['text']}' は見つかりませんでした"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_find_all(text: str, match_case: bool = False) -> str:
    """文書内のテキストの全出現箇所を検索します。"""
    try:
        result = word.find_all(text, match_case=match_case)
        if result["count"] == 0:
            return f"'{result['text']}' は見つかりませんでした"
        lines = [f"'{result['text']}' が {result['count']} 箇所見つかりました:"]
        for loc in result["locations"]:
            lines.append(f"  段落 {loc['paragraph_index']}, 文字位置 {loc['char_position']}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_highlight_found_text(
    text: str,
    highlight_color: int = 7,
    match_case: bool = False,
) -> str:
    """文書内のテキストを検索してハイライトします。highlight_color: WdColorIndex値（7=黄色）。"""
    try:
        result = word.highlight_found_text(text, highlight_color=highlight_color, match_case=match_case)
        return f"'{result['text']}' を {result['count']} 箇所ハイライトしました（色: {result['highlight_color']}）"
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 17. Mail Merge Support
# ---------------------------------------------------------------------------

@mcp.tool()
def word_start_mail_merge(data_source_path: str) -> str:
    """差し込み印刷を開始します。data_source_pathにデータソース（CSV/Excelファイル）のパスを指定。"""
    try:
        result = word.start_mail_merge(data_source_path)
        return f"差し込み印刷を開始しました（データソース: {result['data_source']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_merge_field(field_name: str) -> str:
    """カーソル位置に差し込みフィールドを挿入します。field_nameでフィールド名を指定。"""
    try:
        result = word.insert_merge_field(field_name)
        return f"差し込みフィールド '{result['field_name']}' を挿入しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_execute_mail_merge(output_path: str | None = None) -> str:
    """差し込み印刷を実行します。output_pathを指定すると結果を保存します。"""
    try:
        result = word.execute_mail_merge(output_path=output_path)
        msg = f"差し込み印刷を実行しました（結果: {result['output_name']}）"
        if result["output_path"]:
            msg += f" 保存先: {result['output_path']}"
        return msg
    except Exception as e:
        return f"エラー: {e}"


# ---------------------------------------------------------------------------
# 18. Remaining Pro Features
# ---------------------------------------------------------------------------

@mcp.tool()
def word_get_word_count() -> str:
    """文書の文字数・単語数を取得します。"""
    try:
        result = word.get_word_count()
        lines = [
            f"単語数: {result['word_count']}",
            f"文字数: {result['character_count']}",
            f"段落数: {result['paragraph_count']}",
            f"ページ数: {result['page_count']}",
        ]
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_date(
    paragraph_index: int | None = None,
    format_str: str | None = None,
) -> str:
    """現在の日付を挿入します。paragraph_indexで挿入位置を指定（省略時は末尾）。"""
    try:
        result = word.insert_date(paragraph_index=paragraph_index, format_str=format_str)
        return f"日付 '{result['date_text']}' を挿入しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_line_numbers(
    start_value: int = 1,
    count_by: int = 1,
    restart_each_page: bool = True,
) -> str:
    """行番号を設定します。count_byで表示間隔を指定（1=毎行、5=5行ごと）。"""
    try:
        result = word.set_line_numbers(
            start_value=start_value, count_by=count_by,
            restart_each_page=restart_each_page,
        )
        restart = "ページごとにリセット" if result["restart_each_page"] else "連続"
        return f"行番号を設定しました（開始: {result['start_value']}, 間隔: {result['count_by']}, {restart}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_compare_documents(
    original_path: str,
    revised_path: str,
    output_path: str | None = None,
) -> str:
    """2つの文書を比較します。output_pathを指定すると比較結果を保存します。"""
    try:
        result = word.compare_documents(original_path, revised_path, output_path=output_path)
        msg = f"文書を比較しました（結果: {result['result_name']}）"
        if result["output_path"]:
            msg += f" 保存先: {result['output_path']}"
        return msg
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_paragraph_indentation(
    paragraph_index: int,
    left: float | None = None,
    right: float | None = None,
    first_line: float | None = None,
    hanging: float | None = None,
) -> str:
    """段落のインデントをポイント単位で設定します。"""
    try:
        result = word.set_paragraph_indentation(
            paragraph_index, left=left, right=right,
            first_line=first_line, hanging=hanging,
        )
        return f"段落 {result['paragraph_index']} のインデントを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_duplicate_document(file_path: str) -> str:
    """現在の文書のコピーを新しいパスに保存します。"""
    try:
        result = word.duplicate_document(file_path)
        return f"文書のコピーを保存しました: {result['file_path']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_special_character(paragraph_index: int, char_type: str) -> str:
    """特殊文字を段落の先頭に挿入します。char_type: "em_dash", "en_dash", "nonbreaking_space", "copyright", "registered", "trademark", "bullet", "section", "paragraph"。"""
    try:
        result = word.insert_special_character(paragraph_index, char_type)
        return f"段落 {result['paragraph_index']} に特殊文字 '{result['char_type']}' を挿入しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_set_default_font(font_name: str, font_size: float | None = None) -> str:
    """文書のデフォルトフォントを設定します。"""
    try:
        result = word.set_default_font(font_name, font_size=font_size)
        msg = f"デフォルトフォントを '{result['font_name']}' に設定しました"
        if result["font_size"]:
            msg += f"（サイズ: {result['font_size']}pt）"
        return msg
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_table_of_authorities(category: int | None = None) -> str:
    """判例目録（Table of Authorities）を挿入します。categoryでカテゴリを絞り込み可能。"""
    try:
        result = word.add_table_of_authorities(category=category)
        msg = "判例目録を挿入しました"
        if result["category"]:
            msg += f"（カテゴリ: {result['category']}）"
        return msg
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_clear_all_formatting(paragraph_index: int | None = None) -> str:
    """段落または文書全体の書式をクリアします。paragraph_indexを省略すると文書全体が対象。"""
    try:
        result = word.clear_all_formatting_word(paragraph_index=paragraph_index)
        return f"書式をクリアしました（対象: {result['scope']}）"
    except Exception as e:
        return f"エラー: {e}"
