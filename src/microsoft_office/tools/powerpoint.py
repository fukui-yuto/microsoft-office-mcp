"""PowerPoint MCP tool definitions."""

import json

from microsoft_office.server import mcp
from microsoft_office.office import powerpoint as ppt


@mcp.tool()
def powerpoint_create(file_path: str | None = None) -> str:
    """PowerPointの新規プレゼンテーションを作成します。file_pathを指定すると即座に保存します。"""
    try:
        result = ppt.create_presentation(file_path)
        msg = f"プレゼンテーション '{result['name']}' を作成しました"
        if file_path:
            msg += f" (保存先: {file_path})"
        return msg
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_open(file_path: str) -> str:
    """既存のPowerPointファイル(.pptx)を開きます。"""
    try:
        result = ppt.open_presentation(file_path)
        return f"'{result['name']}' を開きました（スライド数: {result['slide_count']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_save(file_path: str | None = None) -> str:
    """アクティブなプレゼンテーションを保存します。file_pathを指定すると別名保存します。"""
    try:
        result = ppt.save_presentation(file_path)
        return f"'{result['name']}' を保存しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_close() -> str:
    """アクティブなプレゼンテーションを閉じます。"""
    try:
        result = ppt.close_presentation()
        return f"'{result['name']}' を閉じました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_get_info() -> str:
    """アクティブなプレゼンテーションの全スライド情報（スライド数、各スライドのテキスト）を取得します。"""
    try:
        result = ppt.get_info()
        lines = [f"プレゼンテーション: {result['name']} (スライド数: {result['slide_count']})"]
        for slide in result["slides"]:
            lines.append(f"\n--- スライド {slide['slide_number']} ---")
            for text in slide["texts"]:
                lines.append(f"  {text}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_slide(layout: int = 2, position: int | None = None) -> str:
    """スライドを追加します。layout: 1=タイトル, 2=タイトルとコンテンツ, 7=白紙。positionで挿入位置を指定（省略時は末尾）。"""
    try:
        result = ppt.add_slide(layout, position)
        return f"スライド {result['slide_number']} を追加しました（レイアウト: {result['layout']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_delete_slide(slide_number: int) -> str:
    """指定したスライドを削除します（1始まり）。"""
    try:
        result = ppt.delete_slide(slide_number)
        return f"スライド {result['deleted']} を削除しました（残り: {result['remaining']}枚）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_move_slide(from_position: int, to_position: int) -> str:
    """スライドの順番を変更します。from_positionからto_positionへ移動します。"""
    try:
        result = ppt.move_slide(from_position, to_position)
        return f"スライドを {result['from']} → {result['to']} に移動しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_title(slide_number: int, title: str) -> str:
    """指定スライドのタイトルを設定します。"""
    try:
        result = ppt.set_title(slide_number, title)
        return f"スライド {result['slide_number']} のタイトルを '{result['title']}' に設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_body(slide_number: int, body: str) -> str:
    """指定スライドの本文テキストを設定します。改行で複数行を指定できます。"""
    try:
        result = ppt.set_body(slide_number, body)
        return f"スライド {result['slide_number']} の本文を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_get_slide_content(slide_number: int) -> str:
    """指定スライドの全テキストとシェイプ情報（位置・サイズ含む）を取得します。"""
    try:
        result = ppt.get_slide_content(slide_number)
        lines = [f"スライド {result['slide_number']}:"]
        for shape in result["shapes"]:
            if shape["has_text"]:
                lines.append(f"  [{shape['name']}] {shape['text']}")
            else:
                lines.append(f"  [{shape['name']}] (テキストなし)")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_text_format(
    slide_number: int,
    shape_index: int,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_name: str | None = None,
    font_color: list[int] | None = None,
) -> str:
    """指定シェイプのテキスト書式を設定します。shape_indexは1始まり。font_colorは[R,G,B]形式（0-255）。"""
    try:
        result = ppt.set_text_format(
            slide_number, shape_index,
            bold=bold, italic=italic, font_size=font_size, font_name=font_name,
            font_color_rgb=tuple(font_color) if font_color else None,
        )
        if "error" in result:
            return f"エラー: {result['error']}"
        return f"スライド {result['slide_number']} のシェイプ {result['shape_index']} の書式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_background_solid(slide_number: int, color: list[int]) -> str:
    """スライドの背景色を単色で設定します。color=[R,G,B]形式（0-255）。例: [0,51,102]"""
    try:
        result = ppt.set_background_solid(slide_number, *tuple(color))
        return f"スライド {result['slide_number']} の背景色を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_background_gradient(
    slide_number: int,
    color1: list[int],
    color2: list[int],
    direction: int = 1,
) -> str:
    """スライドの背景をグラデーションで設定します。color1/color2=[R,G,B]形式。direction: 1=横, 2=縦, 3=斜め上, 4=斜め下。"""
    try:
        result = ppt.set_background_gradient(
            slide_number,
            color1_rgb=tuple(color1),
            color2_rgb=tuple(color2),
            direction=direction,
        )
        return f"スライド {result['slide_number']} の背景グラデーションを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_design_slide(
    slide_number: int,
    bg_color: list[int] | None = None,
    bg_gradient_color1: list[int] | None = None,
    bg_gradient_color2: list[int] | None = None,
    bg_gradient_direction: int | None = None,
    title_font_name: str | None = None,
    title_font_size: float | None = None,
    title_bold: bool | None = None,
    title_color: list[int] | None = None,
    title_alignment: int | None = None,
    body_font_name: str | None = None,
    body_font_size: float | None = None,
    body_color: list[int] | None = None,
    body_line_spacing: float | None = None,
    body_alignment: int | None = None,
) -> str:
    """スライドに包括的なデザインを適用します。色は[R,G,B]形式（0-255）。
    alignment: 1=左揃え, 2=中央, 3=右揃え。
    bg_gradient_color1/color2/directionを指定するとグラデーション背景になります。"""
    try:
        bg_gradient = None
        if bg_gradient_color1 and bg_gradient_color2:
            bg_gradient = (
                tuple(bg_gradient_color1),
                tuple(bg_gradient_color2),
                bg_gradient_direction if bg_gradient_direction else 1,
            )
        result = ppt.design_slide(
            slide_number,
            bg_rgb=tuple(bg_color) if bg_color else None,
            bg_gradient=bg_gradient,
            title_font_name=title_font_name,
            title_font_size=title_font_size,
            title_bold=title_bold,
            title_color_rgb=tuple(title_color) if title_color else None,
            title_alignment=title_alignment,
            body_font_name=body_font_name,
            body_font_size=body_font_size,
            body_color_rgb=tuple(body_color) if body_color else None,
            body_line_spacing=body_line_spacing,
            body_alignment=body_alignment,
        )
        if "error" in result:
            return f"エラー: {result['error']}"
        return f"スライド {result['slide_number']} にデザインを適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_shape(
    slide_number: int,
    shape_type: int,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color: list[int] | None = None,
    fill_transparency: float = 0.0,
    line_color: list[int] | None = None,
    line_weight: float | None = None,
    shadow: bool = False,
    rotation: float = 0.0,
    z_order_back: bool = False,
) -> str:
    """スライドに図形を追加します。shape_type: 1=四角, 5=角丸四角, 9=楕円, 7=三角。
    座標はポイント単位（16:9スライドは約960x540pt）。fill_color/line_color=[R,G,B]。"""
    try:
        result = ppt.add_shape(
            slide_number, shape_type, left, top, width, height,
            fill_rgb=tuple(fill_color) if fill_color else None,
            fill_transparency=fill_transparency,
            line_rgb=tuple(line_color) if line_color else None,
            line_weight=line_weight,
            shadow=shadow,
            rotation=rotation,
            z_order_back=z_order_back,
        )
        return f"スライド {result['slide_number']} に図形 '{result['shape_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_shape_with_gradient(
    slide_number: int,
    shape_type: int,
    left: float,
    top: float,
    width: float,
    height: float,
    color1: list[int],
    color2: list[int],
    gradient_direction: int = 1,
    transparency: float = 0.0,
    shadow: bool = False,
    z_order_back: bool = False,
) -> str:
    """グラデーション塗りつぶしの図形を追加します。color1/color2=[R,G,B]。
    gradient_direction: 1=横, 2=縦, 3=斜め上, 4=斜め下。"""
    try:
        result = ppt.add_shape_with_gradient(
            slide_number, shape_type, left, top, width, height,
            color1_rgb=tuple(color1),
            color2_rgb=tuple(color2),
            gradient_direction=gradient_direction,
            transparency=transparency,
            shadow=shadow,
            z_order_back=z_order_back,
        )
        return f"スライド {result['slide_number']} にグラデーション図形 '{result['shape_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_textbox(
    slide_number: int,
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color: list[int] | None = None,
    bold: bool = False,
    italic: bool = False,
    alignment: int = 1,
    fill_color: list[int] | None = None,
    fill_transparency: float = 0.0,
) -> str:
    """テキストボックスを追加します。座標はポイント単位。alignment: 1=左, 2=中央, 3=右。
    font_color/fill_color=[R,G,B]。"""
    try:
        result = ppt.add_textbox(
            slide_number, left, top, width, height, text,
            font_name=font_name,
            font_size=font_size,
            font_color_rgb=tuple(font_color) if font_color else None,
            bold=bold,
            italic=italic,
            alignment=alignment,
            fill_rgb=tuple(fill_color) if fill_color else None,
            fill_transparency=fill_transparency,
        )
        return f"スライド {result['slide_number']} にテキストボックス '{result['shape_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_line(
    slide_number: int,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    color: list[int] | None = None,
    weight: float = 1.0,
    dash_style: int = 1,
) -> str:
    """スライドに線を追加します。(x1,y1)から(x2,y2)へ。dash_style: 1=実線, 2=点線, 3=破線, 4=一点鎖線。
    color=[R,G,B]（省略時は黒）。"""
    try:
        result = ppt.add_line(
            slide_number, x1, y1, x2, y2,
            color_rgb=tuple(color) if color else (0, 0, 0),
            weight=weight,
            dash_style=dash_style,
        )
        return f"スライド {result['slide_number']} に線 '{result['shape_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_connector(
    slide_number: int,
    connector_type: int,
    begin_x: float,
    begin_y: float,
    end_x: float,
    end_y: float,
    color: list[int] | None = None,
    weight: float = 1.0,
    dash_style: int = 1,
    begin_arrow: bool = False,
    end_arrow: bool = False,
) -> str:
    """コネクタ（接続線）を追加します。connector_type: 1=直線, 2=カギ線, 3=曲線。
    color=[R,G,B]。begin_arrow/end_arrowで矢印の有無を指定。"""
    try:
        result = ppt.add_connector(
            slide_number, connector_type, begin_x, begin_y, end_x, end_y,
            color_rgb=tuple(color) if color else None,
            weight=weight,
            dash_style=dash_style,
            begin_arrow=begin_arrow,
            end_arrow=end_arrow,
        )
        return f"スライド {result['slide_number']} にコネクタ '{result['shape_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_bullets(
    slide_number: int,
    bullet_char: int = 8226,
    bullet_color: list[int] | None = None,
    bullet_size: float | None = None,
    exclude_texts: list[str] | None = None,
) -> str:
    """本文テキストに箇条書きを設定します。bullet_char: 文字コード（8226=●）。
    bullet_color=[R,G,B]。bullet_size: 相対サイズ。exclude_textsで除外するテキストを指定。"""
    try:
        result = ppt.set_bullets(
            slide_number,
            bullet_char=bullet_char,
            bullet_color_rgb=tuple(bullet_color) if bullet_color else None,
            bullet_size=bullet_size,
            exclude_texts=exclude_texts,
        )
        if "error" in result:
            return f"エラー: {result['error']}"
        return f"スライド {result['slide_number']} に箇条書きを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_clear_extra_shapes(slide_number: int) -> str:
    """タイトルと本文以外の余分な図形をすべて削除します。"""
    try:
        result = ppt.clear_extra_shapes(slide_number)
        return f"スライド {result['slide_number']} から {result['deleted']} 個の図形を削除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_reposition_shape(
    slide_number: int,
    shape_index: int,
    left: float | None = None,
    top: float | None = None,
    width: float | None = None,
    height: float | None = None,
) -> str:
    """図形の位置やサイズを変更します。shape_indexは1始まり。座標はポイント単位。変更したい値のみ指定。"""
    try:
        result = ppt.reposition_shape(
            slide_number, shape_index,
            left=left, top=top, width=width, height=height,
        )
        return f"スライド {result['slide_number']} の図形 '{result['shape_name']}' の位置を変更しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_shadow(
    slide_number: int,
    shape_index: int,
    blur: float = 8,
    offset_x: float = 3,
    offset_y: float = 3,
    transparency: float = 0.6,
) -> str:
    """図形に影効果を追加します。blur: ぼかし量、offset_x/y: オフセット、transparency: 透明度（0-1）。"""
    try:
        result = ppt.set_shape_shadow(
            slide_number, shape_index,
            blur=blur, offset_x=offset_x, offset_y=offset_y,
            transparency=transparency,
        )
        return f"スライド {result['slide_number']} の図形に影効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_glow(
    slide_number: int,
    shape_index: int,
    color: list[int],
    radius: float = 10.0,
    transparency: float = 0.0,
) -> str:
    """図形に光彩（グロー）効果を追加します。color=[R,G,B]。radius: 光彩の半径、transparency: 透明度（0-1）。"""
    try:
        result = ppt.set_shape_glow(
            slide_number, shape_index,
            color_rgb=tuple(color),
            radius=radius,
            transparency=transparency,
        )
        return f"スライド {result['slide_number']} の図形に光彩効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_reflection(
    slide_number: int,
    shape_index: int,
    reflection_type: int = 1,
) -> str:
    """図形に反射効果を追加します。reflection_type: 1-9の反射スタイル。"""
    try:
        result = ppt.set_shape_reflection(
            slide_number, shape_index,
            reflection_type=reflection_type,
        )
        return f"スライド {result['slide_number']} の図形に反射効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_soft_edges(
    slide_number: int,
    shape_index: int,
    radius: float = 10.0,
) -> str:
    """図形にぼかし（ソフトエッジ）効果を追加します。radius: ぼかしの半径（ポイント）。"""
    try:
        result = ppt.set_shape_soft_edges(
            slide_number, shape_index,
            radius=radius,
        )
        return f"スライド {result['slide_number']} の図形にぼかし効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_3d(
    slide_number: int,
    shape_index: int,
    bevel_type: int = 3,
    bevel_width: float = 6,
    bevel_height: float = 6,
    depth: float = 0,
    material: int = 1,
) -> str:
    """図形に3D効果を追加します。bevel_type: 面取りの種類（1-12）。material: 質感（1=つや消し, 2=光沢, 3=メタル）。"""
    try:
        result = ppt.set_shape_3d(
            slide_number, shape_index,
            bevel_type=bevel_type,
            bevel_width=bevel_width,
            bevel_height=bevel_height,
            depth=depth,
            material=material,
        )
        return f"スライド {result['slide_number']} の図形に3D効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_table(
    slide_number: int,
    rows: int,
    cols: int,
    left: float,
    top: float,
    width: float,
    height: float,
    data: list[list[str]] | None = None,
    header_fill_color: list[int] | None = None,
    header_font_color: list[int] | None = None,
    border_color: list[int] | None = None,
) -> str:
    """スライドに表を挿入します。dataは2次元配列（行×列）。
    header_fill_color/header_font_color/border_color=[R,G,B]。"""
    try:
        result = ppt.add_table(
            slide_number, rows, cols, left, top, width, height,
            data=data,
            header_fill_color_rgb=tuple(header_fill_color) if header_fill_color else None,
            header_font_color_rgb=tuple(header_font_color) if header_font_color else None,
            border_color_rgb=tuple(border_color) if border_color else None,
        )
        return f"スライド {result['slide_number']} に {rows}行×{cols}列 の表を追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_chart(
    slide_number: int,
    chart_type: int,
    left: float,
    top: float,
    width: float,
    height: float,
    categories: list[str],
    series_data: list[dict],
    title: str | None = None,
    has_legend: bool = True,
) -> str:
    """スライドにグラフを追加します。chart_type: 51=棒グラフ, 4=折れ線, 5=円, 57=横棒。
    series_data=[{"name": "系列名", "values": [1,2,3]}, ...]。categories=["A","B","C"]。"""
    try:
        result = ppt.add_chart(
            slide_number, chart_type, left, top, width, height,
            categories=categories,
            series_data=series_data,
            title=title,
            has_legend=has_legend,
        )
        return f"スライド {result['slide_number']} にグラフを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_speaker_notes(slide_number: int, notes_text: str) -> str:
    """指定スライドにスピーカーノート（発表者ノート）を設定します。"""
    try:
        result = ppt.set_speaker_notes(slide_number, notes_text)
        return f"スライド {result['slide_number']} にノートを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_get_speaker_notes(slide_number: int) -> str:
    """指定スライドのスピーカーノート（発表者ノート）を取得します。"""
    try:
        result = ppt.get_speaker_notes(slide_number)
        return f"スライド {result['slide_number']} のノート:\n{result['notes']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_export_to_pdf(output_path: str) -> str:
    """アクティブなプレゼンテーションをPDFファイルとして出力します。"""
    try:
        result = ppt.export_to_pdf(output_path)
        return f"PDFを出力しました: {result['output_path']}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_slide_transition(
    slide_number: int,
    transition_type: int = 0,
    duration: float = 1.0,
    advance_on_click: bool = True,
    advance_time: float | None = None,
) -> str:
    """スライドの画面切り替え効果を設定します。transition_type: 0=なし, 1=カット, 2=フェード等。
    duration: 切り替え時間（秒）。advance_time: 自動切り替え時間（秒）。"""
    try:
        result = ppt.set_slide_transition(
            slide_number,
            transition_type=transition_type,
            duration=duration,
            advance_on_click=advance_on_click,
            advance_time=advance_time,
        )
        return f"スライド {result['slide_number']} に画面切り替え効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_animation(
    slide_number: int,
    shape_index: int,
    effect_type: int,
    trigger: int = 1,
    duration: float = 0.5,
    delay: float = 0.0,
) -> str:
    """図形にアニメーション効果を追加します。effect_type: アニメーションの種類。
    trigger: 1=クリック時, 2=前の動作と同時, 3=前の動作の後。duration: 再生時間（秒）。delay: 遅延（秒）。"""
    try:
        result = ppt.add_animation(
            slide_number, shape_index,
            effect_type=effect_type,
            trigger=trigger,
            duration=duration,
            delay=delay,
        )
        return f"スライド {result['slide_number']} の図形にアニメーションを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_group_shapes(slide_number: int, shape_indices: list[int]) -> str:
    """複数の図形をグループ化します。shape_indicesは1始まりのインデックスのリスト。例: [1,2,3]"""
    try:
        result = ppt.group_shapes(slide_number, shape_indices)
        return f"スライド {result['slide_number']} で {len(shape_indices)} 個の図形をグループ化しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_ungroup_shapes(slide_number: int, shape_index: int) -> str:
    """グループ化された図形を解除します。shape_indexはグループのインデックス（1始まり）。"""
    try:
        result = ppt.ungroup_shapes(slide_number, shape_index)
        return f"スライド {result['slide_number']} のグループを解除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_duplicate_slide(slide_number: int) -> str:
    """指定スライドを複製します。複製はすぐ後ろに挿入されます。"""
    try:
        result = ppt.duplicate_slide(slide_number)
        return f"スライド {result['slide_number']} を複製しました（新しいスライド: {result['new_slide_number']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_image(
    slide_number: int,
    file_path: str,
    left: float,
    top: float,
    width: float | None = None,
    height: float | None = None,
) -> str:
    """スライドに画像を挿入します。座標はポイント単位。width/heightを省略すると元のサイズで挿入。"""
    try:
        result = ppt.add_image(
            slide_number, file_path, left, top,
            width=width, height=height,
        )
        return f"スライド {result['slide_number']} に画像 '{result['shape_name']}' を挿入しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_crop_image(
    slide_number: int,
    shape_index: int,
    crop_left: float = 0,
    crop_right: float = 0,
    crop_top: float = 0,
    crop_bottom: float = 0,
) -> str:
    """画像をトリミングします。各値は0-1の割合（例: 0.1=10%トリミング）。"""
    try:
        result = ppt.crop_image(
            slide_number, shape_index,
            crop_left=crop_left, crop_right=crop_right,
            crop_top=crop_top, crop_bottom=crop_bottom,
        )
        return f"スライド {result['slide_number']} の画像をトリミングしました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_image_effects(
    slide_number: int,
    shape_index: int,
    brightness: float | None = None,
    contrast: float | None = None,
) -> str:
    """画像の明るさとコントラストを調整します。値は-1.0〜1.0（0=変更なし）。"""
    try:
        result = ppt.set_image_effects(
            slide_number, shape_index,
            brightness=brightness, contrast=contrast,
        )
        return f"スライド {result['slide_number']} の画像効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_slide_size(width: float, height: float) -> str:
    """スライドのサイズを変更します。単位はポイント。標準16:9=(960,540)、4:3=(720,540)。"""
    try:
        result = ppt.set_slide_size(width, height)
        return f"スライドサイズを {result['width']}x{result['height']} pt に変更しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_section(name: str, slide_number: int) -> str:
    """指定スライドの位置にセクションを追加します。"""
    try:
        result = ppt.add_section(name, slide_number)
        return f"セクション '{result['name']}' をスライド {result['slide_number']} に追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_get_sections() -> str:
    """プレゼンテーションのセクション一覧を取得します。"""
    try:
        result = ppt.get_sections()
        if not result["sections"]:
            return "セクションはありません"
        lines = ["セクション一覧:"]
        for section in result["sections"]:
            lines.append(f"  {section['index']}. {section['name']} (スライド数: {section['slide_count']})")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_text(slide_number: int, shape_index: int, text: str) -> str:
    """図形内のテキストを設定します。shape_indexは1始まり。"""
    try:
        result = ppt.set_shape_text(slide_number, shape_index, text)
        return f"スライド {result['slide_number']} のシェイプ {result['shape_index']} にテキストを設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Shape Management
# ============================================================

@mcp.tool()
def powerpoint_align_shapes(slide_number: int, shape_indices: list[int], alignment: str) -> str:
    """複数の図形を整列します。shape_indicesは1始まりのインデックスのリスト。
    alignment: "left"=左揃え, "center"=中央揃え, "right"=右揃え, "top"=上揃え, "middle"=中央揃え(縦), "bottom"=下揃え。"""
    try:
        result = ppt.align_shapes(slide_number, shape_indices, alignment)
        return f"スライド {result['slide_number']} の図形を {result['alignment']} に整列しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_distribute_shapes(slide_number: int, shape_indices: list[int], direction: str) -> str:
    """複数の図形を均等に配置します。direction: "horizontal"=水平, "vertical"=垂直。"""
    try:
        result = ppt.distribute_shapes(slide_number, shape_indices, direction)
        return f"スライド {result['slide_number']} の図形を {result['direction']} に均等配置しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_z_order(slide_number: int, shape_index: int, order: str) -> str:
    """図形のZオーダー（重なり順）を変更します。order: "front"=最前面, "back"=最背面, "forward"=1つ前面, "backward"=1つ背面。"""
    try:
        result = ppt.set_shape_z_order(slide_number, shape_index, order)
        return f"スライド {result['slide_number']} の図形のZオーダーを {result['order']} に変更しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_duplicate_shape(slide_number: int, shape_index: int) -> str:
    """図形を同じスライド上に複製します。shape_indexは1始まり。"""
    try:
        result = ppt.duplicate_shape(slide_number, shape_index)
        return f"スライド {result['slide_number']} の図形を複製しました（新しい図形: {result['new_shape_name']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_delete_shape(slide_number: int, shape_index: int) -> str:
    """指定した図形を削除します。shape_indexは1始まり。"""
    try:
        result = ppt.delete_shape(slide_number, shape_index)
        return f"スライド {result['slide_number']} の図形 '{result['shape_name']}' を削除しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_get_shape_properties(slide_number: int, shape_index: int) -> str:
    """図形の詳細プロパティ（位置、サイズ、タイプ、テキスト、塗りつぶし情報）を取得します。shape_indexは1始まり。"""
    try:
        result = ppt.get_shape_properties(slide_number, shape_index)
        lines = [f"スライド {result['slide_number']} のシェイプ {result['shape_index']}:"]
        lines.append(f"  名前: {result['name']}")
        lines.append(f"  位置: Left={result['left']}, Top={result['top']}")
        lines.append(f"  サイズ: Width={result['width']}, Height={result['height']}")
        lines.append(f"  回転: {result['rotation']}°")
        lines.append(f"  タイプ: {result['shape_type']}")
        if result.get("text"):
            lines.append(f"  テキスト: {result['text']}")
        if "fill_type" in result:
            lines.append(f"  塗りつぶしタイプ: {result['fill_type']}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_copy_shape_to_slide(source_slide: int, shape_index: int, target_slide: int) -> str:
    """図形を別のスライドにコピーします。shape_indexは1始まり。"""
    try:
        result = ppt.copy_shape_to_slide(source_slide, shape_index, target_slide)
        return f"スライド {result['source_slide']} の図形をスライド {result['target_slide']} にコピーしました（新しい図形: {result['new_shape_name']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_size(slide_number: int, shape_index: int, width: float, height: float) -> str:
    """図形のサイズを変更します。shape_indexは1始まり。width/heightはポイント単位。"""
    try:
        result = ppt.set_shape_size(slide_number, shape_index, width, height)
        return f"スライド {result['slide_number']} の図形 '{result['shape_name']}' のサイズを変更しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Advanced Text Effects
# ============================================================

@mcp.tool()
def powerpoint_set_text_shadow(
    slide_number: int,
    shape_index: int,
    blur_radius: float = 5,
    distance: float = 3,
    angle: float = 45,
    color: list[int] | None = None,
    transparency: float = 60,
) -> str:
    """テキストに影効果を追加します。color=[R,G,B]。transparency: 透明度（0-100）。"""
    try:
        result = ppt.set_text_shadow(
            slide_number, shape_index,
            blur_radius=blur_radius,
            distance=distance,
            angle=angle,
            color_rgb=tuple(color) if color else (0, 0, 0),
            transparency=transparency,
        )
        return f"スライド {result['slide_number']} のテキストに影効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_text_glow(
    slide_number: int,
    shape_index: int,
    radius: float = 5,
    color: list[int] | None = None,
    transparency: float = 40,
) -> str:
    """テキストに光彩（グロー）効果を追加します。color=[R,G,B]。transparency: 透明度（0-100）。"""
    try:
        result = ppt.set_text_glow(
            slide_number, shape_index,
            radius=radius,
            color_rgb=tuple(color) if color else (255, 255, 0),
            transparency=transparency,
        )
        return f"スライド {result['slide_number']} のテキストに光彩効果を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_text_outline(
    slide_number: int,
    shape_index: int,
    color: list[int] | None = None,
    weight: float = 1.0,
) -> str:
    """テキストにアウトライン（縁取り）を追加します。color=[R,G,B]。"""
    try:
        result = ppt.set_text_outline(
            slide_number, shape_index,
            color_rgb=tuple(color) if color else (0, 0, 0),
            weight=weight,
        )
        return f"スライド {result['slide_number']} のテキストにアウトラインを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_text_gradient_fill(
    slide_number: int,
    shape_index: int,
    color1: list[int] | None = None,
    color2: list[int] | None = None,
    angle: float = 0,
) -> str:
    """テキストにグラデーション塗りつぶしを適用します。color1/color2=[R,G,B]。"""
    try:
        result = ppt.set_text_gradient_fill(
            slide_number, shape_index,
            color1_rgb=tuple(color1) if color1 else (0, 0, 255),
            color2_rgb=tuple(color2) if color2 else (255, 0, 255),
            angle=angle,
        )
        return f"スライド {result['slide_number']} のテキストにグラデーションを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_rich_textbox(
    slide_number: int,
    left: float,
    top: float,
    width: float,
    height: float,
    runs: list[dict],
) -> str:
    """リッチテキストボックスを追加します。runsはランごとの書式指定リスト。
    runs=[{"text": "Hello ", "font_size": 24, "bold": true, "font_color": [255,0,0]}, {"text": "World", "font_size": 18, "italic": true}]"""
    try:
        # Convert font_color lists to tuples in runs
        converted_runs = []
        for run in runs:
            r = dict(run)
            if "font_color" in r and r["font_color"] is not None:
                r["font_color"] = tuple(r["font_color"])
            converted_runs.append(r)
        result = ppt.add_rich_textbox(
            slide_number, left, top, width, height, converted_runs,
        )
        return f"スライド {result['slide_number']} にリッチテキストボックス '{result['shape_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Advanced Fills
# ============================================================

@mcp.tool()
def powerpoint_set_shape_pattern_fill(
    slide_number: int,
    shape_index: int,
    pattern_type: int,
    fore_color: list[int],
    back_color: list[int],
) -> str:
    """図形にパターン塗りつぶしを適用します。pattern_type: msoPatternの値。fore_color/back_color=[R,G,B]。"""
    try:
        result = ppt.set_shape_pattern_fill(
            slide_number, shape_index,
            pattern_type=pattern_type,
            fore_color_rgb=tuple(fore_color),
            back_color_rgb=tuple(back_color),
        )
        return f"スライド {result['slide_number']} の図形にパターン塗りつぶしを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_texture_fill(
    slide_number: int,
    shape_index: int,
    texture_path: str,
) -> str:
    """図形に画像ファイルからのテクスチャ塗りつぶしを適用します。"""
    try:
        result = ppt.set_shape_texture_fill(slide_number, shape_index, texture_path)
        return f"スライド {result['slide_number']} の図形にテクスチャ塗りつぶしを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_picture_fill(
    slide_number: int,
    shape_index: int,
    image_path: str,
    stretch: bool = True,
) -> str:
    """図形に画像塗りつぶしを適用します。stretch=Trueで画像を伸縮して塗りつぶし。"""
    try:
        result = ppt.set_shape_picture_fill(
            slide_number, shape_index, image_path, stretch=stretch,
        )
        return f"スライド {result['slide_number']} の図形に画像塗りつぶしを設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Freeform Shapes
# ============================================================

@mcp.tool()
def powerpoint_add_freeform_shape(
    slide_number: int,
    points: list[list[float]],
    fill_color: list[int] | None = None,
    line_color: list[int] | None = None,
    line_weight: float = 1.0,
    closed: bool = True,
) -> str:
    """フリーフォーム（自由形状）図形を追加します。pointsは[[x,y],...]形式の座標リスト。
    fill_color/line_color=[R,G,B]。closed=Trueでパスを閉じます。"""
    try:
        result = ppt.add_freeform_shape(
            slide_number, points,
            fill_color_rgb=tuple(fill_color) if fill_color else None,
            line_color_rgb=tuple(line_color) if line_color else None,
            line_weight=line_weight,
            closed=closed,
        )
        return f"スライド {result['slide_number']} にフリーフォーム図形 '{result['shape_name']}' を追加しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Advanced Table
# ============================================================

@mcp.tool()
def powerpoint_format_table_cell(
    slide_number: int,
    shape_index: int,
    row: int,
    col: int,
    fill_color: list[int] | None = None,
    font_color: list[int] | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    alignment: int | None = None,
) -> str:
    """テーブルの個別セルの書式を設定します。row/colは1始まり。fill_color/font_color=[R,G,B]。
    alignment: 1=左揃え, 2=中央, 3=右揃え。"""
    try:
        result = ppt.format_table_cell(
            slide_number, shape_index, row, col,
            fill_color_rgb=tuple(fill_color) if fill_color else None,
            font_color_rgb=tuple(font_color) if font_color else None,
            font_size=font_size,
            bold=bold,
            alignment=alignment,
        )
        return f"スライド {result['slide_number']} のテーブルセル({row},{col})の書式を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_merge_table_cells(
    slide_number: int,
    shape_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> str:
    """テーブルのセルを結合します。start_row/start_col/end_row/end_colは1始まり。"""
    try:
        result = ppt.merge_table_cells(
            slide_number, shape_index,
            start_row, start_col, end_row, end_col,
        )
        return f"スライド {result['slide_number']} のテーブルセルを結合しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_table_border(
    slide_number: int,
    shape_index: int,
    border_type: str,
    color: list[int] | None = None,
    weight: float | None = None,
) -> str:
    """テーブルの罫線を設定します。border_type: "all"=すべて, "outside"=外枠, "inside"=内側, "none"=なし。
    color=[R,G,B]。"""
    try:
        result = ppt.set_table_border(
            slide_number, shape_index, border_type,
            color_rgb=tuple(color) if color else None,
            weight=weight,
        )
        return f"スライド {result['slide_number']} のテーブル罫線を設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Slide Master & Layout
# ============================================================

@mcp.tool()
def powerpoint_get_slide_layouts() -> str:
    """利用可能なスライドレイアウトの一覧を取得します。"""
    try:
        result = ppt.get_slide_layouts()
        if not result["layouts"]:
            return "利用可能なレイアウトはありません"
        lines = ["スライドレイアウト一覧:"]
        for layout in result["layouts"]:
            lines.append(f"  {layout['index']}. {layout['name']}")
        return "\n".join(lines)
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_apply_slide_layout(slide_number: int, layout_index: int) -> str:
    """スライドにレイアウトを適用します。layout_indexはget_slide_layoutsで取得したインデックス。"""
    try:
        result = ppt.apply_slide_layout(slide_number, layout_index)
        return f"スライド {result['slide_number']} にレイアウト {result['layout_index']} を適用しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_slide_number_visibility(visible: bool = True, start_number: int = 1) -> str:
    """スライド番号の表示/非表示を切り替えます。start_numberで開始番号を指定。"""
    try:
        result = ppt.set_slide_number_visibility(visible=visible, start_number=start_number)
        status = "表示" if result["visible"] else "非表示"
        return f"スライド番号を{status}に設定しました（開始番号: {result['start_number']}）"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Media
# ============================================================

@mcp.tool()
def powerpoint_add_video(
    slide_number: int,
    file_path: str,
    left: float,
    top: float,
    width: float,
    height: float,
) -> str:
    """スライドに動画を埋め込みます。座標はポイント単位。"""
    try:
        result = ppt.add_video(slide_number, file_path, left, top, width, height)
        return f"スライド {result['slide_number']} に動画 '{result['shape_name']}' を埋め込みました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_add_audio(
    slide_number: int,
    file_path: str,
    left: float = 0,
    top: float = 0,
    play_across_slides: bool = False,
) -> str:
    """スライドに音声ファイルを埋め込みます。play_across_slides=Trueでスライド間再生。"""
    try:
        result = ppt.add_audio(
            slide_number, file_path,
            left=left, top=top,
            play_across_slides=play_across_slides,
        )
        return f"スライド {result['slide_number']} に音声 '{result['shape_name']}' を埋め込みました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Advanced Animation
# ============================================================

@mcp.tool()
def powerpoint_add_motion_path(
    slide_number: int,
    shape_index: int,
    path_type: str,
    duration: float = 1.0,
) -> str:
    """図形にモーションパスアニメーションを追加します。path_type: "line", "arc", "circle", "diamond", "custom"。"""
    try:
        result = ppt.add_motion_path(
            slide_number, shape_index,
            path_type=path_type,
            duration=duration,
        )
        return f"スライド {result['slide_number']} の図形にモーションパスを追加しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_animation_trigger(
    slide_number: int,
    shape_index: int,
    trigger_type: str,
    trigger_shape_index: int | None = None,
) -> str:
    """アニメーションのトリガーを設定します。trigger_type: "on_click", "with_previous", "after_previous"。"""
    try:
        result = ppt.set_animation_trigger(
            slide_number, shape_index,
            trigger_type=trigger_type,
            trigger_shape_index=trigger_shape_index,
        )
        return f"スライド {result['slide_number']} のアニメーショントリガーを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_animation_order(
    slide_number: int,
    effect_index: int,
    new_position: int,
) -> str:
    """アニメーションの再生順序を変更します。effect_index/new_positionは1始まり。"""
    try:
        result = ppt.set_animation_order(
            slide_number, effect_index, new_position,
        )
        return f"スライド {result['slide_number']} のアニメーション順序を変更しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Advanced Shape Effects
# ============================================================

@mcp.tool()
def powerpoint_set_shape_border(
    slide_number: int,
    shape_index: int,
    color: list[int] | None = None,
    weight: float | None = None,
    dash_style: str | None = None,
) -> str:
    """図形の枠線を設定します。color=[R,G,B]。dash_style: "solid", "dash", "dot", "dash_dot"。"""
    try:
        result = ppt.set_shape_border(
            slide_number, shape_index,
            color_rgb=tuple(color) if color else None,
            weight=weight,
            dash_style=dash_style,
        )
        return f"スライド {result['slide_number']} の図形の枠線を設定しました"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Export (Single Slide)
# ============================================================

@mcp.tool()
def powerpoint_export_slide_as_image(
    slide_number: int,
    file_path: str,
    width: int = 1920,
    height: int = 1080,
) -> str:
    """スライドを画像ファイル（PNG/JPG）としてエクスポートします。"""
    try:
        result = ppt.export_slide_as_image(
            slide_number, file_path,
            width=width, height=height,
        )
        return f"スライド {result['slide_number']} を画像として出力しました: {result['file_path']}"
    except Exception as e:
        return f"エラー: {e}"


# ============================================================
# Shape Text Advanced
# ============================================================

@mcp.tool()
def powerpoint_set_shape_text_vertical(
    slide_number: int,
    shape_index: int,
    orientation: str,
) -> str:
    """テキストの方向を設定します。orientation: "horizontal"=横書き, "vertical"=縦書き, "vertical270"=270度回転, "stacked"=縦中横。"""
    try:
        result = ppt.set_shape_text_vertical(
            slide_number, shape_index, orientation,
        )
        return f"スライド {result['slide_number']} の図形のテキスト方向を設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_text_margin(
    slide_number: int,
    shape_index: int,
    left: float = 7.2,
    top: float = 3.6,
    right: float = 7.2,
    bottom: float = 3.6,
) -> str:
    """図形の内部テキストマージンを設定します。単位はポイント。"""
    try:
        result = ppt.set_shape_text_margin(
            slide_number, shape_index,
            left=left, top=top, right=right, bottom=bottom,
        )
        return f"スライド {result['slide_number']} の図形のテキストマージンを設定しました"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def powerpoint_set_shape_autofit(
    slide_number: int,
    shape_index: int,
    autofit_type: str,
) -> str:
    """テキストの自動調整を設定します。autofit_type: "none"=調整なし, "shrink"=テキストを縮小, "resize_shape"=図形をテキストに合わせる。"""
    try:
        result = ppt.set_shape_autofit(
            slide_number, shape_index, autofit_type,
        )
        return f"スライド {result['slide_number']} の図形の自動調整を設定しました"
    except Exception as e:
        return f"エラー: {e}"
