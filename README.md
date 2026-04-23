# Microsoft Office MCP Server

Claude Code から PowerPoint・Word・Excel をリアルタイムで操作する MCP サーバーです。

Windows の COM オートメーション（pywin32）を使い、**開いている Office アプリを直接制御**します。
Claude と会話しながら、目の前で資料が作られていく体験を実現します。

## デモ

```
あなた: 「3枚のスライドで自己紹介プレゼンを作って」
Claude: PowerPointが開き、スライドが1枚ずつ追加されていく
あなた: 「2枚目のタイトルを変えて」
Claude: そのまま修正される
```

## 必要な環境

- Windows 10/11
- Microsoft Office（PowerPoint / Word / Excel）がインストール済み
- Python 3.12+
- [uv](https://docs.astral.sh/uv/)
- [Claude Code](https://docs.anthropic.com/en/docs/claude-code)

## セットアップ

```bash
# リポジトリをクローン
git clone https://github.com/fukui-yuto/microsoft-office-mcp.git
cd microsoft-office-mcp

# 依存関係をインストール
uv sync
```

## Claude Code への登録

### 方法1: グローバル設定（全プロジェクトで使用可能）

`~/.claude/settings.json` に追加：

```json
{
  "mcpServers": {
    "microsoft-office": {
      "command": "uv",
      "args": [
        "run",
        "--directory",
        "C:\\Users\\<ユーザー名>\\path\\to\\microsoft-office-mcp",
        "microsoft-office"
      ]
    }
  }
}
```

### 方法2: プロジェクト設定（このリポジトリ内のみ）

リポジトリ直下の `.mcp.json` に既に設定済みです。

登録後、**Claude Code を再起動**してください。

## 使い方

Claude Code で普通に会話するだけです：

### PowerPoint

```
「新しいプレゼンを作って」
「5枚のスライドで会社紹介資料を作って、デザインも凝って」
「3枚目にグラフを追加して」
「スライドにフェードのアニメーションを付けて」
「PDFで出力して」
```

### Word

```
「議事録のテンプレートを作って」
「ヘッダーとフッターを追加して」
「目次を挿入して」
「透かしを入れて」
「PDFで出力して」
```

### Excel

```
「売上管理表を作って、罫線と書式も整えて」
「棒グラフを追加して」
「条件付き書式を設定して」
「ピボットテーブルを作成して」
「PDFで出力して」
```

## ツール一覧（120ツール）

### PowerPoint（45ツール）

| カテゴリ | ツール |
|----------|--------|
| 基本操作 | `create` `open` `save` `close` `get_info` |
| スライド管理 | `add_slide` `delete_slide` `move_slide` `duplicate_slide` `set_slide_size` |
| テキスト | `set_title` `set_body` `get_slide_content` `set_text_format` `set_shape_text` `set_bullets` |
| デザイン | `design_slide` `set_background_solid` `set_background_gradient` |
| 図形 | `add_shape` `add_shape_with_gradient` `add_textbox` `add_line` `add_connector` `reposition_shape` `clear_extra_shapes` `group_shapes` `ungroup_shapes` |
| 図形効果 | `set_shape_shadow` `set_shape_glow` `set_shape_reflection` `set_shape_soft_edges` `set_shape_3d` |
| 表・グラフ | `add_table` `add_chart` |
| 画像 | `add_image` `crop_image` `set_image_effects` |
| アニメーション | `set_slide_transition` `add_animation` |
| ノート・セクション | `set_speaker_notes` `get_speaker_notes` `add_section` `get_sections` |
| エクスポート | `export_to_pdf` |

### Word（36ツール）

| カテゴリ | ツール |
|----------|--------|
| 基本操作 | `create` `open` `save` `close` `get_content` |
| 段落・見出し | `add_paragraph` `add_heading` `edit_paragraph` `delete_paragraph` `insert_paragraph_at` |
| 書式設定 | `set_font` `set_paragraph_format_extended` |
| リスト | `add_list` |
| 表 | `insert_table` `format_table` |
| ページ設定 | `set_page_setup` `add_header` `add_footer` `add_page_numbers` `set_columns` `add_section_break` |
| 挿入 | `insert_image` `add_text_box` `add_hyperlink` `add_bookmark` `get_bookmarks` `add_table_of_contents` |
| 検索・置換 | `find_and_replace` |
| レビュー | `add_comment` `get_comments` `enable_track_changes` `accept_all_changes` `reject_all_changes` |
| 装飾 | `add_watermark` `set_page_borders` |
| エクスポート | `export_to_pdf` |

### Excel（39ツール）

| カテゴリ | ツール |
|----------|--------|
| 基本操作 | `create` `open` `save` `close` |
| シート管理 | `get_sheets` `add_sheet` `delete_sheet` `rename_sheet` |
| データ読み書き | `read_cell` `write_cell` `read_range` `write_range` `get_sheet_data` |
| 書式設定 | `set_cell_format` `set_cell_fill` `set_cell_borders` `set_cell_alignment` `set_number_format` |
| レイアウト | `merge_cells` `unmerge_cells` `set_column_width` `set_row_height` `freeze_panes` |
| 数式 | `set_formula` |
| 名前付き範囲 | `add_named_range` `delete_named_range` `get_named_ranges` |
| データ機能 | `sort_range` `set_auto_filter` `set_conditional_formatting` `add_data_validation` |
| グラフ | `add_chart` |
| ピボット | `create_pivot_table` |
| コメント | `add_comment` `get_comments` |
| 印刷・出力 | `set_print_setup` `export_to_pdf` |
| 保護 | `protect_sheet` `unprotect_sheet` |

## 技術構成

```
src/microsoft_office/
├── server.py        # FastMCP サーバー（エントリポイント）
├── com_utils.py     # COM 共通ユーティリティ
├── office/          # COM オートメーションロジック（3,628行）
│   ├── powerpoint.py   # 1,599行 - 33関数
│   ├── word.py         # 851行 - 36関数
│   └── excel.py        # 1,178行 - 39関数
└── tools/           # MCP ツール定義（1,934行）
    ├── powerpoint.py   # 816行 - 45ツール
    ├── word.py         # 524行 - 36ツール
    └── excel.py        # 594行 - 39ツール
```

- **MCP SDK**: `mcp[cli]` (FastMCP)
- **Office 制御**: `pywin32` (win32com.client) による COM オートメーション
- **パッケージ管理**: `uv`

## 注意事項

- Windows 専用です（COM オートメーション依存）
- Office アプリが起動していない場合、自動で起動します
- 同じ Office アプリの複数インスタンスでの同時操作は避けてください
- ファイルパスは絶対パスが推奨です（相対パスは自動変換されます）

## ライセンス

MIT
