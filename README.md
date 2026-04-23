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
「モダンなタイトルスライドをデザインして」
「プロセスフローを矢印スタイルで作って」
「比較スライドをVS形式で作って」
「チームメンバー紹介スライドを作って」
「スライドにフェードのアニメーションを付けて」
「PDFで出力して」
```

### Word

```
「議事録のテンプレートを作って」
「ビジネスレターを作成して」
「請求書を作って」
「レポートテンプレートを生成して」
「履歴書をモダンスタイルで作って」
「契約書テンプレートを作って」
「PDFで出力して」
```

### Excel

```
「売上管理表を作って、罫線と書式も整えて」
「ダッシュボードヘッダーを作って」
「KPIカードを並べて」
「ガントチャートを作成して」
「スコアカードを信号機スタイルで作って」
「財務レポートのレイアウトを作って」
「PDFで出力して」
```

## ツール一覧（484ツール）

### PowerPoint 基本（115ツール）

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
| 図形管理 | `align_shapes` `distribute_shapes` `set_shape_z_order` `duplicate_shape` `delete_shape` `get_shape_properties` `copy_shape_to_slide` `set_shape_size` |
| テキスト効果 | `set_text_shadow` `set_text_glow` `set_text_outline` `set_text_gradient_fill` `add_rich_textbox` |
| 高度な塗りつぶし | `set_shape_pattern_fill` `set_shape_texture_fill` `set_shape_picture_fill` |
| フリーフォーム | `add_freeform_shape` |
| 表の高度な操作 | `format_table_cell` `merge_table_cells` `set_table_border` |
| マスター・レイアウト | `get_slide_layouts` `apply_slide_layout` `set_slide_number_visibility` |
| メディア | `add_video` `add_audio` |
| 高度なアニメーション | `add_motion_path` `set_animation_trigger` `set_animation_order` |
| 図形の装飾 | `set_shape_border` |
| スライドエクスポート | `export_slide_as_image` |
| テキスト配置 | `set_shape_text_vertical` `set_shape_text_margin` `set_shape_autofit` |
| スライド操作 | `set_slide_background_image` `get_slide_count` `clear_slide` |
| 図形詳細 | `set_shape_opacity` `rotate_shape` `flip_shape` `set_shape_name` `find_shape_by_name` `list_shapes` |
| テキスト詳細 | `add_superscript` `add_subscript` `set_paragraph_spacing` `set_text_columns` `clear_shape_text` |
| グラフ書式 | `format_chart_title` `format_chart_axis` `format_chart_legend` `set_chart_style` `format_chart_data_labels` `add_chart_data_table` |
| ダイアグラム | `create_org_chart` `create_pyramid_diagram` `create_circular_diagram` `create_matrix_diagram` |
| プレゼン全体 | `set_all_slides_background` `apply_font_to_all` `get_presentation_summary` |
| グラデーション | `set_shape_gradient` |
| テキスト操作 | `replace_text` `replace_text_all_slides` `get_all_text` |
| リンク・メディア | `set_shape_hyperlink` `add_qr_code_shape` |
| ヘッダー・フッター | `add_header_footer` `set_slide_notes_format` |
| ファイル操作 | `duplicate_presentation` `insert_slides_from` |

### PowerPoint デザインプリセット（39ツール）

| カテゴリ | ツール |
|----------|--------|
| テーマ・マスター | `apply_theme_colors` `set_master_font` `apply_consistent_branding` |
| 透かし | `add_watermark` |
| スライドデザイン | `create_title_slide_design` `create_content_slide_design` `create_section_divider` `create_closing_slide` |
| 装飾要素 | `add_decorative_element` `add_progress_bar` |
| 構造化スライド | `create_agenda_slide` `create_comparison_slide` `create_process_flow` `create_quote_slide` `create_team_slide` |
| スライド番号 | `add_slide_number_footer` |
| データ可視化 | `create_chart_slide` `create_dashboard_slide` `create_stat_highlight` |
| レイアウト | `create_two_column_slide` `create_image_text_slide` `create_three_column_slide` `create_icon_list_slide` |
| インフォグラフィック | `create_timeline_slide` `create_funnel_diagram` |
| ビジネステンプレート | `create_swot_slide` `create_roadmap_slide` `create_pricing_table` `create_testimonial_slide` |
| 比較・分析 | `create_before_after_slide` `create_problem_solution_slide` |
| ショーケース | `create_feature_showcase` `create_data_table_slide` `create_thank_you_slide` |
| ダイアグラム | `create_mind_map` `create_hierarchy_slide` `create_venn_diagram` |
| ダッシュボード | `create_metrics_grid` `create_workflow_slide` |

### Word 基本（119ツール）

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
| スタイル管理 | `create_style` `apply_style` `get_styles` `modify_style` |
| 脚注・文末脚注 | `add_footnote` `add_endnote` `get_footnotes` `get_endnotes` |
| 表の高度な操作 | `merge_table_cells` `split_table_cell` `format_table_cell` `set_table_cell_borders` `set_table_width` |
| 図形 | `add_shape` `add_line` |
| テキスト装飾 | `add_drop_cap` `set_text_highlight` `set_character_spacing` `add_text_effect` |
| ページレイアウト | `insert_page_break` `insert_column_break` `set_paragraph_borders` `set_paragraph_shading` `add_horizontal_line` |
| 文書プロパティ | `set_document_properties` `get_document_properties` `get_document_statistics` |
| 高度な機能 | `add_cross_reference` `insert_field` `set_tab_stops` `protect_document` `unprotect_document` |
| コンテンツコントロール | `add_content_control` `insert_building_block` |
| リスト高度 | `set_list_level` `restart_list_numbering` |
| テキスト装飾詳細 | `set_text_color_range` `set_text_size_range` `add_strikethrough` `set_underline_style` `add_small_caps` `add_all_caps` |
| 段落詳細 | `set_keep_with_next` `set_keep_together` `set_page_break_before` `set_widow_orphan_control` `set_outline_level` `get_paragraph_count` |
| 表詳細 | `set_table_row_height` `set_table_cell_width` `set_table_alignment` `add_table_row` `delete_table_row` `add_table_column` `delete_table_column` `get_table_data` `set_table_repeat_header` |
| セクション管理 | `get_section_count` `set_section_page_setup` `set_section_header` `set_section_footer` `link_section_header` |
| ナビゲーション | `go_to_page` `get_page_count` `insert_text_at_bookmark` `get_paragraph_text` `get_paragraph_range` |
| 検索詳細 | `find_text` `find_all` `highlight_found_text` |
| 差し込み印刷 | `start_mail_merge` `insert_merge_field` `execute_mail_merge` |
| 文書情報 | `get_word_count` `insert_date` `set_line_numbers` |
| 文書操作 | `compare_documents` `duplicate_document` `clear_all_formatting_word` |
| インデント | `set_paragraph_indentation` |
| 特殊文字 | `insert_special_character` |
| デフォルト設定 | `set_default_font` |
| 法務 | `add_table_of_authorities` |

### Word テンプレート（33ツール）

| カテゴリ | ツール |
|----------|--------|
| カバーページ | `create_cover_page` `add_cover_page_image` |
| テーマ | `setup_document_theme` |
| テンプレート | `create_meeting_minutes` `create_report_template` `create_letter` `create_invoice` `create_resume` `create_newsletter` `create_contract` |
| 書式一括設定 | `format_all_headings` |
| 装飾要素 | `add_sidebar` `add_callout_box` |
| 参照 | `create_table_of_figures` `insert_caption` |
| ビジネス文書 | `create_proposal` `create_sop` `create_project_charter` `create_meeting_agenda` |
| HR・管理 | `create_employee_handbook_section` `create_faq_document` `create_checklist_document` |
| 署名 | `add_signature_block` |
| レポート | `create_executive_summary` `create_status_report` |
| 申請・報告 | `create_change_request` `create_incident_report` |
| 教育 | `create_training_manual` `create_user_guide` |
| 規程・仕様 | `create_policy_document` `create_technical_specification` |
| 広報・事例 | `create_press_release` `create_case_study` |

### Excel 基本（140ツール）

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
| スパークライン | `add_sparkline` `format_sparkline` |
| 高度な条件付き書式 | `set_conditional_formatting_color_scale` `set_conditional_formatting_data_bar` `set_conditional_formatting_icon_set` |
| グループ化 | `group_rows` `ungroup_rows` `group_columns` `ungroup_columns` `set_outline_level` `collapse_group` |
| ハイパーリンク | `add_hyperlink` `add_internal_link` |
| 画像・図形 | `insert_image` `add_shape` `add_textbox` |
| 高度な書式 | `set_cell_style` `auto_fit_columns` `auto_fit_rows` `set_cell_indent` `set_text_rotation` `set_wrap_text` `set_cell_pattern` |
| データ操作 | `copy_range` `clear_range` `find_value` `replace_value` `remove_duplicates` `text_to_columns` `transpose_range` |
| グラフ高度 | `format_chart` `format_chart_series` `add_chart_trendline` `set_chart_area_format` |
| ピボット高度 | `format_pivot_table` `add_pivot_field` `refresh_pivot_table` |
| ブック管理 | `set_workbook_properties` `get_workbook_statistics` `set_tab_color` `hide_sheet` `unhide_sheet` `copy_sheet` `move_sheet` |
| 印刷高度 | `set_print_area` `set_print_titles` `add_page_break` |
| 保護高度 | `protect_workbook` `unprotect_workbook` `lock_cells` |
| 配列数式 | `set_array_formula` `evaluate_formula` `set_formula_range` `create_named_formula` |
| データ分析 | `calculate_statistics` `create_frequency_distribution` `goal_seek` `create_data_table_analysis` |
| 高度な書式2 | `set_rich_text_cell` `add_cell_dropdown` `set_cell_hyperlink_format` `clear_all_formatting` |
| チャート拡張 | `add_combo_chart` `add_pie_chart` `add_scatter_chart` `add_stock_chart` `add_radar_chart` |
| ブックナビ | `activate_sheet` `get_active_sheet` `get_used_range` `get_last_row` `get_last_column` |
| データ操作2 | `fill_series` `concatenate_range` `split_text_by_rows` `apply_formula_to_range` `create_sequence` |
| 条件操作 | `highlight_cells` `count_if` `sum_if` |
| 表示設定 | `set_gridlines_visible` `set_headings_visible` `set_zoom_level` `set_sheet_direction` |
| バリデーション拡張 | `add_number_validation` `add_date_validation` `add_text_length_validation` |
| 監査 | `trace_precedents` `trace_dependents` `check_errors` `get_cell_formula` |
| 小計・集計 | `insert_subtotals` `create_dropdown_list` `set_conditional_icon` |
| チャート装飾 | `add_error_bars` `set_chart_gradient` |
| スタイル | `create_named_style` `apply_alternating_colors` |
| データ取得 | `get_distinct_values` `vlookup` `create_summary_sheet` |

### Excel テンプレート（38ツール）

| カテゴリ | ツール |
|----------|--------|
| ダッシュボード | `create_dashboard_header` `create_kpi_cards` `add_data_summary` `create_heatmap` |
| テーブル | `create_data_table` `create_summary_row` `apply_table_theme` `create_comparison_table` |
| テンプレート | `create_input_form` `create_calendar` `create_gantt_chart` `create_scorecard` `create_financial_report` `create_invoice_template` `create_timesheet` `create_checklist` |
| ナビゲーション | `add_sheet_navigation` |
| 印刷準備 | `setup_print_ready` |
| 数値書式 | `format_as_currency` `format_as_percentage` |
| ビジネステンプレート | `create_budget_template` `create_project_tracker` `create_expense_report` `create_sales_report` |
| 管理テンプレート | `create_inventory_tracker` `create_employee_roster` `create_risk_matrix` `create_attendance_tracker` |
| ダッシュボード高度 | `create_kpi_dashboard` `create_vendor_comparison` |
| 財務 | `create_cash_flow_statement` |
| ライフスタイル | `create_workout_tracker` `create_meal_planner` |
| 計算・分析 | `create_loan_calculator` `create_grade_book` `create_survey_results` |
| カタログ | `create_price_list` `create_conversion_table` |

## 技術構成

```
src/microsoft_office/
├── server.py              # FastMCP サーバー（エントリポイント）
├── com_utils.py           # COM 共通ユーティリティ
├── office/                # COM オートメーションロジック（16,422行）
│   ├── powerpoint.py        # 3,768行 - 109関数
│   ├── powerpoint_advanced.py # 2,362行 - 15関数
│   ├── word.py              # 2,312行 - 110関数
│   ├── word_advanced.py     # 1,910行 - 15関数
│   ├── excel.py             # 4,006行 - 133関数
│   └── excel_advanced.py    # 2,064行 - 20関数
└── tools/                 # MCP ツール定義（6,474行）
    ├── powerpoint.py        # 1,871行 - 105ツール
    ├── powerpoint_advanced.py # 270行 - 15ツール
    ├── word.py              # 1,607行 - 109ツール
    ├── word_advanced.py     # 289行 - 15ツール
    ├── excel.py             # 2,105行 - 130ツール
    └── excel_advanced.py    # 332行 - 20ツール
```

- **総コード行数**: 34,386行+
- **総ツール数**: 484ツール
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
