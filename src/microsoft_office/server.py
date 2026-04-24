"""Microsoft Office MCP Server - PowerPoint, Word, Excel をCOMオートメーションで制御"""

from mcp.server.fastmcp import FastMCP

mcp = FastMCP(
    "Microsoft Office",
    instructions="""PowerPoint、Word、ExcelをCOMオートメーションでリアルタイム制御するMCPサーバー

## PowerPoint デザインガイドライン

### 推奨: pptx_* セーフAPIを使う

プレゼンテーション作成時は、デザイントークン制約付きの **pptx_*** ツールを優先すること。
これらは色・フォント・サイズをすべてテーマトークンで管理し、WCAG AAコントラスト比を自動検証する。

**利用可能なテーマ:** business, minimal, warm
**利用可能なツール:**
- `pptx_add_title_slide(title, subtitle, theme)` — タイトルスライド
- `pptx_add_section_header(title, accent, theme)` — セクション見出し
- `pptx_add_bullet_slide(title, bullets, theme)` — 箇条書き（最大5項目）
- `pptx_add_two_column_slide(title, left, right, theme)` — 2カラム
- `pptx_add_image_slide(title, image_path, caption, theme)` — 画像
- `pptx_add_comparison_slide(title, items, theme)` — 比較（2〜4項目）

### テーマ選択の指針
- **business**: 青系アクセント。提案書・営業・報告書に最適
- **minimal**: モノクロ + 黒アクセント。技術報告・データ重視の発表に
- **warm**: 暖色系アクセント。教育・研修・クリエイティブ提案に

### 設計原則
- 色はテーマトークンのみ使用（RGB直指定禁止）
- グラデーション・ドロップシャドウ・3D効果・WordArtは使用しない
- 1スライドのアクセント色使用は1要素まで
- 箇条書きは5項目以内、超える場合はスライドを分割する
- フォントサイズは11pt〜36ptの範囲内で使用する
""",
)


def main():
    # Import tool modules to register @mcp.tool() decorators
    from microsoft_office.tools import powerpoint, word, excel, word_advanced, powerpoint_advanced, excel_advanced, powerpoint_safe  # noqa: F401

    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
