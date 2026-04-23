"""Word advanced design & template MCP tool definitions."""

from microsoft_office.server import mcp
from microsoft_office.office import word_advanced


@mcp.tool()
def word_create_cover_page(
    title: str,
    subtitle: str | None = None,
    author: str | None = None,
    date: str | None = None,
    company: str | None = None,
    style: str = "modern",
) -> str:
    """プロフェッショナルな表紙ページを作成します。style: modern(モダン), executive(エグゼクティブ), creative(クリエイティブ), minimal(ミニマル), academic(学術)。"""
    try:
        result = word_advanced.create_cover_page(
            title, subtitle=subtitle, author=author,
            date=date, company=company, style=style,
        )
        return f"表紙を作成しました（スタイル: {result['style']}、段落数: {result['paragraphs']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_setup_document_theme(theme_name: str) -> str:
    """文書全体にテーマを適用します。フォント・色・見出しスタイル・行間を一括設定。theme_name: corporate, elegant, modern, academic, creative。"""
    try:
        result = word_advanced.setup_document_theme(theme_name)
        return (
            f"テーマ '{result['theme']}' を適用しました"
            f"（見出しフォント: {result['heading_font']}, 本文フォント: {result['body_font']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_meeting_minutes(
    title: str,
    date: str,
    attendees: list[str],
    agenda_items: list[str],
    action_items: list[dict] | None = None,
    notes: str | None = None,
) -> str:
    """議事録を自動生成します。出席者・議題・アクションアイテムを含むプロフェッショナルな書式で作成。action_items: [{"owner": "名前", "task": "内容", "due": "期日"}]。"""
    try:
        result = word_advanced.create_meeting_minutes(
            title, date, attendees, agenda_items,
            action_items=action_items, notes=notes,
        )
        return (
            f"議事録を作成しました（出席者: {result['attendee_count']}名、"
            f"議題: {result['agenda_count']}件、"
            f"アクションアイテム: {result['action_item_count']}件）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_report_template(
    title: str,
    author: str | None = None,
    sections: list[str] | None = None,
    style: str = "business",
) -> str:
    """レポートテンプレートを作成します。表紙・目次・セクション見出し・ページ番号を含む構造化文書。style: business, technical, executive_summary。"""
    try:
        result = word_advanced.create_report_template(
            title, author=author, sections=sections, style=style,
        )
        return (
            f"レポートテンプレートを作成しました（スタイル: {result['style']}、"
            f"セクション数: {result['section_count']}、セクション: {', '.join(result['sections'])}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_letter(
    recipient_name: str,
    recipient_address: str,
    subject: str,
    body_paragraphs: list[str],
    sender_name: str,
    sender_title: str | None = None,
    company: str | None = None,
    style: str = "formal",
) -> str:
    """ビジネスレターを作成します。style: formal(ブロック形式), semi_formal(修正ブロック), modern(モダンレイアウト)。"""
    try:
        result = word_advanced.create_letter(
            recipient_name, recipient_address, subject,
            body_paragraphs, sender_name,
            sender_title=sender_title, company=company, style=style,
        )
        return (
            f"ビジネスレターを作成しました（宛先: {result['recipient']}、"
            f"件名: {result['subject']}、スタイル: {result['style']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_invoice(
    company_name: str,
    client_name: str,
    items: list[dict],
    invoice_number: str | None = None,
    date: str | None = None,
    due_date: str | None = None,
    notes: str | None = None,
    tax_rate: float | None = None,
) -> str:
    """プロフェッショナルな請求書を作成します。items: [{"description": "説明", "quantity": 数量, "unit_price": 単価}]。tax_rateはパーセント（例: 10）。"""
    try:
        result = word_advanced.create_invoice(
            company_name, client_name, items,
            invoice_number=invoice_number, date=date,
            due_date=due_date, notes=notes, tax_rate=tax_rate,
        )
        return (
            f"請求書を作成しました（請求番号: {result['invoice_number']}、"
            f"品目数: {result['item_count']}、"
            f"小計: ${result['subtotal']:,.2f}、合計: ${result['total']:,.2f}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_resume(
    name: str,
    contact_info: dict,
    sections: list[dict],
    style: str = "modern",
) -> str:
    """プロフェッショナルな履歴書を作成します。contact_info: {"email": "", "phone": "", "address": "", "linkedin": ""}。sections: [{"title": "Experience", "items": [{"title": "役職", "subtitle": "会社名", "date": "2020-2024", "details": ["実績1"]}]}]。style: modern, classic, creative。"""
    try:
        result = word_advanced.create_resume(
            name, contact_info, sections, style=style,
        )
        return (
            f"履歴書を作成しました（名前: {result['name']}、"
            f"スタイル: {result['style']}、セクション数: {result['section_count']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_table_of_figures(label: str = "Figure") -> str:
    """図表目次を挿入します。labelで対象ラベルを指定（例: Figure, Table）。"""
    try:
        result = word_advanced.create_table_of_figures(label=label)
        return f"図表目次を挿入しました（ラベル: {result['label']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_cover_page_image(
    image_path: str,
    style: str = "full_bleed",
) -> str:
    """表紙に画像を追加します。style: full_bleed(全面), centered(中央配置), banner_top(上部バナー), banner_bottom(下部バナー)。"""
    try:
        result = word_advanced.add_cover_page_image(image_path, style=style)
        return f"表紙画像を追加しました（スタイル: {result['style']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_format_all_headings(heading_styles: dict) -> str:
    """文書内の全見出しの書式を一括設定します。heading_styles: {"h1": {"font_size": 24, "color": [0,0,128], "bold": true, "font_name": "Calibri"}, "h2": {...}, "h3": {...}}。"""
    try:
        # Convert color lists to tuples
        converted = {}
        for key, val in heading_styles.items():
            converted[key] = dict(val)
            if "color" in converted[key] and isinstance(converted[key]["color"], list):
                converted[key]["color"] = tuple(converted[key]["color"])
        result = word_advanced.format_all_headings(converted)
        updated_keys = [k for k, v in result["updated"].items() if v is True]
        return f"見出し書式を更新しました: {', '.join(updated_keys)}"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_sidebar(
    text: str,
    position: str = "left",
    width: float = 120,
    fill_color: list[int] | None = None,
    text_color: list[int] | None = None,
    font_size: float = 9,
) -> str:
    """サイドバーを追加します（プルクォート・重要事項用）。position: left/right。fill_color/text_colorは[R,G,B]形式。"""
    try:
        result = word_advanced.add_sidebar(
            text, position=position, width=width,
            fill_color=tuple(fill_color) if fill_color else None,
            text_color=tuple(text_color) if text_color else None,
            font_size=font_size,
        )
        return f"サイドバーを追加しました（位置: {result['position']}、幅: {result['width']}pt）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_newsletter(
    title: str,
    subtitle: str | None = None,
    columns: int = 2,
    articles: list[dict] | None = None,
) -> str:
    """ニュースレターを作成します。マルチカラムレイアウト。articles: [{"title": "見出し", "body": "本文", "author": "著者"}]。"""
    try:
        result = word_advanced.create_newsletter(
            title, subtitle=subtitle, columns=columns, articles=articles,
        )
        return (
            f"ニュースレターを作成しました（タイトル: {result['title']}、"
            f"カラム数: {result['columns']}、記事数: {result['article_count']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_callout_box(
    text: str,
    style: str = "info",
    paragraph_index: int | None = None,
) -> str:
    """コールアウトボックス（注意書き）を追加します。style: info(青), warning(黄), success(緑), error(赤), tip(紫)。paragraph_indexで挿入位置を指定可能。"""
    try:
        result = word_advanced.add_callout_box(
            text, style=style, paragraph_index=paragraph_index,
        )
        return f"コールアウトボックスを追加しました（スタイル: {result['style']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_contract(
    title: str,
    parties: list[str],
    clauses: list[dict],
    date: str | None = None,
    style: str = "standard",
) -> str:
    """契約書テンプレートを作成します。番号付き条項・署名ブロック付き。clauses: [{"title": "条項名", "content": "条項内容"}]。"""
    try:
        result = word_advanced.create_contract(
            title, parties, clauses, date=date, style=style,
        )
        return (
            f"契約書を作成しました（タイトル: {result['title']}、"
            f"当事者: {result['party_count']}名、条項数: {result['clause_count']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_insert_caption(
    text: str,
    label: str = "Figure",
    position: str = "below",
) -> str:
    """図表のキャプションを自動番号付きで挿入します。label: Figure/Table。position: below/above。"""
    try:
        result = word_advanced.insert_caption(
            text, label=label, position=position,
        )
        return f"キャプションを挿入しました（{result['label']}: {result['text']}、位置: {result['position']}）"
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_proposal(
    title: str,
    client_name: str,
    sections: list[dict],
    author: str | None = None,
    date: str | None = None,
    style: str = "professional",
) -> str:
    """ビジネス提案書を作成します。表紙・目次・セクション構成のプロフェッショナルな提案書。sections: [{"title": "セクション名", "content": "内容"}]。style: professional(青系), creative(大胆デザイン), minimal(ミニマル)。"""
    try:
        result = word_advanced.create_proposal(
            title, client_name, sections,
            author=author, date=date, style=style,
        )
        return (
            f"提案書を作成しました（タイトル: {result['title']}、"
            f"クライアント: {result['client']}、スタイル: {result['style']}、"
            f"セクション数: {result['section_count']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_sop(
    title: str,
    purpose: str,
    scope: str,
    procedures: list[dict],
    responsibilities: list[dict] | None = None,
    style: str = "standard",
) -> str:
    """標準作業手順書（SOP）を作成します。目的・範囲・手順・責任を含む構造化文書。procedures: [{"step": 1, "title": "手順名", "details": "詳細", "caution": "注意事項"}]。responsibilities: [{"role": "役割", "responsibility": "責任内容"}]。"""
    try:
        result = word_advanced.create_sop(
            title, purpose, scope, procedures,
            responsibilities=responsibilities, style=style,
        )
        return (
            f"SOPを作成しました（タイトル: {result['title']}、"
            f"手順数: {result['procedure_count']}、"
            f"責任者数: {result['responsibility_count']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_project_charter(
    project_name: str,
    sponsor: str,
    manager: str,
    objectives: list[str],
    scope: str,
    milestones: list[dict],
    budget: str | None = None,
) -> str:
    """プロジェクト憲章を作成します。目標・範囲・マイルストーン・承認欄を含む標準フォーマット。milestones: [{"name": "マイルストーン名", "date": "日付", "deliverable": "成果物"}]。"""
    try:
        result = word_advanced.create_project_charter(
            project_name, sponsor, manager, objectives,
            scope, milestones, budget=budget,
        )
        return (
            f"プロジェクト憲章を作成しました（プロジェクト: {result['project_name']}、"
            f"目標数: {result['objective_count']}、"
            f"マイルストーン数: {result['milestone_count']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_meeting_agenda(
    title: str,
    date: str,
    time: str,
    location: str,
    attendees: list[str],
    agenda_items: list[dict],
    notes: str | None = None,
) -> str:
    """会議アジェンダを作成します。日時・場所・出席者・議題・アクションアイテム表を含むプロフェッショナルな書式。agenda_items: [{"topic": "議題", "presenter": "発表者", "duration": "10min"}]。"""
    try:
        result = word_advanced.create_meeting_agenda(
            title, date, time, location, attendees,
            agenda_items, notes=notes,
        )
        return (
            f"会議アジェンダを作成しました（タイトル: {result['title']}、"
            f"出席者: {result['attendee_count']}名、"
            f"議題数: {result['agenda_item_count']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_employee_handbook_section(
    title: str,
    policies: list[dict],
    effective_date: str | None = None,
) -> str:
    """従業員ハンドブックのセクションを作成します。ポリシー・確認署名欄付き。policies: [{"title": "ポリシー名", "content": "内容", "important": false}]。importantがtrueの場合、強調表示されます。"""
    try:
        result = word_advanced.create_employee_handbook_section(
            title, policies, effective_date=effective_date,
        )
        return (
            f"ハンドブックセクションを作成しました（タイトル: {result['title']}、"
            f"ポリシー数: {result['policy_count']}、"
            f"施行日: {result['effective_date']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_faq_document(
    title: str,
    faqs: list[dict],
    style: str = "accordion",
) -> str:
    """FAQドキュメントを作成します。faqs: [{"question": "質問", "answer": "回答"}]。style: accordion(Q&Aブロック), numbered(番号付きリスト), table(表形式)。"""
    try:
        result = word_advanced.create_faq_document(
            title, faqs, style=style,
        )
        return (
            f"FAQドキュメントを作成しました（タイトル: {result['title']}、"
            f"FAQ数: {result['faq_count']}、スタイル: {result['style']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_create_checklist_document(
    title: str,
    categories: list[dict],
    style: str = "checkbox",
) -> str:
    """チェックリストドキュメントを作成します。categories: [{"name": "カテゴリ名", "items": ["項目1", "項目2"]}]。style: checkbox(チェックボックス付き), numbered(番号付き), bullet(箇条書き)。"""
    try:
        result = word_advanced.create_checklist_document(
            title, categories, style=style,
        )
        return (
            f"チェックリストを作成しました（タイトル: {result['title']}、"
            f"カテゴリ数: {result['category_count']}、"
            f"総項目数: {result['total_items']}、スタイル: {result['style']}）"
        )
    except Exception as e:
        return f"エラー: {e}"


@mcp.tool()
def word_add_signature_block(
    names: list[str],
    titles: list[str] | None = None,
    date_line: bool = True,
    witness: bool = False,
) -> str:
    """プロフェッショナルな署名ブロックを追加します。署名線・名前・肩書・日付欄を含む。witnessがtrueの場合、証人署名欄も追加。"""
    try:
        result = word_advanced.add_signature_block(
            names, titles=titles, date_line=date_line, witness=witness,
        )
        return (
            f"署名ブロックを追加しました（署名者数: {result['signatory_count']}、"
            f"肩書: {'あり' if result['has_titles'] else 'なし'}、"
            f"日付欄: {'あり' if result['date_line'] else 'なし'}、"
            f"証人: {'あり' if result['witness'] else 'なし'}）"
        )
    except Exception as e:
        return f"エラー: {e}"
