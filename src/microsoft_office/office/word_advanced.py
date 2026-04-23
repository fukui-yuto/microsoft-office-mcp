"""Word advanced document design and template automation.

Provides high-level functions for creating professional document designs,
templates, and complex layouts using Word COM automation.
"""

import datetime

from microsoft_office.com_utils import ensure_absolute_path, get_or_create_app, rgb

# ---------------------------------------------------------------------------
# Word enumeration constants
# ---------------------------------------------------------------------------

WD_STYLE_NORMAL = -1
WD_STYLE_HEADING1 = -2
WD_STYLE_HEADING2 = -3
WD_STYLE_HEADING3 = -4
WD_STYLE_HEADING4 = -67

WD_ALIGN_LEFT = 0
WD_ALIGN_CENTER = 1
WD_ALIGN_RIGHT = 2
WD_ALIGN_JUSTIFY = 3

WD_COLLAPSE_START = 1
WD_COLLAPSE_END = 0

WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_1PT5 = 1
WD_LINE_SPACE_DOUBLE = 2
WD_LINE_SPACE_EXACTLY = 4
WD_LINE_SPACE_MULTIPLE = 5

WD_SECTION_BREAK_NEXT_PAGE = 2
WD_SECTION_BREAK_CONTINUOUS = 3
WD_PAGE_BREAK = 7

WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1

WD_BORDER_TOP = -1
WD_BORDER_LEFT = -2
WD_BORDER_BOTTOM = -3
WD_BORDER_RIGHT = -4

WD_CAPTION_POSITION_BELOW = 1
WD_CAPTION_POSITION_ABOVE = 0

# Shape constants
MSO_AUTO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5

# Table of figures field codes
WD_FIELD_TOC = 13

_HEADING_STYLES = {1: WD_STYLE_HEADING1, 2: WD_STYLE_HEADING2, 3: WD_STYLE_HEADING3, 4: WD_STYLE_HEADING4}


# ---------------------------------------------------------------------------
# Theme color palettes
# ---------------------------------------------------------------------------

THEMES = {
    "corporate": {
        "primary": (0, 70, 127),
        "secondary": (0, 112, 192),
        "accent": (0, 176, 240),
        "text": (51, 51, 51),
        "heading_font": "Calibri",
        "body_font": "Calibri",
        "h1_size": 26,
        "h2_size": 18,
        "h3_size": 14,
        "body_size": 11,
        "line_spacing": 1.15,
    },
    "elegant": {
        "primary": (88, 44, 77),
        "secondary": (128, 64, 96),
        "accent": (192, 160, 128),
        "text": (64, 64, 64),
        "heading_font": "Georgia",
        "body_font": "Garamond",
        "h1_size": 28,
        "h2_size": 20,
        "h3_size": 14,
        "body_size": 12,
        "line_spacing": 1.3,
    },
    "modern": {
        "primary": (41, 65, 122),
        "secondary": (68, 114, 196),
        "accent": (237, 125, 49),
        "text": (38, 38, 38),
        "heading_font": "Segoe UI",
        "body_font": "Segoe UI",
        "h1_size": 28,
        "h2_size": 18,
        "h3_size": 13,
        "body_size": 10.5,
        "line_spacing": 1.2,
    },
    "academic": {
        "primary": (0, 0, 0),
        "secondary": (64, 64, 64),
        "accent": (0, 0, 128),
        "text": (0, 0, 0),
        "heading_font": "Times New Roman",
        "body_font": "Times New Roman",
        "h1_size": 16,
        "h2_size": 14,
        "h3_size": 13,
        "body_size": 12,
        "line_spacing": 2.0,
    },
    "creative": {
        "primary": (255, 87, 34),
        "secondary": (33, 150, 243),
        "accent": (76, 175, 80),
        "text": (55, 55, 55),
        "heading_font": "Century Gothic",
        "body_font": "Calibri",
        "h1_size": 30,
        "h2_size": 20,
        "h3_size": 14,
        "body_size": 11,
        "line_spacing": 1.2,
    },
}

COVER_STYLES = {
    "modern": {
        "primary": (41, 65, 122),
        "accent": (237, 125, 49),
        "title_font": "Segoe UI Light",
        "title_size": 44,
        "subtitle_size": 20,
        "meta_size": 12,
    },
    "executive": {
        "primary": (0, 51, 102),
        "accent": (153, 153, 153),
        "title_font": "Garamond",
        "title_size": 36,
        "subtitle_size": 18,
        "meta_size": 12,
    },
    "creative": {
        "primary": (255, 87, 34),
        "accent": (33, 150, 243),
        "title_font": "Century Gothic",
        "title_size": 48,
        "subtitle_size": 22,
        "meta_size": 11,
    },
    "minimal": {
        "primary": (51, 51, 51),
        "accent": (180, 180, 180),
        "title_font": "Calibri Light",
        "title_size": 40,
        "subtitle_size": 18,
        "meta_size": 11,
    },
    "academic": {
        "primary": (0, 0, 0),
        "accent": (0, 0, 128),
        "title_font": "Times New Roman",
        "title_size": 24,
        "subtitle_size": 16,
        "meta_size": 12,
    },
}

CALLOUT_STYLES = {
    "info": {"border_color": (0, 112, 192), "bg_color": (222, 235, 247), "icon": "ℹ"},
    "warning": {"border_color": (255, 185, 0), "bg_color": (255, 243, 205), "icon": "⚠"},
    "success": {"border_color": (76, 175, 80), "bg_color": (232, 245, 233), "icon": "✓"},
    "error": {"border_color": (211, 47, 47), "bg_color": (255, 235, 238), "icon": "✗"},
    "tip": {"border_color": (123, 31, 162), "bg_color": (243, 229, 245), "icon": "💡"},
}


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_app():
    """Get or create the Word application instance."""
    return get_or_create_app("Word.Application")


def _get_active_doc():
    """Get the active Word document."""
    app = _get_app()
    return app.ActiveDocument


def _collapse_to_end(doc):
    """Get a range collapsed to the end of the document."""
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    return rng


def _add_paragraph(doc, text, style=None, alignment=None, font_name=None,
                   font_size=None, font_color=None, bold=None, italic=None,
                   space_before=None, space_after=None, line_spacing=None):
    """Add a paragraph with full formatting options."""
    rng = _collapse_to_end(doc)
    rng.InsertAfter(text + "\r")
    para = doc.Paragraphs(doc.Paragraphs.Count)

    if style is not None:
        try:
            para.Style = style
        except Exception:
            pass

    if alignment is not None:
        para.Alignment = alignment

    fmt = para.Range.Font
    if font_name:
        fmt.Name = font_name
    if font_size:
        fmt.Size = font_size
    if font_color:
        fmt.Color = rgb(*font_color)
    if bold is not None:
        fmt.Bold = bold
    if italic is not None:
        fmt.Italic = italic

    pf = para.Format
    if space_before is not None:
        pf.SpaceBefore = space_before
    if space_after is not None:
        pf.SpaceAfter = space_after
    if line_spacing is not None:
        pf.LineSpacingRule = WD_LINE_SPACE_MULTIPLE
        pf.LineSpacing = line_spacing * 12  # points

    return para


def _add_empty_lines(doc, count=1):
    """Add empty paragraphs."""
    for _ in range(count):
        _add_paragraph(doc, "", font_size=6, space_before=0, space_after=0)


def _insert_page_break(doc):
    """Insert a page break at the end of the document."""
    rng = _collapse_to_end(doc)
    rng.InsertBreak(WD_PAGE_BREAK)


def _insert_section_break(doc, break_type=WD_SECTION_BREAK_NEXT_PAGE):
    """Insert a section break at the end of the document."""
    rng = _collapse_to_end(doc)
    rng.InsertBreak(break_type)


def _add_horizontal_line(doc, color=(180, 180, 180), thickness=1.0):
    """Add a horizontal line using paragraph border."""
    para = _add_paragraph(doc, "", space_before=6, space_after=6)
    border = para.Format.Borders(WD_BORDER_BOTTOM)
    border.LineStyle = 1  # wdLineStyleSingle
    border.LineWidth = thickness * 8  # approximate conversion
    border.Color = rgb(*color)
    return para


def _create_table(doc, rows, cols, data=None):
    """Create a table at the end of the document."""
    rng = _collapse_to_end(doc)
    rng.InsertAfter("\r")
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    table = doc.Tables.Add(rng, rows, cols)
    table.Borders.Enable = True

    if data:
        for r_idx, row_data in enumerate(data):
            for c_idx, cell_value in enumerate(row_data):
                if r_idx < rows and c_idx < cols:
                    table.Cell(r_idx + 1, c_idx + 1).Range.Text = str(cell_value)

    return table


def _format_table_header(table, fill_color=(0, 70, 127), text_color=(255, 255, 255),
                         font_name="Calibri", font_size=11, bold=True):
    """Format the header row of a table."""
    for c in range(1, table.Columns.Count + 1):
        cell = table.Cell(1, c)
        cell.Shading.BackgroundPatternColor = rgb(*fill_color)
        rng = cell.Range
        rng.Font.Color = rgb(*text_color)
        rng.Font.Bold = bold
        if font_name:
            rng.Font.Name = font_name
        if font_size:
            rng.Font.Size = font_size


def _format_table_body(table, font_name="Calibri", font_size=10.5,
                       alt_row_color=None):
    """Format the body rows of a table."""
    for r in range(2, table.Rows.Count + 1):
        for c in range(1, table.Columns.Count + 1):
            cell = table.Cell(r, c)
            rng = cell.Range
            if font_name:
                rng.Font.Name = font_name
            if font_size:
                rng.Font.Size = font_size
            if alt_row_color and r % 2 == 0:
                cell.Shading.BackgroundPatternColor = rgb(*alt_row_color)


def _set_cell_text(table, row, col, text, bold=False, alignment=None,
                   font_name=None, font_size=None, font_color=None):
    """Set text and formatting for a table cell."""
    cell = table.Cell(row, col)
    cell.Range.Text = str(text)
    if bold:
        cell.Range.Font.Bold = True
    if alignment is not None:
        cell.Range.ParagraphFormat.Alignment = alignment
    if font_name:
        cell.Range.Font.Name = font_name
    if font_size:
        cell.Range.Font.Size = font_size
    if font_color:
        cell.Range.Font.Color = rgb(*font_color)


# ---------------------------------------------------------------------------
# 1. Cover Page
# ---------------------------------------------------------------------------

def create_cover_page(title: str, subtitle: str | None = None,
                      author: str | None = None, date: str | None = None,
                      company: str | None = None, style: str = "modern") -> dict:
    """Create a professional cover page on the active document.

    Styles: modern, executive, creative, minimal, academic.
    """
    doc = _get_active_doc()
    cs = COVER_STYLES.get(style, COVER_STYLES["modern"])
    primary = cs["primary"]
    accent = cs["accent"]

    if style == "modern":
        # Clean design with colored accent bar on the left
        _add_empty_lines(doc, 6)

        # Accent bar using a shape
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_START)
        shape = doc.Shapes.AddShape(
            MSO_AUTO_SHAPE_RECTANGLE,
            Left=36, Top=36,
            Width=8, Height=700,
        )
        shape.Fill.ForeColor.RGB = rgb(*primary)
        shape.Line.Visible = False

        # Title
        _add_paragraph(doc, title,
                       font_name=cs["title_font"], font_size=cs["title_size"],
                       font_color=primary, bold=False,
                       space_before=0, space_after=6, alignment=WD_ALIGN_LEFT)

        # Accent line under title
        _add_horizontal_line(doc, color=accent, thickness=2.0)

        if subtitle:
            _add_paragraph(doc, subtitle,
                           font_name="Segoe UI", font_size=cs["subtitle_size"],
                           font_color=(100, 100, 100), bold=False,
                           space_before=12, space_after=6, alignment=WD_ALIGN_LEFT)

        _add_empty_lines(doc, 4)

        # Meta information
        meta_items = []
        if author:
            meta_items.append(author)
        if company:
            meta_items.append(company)
        if date:
            meta_items.append(date)
        else:
            meta_items.append(datetime.date.today().strftime("%Y-%m-%d"))

        for item in meta_items:
            _add_paragraph(doc, item,
                           font_name="Segoe UI", font_size=cs["meta_size"],
                           font_color=(120, 120, 120),
                           space_before=2, space_after=2, alignment=WD_ALIGN_LEFT)

    elif style == "executive":
        # Formal with double border
        _add_empty_lines(doc, 3)

        # Top border line
        _add_horizontal_line(doc, color=primary, thickness=2.5)
        _add_horizontal_line(doc, color=accent, thickness=0.5)

        _add_empty_lines(doc, 4)

        _add_paragraph(doc, title,
                       font_name=cs["title_font"], font_size=cs["title_size"],
                       font_color=primary, bold=True,
                       space_before=0, space_after=12, alignment=WD_ALIGN_CENTER)

        if subtitle:
            _add_paragraph(doc, subtitle,
                           font_name="Garamond", font_size=cs["subtitle_size"],
                           font_color=(80, 80, 80), italic=True,
                           space_before=6, space_after=12, alignment=WD_ALIGN_CENTER)

        _add_empty_lines(doc, 6)

        # Bottom border
        _add_horizontal_line(doc, color=accent, thickness=0.5)
        _add_horizontal_line(doc, color=primary, thickness=2.5)

        _add_empty_lines(doc, 2)

        meta_parts = []
        if author:
            meta_parts.append(f"Prepared by: {author}")
        if company:
            meta_parts.append(company)
        if date:
            meta_parts.append(date)
        else:
            meta_parts.append(datetime.date.today().strftime("%B %d, %Y"))

        for part in meta_parts:
            _add_paragraph(doc, part,
                           font_name="Garamond", font_size=cs["meta_size"],
                           font_color=(100, 100, 100),
                           space_before=2, space_after=2, alignment=WD_ALIGN_CENTER)

    elif style == "creative":
        # Bold colors + geometric elements
        # Add a large colored rectangle at top
        shape = doc.Shapes.AddShape(
            MSO_AUTO_SHAPE_RECTANGLE,
            Left=0, Top=0,
            Width=620, Height=200,
        )
        shape.Fill.ForeColor.RGB = rgb(*primary)
        shape.Line.Visible = False

        # Add accent shape
        shape2 = doc.Shapes.AddShape(
            MSO_AUTO_SHAPE_RECTANGLE,
            Left=0, Top=200,
            Width=620, Height=8,
        )
        shape2.Fill.ForeColor.RGB = rgb(*accent)
        shape2.Line.Visible = False

        _add_empty_lines(doc, 10)

        _add_paragraph(doc, title,
                       font_name=cs["title_font"], font_size=cs["title_size"],
                       font_color=primary, bold=True,
                       space_before=24, space_after=12, alignment=WD_ALIGN_LEFT)

        if subtitle:
            _add_paragraph(doc, subtitle,
                           font_name="Century Gothic", font_size=cs["subtitle_size"],
                           font_color=accent, bold=False,
                           space_before=6, space_after=12, alignment=WD_ALIGN_LEFT)

        _add_empty_lines(doc, 6)

        meta_items = []
        if author:
            meta_items.append(author)
        if company:
            meta_items.append(company)
        if date:
            meta_items.append(date)
        else:
            meta_items.append(datetime.date.today().strftime("%Y/%m/%d"))

        for item in meta_items:
            _add_paragraph(doc, item,
                           font_name="Century Gothic", font_size=cs["meta_size"],
                           font_color=(100, 100, 100),
                           space_before=2, space_after=2, alignment=WD_ALIGN_LEFT)

    elif style == "minimal":
        # Lots of whitespace, very clean
        _add_empty_lines(doc, 12)

        _add_paragraph(doc, title,
                       font_name=cs["title_font"], font_size=cs["title_size"],
                       font_color=primary, bold=False,
                       space_before=0, space_after=18, alignment=WD_ALIGN_CENTER)

        if subtitle:
            _add_paragraph(doc, subtitle,
                           font_name="Calibri Light", font_size=cs["subtitle_size"],
                           font_color=(140, 140, 140),
                           space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)

        _add_empty_lines(doc, 10)

        # Single thin line
        _add_horizontal_line(doc, color=accent, thickness=0.5)

        _add_empty_lines(doc, 1)

        meta_parts = []
        if author:
            meta_parts.append(author)
        if company:
            meta_parts.append(company)
        if date:
            meta_parts.append(date)
        else:
            meta_parts.append(datetime.date.today().strftime("%Y-%m-%d"))

        _add_paragraph(doc, "  |  ".join(meta_parts),
                       font_name="Calibri Light", font_size=cs["meta_size"],
                       font_color=(160, 160, 160),
                       space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)

    elif style == "academic":
        # Traditional academic title page
        _add_empty_lines(doc, 6)

        _add_paragraph(doc, title.upper(),
                       font_name=cs["title_font"], font_size=cs["title_size"],
                       font_color=primary, bold=True,
                       space_before=0, space_after=24,
                       alignment=WD_ALIGN_CENTER, line_spacing=2.0)

        if subtitle:
            _add_paragraph(doc, subtitle,
                           font_name="Times New Roman", font_size=cs["subtitle_size"],
                           font_color=(0, 0, 0), italic=True,
                           space_before=12, space_after=12, alignment=WD_ALIGN_CENTER)

        _add_empty_lines(doc, 4)

        if author:
            _add_paragraph(doc, "by",
                           font_name="Times New Roman", font_size=12,
                           font_color=(0, 0, 0),
                           space_before=12, space_after=6, alignment=WD_ALIGN_CENTER)
            _add_paragraph(doc, author,
                           font_name="Times New Roman", font_size=14,
                           font_color=(0, 0, 0), bold=True,
                           space_before=6, space_after=12, alignment=WD_ALIGN_CENTER)

        _add_empty_lines(doc, 4)

        if company:
            _add_paragraph(doc, company,
                           font_name="Times New Roman", font_size=12,
                           font_color=(0, 0, 0),
                           space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)

        date_str = date or datetime.date.today().strftime("%B %d, %Y")
        _add_paragraph(doc, date_str,
                       font_name="Times New Roman", font_size=12,
                       font_color=(0, 0, 0),
                       space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)

    return {"title": title, "style": style, "paragraphs": doc.Paragraphs.Count}


# ---------------------------------------------------------------------------
# 2. Document Theme
# ---------------------------------------------------------------------------

def setup_document_theme(theme_name: str) -> dict:
    """Apply a complete document theme with fonts, colors, heading styles, and spacing.

    Themes: corporate, elegant, modern, academic, creative.
    """
    doc = _get_active_doc()
    theme = THEMES.get(theme_name, THEMES["corporate"])

    # Configure heading styles
    for level, style_id in _HEADING_STYLES.items():
        try:
            style = doc.Styles(style_id)
            style.Font.Name = theme["heading_font"]
            style.Font.Color = rgb(*theme["primary"])
            style.Font.Bold = True

            size_key = f"h{level}_size"
            if size_key in theme:
                style.Font.Size = theme[size_key]

            # Heading spacing
            style.ParagraphFormat.SpaceBefore = 12 if level > 1 else 18
            style.ParagraphFormat.SpaceAfter = 6
        except Exception:
            pass

    # Configure Normal style
    try:
        normal = doc.Styles(WD_STYLE_NORMAL)
        normal.Font.Name = theme["body_font"]
        normal.Font.Size = theme["body_size"]
        normal.Font.Color = rgb(*theme["text"])
        normal.ParagraphFormat.LineSpacingRule = WD_LINE_SPACE_MULTIPLE
        normal.ParagraphFormat.LineSpacing = theme["line_spacing"] * 12
        normal.ParagraphFormat.SpaceAfter = 8
        normal.ParagraphFormat.SpaceBefore = 0
    except Exception:
        pass

    # Set default document font
    try:
        doc.Content.Font.Name = theme["body_font"]
        doc.Content.Font.Size = theme["body_size"]
        doc.Content.Font.Color = rgb(*theme["text"])
    except Exception:
        pass

    return {
        "theme": theme_name,
        "heading_font": theme["heading_font"],
        "body_font": theme["body_font"],
        "primary_color": list(theme["primary"]),
    }


# ---------------------------------------------------------------------------
# 3. Meeting Minutes
# ---------------------------------------------------------------------------

def create_meeting_minutes(title: str, date: str, attendees: list[str],
                           agenda_items: list[str],
                           action_items: list[dict] | None = None,
                           notes: str | None = None) -> dict:
    """Generate a complete meeting minutes document.

    action_items: list of {"owner": "Name", "task": "Description", "due": "Date"}
    """
    doc = _get_active_doc()
    primary = (0, 70, 127)
    accent = (0, 112, 192)

    # Title
    _add_paragraph(doc, "MEETING MINUTES",
                   font_name="Calibri", font_size=24,
                   font_color=primary, bold=True,
                   space_before=0, space_after=6, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # Meeting title
    _add_paragraph(doc, title,
                   font_name="Calibri", font_size=16,
                   font_color=(51, 51, 51), bold=True,
                   space_before=12, space_after=12, alignment=WD_ALIGN_LEFT)

    # Info table
    info_data = [
        ["Date", date],
        ["Attendees", ", ".join(attendees)],
    ]
    info_table = _create_table(doc, len(info_data), 2, info_data)
    info_table.Borders.Enable = True

    # Style info table
    for r in range(1, info_table.Rows.Count + 1):
        cell1 = info_table.Cell(r, 1)
        cell1.Range.Font.Bold = True
        cell1.Range.Font.Name = "Calibri"
        cell1.Range.Font.Size = 10.5
        cell1.Range.Font.Color = rgb(*primary)
        cell1.Shading.BackgroundPatternColor = rgb(240, 245, 250)
        cell2 = info_table.Cell(r, 2)
        cell2.Range.Font.Name = "Calibri"
        cell2.Range.Font.Size = 10.5

    try:
        info_table.Columns(1).Width = 100
        info_table.Columns(2).Width = 360
    except Exception:
        pass

    # Agenda section
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Agenda",
                   font_name="Calibri", font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=0.5)

    for i, item in enumerate(agenda_items, 1):
        _add_paragraph(doc, f"{i}. {item}",
                       font_name="Calibri", font_size=11,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=4, alignment=WD_ALIGN_LEFT)

    # Notes section
    if notes:
        _add_paragraph(doc, "Notes",
                       font_name="Calibri", font_size=14,
                       font_color=primary, bold=True,
                       space_before=18, space_after=6, alignment=WD_ALIGN_LEFT)
        _add_horizontal_line(doc, color=accent, thickness=0.5)
        _add_paragraph(doc, notes,
                       font_name="Calibri", font_size=11,
                       font_color=(51, 51, 51),
                       space_before=6, space_after=6, alignment=WD_ALIGN_LEFT)

    # Action items section
    if action_items:
        _add_paragraph(doc, "Action Items",
                       font_name="Calibri", font_size=14,
                       font_color=primary, bold=True,
                       space_before=18, space_after=6, alignment=WD_ALIGN_LEFT)
        _add_horizontal_line(doc, color=accent, thickness=0.5)

        header = ["#", "Task", "Owner", "Due Date"]
        rows_data = [header]
        for i, item in enumerate(action_items, 1):
            rows_data.append([
                str(i),
                item.get("task", ""),
                item.get("owner", ""),
                item.get("due", ""),
            ])

        action_table = _create_table(doc, len(rows_data), 4, rows_data)
        _format_table_header(action_table, fill_color=primary)
        _format_table_body(action_table, alt_row_color=(240, 245, 250))

        try:
            action_table.Columns(1).Width = 30
            action_table.Columns(2).Width = 240
            action_table.Columns(3).Width = 100
            action_table.Columns(4).Width = 90
        except Exception:
            pass

    return {
        "title": title,
        "date": date,
        "attendee_count": len(attendees),
        "agenda_count": len(agenda_items),
        "action_item_count": len(action_items) if action_items else 0,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 4. Report Template
# ---------------------------------------------------------------------------

def create_report_template(title: str, author: str | None = None,
                           sections: list[str] | None = None,
                           style: str = "business") -> dict:
    """Create a structured report with cover page, TOC, section headers, page numbers.

    Styles: business, technical, executive_summary.
    """
    doc = _get_active_doc()

    style_config = {
        "business": {
            "primary": (0, 70, 127),
            "accent": (0, 112, 192),
            "font": "Calibri",
            "title_size": 32,
        },
        "technical": {
            "primary": (51, 51, 51),
            "accent": (68, 114, 196),
            "font": "Consolas",
            "title_size": 28,
        },
        "executive_summary": {
            "primary": (88, 44, 77),
            "accent": (192, 160, 128),
            "font": "Georgia",
            "title_size": 30,
        },
    }

    cfg = style_config.get(style, style_config["business"])
    primary = cfg["primary"]
    accent = cfg["accent"]
    font = cfg["font"]

    # --- Cover page ---
    _add_empty_lines(doc, 8)

    _add_paragraph(doc, title,
                   font_name=font, font_size=cfg["title_size"],
                   font_color=primary, bold=True,
                   space_before=0, space_after=12, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    if author:
        _add_empty_lines(doc, 3)
        _add_paragraph(doc, f"Prepared by: {author}",
                       font_name=font, font_size=13,
                       font_color=(100, 100, 100),
                       space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)

    date_str = datetime.date.today().strftime("%B %d, %Y")
    _add_paragraph(doc, date_str,
                   font_name=font, font_size=12,
                   font_color=(120, 120, 120),
                   space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)

    # Page break after cover
    _insert_page_break(doc)

    # --- Table of Contents placeholder ---
    _add_paragraph(doc, "Table of Contents",
                   font_name=font, font_size=20,
                   font_color=primary, bold=True,
                   space_before=12, space_after=12, alignment=WD_ALIGN_LEFT)

    # Insert actual TOC
    rng = _collapse_to_end(doc)
    try:
        doc.TablesOfContents.Add(
            Range=rng,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=3,
        )
    except Exception:
        _add_paragraph(doc, "[Table of Contents - Update field to populate]",
                       font_name=font, font_size=11,
                       font_color=(150, 150, 150), italic=True)

    _insert_page_break(doc)

    # --- Sections ---
    default_sections = ["Introduction", "Background", "Analysis", "Findings", "Recommendations", "Conclusion"]
    section_list = sections or default_sections

    for i, section_title in enumerate(section_list):
        # Add section heading
        rng = _collapse_to_end(doc)
        rng.InsertAfter(section_title + "\r")
        para = doc.Paragraphs(doc.Paragraphs.Count)
        para.Style = WD_STYLE_HEADING1
        para.Range.Font.Name = font
        para.Range.Font.Color = rgb(*primary)

        # Placeholder content
        _add_paragraph(doc, f"[Content for {section_title}]",
                       font_name=font, font_size=11,
                       font_color=(150, 150, 150), italic=True,
                       space_before=6, space_after=12)

        # Page break between sections (except last)
        if i < len(section_list) - 1:
            _insert_page_break(doc)

    # --- Page numbers ---
    try:
        section = doc.Sections(1)
        footer = section.Footers(1)
        footer.PageNumbers.Add(WD_ALIGN_CENTER)
    except Exception:
        pass

    return {
        "title": title,
        "style": style,
        "section_count": len(section_list),
        "sections": section_list,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 5. Business Letter
# ---------------------------------------------------------------------------

def create_letter(recipient_name: str, recipient_address: str,
                  subject: str, body_paragraphs: list[str],
                  sender_name: str, sender_title: str | None = None,
                  company: str | None = None,
                  style: str = "formal") -> dict:
    """Create a business letter.

    Styles: formal (block format), semi_formal (modified block), modern (contemporary).
    """
    doc = _get_active_doc()

    style_config = {
        "formal": {
            "font": "Times New Roman",
            "size": 12,
            "color": (0, 0, 0),
            "align": WD_ALIGN_LEFT,
            "spacing": 1.0,
        },
        "semi_formal": {
            "font": "Calibri",
            "size": 11,
            "color": (51, 51, 51),
            "align": WD_ALIGN_LEFT,
            "spacing": 1.15,
        },
        "modern": {
            "font": "Segoe UI",
            "size": 10.5,
            "color": (38, 38, 38),
            "align": WD_ALIGN_LEFT,
            "spacing": 1.2,
        },
    }

    cfg = style_config.get(style, style_config["formal"])
    font = cfg["font"]
    size = cfg["size"]
    color = cfg["color"]
    alignment = cfg["align"]

    # Sender info (for semi_formal and modern, right-aligned)
    if style in ("semi_formal", "modern"):
        sender_align = WD_ALIGN_RIGHT
    else:
        sender_align = WD_ALIGN_LEFT

    # Company header
    if company:
        _add_paragraph(doc, company,
                       font_name=font, font_size=size + 2,
                       font_color=color, bold=True,
                       space_before=0, space_after=2, alignment=sender_align)

    if sender_title:
        _add_paragraph(doc, f"{sender_name}, {sender_title}",
                       font_name=font, font_size=size,
                       font_color=color,
                       space_before=2, space_after=2, alignment=sender_align)
    else:
        _add_paragraph(doc, sender_name,
                       font_name=font, font_size=size,
                       font_color=color,
                       space_before=2, space_after=2, alignment=sender_align)

    # Date
    _add_empty_lines(doc, 1)
    date_str = datetime.date.today().strftime("%B %d, %Y")
    _add_paragraph(doc, date_str,
                   font_name=font, font_size=size,
                   font_color=color,
                   space_before=6, space_after=6, alignment=sender_align)

    _add_empty_lines(doc, 1)

    # Recipient
    _add_paragraph(doc, recipient_name,
                   font_name=font, font_size=size,
                   font_color=color, bold=True,
                   space_before=0, space_after=2, alignment=WD_ALIGN_LEFT)

    for line in recipient_address.split("\n"):
        _add_paragraph(doc, line.strip(),
                       font_name=font, font_size=size,
                       font_color=color,
                       space_before=0, space_after=2, alignment=WD_ALIGN_LEFT)

    _add_empty_lines(doc, 1)

    # Subject line
    if style == "modern":
        _add_horizontal_line(doc, color=(0, 70, 127), thickness=1.5)
    _add_paragraph(doc, f"Re: {subject}",
                   font_name=font, font_size=size,
                   font_color=color, bold=True,
                   space_before=6, space_after=12, alignment=WD_ALIGN_LEFT)
    if style == "modern":
        _add_horizontal_line(doc, color=(0, 70, 127), thickness=1.5)

    _add_empty_lines(doc, 1)

    # Salutation
    _add_paragraph(doc, f"Dear {recipient_name},",
                   font_name=font, font_size=size,
                   font_color=color,
                   space_before=6, space_after=12, alignment=WD_ALIGN_LEFT)

    # Body paragraphs
    for para_text in body_paragraphs:
        _add_paragraph(doc, para_text,
                       font_name=font, font_size=size,
                       font_color=color,
                       space_before=0, space_after=12,
                       alignment=alignment, line_spacing=cfg["spacing"])

    # Closing
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Sincerely,",
                   font_name=font, font_size=size,
                   font_color=color,
                   space_before=6, space_after=36, alignment=WD_ALIGN_LEFT)

    _add_paragraph(doc, sender_name,
                   font_name=font, font_size=size,
                   font_color=color, bold=True,
                   space_before=0, space_after=2, alignment=WD_ALIGN_LEFT)

    if sender_title:
        _add_paragraph(doc, sender_title,
                       font_name=font, font_size=size,
                       font_color=color,
                       space_before=0, space_after=2, alignment=WD_ALIGN_LEFT)

    if company:
        _add_paragraph(doc, company,
                       font_name=font, font_size=size,
                       font_color=color,
                       space_before=0, space_after=2, alignment=WD_ALIGN_LEFT)

    return {
        "recipient": recipient_name,
        "subject": subject,
        "style": style,
        "body_paragraphs": len(body_paragraphs),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 6. Invoice
# ---------------------------------------------------------------------------

def create_invoice(company_name: str, client_name: str,
                   items: list[dict],
                   invoice_number: str | None = None,
                   date: str | None = None,
                   due_date: str | None = None,
                   notes: str | None = None,
                   tax_rate: float | None = None) -> dict:
    """Create a professional invoice.

    items: list of {"description": str, "quantity": int/float, "unit_price": float}
    """
    doc = _get_active_doc()
    primary = (0, 70, 127)
    accent = (0, 112, 192)

    # Header
    _add_paragraph(doc, "INVOICE",
                   font_name="Calibri", font_size=36,
                   font_color=primary, bold=True,
                   space_before=0, space_after=6, alignment=WD_ALIGN_RIGHT)

    _add_horizontal_line(doc, color=accent, thickness=2.5)

    # Company and Invoice info side by side (using table)
    inv_num = invoice_number or f"INV-{datetime.date.today().strftime('%Y%m%d')}-001"
    inv_date = date or datetime.date.today().strftime("%Y-%m-%d")
    inv_due = due_date or (datetime.date.today() + datetime.timedelta(days=30)).strftime("%Y-%m-%d")

    info_data = [
        [company_name, f"Invoice #: {inv_num}"],
        ["", f"Date: {inv_date}"],
        [f"Bill To: {client_name}", f"Due Date: {inv_due}"],
    ]

    info_table = _create_table(doc, 3, 2, info_data)
    info_table.Borders.Enable = False

    # Format info table
    for r in range(1, 4):
        for c in range(1, 3):
            cell = info_table.Cell(r, c)
            cell.Range.Font.Name = "Calibri"
            cell.Range.Font.Size = 10.5

    info_table.Cell(1, 1).Range.Font.Bold = True
    info_table.Cell(1, 1).Range.Font.Size = 14
    info_table.Cell(1, 1).Range.Font.Color = rgb(*primary)
    info_table.Cell(3, 1).Range.Font.Bold = True

    try:
        info_table.Columns(1).Width = 270
        info_table.Columns(2).Width = 190
    except Exception:
        pass

    for r in range(1, 4):
        info_table.Cell(r, 2).Range.ParagraphFormat.Alignment = WD_ALIGN_RIGHT

    _add_empty_lines(doc, 1)

    # Items table
    header = ["Description", "Qty", "Unit Price", "Amount"]
    rows_data = [header]

    subtotal = 0.0
    for item in items:
        qty = item.get("quantity", 1)
        price = item.get("unit_price", 0)
        amount = qty * price
        subtotal += amount
        rows_data.append([
            item.get("description", ""),
            str(qty),
            f"${price:,.2f}",
            f"${amount:,.2f}",
        ])

    items_table = _create_table(doc, len(rows_data), 4, rows_data)
    _format_table_header(items_table, fill_color=primary, font_name="Calibri", font_size=10.5)
    _format_table_body(items_table, font_name="Calibri", font_size=10.5,
                       alt_row_color=(245, 248, 252))

    # Right-align numeric columns
    for r in range(1, items_table.Rows.Count + 1):
        for c in [2, 3, 4]:
            items_table.Cell(r, c).Range.ParagraphFormat.Alignment = WD_ALIGN_RIGHT

    try:
        items_table.Columns(1).Width = 240
        items_table.Columns(2).Width = 50
        items_table.Columns(3).Width = 85
        items_table.Columns(4).Width = 85
    except Exception:
        pass

    _add_empty_lines(doc, 1)

    # Totals section
    tax_amount = 0.0
    if tax_rate:
        tax_amount = subtotal * (tax_rate / 100)

    total = subtotal + tax_amount

    totals_data = [["Subtotal", f"${subtotal:,.2f}"]]
    if tax_rate:
        totals_data.append([f"Tax ({tax_rate}%)", f"${tax_amount:,.2f}"])
    totals_data.append(["Total", f"${total:,.2f}"])

    totals_table = _create_table(doc, len(totals_data), 2, totals_data)
    totals_table.Borders.Enable = False

    try:
        totals_table.Columns(1).Width = 100
        totals_table.Columns(2).Width = 100
    except Exception:
        pass

    # Align totals table to the right
    for r in range(1, totals_table.Rows.Count + 1):
        totals_table.Cell(r, 1).Range.Font.Name = "Calibri"
        totals_table.Cell(r, 1).Range.Font.Size = 11
        totals_table.Cell(r, 1).Range.Font.Bold = True
        totals_table.Cell(r, 1).Range.ParagraphFormat.Alignment = WD_ALIGN_RIGHT
        totals_table.Cell(r, 2).Range.Font.Name = "Calibri"
        totals_table.Cell(r, 2).Range.Font.Size = 11
        totals_table.Cell(r, 2).Range.ParagraphFormat.Alignment = WD_ALIGN_RIGHT

    # Bold the total row
    last_row = totals_table.Rows.Count
    totals_table.Cell(last_row, 1).Range.Font.Size = 14
    totals_table.Cell(last_row, 1).Range.Font.Color = rgb(*primary)
    totals_table.Cell(last_row, 2).Range.Font.Size = 14
    totals_table.Cell(last_row, 2).Range.Font.Bold = True
    totals_table.Cell(last_row, 2).Range.Font.Color = rgb(*primary)

    # Right-indent the totals table by using paragraph indentation
    totals_table.Rows.Alignment = WD_ALIGN_RIGHT

    # Notes
    if notes:
        _add_empty_lines(doc, 2)
        _add_horizontal_line(doc, color=(200, 200, 200), thickness=0.5)
        _add_paragraph(doc, "Notes:",
                       font_name="Calibri", font_size=10,
                       font_color=primary, bold=True,
                       space_before=6, space_after=4)
        _add_paragraph(doc, notes,
                       font_name="Calibri", font_size=9.5,
                       font_color=(100, 100, 100),
                       space_before=0, space_after=6)

    # Footer
    _add_empty_lines(doc, 2)
    _add_paragraph(doc, "Thank you for your business!",
                   font_name="Calibri", font_size=11,
                   font_color=accent, italic=True,
                   space_before=12, space_after=6, alignment=WD_ALIGN_CENTER)

    return {
        "company": company_name,
        "client": client_name,
        "invoice_number": inv_num,
        "item_count": len(items),
        "subtotal": subtotal,
        "tax": tax_amount,
        "total": total,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 7. Resume
# ---------------------------------------------------------------------------

def create_resume(name: str, contact_info: dict,
                  sections: list[dict],
                  style: str = "modern") -> dict:
    """Create a professional resume.

    contact_info: {"email": "", "phone": "", "address": "", "linkedin": ""}
    sections: [{"title": "Experience", "items": [{"title": "Job Title",
                "subtitle": "Company", "date": "2020-2024",
                "details": ["bullet1", "bullet2"]}]}]
    Styles: modern, classic, creative.
    """
    doc = _get_active_doc()

    style_config = {
        "modern": {
            "primary": (41, 65, 122),
            "accent": (68, 114, 196),
            "name_font": "Segoe UI Light",
            "heading_font": "Segoe UI Semibold",
            "body_font": "Segoe UI",
            "name_size": 28,
            "heading_size": 13,
            "body_size": 10,
        },
        "classic": {
            "primary": (0, 0, 0),
            "accent": (0, 0, 128),
            "name_font": "Times New Roman",
            "heading_font": "Times New Roman",
            "body_font": "Times New Roman",
            "name_size": 24,
            "heading_size": 13,
            "body_size": 11,
        },
        "creative": {
            "primary": (255, 87, 34),
            "accent": (33, 150, 243),
            "name_font": "Century Gothic",
            "heading_font": "Century Gothic",
            "body_font": "Calibri",
            "name_size": 30,
            "heading_size": 14,
            "body_size": 10,
        },
    }

    cfg = style_config.get(style, style_config["modern"])
    primary = cfg["primary"]
    accent = cfg["accent"]

    # Set tight margins for resume
    ps = doc.PageSetup
    ps.TopMargin = 36  # 0.5 inch
    ps.BottomMargin = 36
    ps.LeftMargin = 54  # 0.75 inch
    ps.RightMargin = 54

    # Name header
    _add_paragraph(doc, name,
                   font_name=cfg["name_font"], font_size=cfg["name_size"],
                   font_color=primary, bold=False,
                   space_before=0, space_after=4, alignment=WD_ALIGN_CENTER)

    # Contact info line
    contact_parts = []
    if "email" in contact_info:
        contact_parts.append(contact_info["email"])
    if "phone" in contact_info:
        contact_parts.append(contact_info["phone"])
    if "address" in contact_info:
        contact_parts.append(contact_info["address"])
    if "linkedin" in contact_info:
        contact_parts.append(contact_info["linkedin"])

    contact_line = "  |  ".join(contact_parts)
    _add_paragraph(doc, contact_line,
                   font_name=cfg["body_font"], font_size=9,
                   font_color=(100, 100, 100),
                   space_before=0, space_after=6, alignment=WD_ALIGN_CENTER)

    # Accent line under header
    if style == "modern":
        _add_horizontal_line(doc, color=accent, thickness=2.0)
    elif style == "classic":
        _add_horizontal_line(doc, color=(0, 0, 0), thickness=1.5)
        _add_horizontal_line(doc, color=(0, 0, 0), thickness=0.5)
    elif style == "creative":
        # Colored bar
        shape = doc.Shapes.AddShape(
            MSO_AUTO_SHAPE_RECTANGLE,
            Left=54, Top=0, Width=490, Height=4,
        )
        shape.Fill.ForeColor.RGB = rgb(*primary)
        shape.Line.Visible = False
        _add_empty_lines(doc, 1)

    # Sections
    for section in sections:
        section_title = section.get("title", "")

        # Section heading
        _add_paragraph(doc, section_title.upper(),
                       font_name=cfg["heading_font"], font_size=cfg["heading_size"],
                       font_color=primary, bold=True,
                       space_before=14, space_after=2, alignment=WD_ALIGN_LEFT)

        # Section line
        _add_horizontal_line(doc, color=accent, thickness=0.75)

        items = section.get("items", [])
        for item in items:
            item_title = item.get("title", "")
            item_subtitle = item.get("subtitle", "")
            item_date = item.get("date", "")
            details = item.get("details", [])

            # Title and date on same conceptual line
            if item_date:
                title_text = f"{item_title}    —    {item_date}"
            else:
                title_text = item_title

            _add_paragraph(doc, title_text,
                           font_name=cfg["body_font"], font_size=cfg["body_size"] + 0.5,
                           font_color=(38, 38, 38), bold=True,
                           space_before=8, space_after=1, alignment=WD_ALIGN_LEFT)

            if item_subtitle:
                _add_paragraph(doc, item_subtitle,
                               font_name=cfg["body_font"], font_size=cfg["body_size"],
                               font_color=(100, 100, 100), italic=True,
                               space_before=0, space_after=3, alignment=WD_ALIGN_LEFT)

            for detail in details:
                _add_paragraph(doc, f"• {detail}",
                               font_name=cfg["body_font"], font_size=cfg["body_size"],
                               font_color=(64, 64, 64),
                               space_before=1, space_after=1, alignment=WD_ALIGN_LEFT)

    return {
        "name": name,
        "style": style,
        "section_count": len(sections),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 8. Table of Figures
# ---------------------------------------------------------------------------

def create_table_of_figures(label: str = "Figure") -> dict:
    """Insert a table of figures/tables at the end of the document."""
    doc = _get_active_doc()
    rng = _collapse_to_end(doc)

    # Add heading
    rng.InsertAfter(f"Table of {label}s\r")
    para = doc.Paragraphs(doc.Paragraphs.Count)
    para.Style = WD_STYLE_HEADING1

    rng = _collapse_to_end(doc)

    try:
        doc.TablesOfFigures.Add(
            Range=rng,
            Caption=label,
        )
    except Exception:
        # Fallback: insert a TOC field code for captions
        try:
            rng.Fields.Add(
                Range=rng,
                Type=WD_FIELD_TOC,
                Text=f'\\c "{label}"',
            )
        except Exception:
            rng.InsertAfter(f"[Table of {label}s - Add captions to populate]\r")

    return {"label": label, "paragraphs": doc.Paragraphs.Count}


# ---------------------------------------------------------------------------
# 9. Cover Page Image
# ---------------------------------------------------------------------------

def add_cover_page_image(image_path: str, style: str = "full_bleed") -> dict:
    """Add an image to the cover page.

    Styles: full_bleed, centered, banner_top, banner_bottom.
    """
    doc = _get_active_doc()
    abs_path = ensure_absolute_path(image_path)
    ps = doc.PageSetup

    page_width = ps.PageWidth
    page_height = ps.PageHeight

    if style == "full_bleed":
        shape = doc.Shapes.AddPicture(
            FileName=abs_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0, Top=0,
            Width=page_width,
            Height=page_height,
        )
        shape.WrapFormat.Type = 3  # wdWrapBehindText
        shape.ZOrder(1)  # Send behind text

    elif style == "centered":
        img_width = page_width * 0.5
        img_height = page_height * 0.3
        left = (page_width - img_width) / 2
        top = page_height * 0.2

        shape = doc.Shapes.AddPicture(
            FileName=abs_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left, Top=top,
            Width=img_width,
            Height=img_height,
        )
        shape.WrapFormat.Type = 3  # wdWrapBehindText

    elif style == "banner_top":
        shape = doc.Shapes.AddPicture(
            FileName=abs_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0, Top=0,
            Width=page_width,
            Height=page_height * 0.3,
        )
        shape.WrapFormat.Type = 3

    elif style == "banner_bottom":
        shape = doc.Shapes.AddPicture(
            FileName=abs_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0, Top=page_height * 0.7,
            Width=page_width,
            Height=page_height * 0.3,
        )
        shape.WrapFormat.Type = 3

    else:
        return {"error": f"Unknown style: {style}"}

    return {"image": abs_path, "style": style}


# ---------------------------------------------------------------------------
# 10. Format All Headings
# ---------------------------------------------------------------------------

def format_all_headings(heading_styles: dict) -> dict:
    """Apply consistent heading formatting across the document.

    heading_styles: {"h1": {"font_size": 24, "color": [0,0,128], "bold": True, "font_name": "Calibri"},
                     "h2": {...}, "h3": {...}}
    """
    doc = _get_active_doc()
    updated = {}

    level_map = {"h1": WD_STYLE_HEADING1, "h2": WD_STYLE_HEADING2,
                 "h3": WD_STYLE_HEADING3, "h4": WD_STYLE_HEADING4}

    for key, style_id in level_map.items():
        if key not in heading_styles:
            continue

        cfg = heading_styles[key]
        try:
            style = doc.Styles(style_id)

            if "font_size" in cfg:
                style.Font.Size = cfg["font_size"]
            if "color" in cfg:
                c = cfg["color"]
                style.Font.Color = rgb(c[0], c[1], c[2])
            if "bold" in cfg:
                style.Font.Bold = cfg["bold"]
            if "italic" in cfg:
                style.Font.Italic = cfg["italic"]
            if "font_name" in cfg:
                style.Font.Name = cfg["font_name"]
            if "space_before" in cfg:
                style.ParagraphFormat.SpaceBefore = cfg["space_before"]
            if "space_after" in cfg:
                style.ParagraphFormat.SpaceAfter = cfg["space_after"]
            if "alignment" in cfg:
                style.ParagraphFormat.Alignment = cfg["alignment"]

            updated[key] = True
        except Exception as e:
            updated[key] = f"error: {e}"

    return {"updated": updated}


# ---------------------------------------------------------------------------
# 11. Sidebar
# ---------------------------------------------------------------------------

def add_sidebar(text: str, position: str = "left", width: float = 120,
                fill_color: tuple | None = None,
                text_color: tuple | None = None,
                font_size: float = 9) -> dict:
    """Add a colored sidebar with text for pull quotes or key facts.

    position: left or right.
    """
    doc = _get_active_doc()
    ps = doc.PageSetup

    default_fill = (0, 70, 127)
    default_text = (255, 255, 255)

    bg = fill_color or default_fill
    fg = text_color or default_text

    if position == "right":
        left = ps.PageWidth - ps.RightMargin - width + 20
    else:
        left = ps.LeftMargin - 20

    top = ps.TopMargin + 50

    # Create the sidebar shape
    shape = doc.Shapes.AddShape(
        MSO_AUTO_SHAPE_RECTANGLE,
        Left=left, Top=top,
        Width=width,
        Height=400,
    )

    shape.Fill.ForeColor.RGB = rgb(*bg)
    shape.Line.Visible = False

    # Add text to shape
    tf = shape.TextFrame
    tf.MarginLeft = 10
    tf.MarginRight = 10
    tf.MarginTop = 15
    tf.MarginBottom = 15
    tf.WordWrap = True

    tr = tf.TextRange
    tr.Text = text
    tr.Font.Color = rgb(*fg)
    tr.Font.Size = font_size
    tr.Font.Name = "Calibri"
    tr.ParagraphFormat.Alignment = WD_ALIGN_LEFT

    shape.WrapFormat.Type = 2  # wdWrapSquare

    return {"position": position, "width": width}


# ---------------------------------------------------------------------------
# 12. Newsletter
# ---------------------------------------------------------------------------

def create_newsletter(title: str, subtitle: str | None = None,
                      columns: int = 2,
                      articles: list[dict] | None = None) -> dict:
    """Create a newsletter layout with headline and articles.

    articles: [{"title": str, "body": str, "author": str}]
    """
    doc = _get_active_doc()
    primary = (0, 70, 127)
    accent = (237, 125, 49)

    # Newsletter header banner
    shape = doc.Shapes.AddShape(
        MSO_AUTO_SHAPE_RECTANGLE,
        Left=0, Top=0,
        Width=620, Height=72,
    )
    shape.Fill.ForeColor.RGB = rgb(*primary)
    shape.Line.Visible = False

    tf = shape.TextFrame
    tf.MarginLeft = 36
    tf.MarginRight = 36
    tf.MarginTop = 12
    tf.MarginBottom = 8

    tr = tf.TextRange
    tr.Text = title
    tr.Font.Color = rgb(255, 255, 255)
    tr.Font.Size = 28
    tr.Font.Name = "Calibri"
    tr.Font.Bold = True
    tr.ParagraphFormat.Alignment = WD_ALIGN_LEFT

    _add_empty_lines(doc, 3)

    # Date and subtitle bar
    date_str = datetime.date.today().strftime("%B %Y")

    if subtitle:
        _add_paragraph(doc, f"{subtitle}  |  {date_str}",
                       font_name="Calibri", font_size=10,
                       font_color=(100, 100, 100),
                       space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)
    else:
        _add_paragraph(doc, date_str,
                       font_name="Calibri", font_size=10,
                       font_color=(100, 100, 100),
                       space_before=6, space_after=6, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # Set up columns for the content section
    _insert_section_break(doc, WD_SECTION_BREAK_CONTINUOUS)

    try:
        section = doc.Sections(doc.Sections.Count)
        section.PageSetup.TextColumns.SetCount(columns)
        section.PageSetup.TextColumns.EvenlySpaced = True
        section.PageSetup.TextColumns.LineBetween = True
    except Exception:
        pass

    # Articles
    article_list = articles or []
    for i, article in enumerate(article_list):
        art_title = article.get("title", "")
        art_body = article.get("body", "")
        art_author = article.get("author", "")

        # Article title
        _add_paragraph(doc, art_title,
                       font_name="Calibri", font_size=14,
                       font_color=primary, bold=True,
                       space_before=12, space_after=4, alignment=WD_ALIGN_LEFT)

        if art_author:
            _add_paragraph(doc, f"By {art_author}",
                           font_name="Calibri", font_size=8.5,
                           font_color=(130, 130, 130), italic=True,
                           space_before=0, space_after=6, alignment=WD_ALIGN_LEFT)

        _add_paragraph(doc, art_body,
                       font_name="Calibri", font_size=10,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=8,
                       alignment=WD_ALIGN_JUSTIFY, line_spacing=1.15)

        if i < len(article_list) - 1:
            _add_horizontal_line(doc, color=(200, 200, 200), thickness=0.5)

    return {
        "title": title,
        "columns": columns,
        "article_count": len(article_list),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 13. Callout Box
# ---------------------------------------------------------------------------

def add_callout_box(text: str, style: str = "info",
                    paragraph_index: int | None = None) -> dict:
    """Add a callout/alert box with colored left border and light background.

    Styles: info (blue), warning (yellow), success (green), error (red), tip (purple).
    """
    doc = _get_active_doc()
    cs = CALLOUT_STYLES.get(style, CALLOUT_STYLES["info"])

    border_color = cs["border_color"]
    bg_color = cs["bg_color"]
    icon = cs["icon"]

    # Build the callout text
    callout_text = f"{icon}  {text}"

    # Add paragraph at specified location or end
    if paragraph_index is not None and paragraph_index <= doc.Paragraphs.Count:
        para = doc.Paragraphs(paragraph_index)
        rng = para.Range
        rng.Collapse(WD_COLLAPSE_START)
        rng.InsertBefore(callout_text + "\r")
        callout_para = doc.Paragraphs(paragraph_index)
    else:
        callout_para = _add_paragraph(doc, callout_text)

    # Format the callout paragraph
    callout_para.Range.Font.Name = "Calibri"
    callout_para.Range.Font.Size = 10.5
    callout_para.Range.Font.Color = rgb(51, 51, 51)

    pf = callout_para.Format
    pf.SpaceBefore = 8
    pf.SpaceAfter = 8

    # Left indent to create visual space
    pf.LeftIndent = 18
    pf.RightIndent = 18

    # Colored left border
    border = pf.Borders(WD_BORDER_LEFT)
    border.LineStyle = 1  # wdLineStyleSingle
    border.LineWidth = 24  # thick left border
    border.Color = rgb(*border_color)

    # Background shading
    callout_para.Shading.BackgroundPatternColor = rgb(*bg_color)

    # Top and bottom padding borders (thin, same as bg)
    for border_id in [WD_BORDER_TOP, WD_BORDER_BOTTOM]:
        b = pf.Borders(border_id)
        b.LineStyle = 1
        b.LineWidth = 2
        b.Color = rgb(*bg_color)

    return {"style": style, "text": text[:50] + "..." if len(text) > 50 else text}


# ---------------------------------------------------------------------------
# 14. Contract
# ---------------------------------------------------------------------------

def create_contract(title: str, parties: list[str],
                    clauses: list[dict],
                    date: str | None = None,
                    style: str = "standard") -> dict:
    """Create a legal contract template.

    parties: list of party names
    clauses: [{"title": "Term", "content": "This agreement shall..."}]
    """
    doc = _get_active_doc()
    primary = (0, 0, 0)
    accent = (64, 64, 64)
    font = "Times New Roman"

    # Title
    _add_paragraph(doc, title.upper(),
                   font_name=font, font_size=18,
                   font_color=primary, bold=True,
                   space_before=24, space_after=18, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=primary, thickness=1.5)

    # Preamble
    contract_date = date or datetime.date.today().strftime("%B %d, %Y")

    preamble = f"This {title} (\"Agreement\") is entered into as of {contract_date}, by and between:"
    _add_paragraph(doc, preamble,
                   font_name=font, font_size=11,
                   font_color=primary,
                   space_before=18, space_after=12,
                   alignment=WD_ALIGN_JUSTIFY, line_spacing=1.5)

    # Parties
    for i, party in enumerate(parties):
        party_label = f"({chr(65 + i)}) {party}"
        if i < len(parties) - 1:
            party_label += ","
        else:
            party_label += "."
        _add_paragraph(doc, party_label,
                       font_name=font, font_size=11,
                       font_color=primary, bold=True,
                       space_before=4, space_after=4,
                       alignment=WD_ALIGN_LEFT)

    _add_empty_lines(doc, 1)

    intro = "NOW, THEREFORE, in consideration of the mutual covenants and agreements herein contained, the parties agree as follows:"
    _add_paragraph(doc, intro,
                   font_name=font, font_size=11,
                   font_color=primary,
                   space_before=12, space_after=18,
                   alignment=WD_ALIGN_JUSTIFY, line_spacing=1.5)

    # Clauses
    for i, clause in enumerate(clauses, 1):
        clause_title = clause.get("title", f"Section {i}")
        clause_content = clause.get("content", "")

        # Clause heading
        _add_paragraph(doc, f"{i}. {clause_title.upper()}",
                       font_name=font, font_size=12,
                       font_color=primary, bold=True,
                       space_before=14, space_after=6,
                       alignment=WD_ALIGN_LEFT)

        # Clause content
        _add_paragraph(doc, clause_content,
                       font_name=font, font_size=11,
                       font_color=primary,
                       space_before=4, space_after=8,
                       alignment=WD_ALIGN_JUSTIFY, line_spacing=1.5)

    # Signature blocks
    _add_empty_lines(doc, 2)
    _add_paragraph(doc, "IN WITNESS WHEREOF, the parties have executed this Agreement as of the date first above written.",
                   font_name=font, font_size=11,
                   font_color=primary,
                   space_before=18, space_after=24,
                   alignment=WD_ALIGN_JUSTIFY, line_spacing=1.5)

    for party in parties:
        _add_empty_lines(doc, 2)

        # Signature line
        _add_paragraph(doc, "_" * 40,
                       font_name=font, font_size=11,
                       font_color=primary,
                       space_before=36, space_after=2, alignment=WD_ALIGN_LEFT)

        _add_paragraph(doc, f"Name: {party}",
                       font_name=font, font_size=11,
                       font_color=primary,
                       space_before=2, space_after=2, alignment=WD_ALIGN_LEFT)

        _add_paragraph(doc, "Title: _________________________",
                       font_name=font, font_size=11,
                       font_color=primary,
                       space_before=2, space_after=2, alignment=WD_ALIGN_LEFT)

        _add_paragraph(doc, "Date: _________________________",
                       font_name=font, font_size=11,
                       font_color=primary,
                       space_before=2, space_after=12, alignment=WD_ALIGN_LEFT)

    # Page numbers
    try:
        section = doc.Sections(1)
        footer = section.Footers(1)
        footer.PageNumbers.Add(WD_ALIGN_CENTER)
    except Exception:
        pass

    return {
        "title": title,
        "party_count": len(parties),
        "clause_count": len(clauses),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 15. Insert Caption
# ---------------------------------------------------------------------------

def insert_caption(text: str, label: str = "Figure",
                   position: str = "below") -> dict:
    """Insert a caption for figures/tables with auto-numbering.

    position: below or above.
    """
    doc = _get_active_doc()

    # Use Word's built-in caption functionality
    try:
        # CaptionLabels - try to use existing or add new
        try:
            doc.Application.CaptionLabels(label)
        except Exception:
            doc.Application.CaptionLabels.Add(label)

        # Insert caption at the end of the document
        rng = _collapse_to_end(doc)
        rng.InsertCaption(
            Label=label,
            Title=f": {text}",
            Position=WD_CAPTION_POSITION_BELOW if position == "below" else WD_CAPTION_POSITION_ABOVE,
        )
    except Exception:
        # Fallback: manual caption
        caption_text = f"{label}: {text}"
        para = _add_paragraph(doc, caption_text,
                              font_name="Calibri", font_size=9,
                              font_color=(100, 100, 100), italic=True,
                              space_before=4, space_after=8,
                              alignment=WD_ALIGN_CENTER)

    return {"label": label, "text": text, "position": position}


# ---------------------------------------------------------------------------
# 16. Business Proposal
# ---------------------------------------------------------------------------

PROPOSAL_STYLES = {
    "professional": {
        "primary": (0, 70, 127),
        "accent": (0, 112, 192),
        "light": (217, 226, 243),
        "title_font": "Calibri",
        "body_font": "Calibri",
        "title_size": 36,
    },
    "creative": {
        "primary": (255, 87, 34),
        "accent": (33, 150, 243),
        "light": (255, 235, 238),
        "title_font": "Century Gothic",
        "body_font": "Segoe UI",
        "title_size": 40,
    },
    "minimal": {
        "primary": (51, 51, 51),
        "accent": (160, 160, 160),
        "light": (245, 245, 245),
        "title_font": "Calibri Light",
        "body_font": "Calibri",
        "title_size": 34,
    },
}


def create_proposal(title: str, client_name: str, sections: list[dict],
                    author: str | None = None, date: str | None = None,
                    style: str = "professional") -> dict:
    """Create a business proposal document.

    Args:
        title: Proposal title.
        client_name: Name of the client.
        sections: [{"title": "...", "content": "..."}]
        author: Author name.
        date: Date string.
        style: professional, creative, minimal.

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    ps = PROPOSAL_STYLES.get(style, PROPOSAL_STYLES["professional"])
    primary = ps["primary"]
    accent = ps["accent"]
    light = ps["light"]
    title_font = ps["title_font"]
    body_font = ps["body_font"]
    title_size = ps["title_size"]
    proposal_date = date or datetime.date.today().strftime("%B %d, %Y")

    # --- Cover page ---
    _add_empty_lines(doc, 4)

    # Accent bar via shape
    shape = doc.Shapes.AddShape(
        MSO_AUTO_SHAPE_RECTANGLE,
        Left=0, Top=0,
        Width=620, Height=8,
    )
    shape.Fill.ForeColor.RGB = rgb(*accent)
    shape.Line.Visible = False

    # Title
    _add_paragraph(doc, title.upper(),
                   font_name=title_font, font_size=title_size,
                   font_color=primary, bold=True,
                   space_before=60, space_after=12, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # Client name
    _add_paragraph(doc, f"Prepared for: {client_name}",
                   font_name=body_font, font_size=14,
                   font_color=primary,
                   space_before=18, space_after=6, alignment=WD_ALIGN_LEFT)

    if author:
        _add_paragraph(doc, f"Prepared by: {author}",
                       font_name=body_font, font_size=12,
                       font_color=(100, 100, 100),
                       space_before=4, space_after=4, alignment=WD_ALIGN_LEFT)

    _add_paragraph(doc, proposal_date,
                   font_name=body_font, font_size=12,
                   font_color=(100, 100, 100),
                   space_before=4, space_after=24, alignment=WD_ALIGN_LEFT)

    _insert_page_break(doc)

    # --- Table of Contents placeholder ---
    _add_paragraph(doc, "TABLE OF CONTENTS",
                   font_name=title_font, font_size=18,
                   font_color=primary, bold=True,
                   space_before=24, space_after=18, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=1.0)

    for i, section in enumerate(sections, 1):
        sec_title = section.get("title", f"Section {i}")
        _add_paragraph(doc, f"{i}.  {sec_title}",
                       font_name=body_font, font_size=12,
                       font_color=primary,
                       space_before=6, space_after=6, alignment=WD_ALIGN_LEFT)

    _insert_page_break(doc)

    # --- Sections ---
    for i, section in enumerate(sections, 1):
        sec_title = section.get("title", f"Section {i}")
        sec_content = section.get("content", "")

        # Section heading
        _add_paragraph(doc, f"{i}. {sec_title}",
                       style=WD_STYLE_HEADING1,
                       font_name=title_font, font_size=20,
                       font_color=primary, bold=True,
                       space_before=24, space_after=8, alignment=WD_ALIGN_LEFT)

        # Accent underline
        _add_horizontal_line(doc, color=accent, thickness=1.5)

        # Content
        _add_paragraph(doc, sec_content,
                       font_name=body_font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=8, space_after=16,
                       alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

    # --- Footer with page numbers ---
    try:
        section_obj = doc.Sections(1)
        footer = section_obj.Footers(1)
        footer.PageNumbers.Add(WD_ALIGN_CENTER)
    except Exception:
        pass

    return {
        "title": title,
        "client": client_name,
        "style": style,
        "section_count": len(sections),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 17. Standard Operating Procedure (SOP)
# ---------------------------------------------------------------------------

def create_sop(title: str, purpose: str, scope: str,
               procedures: list[dict],
               responsibilities: list[dict] | None = None,
               style: str = "standard") -> dict:
    """Create a Standard Operating Procedure document.

    Args:
        title: SOP title.
        purpose: Purpose statement.
        scope: Scope description.
        procedures: [{"step": 1, "title": "...", "details": "...", "caution": "..."}]
        responsibilities: [{"role": "...", "responsibility": "..."}]
        style: standard.

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    primary = (0, 70, 127)
    accent = (0, 112, 192)
    caution_color = (204, 102, 0)
    font = "Calibri"

    # Header block
    shape = doc.Shapes.AddShape(
        MSO_AUTO_SHAPE_RECTANGLE,
        Left=0, Top=0,
        Width=620, Height=80,
    )
    shape.Fill.ForeColor.RGB = rgb(*primary)
    shape.Line.Visible = False
    tf = shape.TextFrame
    tf.MarginLeft = 36
    tf.MarginTop = 14
    tr = tf.TextRange
    tr.Text = "STANDARD OPERATING PROCEDURE"
    tr.Font.Color = rgb(255, 255, 255)
    tr.Font.Size = 14
    tr.Font.Name = font
    tr.Font.Bold = True

    _add_empty_lines(doc, 3)

    # Title
    _add_paragraph(doc, title,
                   font_name=font, font_size=22,
                   font_color=primary, bold=True,
                   space_before=24, space_after=12, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # Metadata table
    sop_date = datetime.date.today().strftime("%B %d, %Y")
    meta_data = [
        ["Document Title:", title],
        ["Effective Date:", sop_date],
        ["Version:", "1.0"],
    ]
    meta_table = _create_table(doc, len(meta_data), 2, meta_data)
    for r in range(1, len(meta_data) + 1):
        meta_table.Cell(r, 1).Range.Font.Bold = True
        meta_table.Cell(r, 1).Range.Font.Name = font
        meta_table.Cell(r, 1).Range.Font.Size = 10
        meta_table.Cell(r, 1).Shading.BackgroundPatternColor = rgb(234, 240, 250)
        meta_table.Cell(r, 2).Range.Font.Name = font
        meta_table.Cell(r, 2).Range.Font.Size = 10
    try:
        meta_table.Columns(1).Width = 130
        meta_table.Columns(2).Width = 350
    except Exception:
        pass

    _add_empty_lines(doc, 1)

    # Purpose section
    _add_paragraph(doc, "1. PURPOSE",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_paragraph(doc, purpose,
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=4, space_after=12,
                   alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

    # Scope section
    _add_paragraph(doc, "2. SCOPE",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_paragraph(doc, scope,
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=4, space_after=12,
                   alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

    # Responsibilities section
    if responsibilities:
        _add_paragraph(doc, "3. RESPONSIBILITIES",
                       font_name=font, font_size=14,
                       font_color=primary, bold=True,
                       space_before=18, space_after=6, alignment=WD_ALIGN_LEFT)

        resp_data = [["Role", "Responsibility"]]
        for resp in responsibilities:
            resp_data.append([resp.get("role", ""), resp.get("responsibility", "")])

        resp_table = _create_table(doc, len(resp_data), 2, resp_data)
        _format_table_header(resp_table, fill_color=primary)
        _format_table_body(resp_table, alt_row_color=(234, 240, 250))
        try:
            resp_table.Columns(1).Width = 150
            resp_table.Columns(2).Width = 330
        except Exception:
            pass
        _add_empty_lines(doc, 1)

    # Procedures section
    proc_heading_num = "4" if responsibilities else "3"
    _add_paragraph(doc, f"{proc_heading_num}. PROCEDURES",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=8, alignment=WD_ALIGN_LEFT)

    for proc in procedures:
        step_num = proc.get("step", "")
        step_title = proc.get("title", "")
        step_details = proc.get("details", "")
        step_caution = proc.get("caution", "")

        # Step heading
        _add_paragraph(doc, f"Step {step_num}: {step_title}",
                       font_name=font, font_size=12,
                       font_color=accent, bold=True,
                       space_before=12, space_after=4, alignment=WD_ALIGN_LEFT)

        # Details
        _add_paragraph(doc, step_details,
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=6,
                       alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

        # Caution note
        if step_caution:
            caution_para = _add_paragraph(doc, f"⚠ CAUTION: {step_caution}",
                                          font_name=font, font_size=10,
                                          font_color=caution_color, bold=True,
                                          space_before=4, space_after=8,
                                          alignment=WD_ALIGN_LEFT)
            caution_para.Shading.BackgroundPatternColor = rgb(255, 243, 205)
            pf = caution_para.Format
            pf.LeftIndent = 18
            pf.RightIndent = 18
            border = pf.Borders(WD_BORDER_LEFT)
            border.LineStyle = 1
            border.LineWidth = 16
            border.Color = rgb(255, 185, 0)

    # Footer
    try:
        section_obj = doc.Sections(1)
        footer = section_obj.Footers(1)
        footer.PageNumbers.Add(WD_ALIGN_CENTER)
    except Exception:
        pass

    return {
        "title": title,
        "procedure_count": len(procedures),
        "responsibility_count": len(responsibilities) if responsibilities else 0,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 18. Project Charter
# ---------------------------------------------------------------------------

def create_project_charter(project_name: str, sponsor: str, manager: str,
                           objectives: list[str], scope: str,
                           milestones: list[dict],
                           budget: str | None = None) -> dict:
    """Create a project charter document.

    Args:
        project_name: Project name.
        sponsor: Project sponsor name.
        manager: Project manager name.
        objectives: List of objective strings.
        scope: Project scope description.
        milestones: [{"name": "...", "date": "...", "deliverable": "..."}]
        budget: Optional budget string.

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    primary = (0, 51, 102)
    accent = (0, 112, 192)
    font = "Calibri"
    charter_date = datetime.date.today().strftime("%B %d, %Y")

    # Title banner
    shape = doc.Shapes.AddShape(
        MSO_AUTO_SHAPE_RECTANGLE,
        Left=0, Top=0,
        Width=620, Height=72,
    )
    shape.Fill.ForeColor.RGB = rgb(*primary)
    shape.Line.Visible = False
    tf = shape.TextFrame
    tf.MarginLeft = 36
    tf.MarginTop = 14
    tr = tf.TextRange
    tr.Text = "PROJECT CHARTER"
    tr.Font.Color = rgb(255, 255, 255)
    tr.Font.Size = 24
    tr.Font.Name = font
    tr.Font.Bold = True

    _add_empty_lines(doc, 3)

    # Project name
    _add_paragraph(doc, project_name,
                   font_name=font, font_size=26,
                   font_color=primary, bold=True,
                   space_before=24, space_after=12, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # Project info table
    info_data = [
        ["Project Name:", project_name],
        ["Project Sponsor:", sponsor],
        ["Project Manager:", manager],
        ["Date:", charter_date],
    ]
    info_table = _create_table(doc, len(info_data), 2, info_data)
    for r in range(1, len(info_data) + 1):
        info_table.Cell(r, 1).Range.Font.Bold = True
        info_table.Cell(r, 1).Range.Font.Name = font
        info_table.Cell(r, 1).Range.Font.Size = 10.5
        info_table.Cell(r, 1).Shading.BackgroundPatternColor = rgb(217, 226, 243)
        info_table.Cell(r, 2).Range.Font.Name = font
        info_table.Cell(r, 2).Range.Font.Size = 10.5
    try:
        info_table.Columns(1).Width = 140
        info_table.Columns(2).Width = 340
    except Exception:
        pass

    _add_empty_lines(doc, 1)

    # Objectives
    _add_paragraph(doc, "PROJECT OBJECTIVES",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    for obj in objectives:
        _add_paragraph(doc, f"•  {obj}",
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=4,
                       alignment=WD_ALIGN_LEFT, line_spacing=1.2)

    _add_empty_lines(doc, 1)

    # Scope
    _add_paragraph(doc, "PROJECT SCOPE",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    _add_paragraph(doc, scope,
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=6, space_after=12,
                   alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

    # Milestones
    _add_paragraph(doc, "KEY MILESTONES",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    ms_data = [["Milestone", "Target Date", "Deliverable"]]
    for ms in milestones:
        ms_data.append([
            ms.get("name", ""),
            ms.get("date", ""),
            ms.get("deliverable", ""),
        ])

    ms_table = _create_table(doc, len(ms_data), 3, ms_data)
    _format_table_header(ms_table, fill_color=primary)
    _format_table_body(ms_table, alt_row_color=(234, 240, 250))
    try:
        ms_table.Columns(1).Width = 160
        ms_table.Columns(2).Width = 110
        ms_table.Columns(3).Width = 210
    except Exception:
        pass

    _add_empty_lines(doc, 1)

    # Budget (if provided)
    if budget:
        _add_paragraph(doc, "BUDGET",
                       font_name=font, font_size=14,
                       font_color=primary, bold=True,
                       space_before=18, space_after=8, alignment=WD_ALIGN_LEFT)
        _add_horizontal_line(doc, color=accent, thickness=1.0)
        _add_paragraph(doc, budget,
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=6, space_after=12,
                       alignment=WD_ALIGN_LEFT, line_spacing=1.3)

    # Approval signatures
    _add_empty_lines(doc, 2)
    _add_paragraph(doc, "APPROVALS",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=12, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    for person_label, person_name in [("Project Sponsor", sponsor), ("Project Manager", manager)]:
        _add_empty_lines(doc, 1)
        _add_paragraph(doc, "_" * 40,
                       font_name=font, font_size=11,
                       space_before=30, space_after=2, alignment=WD_ALIGN_LEFT)
        _add_paragraph(doc, f"{person_label}: {person_name}",
                       font_name=font, font_size=10,
                       font_color=(80, 80, 80),
                       space_before=2, space_after=2, alignment=WD_ALIGN_LEFT)
        _add_paragraph(doc, "Date: _________________________",
                       font_name=font, font_size=10,
                       font_color=(80, 80, 80),
                       space_before=2, space_after=8, alignment=WD_ALIGN_LEFT)

    # Page numbers
    try:
        section_obj = doc.Sections(1)
        footer = section_obj.Footers(1)
        footer.PageNumbers.Add(WD_ALIGN_CENTER)
    except Exception:
        pass

    return {
        "project_name": project_name,
        "sponsor": sponsor,
        "manager": manager,
        "objective_count": len(objectives),
        "milestone_count": len(milestones),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 19. Meeting Agenda
# ---------------------------------------------------------------------------

def create_meeting_agenda(title: str, date: str, time: str, location: str,
                          attendees: list[str], agenda_items: list[dict],
                          notes: str | None = None) -> dict:
    """Create a meeting agenda document.

    Args:
        title: Meeting title.
        date: Meeting date string.
        time: Meeting time string.
        location: Meeting location.
        attendees: List of attendee names.
        agenda_items: [{"topic": "...", "presenter": "...", "duration": "10min"}]
        notes: Optional additional notes.

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    primary = (0, 70, 127)
    accent = (0, 112, 192)
    font = "Calibri"

    # Header banner
    shape = doc.Shapes.AddShape(
        MSO_AUTO_SHAPE_RECTANGLE,
        Left=0, Top=0,
        Width=620, Height=60,
    )
    shape.Fill.ForeColor.RGB = rgb(*primary)
    shape.Line.Visible = False
    tf = shape.TextFrame
    tf.MarginLeft = 36
    tf.MarginTop = 12
    tr = tf.TextRange
    tr.Text = "MEETING AGENDA"
    tr.Font.Color = rgb(255, 255, 255)
    tr.Font.Size = 20
    tr.Font.Name = font
    tr.Font.Bold = True

    _add_empty_lines(doc, 2)

    # Meeting title
    _add_paragraph(doc, title,
                   font_name=font, font_size=20,
                   font_color=primary, bold=True,
                   space_before=18, space_after=12, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=1.5)

    # Meeting details table
    details_data = [
        ["Date:", date],
        ["Time:", time],
        ["Location:", location],
        ["Attendees:", ", ".join(attendees)],
    ]
    details_table = _create_table(doc, len(details_data), 2, details_data)
    for r in range(1, len(details_data) + 1):
        details_table.Cell(r, 1).Range.Font.Bold = True
        details_table.Cell(r, 1).Range.Font.Name = font
        details_table.Cell(r, 1).Range.Font.Size = 10.5
        details_table.Cell(r, 1).Shading.BackgroundPatternColor = rgb(217, 226, 243)
        details_table.Cell(r, 2).Range.Font.Name = font
        details_table.Cell(r, 2).Range.Font.Size = 10.5
    try:
        details_table.Columns(1).Width = 100
        details_table.Columns(2).Width = 380
    except Exception:
        pass
    details_table.Borders.Enable = True

    _add_empty_lines(doc, 1)

    # Agenda items table
    _add_paragraph(doc, "AGENDA",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=8, alignment=WD_ALIGN_LEFT)

    agenda_data = [["#", "Topic", "Presenter", "Duration"]]
    for i, item in enumerate(agenda_items, 1):
        agenda_data.append([
            str(i),
            item.get("topic", ""),
            item.get("presenter", ""),
            item.get("duration", ""),
        ])

    agenda_table = _create_table(doc, len(agenda_data), 4, agenda_data)
    _format_table_header(agenda_table, fill_color=primary)
    _format_table_body(agenda_table, alt_row_color=(234, 240, 250))
    try:
        agenda_table.Columns(1).Width = 35
        agenda_table.Columns(2).Width = 250
        agenda_table.Columns(3).Width = 110
        agenda_table.Columns(4).Width = 85
    except Exception:
        pass

    # Notes section
    if notes:
        _add_empty_lines(doc, 1)
        _add_paragraph(doc, "NOTES",
                       font_name=font, font_size=14,
                       font_color=primary, bold=True,
                       space_before=14, space_after=8, alignment=WD_ALIGN_LEFT)
        _add_paragraph(doc, notes,
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=12,
                       alignment=WD_ALIGN_LEFT, line_spacing=1.3)

    # Action items placeholder
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "ACTION ITEMS",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=8, alignment=WD_ALIGN_LEFT)

    action_data = [["#", "Action Item", "Owner", "Due Date", "Status"]]
    for i in range(1, 4):
        action_data.append([str(i), "", "", "", ""])

    action_table = _create_table(doc, len(action_data), 5, action_data)
    _format_table_header(action_table, fill_color=accent)
    try:
        action_table.Columns(1).Width = 30
        action_table.Columns(2).Width = 200
        action_table.Columns(3).Width = 90
        action_table.Columns(4).Width = 80
        action_table.Columns(5).Width = 80
    except Exception:
        pass

    return {
        "title": title,
        "attendee_count": len(attendees),
        "agenda_item_count": len(agenda_items),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 20. Employee Handbook Section
# ---------------------------------------------------------------------------

def create_employee_handbook_section(title: str, policies: list[dict],
                                     effective_date: str | None = None) -> dict:
    """Create an HR handbook section.

    Args:
        title: Section title.
        policies: [{"title": "...", "content": "...", "important": False}]
        effective_date: Optional effective date string.

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    primary = (0, 51, 102)
    accent = (0, 112, 192)
    important_bg = (255, 243, 205)
    important_border = (255, 185, 0)
    font = "Calibri"
    eff_date = effective_date or datetime.date.today().strftime("%B %d, %Y")

    # Header
    shape = doc.Shapes.AddShape(
        MSO_AUTO_SHAPE_RECTANGLE,
        Left=0, Top=0,
        Width=620, Height=60,
    )
    shape.Fill.ForeColor.RGB = rgb(*primary)
    shape.Line.Visible = False
    tf = shape.TextFrame
    tf.MarginLeft = 36
    tf.MarginTop = 12
    tr = tf.TextRange
    tr.Text = "EMPLOYEE HANDBOOK"
    tr.Font.Color = rgb(255, 255, 255)
    tr.Font.Size = 18
    tr.Font.Name = font
    tr.Font.Bold = True

    _add_empty_lines(doc, 3)

    # Section title
    _add_paragraph(doc, title,
                   font_name=font, font_size=22,
                   font_color=primary, bold=True,
                   space_before=24, space_after=8, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # Effective date
    _add_paragraph(doc, f"Effective Date: {eff_date}",
                   font_name=font, font_size=10,
                   font_color=(100, 100, 100), italic=True,
                   space_before=4, space_after=16, alignment=WD_ALIGN_LEFT)

    # Policies
    for i, policy in enumerate(policies, 1):
        pol_title = policy.get("title", f"Policy {i}")
        pol_content = policy.get("content", "")
        is_important = policy.get("important", False)

        # Policy heading
        _add_paragraph(doc, f"{i}. {pol_title}",
                       font_name=font, font_size=14,
                       font_color=primary, bold=True,
                       space_before=16, space_after=6, alignment=WD_ALIGN_LEFT)

        # Policy content
        content_para = _add_paragraph(doc, pol_content,
                                      font_name=font, font_size=11,
                                      font_color=(51, 51, 51),
                                      space_before=4, space_after=10,
                                      alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

        # Highlight important policies
        if is_important:
            content_para.Shading.BackgroundPatternColor = rgb(*important_bg)
            pf = content_para.Format
            pf.LeftIndent = 18
            pf.RightIndent = 18
            border = pf.Borders(WD_BORDER_LEFT)
            border.LineStyle = 1
            border.LineWidth = 20
            border.Color = rgb(*important_border)

            _add_paragraph(doc, "⚠ This is a mandatory policy. All employees must comply.",
                           font_name=font, font_size=9,
                           font_color=(180, 100, 0), bold=True,
                           space_before=2, space_after=8, alignment=WD_ALIGN_LEFT)

    # Acknowledgment section
    _add_empty_lines(doc, 2)
    _add_paragraph(doc, "ACKNOWLEDGMENT",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=18, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    _add_paragraph(doc, "I acknowledge that I have received, read, and understand the policies outlined above. I agree to abide by these policies.",
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=8, space_after=24,
                   alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

    # Signature
    _add_paragraph(doc, "_" * 40 + "          " + "_" * 20,
                   font_name=font, font_size=11,
                   space_before=24, space_after=2, alignment=WD_ALIGN_LEFT)
    _add_paragraph(doc, "Employee Signature                                              Date",
                   font_name=font, font_size=9,
                   font_color=(100, 100, 100),
                   space_before=2, space_after=12, alignment=WD_ALIGN_LEFT)

    _add_paragraph(doc, "_" * 40,
                   font_name=font, font_size=11,
                   space_before=18, space_after=2, alignment=WD_ALIGN_LEFT)
    _add_paragraph(doc, "Printed Name",
                   font_name=font, font_size=9,
                   font_color=(100, 100, 100),
                   space_before=2, space_after=8, alignment=WD_ALIGN_LEFT)

    return {
        "title": title,
        "policy_count": len(policies),
        "effective_date": eff_date,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 21. FAQ Document
# ---------------------------------------------------------------------------

def create_faq_document(title: str, faqs: list[dict],
                        style: str = "accordion") -> dict:
    """Create a FAQ document.

    Args:
        title: Document title.
        faqs: [{"question": "...", "answer": "..."}]
        style: accordion (Q&A blocks), numbered (numbered list), table (table format).

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    primary = (0, 70, 127)
    accent = (0, 112, 192)
    q_color = (0, 70, 127)
    a_color = (51, 51, 51)
    font = "Calibri"

    # Title
    _add_paragraph(doc, title,
                   font_name=font, font_size=24,
                   font_color=primary, bold=True,
                   space_before=24, space_after=8, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    _add_paragraph(doc, f"Last Updated: {datetime.date.today().strftime('%B %d, %Y')}",
                   font_name=font, font_size=9,
                   font_color=(130, 130, 130), italic=True,
                   space_before=4, space_after=16, alignment=WD_ALIGN_CENTER)

    if style == "accordion":
        for i, faq in enumerate(faqs, 1):
            question = faq.get("question", "")
            answer = faq.get("answer", "")

            # Question block with colored background
            q_para = _add_paragraph(doc, f"Q{i}: {question}",
                                    font_name=font, font_size=12,
                                    font_color=q_color, bold=True,
                                    space_before=12, space_after=4,
                                    alignment=WD_ALIGN_LEFT)
            q_para.Shading.BackgroundPatternColor = rgb(217, 226, 243)
            pf = q_para.Format
            pf.LeftIndent = 12
            pf.RightIndent = 12
            border = pf.Borders(WD_BORDER_LEFT)
            border.LineStyle = 1
            border.LineWidth = 16
            border.Color = rgb(*accent)

            # Answer
            _add_paragraph(doc, answer,
                           font_name=font, font_size=11,
                           font_color=a_color,
                           space_before=6, space_after=10,
                           alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

    elif style == "numbered":
        for i, faq in enumerate(faqs, 1):
            question = faq.get("question", "")
            answer = faq.get("answer", "")

            _add_paragraph(doc, f"{i}. {question}",
                           font_name=font, font_size=12,
                           font_color=q_color, bold=True,
                           space_before=12, space_after=4,
                           alignment=WD_ALIGN_LEFT)

            _add_paragraph(doc, answer,
                           font_name=font, font_size=11,
                           font_color=a_color,
                           space_before=4, space_after=8,
                           alignment=WD_ALIGN_JUSTIFY, line_spacing=1.3)

            if i < len(faqs):
                _add_horizontal_line(doc, color=(220, 220, 220), thickness=0.5)

    elif style == "table":
        faq_data = [["#", "Question", "Answer"]]
        for i, faq in enumerate(faqs, 1):
            faq_data.append([
                str(i),
                faq.get("question", ""),
                faq.get("answer", ""),
            ])

        faq_table = _create_table(doc, len(faq_data), 3, faq_data)
        _format_table_header(faq_table, fill_color=primary)
        _format_table_body(faq_table, alt_row_color=(234, 240, 250))
        try:
            faq_table.Columns(1).Width = 30
            faq_table.Columns(2).Width = 200
            faq_table.Columns(3).Width = 250
        except Exception:
            pass

    return {
        "title": title,
        "faq_count": len(faqs),
        "style": style,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 22. Checklist Document
# ---------------------------------------------------------------------------

def create_checklist_document(title: str, categories: list[dict],
                              style: str = "checkbox") -> dict:
    """Create a checklist document.

    Args:
        title: Document title.
        categories: [{"name": "...", "items": ["..."]}]
        style: checkbox (with boxes), numbered, bullet.

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    primary = (0, 70, 127)
    accent = (0, 112, 192)
    font = "Calibri"

    # Title
    _add_paragraph(doc, title,
                   font_name=font, font_size=22,
                   font_color=primary, bold=True,
                   space_before=24, space_after=8, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    _add_paragraph(doc, f"Date: {datetime.date.today().strftime('%B %d, %Y')}",
                   font_name=font, font_size=10,
                   font_color=(100, 100, 100),
                   space_before=4, space_after=16, alignment=WD_ALIGN_RIGHT)

    total_items = 0

    for cat in categories:
        cat_name = cat.get("name", "")
        items = cat.get("items", [])
        total_items += len(items)

        # Category heading
        cat_para = _add_paragraph(doc, cat_name,
                                  font_name=font, font_size=14,
                                  font_color=primary, bold=True,
                                  space_before=16, space_after=6,
                                  alignment=WD_ALIGN_LEFT)
        cat_para.Shading.BackgroundPatternColor = rgb(217, 226, 243)
        pf = cat_para.Format
        pf.LeftIndent = 8
        pf.RightIndent = 8

        # Items
        for j, item in enumerate(items, 1):
            if style == "checkbox":
                prefix = "☐  "
            elif style == "numbered":
                prefix = f"{j}.  "
            else:  # bullet
                prefix = "•  "

            _add_paragraph(doc, f"{prefix}{item}",
                           font_name=font, font_size=11,
                           font_color=(51, 51, 51),
                           space_before=3, space_after=3,
                           alignment=WD_ALIGN_LEFT, line_spacing=1.2)

    # Summary
    _add_empty_lines(doc, 1)
    _add_horizontal_line(doc, color=accent, thickness=1.0)
    _add_paragraph(doc, f"Total Items: {total_items}",
                   font_name=font, font_size=11,
                   font_color=primary, bold=True,
                   space_before=8, space_after=4, alignment=WD_ALIGN_RIGHT)

    # Completion signature
    _add_empty_lines(doc, 2)
    _add_paragraph(doc, "Completed by: _________________________     Date: _______________",
                   font_name=font, font_size=10,
                   font_color=(80, 80, 80),
                   space_before=12, space_after=8, alignment=WD_ALIGN_LEFT)

    return {
        "title": title,
        "category_count": len(categories),
        "total_items": total_items,
        "style": style,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 23. Signature Block
# ---------------------------------------------------------------------------

def add_signature_block(names: list[str], titles: list[str] | None = None,
                        date_line: bool = True, witness: bool = False) -> dict:
    """Add a professional signature block.

    Args:
        names: List of signatory names.
        titles: Optional list of titles corresponding to names.
        date_line: Whether to include a date line.
        witness: Whether to include a witness signature.

    Returns:
        dict with status info.
    """
    doc = _get_active_doc()
    font = "Calibri"
    primary = (0, 51, 102)
    line_color = (80, 80, 80)
    titles_list = titles or [None] * len(names)

    _add_empty_lines(doc, 1)
    _add_horizontal_line(doc, color=(180, 180, 180), thickness=1.0)

    for i, name in enumerate(names):
        _add_empty_lines(doc, 1)

        # Signature line
        _add_paragraph(doc, "_" * 45,
                       font_name=font, font_size=11,
                       font_color=line_color,
                       space_before=36, space_after=2, alignment=WD_ALIGN_LEFT)

        # Name
        _add_paragraph(doc, name,
                       font_name=font, font_size=11,
                       font_color=primary, bold=True,
                       space_before=2, space_after=2, alignment=WD_ALIGN_LEFT)

        # Title
        if i < len(titles_list) and titles_list[i]:
            _add_paragraph(doc, titles_list[i],
                           font_name=font, font_size=10,
                           font_color=(100, 100, 100),
                           space_before=1, space_after=2, alignment=WD_ALIGN_LEFT)

        # Date line
        if date_line:
            _add_paragraph(doc, "Date: _________________________",
                           font_name=font, font_size=10,
                           font_color=line_color,
                           space_before=4, space_after=8, alignment=WD_ALIGN_LEFT)

    # Witness block
    if witness:
        _add_empty_lines(doc, 1)
        _add_paragraph(doc, "WITNESS:",
                       font_name=font, font_size=10,
                       font_color=primary, bold=True,
                       space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)

        _add_paragraph(doc, "_" * 45,
                       font_name=font, font_size=11,
                       font_color=line_color,
                       space_before=30, space_after=2, alignment=WD_ALIGN_LEFT)

        _add_paragraph(doc, "Witness Name: _________________________",
                       font_name=font, font_size=10,
                       font_color=line_color,
                       space_before=2, space_after=2, alignment=WD_ALIGN_LEFT)

        if date_line:
            _add_paragraph(doc, "Date: _________________________",
                           font_name=font, font_size=10,
                           font_color=line_color,
                           space_before=4, space_after=8, alignment=WD_ALIGN_LEFT)

    return {
        "signatory_count": len(names),
        "has_titles": titles is not None,
        "date_line": date_line,
        "witness": witness,
    }


# ---------------------------------------------------------------------------
# 24. Executive Summary
# ---------------------------------------------------------------------------

def create_executive_summary(
    title: str,
    key_findings: list[str],
    recommendations: list[str],
    conclusion: str | None = None,
    style: str = "professional",
) -> dict:
    """Create a professional executive summary document.

    Args:
        title: Document title.
        key_findings: List of key finding strings.
        recommendations: List of recommendation strings.
        conclusion: Optional concluding paragraph.
        style: "professional" (corporate blue), "modern" (clean design),
               "minimal" (simple layout).

    Returns:
        dict with title, style, finding_count, recommendation_count, paragraphs.
    """
    doc = _get_active_doc()

    styles = {
        "professional": {"primary": (0, 51, 102), "accent": (0, 112, 192),
                         "font": "Calibri", "title_size": 26, "heading_size": 16},
        "modern": {"primary": (41, 65, 122), "accent": (68, 114, 196),
                   "font": "Segoe UI", "title_size": 28, "heading_size": 17},
        "minimal": {"primary": (50, 50, 50), "accent": (130, 130, 130),
                    "font": "Calibri Light", "title_size": 24, "heading_size": 15},
    }
    s = styles.get(style, styles["professional"])
    primary = s["primary"]
    accent = s["accent"]
    font = s["font"]

    # Title
    _add_paragraph(doc, title,
                   font_name=font, font_size=s["title_size"],
                   font_color=primary, bold=True,
                   space_before=0, space_after=4, alignment=WD_ALIGN_LEFT)

    # Accent line
    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # "EXECUTIVE SUMMARY" label
    _add_paragraph(doc, "EXECUTIVE SUMMARY",
                   font_name=font, font_size=10,
                   font_color=accent, bold=True,
                   space_before=8, space_after=16, alignment=WD_ALIGN_LEFT)

    # Date
    _add_paragraph(doc, f"Date: {datetime.date.today().strftime('%B %d, %Y')}",
                   font_name=font, font_size=10,
                   font_color=(100, 100, 100),
                   space_before=0, space_after=16, alignment=WD_ALIGN_LEFT)

    # Key Findings section
    _add_paragraph(doc, "Key Findings",
                   font_name=font, font_size=s["heading_size"],
                   font_color=primary, bold=True,
                   space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=0.5)

    for i, finding in enumerate(key_findings, 1):
        _add_paragraph(doc, f"{i}. {finding}",
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=6, space_after=4, alignment=WD_ALIGN_LEFT)

    # Recommendations section
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Recommendations",
                   font_name=font, font_size=s["heading_size"],
                   font_color=primary, bold=True,
                   space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=0.5)

    for i, rec in enumerate(recommendations, 1):
        _add_paragraph(doc, f"{i}. {rec}",
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=6, space_after=4, alignment=WD_ALIGN_LEFT)

    # Conclusion
    if conclusion:
        _add_empty_lines(doc, 1)
        _add_paragraph(doc, "Conclusion",
                       font_name=font, font_size=s["heading_size"],
                       font_color=primary, bold=True,
                       space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)

        _add_horizontal_line(doc, color=accent, thickness=0.5)

        _add_paragraph(doc, conclusion,
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=6, space_after=8, alignment=WD_ALIGN_JUSTIFY)

    return {
        "title": title,
        "style": style,
        "finding_count": len(key_findings),
        "recommendation_count": len(recommendations),
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 25. Status Report
# ---------------------------------------------------------------------------

def create_status_report(
    project_name: str,
    period: str,
    overall_status: str,
    accomplishments: list[str],
    issues: list[str],
    next_steps: list[str],
    metrics: list[dict] | None = None,
) -> dict:
    """Create a project status report.

    Args:
        project_name: Name of the project.
        period: Reporting period (e.g. "Q1 2024").
        overall_status: "on_track", "at_risk", "delayed".
        accomplishments: List of accomplishment strings.
        issues: List of issue/risk strings.
        next_steps: List of planned next steps.
        metrics: Optional list of {"name": "...", "value": "...", "target": "..."}.

    Returns:
        dict with project_name, status, period, paragraphs.
    """
    doc = _get_active_doc()
    font = "Calibri"
    primary = (0, 51, 102)
    accent = (0, 112, 192)

    status_colors = {
        "on_track": (40, 167, 69),
        "at_risk": (255, 193, 7),
        "delayed": (220, 53, 69),
    }
    status_labels = {
        "on_track": "ON TRACK",
        "at_risk": "AT RISK",
        "delayed": "DELAYED",
    }
    status_color = status_colors.get(overall_status, (100, 100, 100))
    status_label = status_labels.get(overall_status, overall_status.upper())

    # Title
    _add_paragraph(doc, "PROJECT STATUS REPORT",
                   font_name=font, font_size=24,
                   font_color=primary, bold=True,
                   space_before=0, space_after=4, alignment=WD_ALIGN_LEFT)

    _add_horizontal_line(doc, color=accent, thickness=2.0)

    # Project info table
    info_data = [
        ["Project:", project_name, "Period:", period],
        ["Status:", status_label, "Date:", datetime.date.today().strftime("%Y-%m-%d")],
    ]
    table = _create_table(doc, 2, 4, info_data)
    _format_table_body(table, font_name=font, font_size=11)

    for r in range(1, 3):
        for c in [1, 3]:
            _set_cell_text(table, r, c, table.Cell(r, c).Range.Text.strip(),
                           bold=True, font_name=font, font_size=10,
                           font_color=primary)

    # Status cell color
    table.Cell(2, 2).Range.Font.Color = rgb(*status_color)
    table.Cell(2, 2).Range.Font.Bold = True

    try:
        table.Columns(1).Width = 72
        table.Columns(2).Width = 200
        table.Columns(3).Width = 72
        table.Columns(4).Width = 150
    except Exception:
        pass

    # Accomplishments
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Accomplishments",
                   font_name=font, font_size=15,
                   font_color=primary, bold=True,
                   space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=(40, 167, 69), thickness=1.0)

    for item in accomplishments:
        _add_paragraph(doc, f"  {item}",
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=3, alignment=WD_ALIGN_LEFT)

    # Issues & Risks
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Issues & Risks",
                   font_name=font, font_size=15,
                   font_color=primary, bold=True,
                   space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=(220, 53, 69), thickness=1.0)

    for item in issues:
        _add_paragraph(doc, f"  {item}",
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=3, alignment=WD_ALIGN_LEFT)

    # Next Steps
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Next Steps",
                   font_name=font, font_size=15,
                   font_color=primary, bold=True,
                   space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    for item in next_steps:
        _add_paragraph(doc, f"  {item}",
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=4, space_after=3, alignment=WD_ALIGN_LEFT)

    # Metrics table (optional)
    if metrics:
        _add_empty_lines(doc, 1)
        _add_paragraph(doc, "Key Metrics",
                       font_name=font, font_size=15,
                       font_color=primary, bold=True,
                       space_before=16, space_after=8, alignment=WD_ALIGN_LEFT)
        _add_horizontal_line(doc, color=accent, thickness=1.0)

        m_data = [["Metric", "Value", "Target"]]
        for m in metrics:
            m_data.append([m.get("name", ""), m.get("value", ""), m.get("target", "")])

        m_table = _create_table(doc, len(m_data), 3, m_data)
        _format_table_header(m_table, fill_color=primary)
        _format_table_body(m_table, font_name=font, font_size=10.5,
                          alt_row_color=(240, 245, 250))

    return {
        "project_name": project_name,
        "status": overall_status,
        "period": period,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 26. Change Request
# ---------------------------------------------------------------------------

def create_change_request(
    title: str,
    requester: str,
    description: str,
    justification: str,
    impact: str,
    priority: str = "medium",
) -> dict:
    """Create a change request form document.

    Args:
        title: Change request title.
        requester: Name of the requester.
        description: Detailed description of the change.
        justification: Business justification.
        impact: Impact assessment.
        priority: "low", "medium", "high", "critical".

    Returns:
        dict with title, requester, priority, paragraphs.
    """
    doc = _get_active_doc()
    font = "Calibri"
    primary = (0, 51, 102)
    accent = (0, 112, 192)

    priority_colors = {
        "low": (40, 167, 69),
        "medium": (255, 193, 7),
        "high": (255, 120, 0),
        "critical": (220, 53, 69),
    }
    p_color = priority_colors.get(priority, (100, 100, 100))

    # Header
    _add_paragraph(doc, "CHANGE REQUEST",
                   font_name=font, font_size=26,
                   font_color=primary, bold=True,
                   space_before=0, space_after=4, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=accent, thickness=2.5)

    # Reference info
    cr_id = f"CR-{datetime.date.today().strftime('%Y%m%d')}-001"
    info_data = [
        ["CR ID:", cr_id, "Date:", datetime.date.today().strftime("%Y-%m-%d")],
        ["Requester:", requester, "Priority:", priority.upper()],
    ]
    table = _create_table(doc, 2, 4, info_data)
    _format_table_body(table, font_name=font, font_size=11)

    for r in range(1, 3):
        for c in [1, 3]:
            _set_cell_text(table, r, c, table.Cell(r, c).Range.Text.strip(),
                           bold=True, font_name=font, font_size=10,
                           font_color=primary)

    # Priority color
    table.Cell(2, 4).Range.Font.Color = rgb(*p_color)
    table.Cell(2, 4).Range.Font.Bold = True

    # Title
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Change Title",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_paragraph(doc, title,
                   font_name=font, font_size=12,
                   font_color=(40, 40, 40),
                   space_before=4, space_after=8, alignment=WD_ALIGN_LEFT)

    # Description
    _add_paragraph(doc, "Description",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=(200, 210, 220), thickness=0.5)
    _add_paragraph(doc, description,
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=6, space_after=8, alignment=WD_ALIGN_JUSTIFY)

    # Justification
    _add_paragraph(doc, "Business Justification",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=(200, 210, 220), thickness=0.5)
    _add_paragraph(doc, justification,
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=6, space_after=8, alignment=WD_ALIGN_JUSTIFY)

    # Impact Assessment
    _add_paragraph(doc, "Impact Assessment",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=(200, 210, 220), thickness=0.5)
    _add_paragraph(doc, impact,
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=6, space_after=8, alignment=WD_ALIGN_JUSTIFY)

    # Approval section
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Approval",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=8, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    approval_data = [
        ["Role", "Name", "Signature", "Date"],
        ["Requester", requester, "", ""],
        ["Manager", "", "", ""],
        ["Approver", "", "", ""],
    ]
    a_table = _create_table(doc, 4, 4, approval_data)
    _format_table_header(a_table, fill_color=primary)
    _format_table_body(a_table, font_name=font, font_size=10.5,
                      alt_row_color=(240, 245, 250))

    return {
        "title": title,
        "requester": requester,
        "priority": priority,
        "cr_id": cr_id,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 27. Incident Report
# ---------------------------------------------------------------------------

def create_incident_report(
    title: str,
    date: str,
    reported_by: str,
    description: str,
    root_cause: str | None = None,
    corrective_actions: list[str] | None = None,
    severity: str = "medium",
) -> dict:
    """Create an incident report document.

    Args:
        title: Incident title.
        date: Date of the incident.
        reported_by: Name of the reporter.
        description: Detailed description.
        root_cause: Optional root cause analysis.
        corrective_actions: Optional list of corrective actions.
        severity: "low", "medium", "high", "critical".

    Returns:
        dict with title, severity, reported_by, paragraphs.
    """
    doc = _get_active_doc()
    font = "Calibri"
    primary = (0, 51, 102)

    severity_colors = {
        "low": (40, 167, 69),
        "medium": (255, 193, 7),
        "high": (255, 120, 0),
        "critical": (220, 53, 69),
    }
    sev_color = severity_colors.get(severity, (100, 100, 100))

    # Header
    _add_paragraph(doc, "INCIDENT REPORT",
                   font_name=font, font_size=26,
                   font_color=(180, 30, 30), bold=True,
                   space_before=0, space_after=4, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=(180, 30, 30), thickness=2.5)

    # Incident info table
    inc_id = f"INC-{datetime.date.today().strftime('%Y%m%d')}-001"
    info_data = [
        ["Incident ID:", inc_id, "Date:", date],
        ["Reported By:", reported_by, "Severity:", severity.upper()],
        ["Status:", "Open", "Report Date:", datetime.date.today().strftime("%Y-%m-%d")],
    ]
    table = _create_table(doc, 3, 4, info_data)
    _format_table_body(table, font_name=font, font_size=11)

    for r in range(1, 4):
        for c in [1, 3]:
            _set_cell_text(table, r, c, table.Cell(r, c).Range.Text.strip(),
                           bold=True, font_name=font, font_size=10,
                           font_color=primary)

    # Severity coloring
    table.Cell(2, 4).Range.Font.Color = rgb(*sev_color)
    table.Cell(2, 4).Range.Font.Bold = True

    # Incident Title
    _add_empty_lines(doc, 1)
    _add_paragraph(doc, "Incident Title",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_paragraph(doc, title,
                   font_name=font, font_size=12,
                   font_color=(40, 40, 40), bold=True,
                   space_before=4, space_after=8, alignment=WD_ALIGN_LEFT)

    # Description
    _add_paragraph(doc, "Description",
                   font_name=font, font_size=14,
                   font_color=primary, bold=True,
                   space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=(200, 200, 210), thickness=0.5)
    _add_paragraph(doc, description,
                   font_name=font, font_size=11,
                   font_color=(51, 51, 51),
                   space_before=6, space_after=8, alignment=WD_ALIGN_JUSTIFY)

    # Root Cause
    if root_cause:
        _add_paragraph(doc, "Root Cause Analysis",
                       font_name=font, font_size=14,
                       font_color=primary, bold=True,
                       space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
        _add_horizontal_line(doc, color=(200, 200, 210), thickness=0.5)
        _add_paragraph(doc, root_cause,
                       font_name=font, font_size=11,
                       font_color=(51, 51, 51),
                       space_before=6, space_after=8, alignment=WD_ALIGN_JUSTIFY)

    # Corrective Actions
    if corrective_actions:
        _add_paragraph(doc, "Corrective Actions",
                       font_name=font, font_size=14,
                       font_color=primary, bold=True,
                       space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
        _add_horizontal_line(doc, color=(200, 200, 210), thickness=0.5)

        ca_data = [["#", "Action", "Status", "Due Date"]]
        for i, action in enumerate(corrective_actions, 1):
            ca_data.append([str(i), action, "Pending", ""])

        ca_table = _create_table(doc, len(ca_data), 4, ca_data)
        _format_table_header(ca_table, fill_color=(180, 30, 30))
        _format_table_body(ca_table, font_name=font, font_size=10.5,
                          alt_row_color=(255, 245, 245))

        try:
            ca_table.Columns(1).Width = 30
            ca_table.Columns(2).Width = 280
            ca_table.Columns(3).Width = 80
            ca_table.Columns(4).Width = 90
        except Exception:
            pass

    # Sign-off
    _add_empty_lines(doc, 2)
    _add_paragraph(doc, "Reviewed by: _________________________     Date: _______________",
                   font_name=font, font_size=10,
                   font_color=(80, 80, 80),
                   space_before=12, space_after=4, alignment=WD_ALIGN_LEFT)
    _add_paragraph(doc, "Approved by: _________________________     Date: _______________",
                   font_name=font, font_size=10,
                   font_color=(80, 80, 80),
                   space_before=4, space_after=8, alignment=WD_ALIGN_LEFT)

    return {
        "title": title,
        "severity": severity,
        "reported_by": reported_by,
        "incident_id": inc_id,
        "paragraphs": doc.Paragraphs.Count,
    }


# ---------------------------------------------------------------------------
# 28. Training Manual
# ---------------------------------------------------------------------------

def create_training_manual(
    title: str,
    modules: list[dict],
    style: str = "structured",
) -> dict:
    """Create a training manual document.

    Args:
        title: Manual title.
        modules: List of dicts with "title", "objectives" (list), "content" (str),
                 "exercises" (list of str).
        style: "structured" (formal numbered), "casual" (friendly design),
               "technical" (code-style).

    Returns:
        dict with title, style, module_count, paragraphs.
    """
    doc = _get_active_doc()

    style_configs = {
        "structured": {"primary": (0, 51, 102), "accent": (0, 112, 192),
                        "font": "Calibri", "title_size": 30, "h_size": 18,
                        "sub_size": 14},
        "casual": {"primary": (76, 175, 80), "accent": (129, 199, 132),
                   "font": "Segoe UI", "title_size": 32, "h_size": 20,
                   "sub_size": 14},
        "technical": {"primary": (50, 50, 50), "accent": (0, 150, 200),
                      "font": "Consolas", "title_size": 26, "h_size": 16,
                      "sub_size": 13},
    }
    s = style_configs.get(style, style_configs["structured"])
    primary = s["primary"]
    accent = s["accent"]
    font = s["font"]

    # Cover / Title
    _add_paragraph(doc, title,
                   font_name=font, font_size=s["title_size"],
                   font_color=primary, bold=True,
                   space_before=0, space_after=4, alignment=WD_ALIGN_CENTER)

    _add_horizontal_line(doc, color=accent, thickness=3.0)

    _add_paragraph(doc, "TRAINING MANUAL",
                   font_name=font, font_size=14,
                   font_color=accent, bold=True,
                   space_before=8, space_after=4, alignment=WD_ALIGN_CENTER)

    _add_paragraph(doc, f"Last Updated: {datetime.date.today().strftime('%B %d, %Y')}",
                   font_name=font, font_size=10,
                   font_color=(120, 120, 120),
                   space_before=4, space_after=16, alignment=WD_ALIGN_CENTER)

    # Table of Contents header
    _add_paragraph(doc, "Table of Contents",
                   font_name=font, font_size=s["h_size"],
                   font_color=primary, bold=True,
                   space_before=20, space_after=10, alignment=WD_ALIGN_LEFT)
    _add_horizontal_line(doc, color=accent, thickness=1.0)

    for i, mod in enumerate(modules, 1):
        _add_paragraph(doc, f"Module {i}: {mod.get('title', '')}",
                       font_name=font, font_size=11,
                       font_color=accent,
                       space_before=4, space_after=2, alignment=WD_ALIGN_LEFT)

    # Modules
    for i, mod in enumerate(modules, 1):
        _insert_page_break(doc)

        mod_title = mod.get("title", f"Module {i}")

        # Module header
        _add_paragraph(doc, f"MODULE {i}",
                       font_name=font, font_size=10,
                       font_color=accent, bold=True,
                       space_before=0, space_after=4, alignment=WD_ALIGN_LEFT)

        _add_paragraph(doc, mod_title,
                       font_name=font, font_size=s["h_size"],
                       font_color=primary, bold=True,
                       space_before=0, space_after=6, alignment=WD_ALIGN_LEFT)

        _add_horizontal_line(doc, color=accent, thickness=1.5)

        # Learning Objectives
        objectives = mod.get("objectives", [])
        if objectives:
            _add_paragraph(doc, "Learning Objectives",
                           font_name=font, font_size=s["sub_size"],
                           font_color=primary, bold=True,
                           space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)

            for obj in objectives:
                _add_paragraph(doc, f"  {obj}",
                               font_name=font, font_size=11,
                               font_color=(51, 51, 51),
                               space_before=3, space_after=2, alignment=WD_ALIGN_LEFT)

        # Content
        content = mod.get("content", "")
        if content:
            _add_empty_lines(doc, 1)
            _add_paragraph(doc, "Content",
                           font_name=font, font_size=s["sub_size"],
                           font_color=primary, bold=True,
                           space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
            _add_horizontal_line(doc, color=(200, 210, 220), thickness=0.5)

            # Split content into paragraphs
            for para_text in content.split("\n"):
                if para_text.strip():
                    _add_paragraph(doc, para_text.strip(),
                                   font_name=font, font_size=11,
                                   font_color=(51, 51, 51),
                                   space_before=4, space_after=4,
                                   alignment=WD_ALIGN_JUSTIFY)

        # Exercises
        exercises = mod.get("exercises", [])
        if exercises:
            _add_empty_lines(doc, 1)
            _add_paragraph(doc, "Exercises",
                           font_name=font, font_size=s["sub_size"],
                           font_color=primary, bold=True,
                           space_before=14, space_after=6, alignment=WD_ALIGN_LEFT)
            _add_horizontal_line(doc, color=(200, 210, 220), thickness=0.5)

            for j, ex in enumerate(exercises, 1):
                _add_paragraph(doc, f"Exercise {j}: {ex}",
                               font_name=font, font_size=11,
                               font_color=(51, 51, 51),
                               space_before=6, space_after=3, alignment=WD_ALIGN_LEFT)

                # Space for answers
                _add_paragraph(doc, "Answer: _______________________________________________",
                               font_name=font, font_size=10,
                               font_color=(150, 150, 150),
                               space_before=3, space_after=8, alignment=WD_ALIGN_LEFT)

    return {
        "title": title,
        "style": style,
        "module_count": len(modules),
        "paragraphs": doc.Paragraphs.Count,
    }
