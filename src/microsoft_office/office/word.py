"""Word COM automation logic for MCP server.

Provides comprehensive Word document automation including document creation,
formatting, page setup, tables, headers/footers, track changes, and export.
"""

from microsoft_office.com_utils import ensure_absolute_path, get_or_create_app, rgb

# ---------------------------------------------------------------------------
# Word enumeration constants
# ---------------------------------------------------------------------------

# Styles
WD_STYLE_NORMAL = -1
WD_STYLE_HEADING1 = -2
WD_STYLE_HEADING2 = -3
WD_STYLE_HEADING3 = -4

# Orientation
WD_ORIENT_PORTRAIT = 0
WD_ORIENT_LANDSCAPE = 1

# Paper size
WD_PAPER_A4 = 7
WD_PAPER_LETTER = 0

# Alignment
WD_ALIGN_LEFT = 0
WD_ALIGN_CENTER = 1
WD_ALIGN_RIGHT = 2
WD_ALIGN_JUSTIFY = 3

# Collapse direction
WD_COLLAPSE_START = 1
WD_COLLAPSE_END = 0

# Find/Replace
WD_REPLACE_ALL = 2

# Underline
WD_UNDERLINE_SINGLE = 1
WD_UNDERLINE_DOUBLE = 3

# Line spacing
WD_LINE_SPACE_SINGLE = 0
WD_LINE_SPACE_1PT5 = 1
WD_LINE_SPACE_DOUBLE = 2
WD_LINE_SPACE_EXACTLY = 4

# Borders
WD_BORDER_TOP = -1
WD_BORDER_LEFT = -2
WD_BORDER_BOTTOM = -3
WD_BORDER_RIGHT = -4

# Section / page breaks
WD_SECTION_BREAK_NEXT_PAGE = 2
WD_SECTION_BREAK_CONTINUOUS = 3
WD_PAGE_BREAK = 7

# Export
WD_EXPORT_PDF = 17

# Heading style lookup
_HEADING_STYLES = {1: WD_STYLE_HEADING1, 2: WD_STYLE_HEADING2, 3: WD_STYLE_HEADING3}


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_app():
    """Get or create the Word application instance."""
    return get_or_create_app("Word.Application")


# ---------------------------------------------------------------------------
# Basic document operations
# ---------------------------------------------------------------------------

def create_document(file_path: str | None = None) -> dict:
    """Create a new Word document, optionally saving to *file_path*."""
    app = _get_app()
    doc = app.Documents.Add()
    if file_path:
        doc.SaveAs2(ensure_absolute_path(file_path))
    return {"name": doc.Name, "paragraph_count": doc.Paragraphs.Count}


def open_document(file_path: str) -> dict:
    """Open an existing Word document."""
    app = _get_app()
    doc = app.Documents.Open(ensure_absolute_path(file_path))
    return {"name": doc.Name, "paragraph_count": doc.Paragraphs.Count}


def save_document(file_path: str | None = None) -> dict:
    """Save the active document, optionally to a new *file_path* via SaveAs2."""
    app = _get_app()
    doc = app.ActiveDocument
    if file_path:
        doc.SaveAs2(ensure_absolute_path(file_path))
    else:
        doc.Save()
    return {"name": doc.Name}


def close_document() -> dict:
    """Close the active document without saving."""
    app = _get_app()
    doc = app.ActiveDocument
    name = doc.Name
    doc.Close()
    return {"name": name}


def get_content() -> dict:
    """Return all paragraphs with their text and style."""
    app = _get_app()
    doc = app.ActiveDocument
    paragraphs = []
    for i in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(i)
        text = para.Range.Text.rstrip("\r")
        style_name = para.Style.NameLocal
        paragraphs.append({"number": i, "text": text, "style": style_name})
    return {
        "name": doc.Name,
        "paragraph_count": doc.Paragraphs.Count,
        "paragraphs": paragraphs,
    }


def add_paragraph(text: str, style: str | None = None) -> dict:
    """Append a paragraph at the end of the document."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    rng.InsertAfter(text + "\r")
    new_para = doc.Paragraphs(doc.Paragraphs.Count)
    if style:
        try:
            new_para.Style = style
        except Exception:
            pass
    return {"paragraph_number": doc.Paragraphs.Count, "text": text}


def add_heading(text: str, level: int = 1) -> dict:
    """Append a heading paragraph (level 1-3)."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    rng.InsertAfter(text + "\r")
    new_para = doc.Paragraphs(doc.Paragraphs.Count)
    style_id = _HEADING_STYLES.get(level, WD_STYLE_HEADING1)
    new_para.Style = style_id
    return {"paragraph_number": doc.Paragraphs.Count, "text": text, "level": level}


def edit_paragraph(paragraph_number: int, text: str) -> dict:
    """Replace the text of a specific paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_number)
    para.Range.Text = text + "\r"
    return {"paragraph_number": paragraph_number, "text": text}


def delete_paragraph(paragraph_number: int) -> dict:
    """Delete a paragraph by number."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_number)
    para.Range.Delete()
    return {"deleted": paragraph_number, "remaining": doc.Paragraphs.Count}


def insert_table(rows: int, cols: int, data: list[list[str]] | None = None) -> dict:
    """Insert a table at the end of the document with optional cell data."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
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
    return {"rows": rows, "cols": cols, "table_count": doc.Tables.Count}


# ---------------------------------------------------------------------------
# Page setup & structure
# ---------------------------------------------------------------------------

def set_page_setup(
    top_margin: float | None = None,
    bottom_margin: float | None = None,
    left_margin: float | None = None,
    right_margin: float | None = None,
    orientation: int | None = None,
    paper_size: int | None = None,
) -> dict:
    """Configure page margins (in points, 72pt = 1 inch), orientation, and paper size.

    Args:
        orientation: 0 = Portrait, 1 = Landscape.
        paper_size: 7 = A4, 0 = Letter.
    """
    app = _get_app()
    doc = app.ActiveDocument
    ps = doc.PageSetup
    if top_margin is not None:
        ps.TopMargin = top_margin
    if bottom_margin is not None:
        ps.BottomMargin = bottom_margin
    if left_margin is not None:
        ps.LeftMargin = left_margin
    if right_margin is not None:
        ps.RightMargin = right_margin
    if orientation is not None:
        ps.Orientation = orientation
    if paper_size is not None:
        ps.PaperSize = paper_size
    return {
        "top_margin": ps.TopMargin,
        "bottom_margin": ps.BottomMargin,
        "left_margin": ps.LeftMargin,
        "right_margin": ps.RightMargin,
        "orientation": ps.Orientation,
        "paper_size": ps.PaperSize,
    }


def add_header(
    text: str,
    alignment: int = 1,
    font_size: float | None = None,
    section_index: int | None = None,
) -> dict:
    """Set the header text for a section.

    Args:
        alignment: 0 = Left, 1 = Center, 2 = Right.
        section_index: 1-based section number (defaults to last section).
    """
    app = _get_app()
    doc = app.ActiveDocument
    section = doc.Sections(section_index or doc.Sections.Count)
    header = section.Headers(1)  # wdHeaderFooterPrimary
    header.Range.Text = text
    header.Range.ParagraphFormat.Alignment = alignment
    if font_size is not None:
        header.Range.Font.Size = font_size
    return {"section": section_index or doc.Sections.Count, "text": text}


def add_footer(
    text: str,
    alignment: int = 1,
    font_size: float | None = None,
    section_index: int | None = None,
) -> dict:
    """Set the footer text for a section.

    Args:
        alignment: 0 = Left, 1 = Center, 2 = Right.
        section_index: 1-based section number (defaults to last section).
    """
    app = _get_app()
    doc = app.ActiveDocument
    section = doc.Sections(section_index or doc.Sections.Count)
    footer = section.Footers(1)  # wdHeaderFooterPrimary
    footer.Range.Text = text
    footer.Range.ParagraphFormat.Alignment = alignment
    if font_size is not None:
        footer.Range.Font.Size = font_size
    return {"section": section_index or doc.Sections.Count, "text": text}


def add_page_numbers(
    alignment: int = 1,
    in_header: bool = False,
    starting_number: int | None = None,
) -> dict:
    """Add page numbers to the footer (or header).

    Args:
        alignment: 0 = Left, 1 = Center, 2 = Right.
        in_header: If True, add to header instead of footer.
        starting_number: Starting page number override.
    """
    app = _get_app()
    doc = app.ActiveDocument
    section = doc.Sections(1)
    if in_header:
        hf = section.Headers(1)
    else:
        hf = section.Footers(1)
    hf.PageNumbers.Add(alignment)
    if starting_number is not None:
        hf.PageNumbers.StartingNumber = starting_number
    return {"alignment": alignment, "in_header": in_header}


def add_table_of_contents(upper_level: int = 1, lower_level: int = 3) -> dict:
    """Insert a Table of Contents at the beginning of the document."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_START)
    toc = doc.TablesOfContents.Add(
        Range=rng,
        UseHeadingStyles=True,
        UpperHeadingLevel=upper_level,
        LowerHeadingLevel=lower_level,
    )
    return {"upper_level": upper_level, "lower_level": lower_level, "toc_count": doc.TablesOfContents.Count}


def add_section_break(break_type: int = 2, paragraph_number: int | None = None) -> dict:
    """Insert a section or page break.

    Args:
        break_type: 2 = NextPage, 3 = Continuous, 7 = PageBreak.
        paragraph_number: Insert before this paragraph (end of doc if None).
    """
    app = _get_app()
    doc = app.ActiveDocument
    if paragraph_number is not None:
        rng = doc.Paragraphs(paragraph_number).Range
        rng.Collapse(WD_COLLAPSE_START)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    rng.InsertBreak(Type=break_type)
    return {"break_type": break_type, "section_count": doc.Sections.Count}


def set_columns(
    num_columns: int = 2,
    spacing: float | None = None,
    section_index: int | None = None,
) -> dict:
    """Set the number of text columns for a section.

    Args:
        num_columns: Number of columns.
        spacing: Spacing between columns in points.
        section_index: 1-based section number (defaults to last section).
    """
    app = _get_app()
    doc = app.ActiveDocument
    section = doc.Sections(section_index or doc.Sections.Count)
    section.PageSetup.TextColumns.SetCount(num_columns)
    if spacing is not None:
        section.PageSetup.TextColumns.Spacing = spacing
    return {
        "num_columns": num_columns,
        "section": section_index or doc.Sections.Count,
    }


# ---------------------------------------------------------------------------
# Text formatting
# ---------------------------------------------------------------------------

def set_paragraph_format_extended(
    paragraph_number: int,
    font_color_rgb: tuple[int, int, int] | None = None,
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
    highlight_color: int | None = None,
) -> dict:
    """Apply extended formatting to a paragraph.

    Args:
        underline: 1 = Single, 3 = Double.
        alignment: 0 = Left, 1 = Center, 2 = Right, 3 = Justify.
        line_spacing_rule: 0 = Single, 1 = 1.5, 2 = Double, 4 = Exactly.
        highlight_color: Word wdColor highlight constant.
    """
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_number)
    rng = para.Range
    fmt = para.Format

    # Font-level properties
    if font_color_rgb is not None:
        rng.Font.Color = rgb(*font_color_rgb)
    if underline is not None:
        rng.Font.Underline = underline
    if strikethrough is not None:
        rng.Font.StrikeThrough = strikethrough
    if highlight_color is not None:
        rng.HighlightColorIndex = highlight_color

    # Paragraph-level properties
    if alignment is not None:
        fmt.Alignment = alignment
    if line_spacing_rule is not None:
        fmt.LineSpacingRule = line_spacing_rule
    if line_spacing is not None:
        fmt.LineSpacing = line_spacing
    if space_before is not None:
        fmt.SpaceBefore = space_before
    if space_after is not None:
        fmt.SpaceAfter = space_after
    if first_line_indent is not None:
        fmt.FirstLineIndent = first_line_indent
    if left_indent is not None:
        fmt.LeftIndent = left_indent
    if right_indent is not None:
        fmt.RightIndent = right_indent

    return {"paragraph_number": paragraph_number}


def set_font(
    paragraph_number: int,
    font_name: str | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    underline: int | None = None,
) -> dict:
    """Set font properties for a paragraph.

    Args:
        underline: 1 = Single, 3 = Double.
        font_color_rgb: Tuple of (R, G, B) values 0-255.
    """
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_number)
    font = para.Range.Font
    if font_name is not None:
        font.Name = font_name
    if font_size is not None:
        font.Size = font_size
    if bold is not None:
        font.Bold = bold
    if italic is not None:
        font.Italic = italic
    if font_color_rgb is not None:
        font.Color = rgb(*font_color_rgb)
    if underline is not None:
        font.Underline = underline
    return {"paragraph_number": paragraph_number}


# Keep legacy name as alias
set_paragraph_format = set_font


# ---------------------------------------------------------------------------
# Content insertion
# ---------------------------------------------------------------------------

def insert_paragraph_at(paragraph_number: int, text: str, style: str | None = None) -> dict:
    """Insert a paragraph BEFORE the specified paragraph number."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Paragraphs(paragraph_number).Range
    rng.Collapse(WD_COLLAPSE_START)
    rng.InsertBefore(text + "\r")
    new_para = doc.Paragraphs(paragraph_number)
    if style:
        try:
            new_para.Style = style
        except Exception:
            pass
    return {
        "paragraph_number": paragraph_number,
        "text": text,
        "paragraph_count": doc.Paragraphs.Count,
    }


def add_list(items: list[str], list_type: str = "bullet", level: int = 0) -> dict:
    """Add a bulleted or numbered list at the end of the document.

    Args:
        items: List of text strings.
        list_type: 'bullet' or 'number'.
        level: Nesting level (0-based).
    """
    app = _get_app()
    doc = app.ActiveDocument
    start_para = doc.Paragraphs.Count + 1
    for item in items:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
        rng.InsertAfter(item + "\r")
        para = doc.Paragraphs(doc.Paragraphs.Count)
        lf = para.Range.ListFormat
        if list_type == "number":
            lf.ApplyNumberDefault()
        else:
            lf.ApplyBulletDefault()
        if level > 0:
            lf.ListLevelNumber = level + 1
    return {
        "items_added": len(items),
        "list_type": list_type,
        "level": level,
    }


def add_hyperlink(paragraph_number: int, url: str, display_text: str | None = None) -> dict:
    """Add a hyperlink to the specified paragraph.

    Args:
        display_text: Visible link text (defaults to url).
    """
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Paragraphs(paragraph_number).Range
    doc.Hyperlinks.Add(
        Anchor=rng,
        Address=url,
        TextToDisplay=display_text or url,
    )
    return {"paragraph_number": paragraph_number, "url": url}


def add_bookmark(name: str, paragraph_number: int) -> dict:
    """Add a bookmark at the specified paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Paragraphs(paragraph_number).Range
    doc.Bookmarks.Add(Name=name, Range=rng)
    return {"name": name, "paragraph_number": paragraph_number}


def get_bookmarks() -> dict:
    """Return all bookmarks in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    bookmarks = []
    for i in range(1, doc.Bookmarks.Count + 1):
        bm = doc.Bookmarks(i)
        bookmarks.append({"name": bm.Name, "start": bm.Start, "end": bm.End})
    return {"bookmark_count": doc.Bookmarks.Count, "bookmarks": bookmarks}


def insert_image(
    file_path: str,
    width: float | None = None,
    height: float | None = None,
    paragraph_number: int | None = None,
) -> dict:
    """Insert an inline image into the document.

    Args:
        file_path: Path to the image file.
        width: Image width in points.
        height: Image height in points.
        paragraph_number: Insert at this paragraph (end of doc if None).
    """
    app = _get_app()
    doc = app.ActiveDocument
    abs_path = ensure_absolute_path(file_path)
    if paragraph_number is not None:
        rng = doc.Paragraphs(paragraph_number).Range
        rng.Collapse(WD_COLLAPSE_START)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    shape = doc.InlineShapes.AddPicture(FileName=abs_path, Range=rng)
    if width is not None:
        shape.Width = width
    if height is not None:
        shape.Height = height
    return {
        "file_path": abs_path,
        "width": shape.Width,
        "height": shape.Height,
        "inline_shape_count": doc.InlineShapes.Count,
    }


def add_text_box(
    left: float,
    top: float,
    width: float,
    height: float,
    text: str,
    font_name: str | None = None,
    font_size: float | None = None,
    fill_rgb: tuple[int, int, int] | None = None,
    border_rgb: tuple[int, int, int] | None = None,
) -> dict:
    """Add a floating text box shape to the document.

    Args:
        left, top, width, height: Position and size in points.
        fill_rgb: Background fill colour (R, G, B).
        border_rgb: Border line colour (R, G, B).
    """
    app = _get_app()
    doc = app.ActiveDocument
    # msoTextOrientationHorizontal = 1
    shape = doc.Shapes.AddTextbox(1, left, top, width, height)
    shape.TextFrame.TextRange.Text = text
    if font_name is not None:
        shape.TextFrame.TextRange.Font.Name = font_name
    if font_size is not None:
        shape.TextFrame.TextRange.Font.Size = font_size
    if fill_rgb is not None:
        shape.Fill.ForeColor.RGB = rgb(*fill_rgb)
        shape.Fill.Visible = True
    if border_rgb is not None:
        shape.Line.Color.RGB = rgb(*border_rgb)
        shape.Line.Visible = True
    return {"shape_name": shape.Name, "text": text}


def find_and_replace(
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_whole_word: bool = False,
    replace_all: bool = True,
) -> dict:
    """Find and replace text throughout the document."""
    app = _get_app()
    doc = app.ActiveDocument
    find = doc.Content.Find
    replace_mode = WD_REPLACE_ALL if replace_all else 1
    result = find.Execute(
        FindText=find_text,
        ReplaceWith=replace_text,
        MatchCase=match_case,
        MatchWholeWord=match_whole_word,
        Replace=replace_mode,
    )
    return {"find_text": find_text, "replace_text": replace_text, "executed": bool(result)}


# ---------------------------------------------------------------------------
# Table formatting
# ---------------------------------------------------------------------------

def format_table(
    table_index: int,
    header_bold: bool = True,
    header_fill_rgb: tuple[int, int, int] | None = None,
    border_style: int = 1,
    auto_fit: bool = True,
    alternating_row_color_rgb: tuple[int, int, int] | None = None,
) -> dict:
    """Format a table with header styling, borders, and optional alternating rows.

    Args:
        table_index: 1-based table index.
        header_bold: Bold the first row.
        header_fill_rgb: Background colour for the header row (R, G, B).
        border_style: Line style constant for borders.
        auto_fit: Auto-fit table to contents.
        alternating_row_color_rgb: Fill colour for even rows (R, G, B).
    """
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)

    if auto_fit:
        # wdAutoFitContent = 1
        table.AutoFitBehavior(1)

    # Borders
    table.Borders.InsideLineStyle = border_style
    table.Borders.OutsideLineStyle = border_style

    # Header row formatting
    header_row = table.Rows(1)
    if header_bold:
        header_row.Range.Font.Bold = True
    if header_fill_rgb is not None:
        for cell_idx in range(1, table.Columns.Count + 1):
            table.Cell(1, cell_idx).Shading.BackgroundPatternColor = rgb(*header_fill_rgb)

    # Alternating row colours
    if alternating_row_color_rgb is not None:
        color_val = rgb(*alternating_row_color_rgb)
        for row_idx in range(2, table.Rows.Count + 1):
            if row_idx % 2 == 0:
                for col_idx in range(1, table.Columns.Count + 1):
                    table.Cell(row_idx, col_idx).Shading.BackgroundPatternColor = color_val

    return {"table_index": table_index, "rows": table.Rows.Count, "cols": table.Columns.Count}


# ---------------------------------------------------------------------------
# Review & annotations
# ---------------------------------------------------------------------------

def add_comment(paragraph_number: int, comment_text: str) -> dict:
    """Add a comment to the specified paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Paragraphs(paragraph_number).Range
    doc.Comments.Add(Range=rng, Text=comment_text)
    return {"paragraph_number": paragraph_number, "comment_text": comment_text}


def get_comments() -> dict:
    """Return all comments in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    comments = []
    for i in range(1, doc.Comments.Count + 1):
        c = doc.Comments(i)
        comments.append({
            "index": i,
            "author": c.Author,
            "text": c.Range.Text,
            "scope_text": c.Scope.Text[:100] if c.Scope else "",
        })
    return {"comment_count": doc.Comments.Count, "comments": comments}


def add_watermark(
    text: str,
    font_size: float = 72,
    color_rgb: tuple[int, int, int] = (192, 192, 192),
    semitransparent: bool = True,
) -> dict:
    """Add a diagonal text watermark via a text effect in the header.

    Args:
        text: Watermark text (e.g., 'DRAFT', 'CONFIDENTIAL').
        font_size: Font size in points.
        color_rgb: Text colour (R, G, B).
        semitransparent: Whether the watermark is semi-transparent.
    """
    app = _get_app()
    doc = app.ActiveDocument
    section = doc.Sections(1)
    header = section.Headers(1)
    # msoTextEffect1 = 0
    shape = header.Shapes.AddTextEffect(
        PresetTextEffect=0,
        Text=text,
        FontName="Calibri",
        FontSize=font_size,
        FontBold=False,
        FontItalic=False,
        Left=0,
        Top=0,
    )
    shape.Rotation = 315
    shape.Fill.ForeColor.RGB = rgb(*color_rgb)
    if semitransparent:
        shape.Fill.Transparency = 0.5
    shape.Fill.Visible = True
    # Centre on page
    # msoAnchorCenter = 1, relative to page
    shape.Left = -999995  # wdShapeCenter
    shape.Top = -999995   # wdShapeCenter
    shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionPage
    shape.RelativeVerticalPosition = 0    # wdRelativeVerticalPositionPage
    return {"text": text, "font_size": font_size}


def set_page_borders(
    border_style: int = 1,
    color_rgb: tuple[int, int, int] | None = None,
    width: float | None = None,
) -> dict:
    """Set page borders for the active document.

    Args:
        border_style: Line style constant (e.g. 1 = single line).
        color_rgb: Border colour (R, G, B).
        width: Border line width in points.
    """
    app = _get_app()
    doc = app.ActiveDocument
    section = doc.Sections(1)
    for border_id in (WD_BORDER_TOP, WD_BORDER_LEFT, WD_BORDER_BOTTOM, WD_BORDER_RIGHT):
        border = section.Borders(border_id)
        border.LineStyle = border_style
        if color_rgb is not None:
            border.Color = rgb(*color_rgb)
        if width is not None:
            border.LineWidth = width
    return {"border_style": border_style}


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def export_to_pdf(output_path: str) -> dict:
    """Export the active document to PDF.

    Args:
        output_path: Destination file path for the PDF.
    """
    app = _get_app()
    doc = app.ActiveDocument
    abs_path = ensure_absolute_path(output_path)
    doc.ExportAsFixedFormat(abs_path, WD_EXPORT_PDF)
    return {"output_path": abs_path}


# ---------------------------------------------------------------------------
# Track changes
# ---------------------------------------------------------------------------

def enable_track_changes(enable: bool = True) -> dict:
    """Enable or disable track changes on the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    doc.TrackRevisions = enable
    return {"track_changes": enable}


def accept_all_changes() -> dict:
    """Accept all tracked revisions in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    doc.Revisions.AcceptAll()
    return {"accepted": True}


def reject_all_changes() -> dict:
    """Reject all tracked revisions in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    doc.Revisions.RejectAll()
    return {"rejected": True}
