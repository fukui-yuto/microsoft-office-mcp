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


# ---------------------------------------------------------------------------
# Style type constants
# ---------------------------------------------------------------------------

WD_STYLE_TYPE_PARAGRAPH = 1
WD_STYLE_TYPE_CHARACTER = 2
WD_STYLE_TYPE_TABLE = 3
WD_STYLE_TYPE_LIST = 4

_STYLE_TYPE_MAP = {
    "paragraph": WD_STYLE_TYPE_PARAGRAPH,
    "character": WD_STYLE_TYPE_CHARACTER,
    "table": WD_STYLE_TYPE_TABLE,
    "list": WD_STYLE_TYPE_LIST,
}

# Alignment constants (reuse existing ones for paragraph alignment)
_ALIGNMENT_MAP = {
    "left": 0,
    "center": 1,
    "right": 2,
    "justify": 3,
}

# Tab alignment constants
WD_TAB_ALIGN_LEFT = 0
WD_TAB_ALIGN_CENTER = 1
WD_TAB_ALIGN_RIGHT = 2
WD_TAB_ALIGN_DECIMAL = 3

_TAB_ALIGNMENT_MAP = {
    "left": WD_TAB_ALIGN_LEFT,
    "center": WD_TAB_ALIGN_CENTER,
    "right": WD_TAB_ALIGN_RIGHT,
    "decimal": WD_TAB_ALIGN_DECIMAL,
}

# Tab leader constants
WD_TAB_LEADER_NONE = 0
WD_TAB_LEADER_DOTS = 1
WD_TAB_LEADER_DASHES = 2
WD_TAB_LEADER_LINES = 3

_TAB_LEADER_MAP = {
    "none": WD_TAB_LEADER_NONE,
    "dots": WD_TAB_LEADER_DOTS,
    "dashes": WD_TAB_LEADER_DASHES,
    "line": WD_TAB_LEADER_LINES,
}

# Content control type constants
WD_CONTENT_CONTROL_RICH_TEXT = 0
WD_CONTENT_CONTROL_TEXT = 1
WD_CONTENT_CONTROL_COMBO_BOX = 3
WD_CONTENT_CONTROL_DROP_DOWN = 4
WD_CONTENT_CONTROL_DATE = 6
WD_CONTENT_CONTROL_CHECKBOX = 8

_CONTENT_CONTROL_MAP = {
    "rich_text": WD_CONTENT_CONTROL_RICH_TEXT,
    "plain_text": WD_CONTENT_CONTROL_TEXT,
    "combo_box": WD_CONTENT_CONTROL_COMBO_BOX,
    "drop_down": WD_CONTENT_CONTROL_DROP_DOWN,
    "date_picker": WD_CONTENT_CONTROL_DATE,
    "checkbox": WD_CONTENT_CONTROL_CHECKBOX,
}

# Protection type constants
WD_PROTECT_READ_ONLY = 3
WD_PROTECT_COMMENTS = 1
WD_PROTECT_FORMS = 2
WD_PROTECT_TRACKED_CHANGES = 0

_PROTECTION_MAP = {
    "read_only": WD_PROTECT_READ_ONLY,
    "comments": WD_PROTECT_COMMENTS,
    "forms": WD_PROTECT_FORMS,
    "tracked_changes": WD_PROTECT_TRACKED_CHANGES,
}

# Cross-reference type constants
WD_REF_TYPE_HEADING = 1
WD_REF_TYPE_BOOKMARK = 2
WD_REF_TYPE_FOOTNOTE = 5

_REF_TYPE_MAP = {
    "heading": WD_REF_TYPE_HEADING,
    "bookmark": WD_REF_TYPE_BOOKMARK,
    "footnote": WD_REF_TYPE_FOOTNOTE,
}

# Vertical alignment for table cells
WD_CELL_ALIGN_TOP = 0
WD_CELL_ALIGN_CENTER = 1
WD_CELL_ALIGN_BOTTOM = 3

_VERTICAL_ALIGNMENT_MAP = {
    "top": WD_CELL_ALIGN_TOP,
    "center": WD_CELL_ALIGN_CENTER,
    "bottom": WD_CELL_ALIGN_BOTTOM,
}

# Table width type constants
WD_TABLE_WIDTH_AUTO = 1
WD_TABLE_WIDTH_FIXED = 2  # wdPreferredWidthPoints
WD_TABLE_WIDTH_PERCENT = 3  # wdPreferredWidthPercent


# ---------------------------------------------------------------------------
# 1. Style Management
# ---------------------------------------------------------------------------

def create_style(
    name: str,
    style_type: str = "paragraph",
    base_style: str | None = None,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
    alignment: str | None = None,
) -> dict:
    """Create a custom style in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    wdtype = _STYLE_TYPE_MAP.get(style_type, WD_STYLE_TYPE_PARAGRAPH)
    style = doc.Styles.Add(Name=name, Type=wdtype)
    if base_style is not None:
        style.BaseStyle = doc.Styles(base_style)
    if font_name is not None:
        style.Font.Name = font_name
    if font_size is not None:
        style.Font.Size = font_size
    if font_color_rgb is not None:
        style.Font.Color = rgb(*font_color_rgb)
    if bold is not None:
        style.Font.Bold = bold
    if italic is not None:
        style.Font.Italic = italic
    if wdtype == WD_STYLE_TYPE_PARAGRAPH:
        if space_before is not None:
            style.ParagraphFormat.SpaceBefore = space_before
        if space_after is not None:
            style.ParagraphFormat.SpaceAfter = space_after
        if line_spacing is not None:
            style.ParagraphFormat.LineSpacing = line_spacing
        if alignment is not None:
            style.ParagraphFormat.Alignment = _ALIGNMENT_MAP.get(alignment, 0)
    return {"name": name, "style_type": style_type}


def apply_style(paragraph_index: int, style_name: str) -> dict:
    """Apply a style to a paragraph by name."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Style = style_name
    return {"paragraph_index": paragraph_index, "style_name": style_name}


def get_styles() -> dict:
    """List all available styles in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    styles = []
    for i in range(1, doc.Styles.Count + 1):
        try:
            s = doc.Styles(i)
            styles.append({
                "name": s.NameLocal,
                "type": s.Type,
                "built_in": s.BuiltIn,
            })
        except Exception:
            pass
    return {"style_count": len(styles), "styles": styles}


def modify_style(
    style_name: str,
    font_name: str | None = None,
    font_size: float | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    space_before: float | None = None,
    space_after: float | None = None,
    line_spacing: float | None = None,
) -> dict:
    """Modify an existing style."""
    app = _get_app()
    doc = app.ActiveDocument
    style = doc.Styles(style_name)
    if font_name is not None:
        style.Font.Name = font_name
    if font_size is not None:
        style.Font.Size = font_size
    if font_color_rgb is not None:
        style.Font.Color = rgb(*font_color_rgb)
    if bold is not None:
        style.Font.Bold = bold
    if italic is not None:
        style.Font.Italic = italic
    try:
        if space_before is not None:
            style.ParagraphFormat.SpaceBefore = space_before
        if space_after is not None:
            style.ParagraphFormat.SpaceAfter = space_after
        if line_spacing is not None:
            style.ParagraphFormat.LineSpacing = line_spacing
    except Exception:
        pass  # Character/table styles may not have ParagraphFormat
    return {"style_name": style_name}


# ---------------------------------------------------------------------------
# 2. Footnotes & Endnotes
# ---------------------------------------------------------------------------

def add_footnote(paragraph_index: int, text: str) -> dict:
    """Add a footnote at the end of a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Paragraphs(paragraph_index).Range
    rng.Collapse(WD_COLLAPSE_END)
    # Move back one character to stay inside the paragraph (before \r)
    rng.MoveEnd(1, -1)
    fn = doc.Footnotes.Add(Range=rng, Text=text)
    return {"paragraph_index": paragraph_index, "text": text, "footnote_count": doc.Footnotes.Count}


def add_endnote(paragraph_index: int, text: str) -> dict:
    """Add an endnote at the end of a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Paragraphs(paragraph_index).Range
    rng.Collapse(WD_COLLAPSE_END)
    rng.MoveEnd(1, -1)
    en = doc.Endnotes.Add(Range=rng, Text=text)
    return {"paragraph_index": paragraph_index, "text": text, "endnote_count": doc.Endnotes.Count}


def get_footnotes() -> dict:
    """List all footnotes in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    footnotes = []
    for i in range(1, doc.Footnotes.Count + 1):
        fn = doc.Footnotes(i)
        footnotes.append({"index": i, "text": fn.Range.Text})
    return {"footnote_count": doc.Footnotes.Count, "footnotes": footnotes}


def get_endnotes() -> dict:
    """List all endnotes in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    endnotes = []
    for i in range(1, doc.Endnotes.Count + 1):
        en = doc.Endnotes(i)
        endnotes.append({"index": i, "text": en.Range.Text})
    return {"endnote_count": doc.Endnotes.Count, "endnotes": endnotes}


# ---------------------------------------------------------------------------
# 3. Advanced Table Operations
# ---------------------------------------------------------------------------

def merge_table_cells(
    table_index: int,
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
) -> dict:
    """Merge a range of cells in a table."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    cell_start = table.Cell(start_row, start_col)
    cell_end = table.Cell(end_row, end_col)
    cell_start.Merge(cell_end)
    return {
        "table_index": table_index,
        "merged": f"({start_row},{start_col})-({end_row},{end_col})",
    }


def split_table_cell(
    table_index: int,
    row: int,
    col: int,
    num_rows: int,
    num_cols: int,
) -> dict:
    """Split a table cell into multiple rows and columns."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    cell = table.Cell(row, col)
    cell.Split(NumRows=num_rows, NumColumns=num_cols)
    return {
        "table_index": table_index,
        "cell": f"({row},{col})",
        "split_into": f"{num_rows}x{num_cols}",
    }


def format_table_cell(
    table_index: int,
    row: int,
    col: int,
    fill_color_rgb: tuple[int, int, int] | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    font_size: float | None = None,
    bold: bool | None = None,
    alignment: int | None = None,
    vertical_alignment: str | None = None,
) -> dict:
    """Format an individual table cell."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    cell = table.Cell(row, col)
    if fill_color_rgb is not None:
        cell.Shading.BackgroundPatternColor = rgb(*fill_color_rgb)
    if font_color_rgb is not None:
        cell.Range.Font.Color = rgb(*font_color_rgb)
    if font_size is not None:
        cell.Range.Font.Size = font_size
    if bold is not None:
        cell.Range.Font.Bold = bold
    if alignment is not None:
        cell.Range.ParagraphFormat.Alignment = alignment
    if vertical_alignment is not None:
        cell.VerticalAlignment = _VERTICAL_ALIGNMENT_MAP.get(vertical_alignment, 0)
    return {"table_index": table_index, "cell": f"({row},{col})"}


def set_table_cell_borders(
    table_index: int,
    row: int,
    col: int,
    border_type: str,
    color_rgb: tuple[int, int, int] | None = None,
    weight: float | None = None,
    style: int | None = None,
) -> dict:
    """Set borders on a table cell. border_type: 'top','bottom','left','right','all'."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    cell = table.Cell(row, col)

    border_map = {
        "top": [WD_BORDER_TOP],
        "bottom": [WD_BORDER_BOTTOM],
        "left": [WD_BORDER_LEFT],
        "right": [WD_BORDER_RIGHT],
        "all": [WD_BORDER_TOP, WD_BORDER_BOTTOM, WD_BORDER_LEFT, WD_BORDER_RIGHT],
    }
    border_ids = border_map.get(border_type, [WD_BORDER_TOP])
    for bid in border_ids:
        border = cell.Borders(bid)
        if style is not None:
            border.LineStyle = style
        else:
            border.LineStyle = 1  # single line default
        if color_rgb is not None:
            border.Color = rgb(*color_rgb)
        if weight is not None:
            border.LineWidth = weight
    return {"table_index": table_index, "cell": f"({row},{col})", "border_type": border_type}


def set_table_width(
    table_index: int,
    width_type: str,
    width: float | None = None,
) -> dict:
    """Set table width. width_type: 'auto','fixed','percent'."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    if width_type == "auto":
        table.AutoFitBehavior(1)  # wdAutoFitContent
    elif width_type == "fixed":
        table.AutoFitBehavior(0)  # wdAutoFitFixed
        if width is not None:
            table.PreferredWidthType = WD_TABLE_WIDTH_FIXED
            table.PreferredWidth = width
    elif width_type == "percent":
        if width is not None:
            table.PreferredWidthType = WD_TABLE_WIDTH_PERCENT
            table.PreferredWidth = width
    return {"table_index": table_index, "width_type": width_type}


# ---------------------------------------------------------------------------
# 4. Shapes & Drawing
# ---------------------------------------------------------------------------

def add_shape(
    shape_type: int,
    left: float,
    top: float,
    width: float,
    height: float,
    fill_color_rgb: tuple[int, int, int] | None = None,
    line_color_rgb: tuple[int, int, int] | None = None,
) -> dict:
    """Add an AutoShape to the document. shape_type: msoAutoShapeType constant (1=rect, 5=roundrect, 9=oval, etc.)."""
    app = _get_app()
    doc = app.ActiveDocument
    shape = doc.Shapes.AddShape(shape_type, left, top, width, height)
    if fill_color_rgb is not None:
        shape.Fill.ForeColor.RGB = rgb(*fill_color_rgb)
        shape.Fill.Visible = True
    if line_color_rgb is not None:
        shape.Line.Color.RGB = rgb(*line_color_rgb)
        shape.Line.Visible = True
    return {"shape_name": shape.Name, "shape_type": shape_type}


def add_line(
    start_x: float,
    start_y: float,
    end_x: float,
    end_y: float,
    color_rgb: tuple[int, int, int] | None = None,
    weight: float = 1.0,
    dash_style: int | None = None,
) -> dict:
    """Draw a line on the document."""
    app = _get_app()
    doc = app.ActiveDocument
    shape = doc.Shapes.AddLine(start_x, start_y, end_x, end_y)
    shape.Line.Weight = weight
    if color_rgb is not None:
        shape.Line.Color.RGB = rgb(*color_rgb)
    if dash_style is not None:
        shape.Line.DashStyle = dash_style
    return {"shape_name": shape.Name}


# ---------------------------------------------------------------------------
# 5. Text Enhancements
# ---------------------------------------------------------------------------

def add_drop_cap(
    paragraph_index: int,
    lines_to_drop: int = 3,
    font_name: str | None = None,
) -> dict:
    """Add a drop cap effect to a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    dc = para.DropCap
    dc.Position = 1  # wdDropNormal
    dc.LinesToDrop = lines_to_drop
    if font_name is not None:
        dc.FontName = font_name
    return {"paragraph_index": paragraph_index, "lines_to_drop": lines_to_drop}


def set_text_highlight(paragraph_index: int, color_index: int) -> dict:
    """Set highlight color on a paragraph. color_index: WdColorIndex 1-16."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Range.HighlightColorIndex = color_index
    return {"paragraph_index": paragraph_index, "color_index": color_index}


def set_character_spacing(
    paragraph_index: int,
    spacing: float = 0,
    kerning: float | None = None,
    scale: int = 100,
) -> dict:
    """Set character spacing for a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Range.Font.Spacing = spacing
    para.Range.Font.Scaling = scale
    if kerning is not None:
        para.Range.Font.Kerning = kerning
    return {"paragraph_index": paragraph_index, "spacing": spacing, "scale": scale}


def add_text_effect(paragraph_index: int, effect_type: str) -> dict:
    """Add text effect to a paragraph. effect_type: 'shadow','outline','emboss','engrave'."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    font = para.Range.Font
    if effect_type == "shadow":
        font.Shadow = True
    elif effect_type == "outline":
        font.Outline = True
    elif effect_type == "emboss":
        font.Emboss = True
    elif effect_type == "engrave":
        font.Engrave = True
    return {"paragraph_index": paragraph_index, "effect_type": effect_type}


# ---------------------------------------------------------------------------
# 6. Page Layout
# ---------------------------------------------------------------------------

def insert_page_break(paragraph_index: int | None = None) -> dict:
    """Insert a page break after a paragraph (or at the end of the document)."""
    app = _get_app()
    doc = app.ActiveDocument
    if paragraph_index is not None:
        rng = doc.Paragraphs(paragraph_index).Range
        rng.Collapse(WD_COLLAPSE_END)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    rng.InsertBreak(Type=WD_PAGE_BREAK)
    return {"paragraph_index": paragraph_index}


def insert_column_break(paragraph_index: int | None = None) -> dict:
    """Insert a column break after a paragraph (or at the end)."""
    app = _get_app()
    doc = app.ActiveDocument
    if paragraph_index is not None:
        rng = doc.Paragraphs(paragraph_index).Range
        rng.Collapse(WD_COLLAPSE_END)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    # wdColumnBreak = 8
    rng.InsertBreak(Type=8)
    return {"paragraph_index": paragraph_index}


def set_paragraph_borders(
    paragraph_index: int,
    border_type: str,
    color_rgb: tuple[int, int, int] | None = None,
    weight: float | None = None,
    style: int | None = None,
) -> dict:
    """Set borders on a paragraph. border_type: 'top','bottom','left','right','box'."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)

    border_map = {
        "top": [WD_BORDER_TOP],
        "bottom": [WD_BORDER_BOTTOM],
        "left": [WD_BORDER_LEFT],
        "right": [WD_BORDER_RIGHT],
        "box": [WD_BORDER_TOP, WD_BORDER_BOTTOM, WD_BORDER_LEFT, WD_BORDER_RIGHT],
    }
    border_ids = border_map.get(border_type, [WD_BORDER_TOP])
    for bid in border_ids:
        border = para.Borders(bid)
        if style is not None:
            border.LineStyle = style
        else:
            border.LineStyle = 1
        if color_rgb is not None:
            border.Color = rgb(*color_rgb)
        if weight is not None:
            border.LineWidth = weight
    return {"paragraph_index": paragraph_index, "border_type": border_type}


def set_paragraph_shading(paragraph_index: int, color_rgb: tuple[int, int, int]) -> dict:
    """Set paragraph background shading color."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Shading.BackgroundPatternColor = rgb(*color_rgb)
    return {"paragraph_index": paragraph_index}


def add_horizontal_line(paragraph_index: int | None = None) -> dict:
    """Add a decorative horizontal line."""
    app = _get_app()
    doc = app.ActiveDocument
    if paragraph_index is not None:
        rng = doc.Paragraphs(paragraph_index).Range
        rng.Collapse(WD_COLLAPSE_END)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    rng.InsertAfter("\r")
    # Apply a bottom border to act as a horizontal line
    para = doc.Paragraphs(doc.Paragraphs.Count)
    border = para.Borders(WD_BORDER_BOTTOM)
    border.LineStyle = 1
    border.LineWidth = 2  # wdLineWidth150pt
    return {"paragraph_index": doc.Paragraphs.Count}


# ---------------------------------------------------------------------------
# 7. Document Properties
# ---------------------------------------------------------------------------

def set_document_properties(
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    category: str | None = None,
    comments: str | None = None,
) -> dict:
    """Set document metadata properties."""
    app = _get_app()
    doc = app.ActiveDocument
    props = doc.BuiltInDocumentProperties
    if title is not None:
        props("Title").Value = title
    if author is not None:
        props("Author").Value = author
    if subject is not None:
        props("Subject").Value = subject
    if keywords is not None:
        props("Keywords").Value = keywords
    if category is not None:
        props("Category").Value = category
    if comments is not None:
        props("Comments").Value = comments
    return {"updated": True}


def get_document_properties() -> dict:
    """Get document metadata properties."""
    app = _get_app()
    doc = app.ActiveDocument
    props = doc.BuiltInDocumentProperties
    result = {}
    for prop_name in ("Title", "Author", "Subject", "Keywords", "Category", "Comments"):
        try:
            result[prop_name.lower()] = str(props(prop_name).Value)
        except Exception:
            result[prop_name.lower()] = ""
    return result


def get_document_statistics() -> dict:
    """Get document statistics: word count, page count, character count."""
    app = _get_app()
    doc = app.ActiveDocument
    # Repaginate to get accurate page count
    doc.Repaginate()
    stats = doc.ComputeStatistics
    # wdStatisticWords = 0, wdStatisticPages = 2, wdStatisticCharacters = 3
    # wdStatisticParagraphs = 4, wdStatisticLines = 1
    return {
        "word_count": stats(0),
        "line_count": stats(1),
        "page_count": stats(2),
        "character_count": stats(3),
        "paragraph_count": stats(4),
    }


# ---------------------------------------------------------------------------
# 8. Advanced Features
# ---------------------------------------------------------------------------

def add_cross_reference(
    ref_type: str,
    ref_item: int,
    ref_format: int | None = None,
) -> dict:
    """Add a cross-reference. ref_type: 'heading','bookmark','footnote'. ref_item: 1-based index."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    wdtype = _REF_TYPE_MAP.get(ref_type, WD_REF_TYPE_HEADING)
    # Default ref format: wdContentText = 0
    fmt = ref_format if ref_format is not None else 0
    doc.Fields.Add(
        Range=rng,
        Type=-1,  # wdFieldEmpty
        Text=f" REF _Ref{ref_item} \\h ",
        PreserveFormatting=True,
    )
    return {"ref_type": ref_type, "ref_item": ref_item}


def insert_field(
    field_code: str,
    paragraph_index: int | None = None,
) -> dict:
    """Insert a Word field (e.g. 'DATE', 'PAGE', 'NUMPAGES', 'AUTHOR', 'TOC')."""
    app = _get_app()
    doc = app.ActiveDocument
    if paragraph_index is not None:
        rng = doc.Paragraphs(paragraph_index).Range
        rng.Collapse(WD_COLLAPSE_END)
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    doc.Fields.Add(
        Range=rng,
        Type=-1,  # wdFieldEmpty
        Text=f" {field_code} ",
        PreserveFormatting=True,
    )
    return {"field_code": field_code, "field_count": doc.Fields.Count}


def set_tab_stops(
    paragraph_index: int,
    positions: list[float],
    alignments: list[str] | None = None,
    leaders: list[str] | None = None,
) -> dict:
    """Set tab stops on a paragraph. positions in points."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    tabs = para.Format.TabStops
    tabs.ClearAll()
    for idx, pos in enumerate(positions):
        align = WD_TAB_ALIGN_LEFT
        leader = WD_TAB_LEADER_NONE
        if alignments and idx < len(alignments):
            align = _TAB_ALIGNMENT_MAP.get(alignments[idx], WD_TAB_ALIGN_LEFT)
        if leaders and idx < len(leaders):
            leader = _TAB_LEADER_MAP.get(leaders[idx], WD_TAB_LEADER_NONE)
        tabs.Add(Position=pos, Alignment=align, Leader=leader)
    return {"paragraph_index": paragraph_index, "tab_count": len(positions)}


def protect_document(
    password: str | None = None,
    protection_type: str = "read_only",
) -> dict:
    """Protect a document."""
    app = _get_app()
    doc = app.ActiveDocument
    ptype = _PROTECTION_MAP.get(protection_type, WD_PROTECT_READ_ONLY)
    if password:
        doc.Protect(Type=ptype, Password=password)
    else:
        doc.Protect(Type=ptype)
    return {"protection_type": protection_type}


def unprotect_document(password: str | None = None) -> dict:
    """Unprotect a document."""
    app = _get_app()
    doc = app.ActiveDocument
    if password:
        doc.Unprotect(Password=password)
    else:
        doc.Unprotect()
    return {"unprotected": True}


# ---------------------------------------------------------------------------
# 9. Content Controls & Templates
# ---------------------------------------------------------------------------

def add_content_control(
    control_type: str,
    paragraph_index: int | None = None,
    title: str | None = None,
    placeholder_text: str | None = None,
) -> dict:
    """Add a content control. control_type: 'rich_text','plain_text','combo_box','drop_down','date_picker','checkbox'."""
    app = _get_app()
    doc = app.ActiveDocument
    if paragraph_index is not None:
        rng = doc.Paragraphs(paragraph_index).Range
    else:
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
    wdtype = _CONTENT_CONTROL_MAP.get(control_type, WD_CONTENT_CONTROL_RICH_TEXT)
    cc = doc.ContentControls.Add(wdtype, rng)
    if title is not None:
        cc.Title = title
    if placeholder_text is not None:
        cc.SetPlaceholderText(Text=placeholder_text)
    return {"control_type": control_type, "title": title or ""}


def insert_building_block(
    name: str,
    category: str | None = None,
) -> dict:
    """Insert a building block (Quick Part) by name."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    templates = app.Templates
    for t_idx in range(1, templates.Count + 1):
        tmpl = templates(t_idx)
        try:
            bbs = tmpl.BuildingBlockEntries
            bb = bbs.Item(name)
            bb.Insert(Where=rng, RichText=True)
            return {"name": name, "inserted": True}
        except Exception:
            continue
    return {"name": name, "inserted": False, "error": "Building block not found"}


# ---------------------------------------------------------------------------
# 10. Lists Advanced
# ---------------------------------------------------------------------------

def set_list_level(paragraph_index: int, level: int) -> dict:
    """Set list indentation level (0-8) for a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Range.ListFormat.ListLevelNumber = level + 1  # 1-based in COM
    return {"paragraph_index": paragraph_index, "level": level}


def restart_list_numbering(paragraph_index: int) -> dict:
    """Restart list numbering at a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Range.ListFormat.ListValue = 1
    # Use ListFormat to restart
    lf = para.Range.ListFormat
    lf.ApplyNumberDefault()
    # Set the restart property
    try:
        para.Range.ListFormat.List.ApplyListTemplate(
            ListTemplate=para.Range.ListFormat.ListTemplate,
            ContinuePreviousList=False,
        )
    except Exception:
        pass
    return {"paragraph_index": paragraph_index}


# ---------------------------------------------------------------------------
# Underline style constants
# ---------------------------------------------------------------------------

WD_UNDERLINE_NONE = 0
# WD_UNDERLINE_SINGLE = 1  # already defined above
# WD_UNDERLINE_DOUBLE = 3  # already defined above
WD_UNDERLINE_DOTTED = 4
WD_UNDERLINE_DASHED = 7  # wdUnderlineDash
WD_UNDERLINE_WAVY = 11
WD_UNDERLINE_THICK = 6

_UNDERLINE_STYLE_MAP = {
    "none": WD_UNDERLINE_NONE,
    "single": WD_UNDERLINE_SINGLE,
    "double": WD_UNDERLINE_DOUBLE,
    "dotted": WD_UNDERLINE_DOTTED,
    "dashed": WD_UNDERLINE_DASHED,
    "wavy": WD_UNDERLINE_WAVY,
    "thick": WD_UNDERLINE_THICK,
}

# Row height rule constants
WD_ROW_HEIGHT_EXACT = 2
WD_ROW_HEIGHT_AT_LEAST = 1
WD_ROW_HEIGHT_AUTO = 0

_ROW_HEIGHT_RULE_MAP = {
    "exact": WD_ROW_HEIGHT_EXACT,
    "at_least": WD_ROW_HEIGHT_AT_LEAST,
    "auto": WD_ROW_HEIGHT_AUTO,
}

# Table alignment constants
WD_TABLE_ALIGN_LEFT = 0
WD_TABLE_ALIGN_CENTER = 1
WD_TABLE_ALIGN_RIGHT = 2

_TABLE_ALIGNMENT_MAP = {
    "left": WD_TABLE_ALIGN_LEFT,
    "center": WD_TABLE_ALIGN_CENTER,
    "right": WD_TABLE_ALIGN_RIGHT,
}


# ---------------------------------------------------------------------------
# 11. Advanced Text Formatting
# ---------------------------------------------------------------------------

def set_text_color_range(
    paragraph_index: int,
    start_char: int,
    end_char: int,
    color: tuple[int, int, int],
) -> dict:
    """Color specific characters in a paragraph. start_char/end_char are 0-based offsets."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    rng = para.Range
    para_start = rng.Start
    sub_rng = doc.Range(para_start + start_char, para_start + end_char)
    sub_rng.Font.Color = rgb(*color)
    return {
        "paragraph_index": paragraph_index,
        "start_char": start_char,
        "end_char": end_char,
    }


def set_text_size_range(
    paragraph_index: int,
    start_char: int,
    end_char: int,
    font_size: float,
) -> dict:
    """Set font size for specific characters in a paragraph. start_char/end_char are 0-based offsets."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    rng = para.Range
    para_start = rng.Start
    sub_rng = doc.Range(para_start + start_char, para_start + end_char)
    sub_rng.Font.Size = font_size
    return {
        "paragraph_index": paragraph_index,
        "start_char": start_char,
        "end_char": end_char,
        "font_size": font_size,
    }


def add_strikethrough(paragraph_index: int, double: bool = False) -> dict:
    """Add strikethrough to a paragraph. double=True for double strikethrough."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    if double:
        para.Range.Font.DoubleStrikeThrough = True
    else:
        para.Range.Font.StrikeThrough = True
    return {"paragraph_index": paragraph_index, "double": double}


def set_underline_style(paragraph_index: int, style: str) -> dict:
    """Set underline style for a paragraph. style: 'single','double','dotted','dashed','wavy','thick','none'."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    ul_val = _UNDERLINE_STYLE_MAP.get(style, WD_UNDERLINE_SINGLE)
    para.Range.Font.Underline = ul_val
    return {"paragraph_index": paragraph_index, "style": style}


def add_small_caps(paragraph_index: int) -> dict:
    """Set small caps on a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Range.Font.SmallCaps = True
    return {"paragraph_index": paragraph_index}


def add_all_caps(paragraph_index: int) -> dict:
    """Set all caps on a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Range.Font.AllCaps = True
    return {"paragraph_index": paragraph_index}


# ---------------------------------------------------------------------------
# 12. Advanced Paragraph
# ---------------------------------------------------------------------------

def set_keep_with_next(paragraph_index: int, keep: bool = True) -> dict:
    """Keep paragraph with the next paragraph (no page break between)."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Format.KeepWithNext = keep
    return {"paragraph_index": paragraph_index, "keep": keep}


def set_keep_together(paragraph_index: int, keep: bool = True) -> dict:
    """Keep paragraph together (no break in middle)."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Format.KeepTogether = keep
    return {"paragraph_index": paragraph_index, "keep": keep}


def set_page_break_before(paragraph_index: int, break_before: bool = True) -> dict:
    """Set page break before a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Format.PageBreakBefore = break_before
    return {"paragraph_index": paragraph_index, "break_before": break_before}


def set_widow_orphan_control(paragraph_index: int, control: bool = True) -> dict:
    """Set widow/orphan control for a paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    para.Format.WidowControl = control
    return {"paragraph_index": paragraph_index, "control": control}


def set_outline_level(paragraph_index: int, level: int) -> dict:
    """Set outline level for a paragraph. 0=body text, 1-9=outline levels."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    # wdOutlineLevelBodyText = 10, wdOutlineLevel1 = 1, etc.
    if level == 0:
        para.Format.OutlineLevel = 10  # wdOutlineLevelBodyText
    else:
        para.Format.OutlineLevel = level
    return {"paragraph_index": paragraph_index, "level": level}


def get_paragraph_count() -> dict:
    """Get total paragraph count of the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    return {"paragraph_count": doc.Paragraphs.Count}


# ---------------------------------------------------------------------------
# 13. Table Advanced
# ---------------------------------------------------------------------------

def set_table_row_height(
    table_index: int,
    row: int,
    height: float,
    rule: str = "exact",
) -> dict:
    """Set row height in a table. rule: 'exact','at_least','auto'."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    r = table.Rows(row)
    r.HeightRule = _ROW_HEIGHT_RULE_MAP.get(rule, WD_ROW_HEIGHT_EXACT)
    r.Height = height
    return {"table_index": table_index, "row": row, "height": height, "rule": rule}


def set_table_cell_width(
    table_index: int,
    row: int,
    col: int,
    width: float,
) -> dict:
    """Set individual cell width in points."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    cell = table.Cell(row, col)
    cell.Width = width
    return {"table_index": table_index, "row": row, "col": col, "width": width}


def set_table_alignment(table_index: int, alignment: str) -> dict:
    """Set table alignment. alignment: 'left','center','right'."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    table.Rows.Alignment = _TABLE_ALIGNMENT_MAP.get(alignment, WD_TABLE_ALIGN_LEFT)
    return {"table_index": table_index, "alignment": alignment}


def add_table_row(
    table_index: int,
    position: int | None = None,
    values: list[str] | None = None,
) -> dict:
    """Add a row to a table. position is 1-based row index (None=end)."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    if position is not None:
        before_row = table.Rows(position)
        table.Rows.Add(BeforeRow=before_row)
        new_row_idx = position
    else:
        table.Rows.Add()
        new_row_idx = table.Rows.Count
    if values:
        for ci, val in enumerate(values):
            if ci < table.Columns.Count:
                table.Cell(new_row_idx, ci + 1).Range.Text = val
    return {
        "table_index": table_index,
        "row": new_row_idx,
        "row_count": table.Rows.Count,
    }


def delete_table_row(table_index: int, row: int) -> dict:
    """Delete a row from a table. row is 1-based."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    table.Rows(row).Delete()
    return {"table_index": table_index, "deleted_row": row, "row_count": table.Rows.Count}


def add_table_column(table_index: int, position: int | None = None) -> dict:
    """Add a column to a table. position is 1-based column index (None=end)."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    if position is not None:
        before_col = table.Columns(position)
        table.Columns.Add(BeforeColumn=before_col)
    else:
        table.Columns.Add()
    return {"table_index": table_index, "col_count": table.Columns.Count}


def delete_table_column(table_index: int, col: int) -> dict:
    """Delete a column from a table. col is 1-based."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    table.Columns(col).Delete()
    return {"table_index": table_index, "deleted_col": col, "col_count": table.Columns.Count}


def get_table_data(table_index: int) -> dict:
    """Read all table data as a 2D array."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    rows = table.Rows.Count
    cols = table.Columns.Count
    data = []
    for r in range(1, rows + 1):
        row_data = []
        for c in range(1, cols + 1):
            try:
                text = table.Cell(r, c).Range.Text
                # Remove trailing \r\x07 that Word appends to cell text
                text = text.rstrip("\r\x07")
                row_data.append(text)
            except Exception:
                row_data.append("")
        data.append(row_data)
    return {"table_index": table_index, "rows": rows, "cols": cols, "data": data}


def set_table_repeat_header(table_index: int, repeat: bool = True) -> dict:
    """Set first row as a repeating header row on each page."""
    app = _get_app()
    doc = app.ActiveDocument
    table = doc.Tables(table_index)
    table.Rows(1).HeadingFormat = repeat
    return {"table_index": table_index, "repeat": repeat}


# ---------------------------------------------------------------------------
# 14. Section Management
# ---------------------------------------------------------------------------

def get_section_count() -> dict:
    """Get the number of sections in the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    return {"section_count": doc.Sections.Count}


def set_section_page_setup(
    section_index: int,
    orientation: int | None = None,
    width: float | None = None,
    height: float | None = None,
    top_margin: float | None = None,
    bottom_margin: float | None = None,
    left_margin: float | None = None,
    right_margin: float | None = None,
) -> dict:
    """Set page setup for a specific section."""
    app = _get_app()
    doc = app.ActiveDocument
    sec = doc.Sections(section_index)
    ps = sec.PageSetup
    if orientation is not None:
        ps.Orientation = orientation
    if width is not None:
        ps.PageWidth = width
    if height is not None:
        ps.PageHeight = height
    if top_margin is not None:
        ps.TopMargin = top_margin
    if bottom_margin is not None:
        ps.BottomMargin = bottom_margin
    if left_margin is not None:
        ps.LeftMargin = left_margin
    if right_margin is not None:
        ps.RightMargin = right_margin
    return {"section_index": section_index}


def set_section_header(
    section_index: int,
    text: str,
    alignment: int | None = None,
) -> dict:
    """Set header text for a specific section."""
    app = _get_app()
    doc = app.ActiveDocument
    sec = doc.Sections(section_index)
    # wdHeaderFooterPrimary = 1
    header = sec.Headers(1)
    header.Range.Text = text
    if alignment is not None:
        header.Range.ParagraphFormat.Alignment = alignment
    return {"section_index": section_index, "text": text}


def set_section_footer(
    section_index: int,
    text: str,
    alignment: int | None = None,
) -> dict:
    """Set footer text for a specific section."""
    app = _get_app()
    doc = app.ActiveDocument
    sec = doc.Sections(section_index)
    # wdHeaderFooterPrimary = 1
    footer = sec.Footers(1)
    footer.Range.Text = text
    if alignment is not None:
        footer.Range.ParagraphFormat.Alignment = alignment
    return {"section_index": section_index, "text": text}


def link_section_header(section_index: int, link_to_previous: bool = True) -> dict:
    """Link or unlink a section header to/from the previous section."""
    app = _get_app()
    doc = app.ActiveDocument
    sec = doc.Sections(section_index)
    # wdHeaderFooterPrimary = 1
    sec.Headers(1).LinkToPrevious = link_to_previous
    return {"section_index": section_index, "link_to_previous": link_to_previous}


# ---------------------------------------------------------------------------
# 15. Document Navigation & Structure
# ---------------------------------------------------------------------------

def go_to_page(page_number: int) -> dict:
    """Navigate to a specific page."""
    app = _get_app()
    # wdGoToPage = 1, wdGoToAbsolute = 1
    app.Selection.GoTo(What=1, Which=1, Count=page_number)
    return {"page_number": page_number}


def get_page_count() -> dict:
    """Get total page count of the active document."""
    app = _get_app()
    doc = app.ActiveDocument
    doc.Repaginate()
    # wdStatisticPages = 2
    page_count = doc.ComputeStatistics(2)
    return {"page_count": page_count}


def insert_text_at_bookmark(bookmark_name: str, text: str) -> dict:
    """Insert text at a bookmark location."""
    app = _get_app()
    doc = app.ActiveDocument
    if not doc.Bookmarks.Exists(bookmark_name):
        raise ValueError(f"Bookmark '{bookmark_name}' not found")
    bm = doc.Bookmarks(bookmark_name)
    rng = bm.Range
    rng.Text = text
    return {"bookmark_name": bookmark_name, "text": text}


def get_paragraph_text(paragraph_index: int) -> dict:
    """Get text of a specific paragraph."""
    app = _get_app()
    doc = app.ActiveDocument
    para = doc.Paragraphs(paragraph_index)
    text = para.Range.Text
    # Remove trailing paragraph mark
    text = text.rstrip("\r")
    return {"paragraph_index": paragraph_index, "text": text}


def get_paragraph_range(start_index: int, end_index: int) -> dict:
    """Get text of a range of paragraphs."""
    app = _get_app()
    doc = app.ActiveDocument
    paragraphs = []
    for i in range(start_index, end_index + 1):
        para = doc.Paragraphs(i)
        text = para.Range.Text.rstrip("\r")
        paragraphs.append({"index": i, "text": text})
    return {"start_index": start_index, "end_index": end_index, "paragraphs": paragraphs}


# ---------------------------------------------------------------------------
# 16. Advanced Find
# ---------------------------------------------------------------------------

def find_text(
    text: str,
    match_case: bool = False,
    match_whole_word: bool = False,
) -> dict:
    """Find text in the document. Returns paragraph index and position of first match."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    found = find.Execute(
        FindText=text,
        MatchCase=match_case,
        MatchWholeWord=match_whole_word,
        Forward=True,
        Wrap=1,  # wdFindStop
    )
    if found:
        # Determine which paragraph the found range is in
        found_start = rng.Start
        para_index = None
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            if p.Range.Start <= found_start < p.Range.End:
                para_index = i
                break
        char_pos = found_start - doc.Paragraphs(para_index).Range.Start if para_index else 0
        return {
            "found": True,
            "text": text,
            "paragraph_index": para_index,
            "char_position": char_pos,
        }
    return {"found": False, "text": text}


def find_all(text: str, match_case: bool = False) -> dict:
    """Find all occurrences of text. Returns list of locations."""
    app = _get_app()
    doc = app.ActiveDocument
    locations = []
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    while find.Execute(
        FindText=text,
        MatchCase=match_case,
        Forward=True,
        Wrap=0,  # wdFindStop
    ):
        found_start = rng.Start
        para_index = None
        for i in range(1, doc.Paragraphs.Count + 1):
            p = doc.Paragraphs(i)
            if p.Range.Start <= found_start < p.Range.End:
                para_index = i
                break
        char_pos = found_start - doc.Paragraphs(para_index).Range.Start if para_index else 0
        locations.append({
            "paragraph_index": para_index,
            "char_position": char_pos,
            "start": found_start,
        })
        # Move past this match to find next
        rng.Start = rng.End
        rng.End = doc.Content.End
    return {"text": text, "count": len(locations), "locations": locations}


def highlight_found_text(
    text: str,
    highlight_color: int = 7,
    match_case: bool = False,
) -> dict:
    """Find and highlight all occurrences of text. highlight_color: WdColorIndex (7=yellow)."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = doc.Content
    find = rng.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    find.Replacement.Highlight = True
    # Set highlight color
    app.Options.DefaultHighlightColorIndex = highlight_color
    count = 0
    while find.Execute(
        FindText=text,
        MatchCase=match_case,
        Forward=True,
        Wrap=0,  # wdFindStop
    ):
        rng.HighlightColorIndex = highlight_color
        count += 1
        rng.Start = rng.End
        rng.End = doc.Content.End
    return {"text": text, "highlight_color": highlight_color, "count": count}


# ---------------------------------------------------------------------------
# 17. Mail Merge Support
# ---------------------------------------------------------------------------

def start_mail_merge(data_source_path: str) -> dict:
    """Start mail merge with a data source file (CSV/Excel)."""
    app = _get_app()
    doc = app.ActiveDocument
    abs_path = ensure_absolute_path(data_source_path)
    doc.MailMerge.OpenDataSource(Name=abs_path)
    return {"data_source": abs_path, "merge_state": doc.MailMerge.State}


def insert_merge_field(field_name: str) -> dict:
    """Insert a merge field at the current cursor position."""
    app = _get_app()
    doc = app.ActiveDocument
    rng = app.Selection.Range
    doc.MailMerge.Fields.Add(Range=rng, Name=field_name)
    return {"field_name": field_name}


def execute_mail_merge(output_path: str | None = None) -> dict:
    """Execute mail merge and optionally save the result."""
    app = _get_app()
    doc = app.ActiveDocument
    mm = doc.MailMerge
    # wdSendToNewDocument = 0
    mm.Destination = 0
    mm.Execute()
    result_doc = app.ActiveDocument
    if output_path:
        result_doc.SaveAs2(ensure_absolute_path(output_path))
    return {
        "executed": True,
        "output_name": result_doc.Name,
        "output_path": output_path,
    }
