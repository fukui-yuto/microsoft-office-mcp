"""Advanced Excel COM automation for dashboards, templates, and professional formatting.

Provides high-level functions for creating professional spreadsheets, dashboards,
financial reports, and formatted templates using Excel COM automation.
"""

import datetime

from microsoft_office.com_utils import get_or_create_app, rgb

# ---------------------------------------------------------------------------
# Excel Constants (duplicated for independence)
# ---------------------------------------------------------------------------

XL_HALIGN_LEFT = -4131
XL_HALIGN_CENTER = -4108
XL_HALIGN_RIGHT = -4152
XL_VALIGN_TOP = -4160
XL_VALIGN_CENTER = -4108
XL_VALIGN_BOTTOM = -4107
XL_BORDER_LEFT = 7
XL_BORDER_TOP = 8
XL_BORDER_BOTTOM = 9
XL_BORDER_RIGHT = 10
XL_BORDER_INSIDE_VERTICAL = 11
XL_BORDER_INSIDE_HORIZONTAL = 12
XL_LINE_STYLE_CONTINUOUS = 1
XL_LINE_STYLE_NONE = -4142
XL_BORDER_WEIGHT_HAIRLINE = 1
XL_BORDER_WEIGHT_THIN = 2
XL_BORDER_WEIGHT_MEDIUM = -4138
XL_BORDER_WEIGHT_THICK = 4
XL_PATTERN_SOLID = 1
XL_ORIENT_PORTRAIT = 1
XL_ORIENT_LANDSCAPE = 2

# ---------------------------------------------------------------------------
# Color palettes
# ---------------------------------------------------------------------------

COLOR_SCHEMES = {
    "corporate_blue": {
        "primary": (41, 65, 122),
        "secondary": (68, 114, 196),
        "accent": (47, 117, 181),
        "light": (217, 226, 243),
        "text": (255, 255, 255),
        "dark_text": (44, 62, 80),
        "bg_alt": (234, 240, 250),
    },
    "modern_dark": {
        "primary": (33, 37, 41),
        "secondary": (52, 58, 64),
        "accent": (0, 123, 255),
        "light": (233, 236, 239),
        "text": (255, 255, 255),
        "dark_text": (33, 37, 41),
        "bg_alt": (248, 249, 250),
    },
    "forest_green": {
        "primary": (39, 78, 19),
        "secondary": (56, 118, 29),
        "accent": (106, 168, 79),
        "light": (217, 234, 211),
        "text": (255, 255, 255),
        "dark_text": (39, 78, 19),
        "bg_alt": (234, 245, 228),
    },
    "sunset_orange": {
        "primary": (183, 71, 42),
        "secondary": (230, 126, 34),
        "accent": (243, 156, 18),
        "light": (253, 235, 208),
        "text": (255, 255, 255),
        "dark_text": (120, 40, 20),
        "bg_alt": (254, 245, 231),
    },
    "royal_purple": {
        "primary": (74, 20, 140),
        "secondary": (123, 31, 162),
        "accent": (156, 39, 176),
        "light": (237, 222, 245),
        "text": (255, 255, 255),
        "dark_text": (74, 20, 140),
        "bg_alt": (243, 229, 252),
    },
}


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_app():
    """Get or create the Excel application instance."""
    return get_or_create_app("Excel.Application")


def _get_ws(app, sheet=None):
    """Return the target worksheet from the active workbook."""
    wb = app.ActiveWorkbook
    return wb.Worksheets(sheet) if sheet else wb.ActiveSheet


def _apply_fill(rng, color_tuple):
    """Apply solid fill to a range."""
    rng.Interior.Pattern = XL_PATTERN_SOLID
    rng.Interior.Color = rgb(*color_tuple)


def _apply_font(rng, bold=None, italic=None, size=None, name=None, color=None):
    """Apply font properties to a range."""
    f = rng.Font
    if bold is not None:
        f.Bold = bold
    if italic is not None:
        f.Italic = italic
    if size is not None:
        f.Size = size
    if name is not None:
        f.Name = name
    if color is not None:
        f.Color = rgb(*color)


def _apply_borders(rng, edges=None, style=XL_LINE_STYLE_CONTINUOUS,
                   weight=XL_BORDER_WEIGHT_THIN, color=None):
    """Apply borders to a range."""
    if edges is None:
        edges = [XL_BORDER_LEFT, XL_BORDER_TOP, XL_BORDER_BOTTOM,
                 XL_BORDER_RIGHT, XL_BORDER_INSIDE_VERTICAL, XL_BORDER_INSIDE_HORIZONTAL]
    for edge in edges:
        try:
            b = rng.Borders(edge)
            b.LineStyle = style
            b.Weight = weight
            if color:
                b.Color = rgb(*color)
        except Exception:
            pass


def _set_alignment(rng, h_align=None, v_align=None, wrap=None, indent=None):
    """Set alignment properties on a range."""
    if h_align is not None:
        rng.HorizontalAlignment = h_align
    if v_align is not None:
        rng.VerticalAlignment = v_align
    if wrap is not None:
        rng.WrapText = wrap
    if indent is not None:
        rng.IndentLevel = indent


def _col_letter(n):
    """Convert 1-based column number to letter(s). 1->A, 27->AA."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _parse_cell(cell_str):
    """Parse cell address like 'B3' into (col_letter, row_number)."""
    col = ""
    row = ""
    for c in cell_str:
        if c.isalpha():
            col += c
        else:
            row += c
    return col, int(row)


def _col_num(col_letter):
    """Convert column letter(s) to 1-based number. A->1, AA->27."""
    result = 0
    for c in col_letter.upper():
        result = result * 26 + (ord(c) - 64)
    return result


def _get_scheme(color_scheme):
    """Get a color scheme dict, defaulting to corporate_blue."""
    return COLOR_SCHEMES.get(color_scheme, COLOR_SCHEMES["corporate_blue"])


# ---------------------------------------------------------------------------
# 1. Dashboard Header
# ---------------------------------------------------------------------------

def create_dashboard_header(sheet, title, subtitle=None, color_scheme="corporate_blue"):
    """Create a professional dashboard header with colored band, title, subtitle, and date.

    Args:
        sheet: Sheet name or None for active sheet.
        title: Main dashboard title.
        subtitle: Optional subtitle text.
        color_scheme: One of corporate_blue, modern_dark, forest_green, sunset_orange, royal_purple.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    scheme = _get_scheme(color_scheme)

    # Merge and format title row (row 1, A1:L1)
    title_range = ws.Range("A1:L1")
    title_range.Merge()
    ws.Range("A1").Value = title
    _apply_fill(title_range, scheme["primary"])
    _apply_font(title_range, bold=True, size=22, name="Segoe UI", color=scheme["text"])
    _set_alignment(title_range, h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER)
    title_range.IndentLevel = 1
    ws.Rows(1).RowHeight = 50

    # Date in the right side
    date_cell = ws.Range("M1:N1")
    date_cell.Merge()
    date_cell.Value = datetime.date.today().strftime("%Y/%m/%d")
    _apply_fill(date_cell, scheme["primary"])
    _apply_font(date_cell, bold=False, size=11, name="Segoe UI", color=scheme["text"])
    _set_alignment(date_cell, h_align=XL_HALIGN_RIGHT, v_align=XL_VALIGN_CENTER)
    date_cell.RowHeight = 50

    # Subtitle row (row 2)
    if subtitle:
        sub_range = ws.Range("A2:N2")
        sub_range.Merge()
        ws.Range("A2").Value = subtitle
        _apply_fill(sub_range, scheme["secondary"])
        _apply_font(sub_range, bold=False, size=12, name="Segoe UI", color=scheme["text"])
        _set_alignment(sub_range, h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER)
        sub_range.IndentLevel = 1
        ws.Rows(2).RowHeight = 30
    else:
        # Thin accent line
        accent_range = ws.Range("A2:N2")
        accent_range.Merge()
        _apply_fill(accent_range, scheme["accent"])
        ws.Rows(2).RowHeight = 4

    return {"sheet": ws.Name, "title": title, "subtitle": subtitle, "scheme": color_scheme}


# ---------------------------------------------------------------------------
# 2. KPI Cards
# ---------------------------------------------------------------------------

def create_kpi_cards(sheet, kpis, start_row=3, style="modern"):
    """Create KPI indicator cards on the sheet.

    Args:
        sheet: Sheet name or None for active sheet.
        kpis: List of dicts with keys: title, value, change, change_type (positive/negative/neutral).
        start_row: Starting row for KPI cards.
        style: modern, minimal, or bold.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    card_width = 3  # columns per card
    gap = 1         # gap columns between cards

    for i, kpi in enumerate(kpis):
        col_start = 1 + i * (card_width + gap)
        col_end = col_start + card_width - 1

        c_start = _col_letter(col_start)
        c_end = _col_letter(col_end)

        # Card background
        card_range = ws.Range(f"{c_start}{start_row}:{c_end}{start_row + 3}")

        if style == "modern":
            _apply_fill(card_range, (250, 250, 252))
            _apply_borders(card_range, edges=[XL_BORDER_LEFT, XL_BORDER_TOP, XL_BORDER_BOTTOM, XL_BORDER_RIGHT],
                           weight=XL_BORDER_WEIGHT_THIN, color=(220, 220, 230))
            # Top accent
            top_accent = ws.Range(f"{c_start}{start_row}:{c_end}{start_row}")
            top_accent.Merge()
            ws.Rows(start_row).RowHeight = 5
            change_type = kpi.get("change_type", "neutral")
            if change_type == "positive":
                _apply_fill(top_accent, (40, 167, 69))
            elif change_type == "negative":
                _apply_fill(top_accent, (220, 53, 69))
            else:
                _apply_fill(top_accent, (108, 117, 125))

        elif style == "minimal":
            _apply_borders(card_range, edges=[XL_BORDER_BOTTOM],
                           weight=XL_BORDER_WEIGHT_MEDIUM, color=(200, 200, 200))

        elif style == "bold":
            _apply_fill(card_range, (245, 245, 248))
            _apply_borders(card_range, edges=[XL_BORDER_LEFT],
                           weight=XL_BORDER_WEIGHT_THICK, color=(41, 65, 122))

        # Title (row start_row + 1)
        title_cell = ws.Range(f"{c_start}{start_row + 1}:{c_end}{start_row + 1}")
        title_cell.Merge()
        ws.Range(f"{c_start}{start_row + 1}").Value = kpi.get("title", "")
        _apply_font(title_cell, bold=False, size=9, name="Segoe UI", color=(120, 120, 130))
        _set_alignment(title_cell, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_BOTTOM)
        ws.Rows(start_row + 1).RowHeight = 22

        # Value (row start_row + 2)
        value_cell = ws.Range(f"{c_start}{start_row + 2}:{c_end}{start_row + 2}")
        value_cell.Merge()
        ws.Range(f"{c_start}{start_row + 2}").Value = kpi.get("value", "")
        val_size = 24 if style == "bold" else 20
        _apply_font(value_cell, bold=True, size=val_size, name="Segoe UI", color=(33, 37, 41))
        _set_alignment(value_cell, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Rows(start_row + 2).RowHeight = 36

        # Change indicator (row start_row + 3)
        change_cell = ws.Range(f"{c_start}{start_row + 3}:{c_end}{start_row + 3}")
        change_cell.Merge()
        change_type = kpi.get("change_type", "neutral")
        change_text = kpi.get("change", "")
        if change_type == "positive":
            prefix = "▲ "
            c_color = (40, 167, 69)
        elif change_type == "negative":
            prefix = "▼ "
            c_color = (220, 53, 69)
        else:
            prefix = "● "
            c_color = (108, 117, 125)
        ws.Range(f"{c_start}{start_row + 3}").Value = prefix + str(change_text)
        _apply_font(change_cell, bold=False, size=10, name="Segoe UI", color=c_color)
        _set_alignment(change_cell, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_TOP)
        ws.Rows(start_row + 3).RowHeight = 22

        # Column widths
        for c in range(col_start, col_end + 1):
            ws.Columns(c).ColumnWidth = 10

    return {"sheet": ws.Name, "kpi_count": len(kpis), "start_row": start_row, "style": style}


# ---------------------------------------------------------------------------
# 3. Data Table
# ---------------------------------------------------------------------------

def create_data_table(sheet, headers, data, start_cell="A1", style="striped_blue"):
    """Create a professional data table with auto-formatting.

    Args:
        sheet: Sheet name or None for active sheet.
        headers: List of header strings.
        data: 2D list of data values.
        start_cell: Top-left cell for the table.
        style: striped_blue, striped_gray, bordered, minimal, dark_header, colorful.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    col_str, row_num = _parse_cell(start_cell)
    col_start = _col_num(col_str)

    num_cols = len(headers)
    num_rows = len(data)

    # Style definitions
    STYLES = {
        "striped_blue": {
            "header_bg": (41, 65, 122), "header_fg": (255, 255, 255),
            "row_alt": (228, 236, 250), "row_norm": (255, 255, 255),
            "border_color": (180, 198, 231),
        },
        "striped_gray": {
            "header_bg": (89, 89, 89), "header_fg": (255, 255, 255),
            "row_alt": (242, 242, 242), "row_norm": (255, 255, 255),
            "border_color": (200, 200, 200),
        },
        "bordered": {
            "header_bg": (68, 114, 196), "header_fg": (255, 255, 255),
            "row_alt": (255, 255, 255), "row_norm": (255, 255, 255),
            "border_color": (68, 114, 196),
        },
        "minimal": {
            "header_bg": (255, 255, 255), "header_fg": (33, 37, 41),
            "row_alt": (255, 255, 255), "row_norm": (255, 255, 255),
            "border_color": (222, 226, 230),
        },
        "dark_header": {
            "header_bg": (33, 37, 41), "header_fg": (255, 255, 255),
            "row_alt": (248, 249, 250), "row_norm": (255, 255, 255),
            "border_color": (173, 181, 189),
        },
        "colorful": {
            "header_bg": (0, 123, 255), "header_fg": (255, 255, 255),
            "row_alt": (232, 245, 253), "row_norm": (255, 255, 255),
            "border_color": (0, 123, 255),
        },
    }

    st = STYLES.get(style, STYLES["striped_blue"])

    # Write headers
    for j, h in enumerate(headers):
        cell = ws.Cells(row_num, col_start + j)
        cell.Value = h

    # Format header row
    h_start = _col_letter(col_start)
    h_end = _col_letter(col_start + num_cols - 1)
    header_range = ws.Range(f"{h_start}{row_num}:{h_end}{row_num}")
    _apply_fill(header_range, st["header_bg"])
    _apply_font(header_range, bold=True, size=11, name="Segoe UI", color=st["header_fg"])
    _set_alignment(header_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(row_num).RowHeight = 28

    # Write data rows
    for i, row_data in enumerate(data):
        r = row_num + 1 + i
        for j, val in enumerate(row_data):
            ws.Cells(r, col_start + j).Value = val

        # Alternating row colors
        row_range = ws.Range(f"{h_start}{r}:{h_end}{r}")
        bg = st["row_alt"] if i % 2 == 0 else st["row_norm"]
        _apply_fill(row_range, bg)
        _apply_font(row_range, size=10, name="Segoe UI", color=(33, 37, 41))
        _set_alignment(row_range, v_align=XL_VALIGN_CENTER)
        ws.Rows(r).RowHeight = 22

    # Borders on entire table
    table_end_row = row_num + num_rows
    full_range = ws.Range(f"{h_start}{row_num}:{h_end}{table_end_row}")

    if style == "minimal":
        # Only top and bottom borders on header, bottom border on each row
        _apply_borders(header_range, edges=[XL_BORDER_TOP, XL_BORDER_BOTTOM],
                       weight=XL_BORDER_WEIGHT_MEDIUM, color=st["border_color"])
        _apply_borders(full_range, edges=[XL_BORDER_INSIDE_HORIZONTAL],
                       weight=XL_BORDER_WEIGHT_HAIRLINE, color=st["border_color"])
    elif style == "bordered":
        _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=st["border_color"])
    else:
        _apply_borders(full_range, edges=[XL_BORDER_LEFT, XL_BORDER_RIGHT,
                                          XL_BORDER_TOP, XL_BORDER_BOTTOM,
                                          XL_BORDER_INSIDE_HORIZONTAL],
                       weight=XL_BORDER_WEIGHT_HAIRLINE, color=st["border_color"])

    # Auto-fit column widths with minimum
    for j in range(num_cols):
        col = ws.Columns(col_start + j)
        col.AutoFit()
        if col.ColumnWidth < 10:
            col.ColumnWidth = 10

    return {"sheet": ws.Name, "rows": num_rows, "cols": num_cols, "style": style,
            "range": f"{h_start}{row_num}:{h_end}{table_end_row}"}


# ---------------------------------------------------------------------------
# 4. Summary Row
# ---------------------------------------------------------------------------

def create_summary_row(sheet, range_str, summary_type="sum", style="bold_bordered"):
    """Add a summary row with formulas below the specified data range.

    Args:
        sheet: Sheet name or None for active sheet.
        range_str: Data range (e.g., "A1:D10").
        summary_type: sum, average, count, min, max.
        style: bold_bordered, colored_band, double_line.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    rng = ws.Range(range_str)
    start_row = rng.Row
    end_row = rng.Row + rng.Rows.Count - 1
    start_col = rng.Column
    end_col = rng.Column + rng.Columns.Count - 1
    summary_row = end_row + 1

    func_map = {"sum": "SUM", "average": "AVERAGE", "count": "COUNT", "min": "MIN", "max": "MAX"}
    func = func_map.get(summary_type, "SUM")

    # Label in first column
    ws.Cells(summary_row, start_col).Value = summary_type.upper()
    _apply_font(ws.Cells(summary_row, start_col), bold=True, size=10, name="Segoe UI")

    # Formulas in remaining columns
    for c in range(start_col + 1, end_col + 1):
        col_letter = _col_letter(c)
        formula = f"={func}({col_letter}{start_row}:{col_letter}{end_row})"
        ws.Cells(summary_row, c).Formula = formula

    # Style the summary row
    s_start = _col_letter(start_col)
    s_end = _col_letter(end_col)
    summary_range = ws.Range(f"{s_start}{summary_row}:{s_end}{summary_row}")

    if style == "bold_bordered":
        _apply_font(summary_range, bold=True, size=10, name="Segoe UI")
        _apply_borders(summary_range, edges=[XL_BORDER_TOP],
                       weight=XL_BORDER_WEIGHT_MEDIUM, color=(33, 37, 41))
        _apply_borders(summary_range, edges=[XL_BORDER_BOTTOM],
                       weight=XL_BORDER_WEIGHT_THICK, color=(33, 37, 41))
    elif style == "colored_band":
        _apply_fill(summary_range, (41, 65, 122))
        _apply_font(summary_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    elif style == "double_line":
        _apply_font(summary_range, bold=True, size=10, name="Segoe UI")
        _apply_borders(summary_range, edges=[XL_BORDER_TOP],
                       weight=XL_BORDER_WEIGHT_THIN, color=(33, 37, 41))
        _apply_borders(summary_range, edges=[XL_BORDER_BOTTOM],
                       weight=XL_BORDER_WEIGHT_MEDIUM, color=(33, 37, 41))

    ws.Rows(summary_row).RowHeight = 26

    return {"sheet": ws.Name, "summary_row": summary_row, "type": summary_type, "style": style}


# ---------------------------------------------------------------------------
# 5. Table Theme
# ---------------------------------------------------------------------------

def apply_table_theme(sheet, range_str, theme="professional"):
    """Apply a complete visual theme to a data range.

    Assumes first row is headers.

    Args:
        sheet: Sheet name or None for active sheet.
        range_str: Range to theme (e.g., "A1:F20").
        theme: professional, financial, marketing, executive, minimal.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    rng = ws.Range(range_str)
    start_row = rng.Row
    end_row = rng.Row + rng.Rows.Count - 1
    start_col = rng.Column
    end_col = rng.Column + rng.Columns.Count - 1

    sc = _col_letter(start_col)
    ec = _col_letter(end_col)

    THEMES = {
        "professional": {
            "header_bg": (41, 65, 122), "header_fg": (255, 255, 255),
            "alt_bg": (228, 236, 250), "border": (180, 198, 231),
            "font_name": "Segoe UI", "h_align": None,
        },
        "financial": {
            "header_bg": (39, 78, 19), "header_fg": (255, 255, 255),
            "alt_bg": (226, 239, 218), "border": (169, 208, 142),
            "font_name": "Calibri", "h_align": XL_HALIGN_RIGHT,
        },
        "marketing": {
            "header_bg": (0, 123, 255), "header_fg": (255, 255, 255),
            "alt_bg": (232, 245, 253), "border": (0, 123, 255),
            "font_name": "Segoe UI", "h_align": None,
        },
        "executive": {
            "header_bg": (33, 37, 41), "header_fg": (255, 255, 255),
            "alt_bg": (248, 249, 250), "border": (173, 181, 189),
            "font_name": "Calibri", "h_align": None,
        },
        "minimal": {
            "header_bg": (255, 255, 255), "header_fg": (33, 37, 41),
            "alt_bg": (255, 255, 255), "border": (200, 200, 200),
            "font_name": "Segoe UI", "h_align": None,
        },
    }

    t = THEMES.get(theme, THEMES["professional"])

    # Header row
    header_range = ws.Range(f"{sc}{start_row}:{ec}{start_row}")
    _apply_fill(header_range, t["header_bg"])
    _apply_font(header_range, bold=True, size=11, name=t["font_name"], color=t["header_fg"])
    _set_alignment(header_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(start_row).RowHeight = 28

    # Data rows
    for r in range(start_row + 1, end_row + 1):
        row_range = ws.Range(f"{sc}{r}:{ec}{r}")
        if (r - start_row) % 2 == 1:
            _apply_fill(row_range, t["alt_bg"])
        else:
            _apply_fill(row_range, (255, 255, 255))
        _apply_font(row_range, size=10, name=t["font_name"], color=(33, 37, 41))
        if t["h_align"]:
            _set_alignment(row_range, h_align=t["h_align"])
        ws.Rows(r).RowHeight = 22

    # Borders
    full_range = ws.Range(f"{sc}{start_row}:{ec}{end_row}")
    if theme == "minimal":
        _apply_borders(header_range, edges=[XL_BORDER_BOTTOM],
                       weight=XL_BORDER_WEIGHT_MEDIUM, color=t["border"])
        _apply_borders(full_range, edges=[XL_BORDER_INSIDE_HORIZONTAL],
                       weight=XL_BORDER_WEIGHT_HAIRLINE, color=t["border"])
    else:
        _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=t["border"])

    # Auto-fit
    for c in range(start_col, end_col + 1):
        col = ws.Columns(c)
        col.AutoFit()
        if col.ColumnWidth < 10:
            col.ColumnWidth = 10

    return {"sheet": ws.Name, "range": range_str, "theme": theme}


# ---------------------------------------------------------------------------
# 6. Input Form
# ---------------------------------------------------------------------------

def create_input_form(sheet, fields, start_row=1, style="bordered"):
    """Create a data entry form with labels, input cells, and validation hints.

    Args:
        sheet: Sheet name or None for active sheet.
        fields: List of dicts with keys: label, type, required, validation, hint.
        start_row: Starting row.
        style: bordered, shaded, minimal.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    # Title
    title_range = ws.Range(f"A{start_row}:E{start_row}")
    title_range.Merge()
    ws.Range(f"A{start_row}").Value = "入力フォーム / Input Form"
    _apply_fill(title_range, (41, 65, 122))
    _apply_font(title_range, bold=True, size=14, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(start_row).RowHeight = 36

    current_row = start_row + 2

    for field in fields:
        label = field.get("label", "")
        required = field.get("required", False)
        hint = field.get("hint", "")

        # Label column (A-B)
        label_range = ws.Range(f"A{current_row}:B{current_row}")
        label_range.Merge()
        display_label = f"* {label}" if required else f"  {label}"
        ws.Range(f"A{current_row}").Value = display_label
        _apply_font(label_range, bold=True, size=10, name="Segoe UI", color=(33, 37, 41))
        _set_alignment(label_range, h_align=XL_HALIGN_RIGHT, v_align=XL_VALIGN_CENTER)

        if required:
            ws.Range(f"A{current_row}").Font.Color = rgb(192, 0, 0)

        # Input cell (C-D)
        input_range = ws.Range(f"C{current_row}:D{current_row}")
        input_range.Merge()

        if style == "bordered":
            _apply_borders(input_range, weight=XL_BORDER_WEIGHT_THIN, color=(100, 100, 100))
            _apply_fill(input_range, (255, 255, 255))
        elif style == "shaded":
            _apply_fill(input_range, (245, 245, 248))
            _apply_borders(input_range, edges=[XL_BORDER_BOTTOM],
                           weight=XL_BORDER_WEIGHT_THIN, color=(150, 150, 150))
        elif style == "minimal":
            _apply_borders(input_range, edges=[XL_BORDER_BOTTOM],
                           weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 200))

        # Hint column (E)
        if hint:
            hint_cell = ws.Range(f"E{current_row}")
            hint_cell.Value = hint
            _apply_font(hint_cell, italic=True, size=9, name="Segoe UI", color=(150, 150, 160))

        ws.Rows(current_row).RowHeight = 26
        current_row += 1

    # Set column widths
    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 15
    ws.Columns("E").ColumnWidth = 30

    return {"sheet": ws.Name, "fields": len(fields), "style": style}


# ---------------------------------------------------------------------------
# 7. Calendar
# ---------------------------------------------------------------------------

def create_calendar(sheet, year, month, start_cell="A1", style="modern"):
    """Create a monthly calendar.

    Args:
        sheet: Sheet name or None for active sheet.
        year: Year (e.g., 2024).
        month: Month (1-12).
        start_cell: Top-left cell.
        style: modern, minimal, colorful.

    Returns:
        dict with status info.
    """
    import calendar

    app = _get_app()
    ws = _get_ws(app, sheet)

    col_str, row_num = _parse_cell(start_cell)
    col_start = _col_num(col_str)

    STYLES = {
        "modern": {
            "header_bg": (41, 65, 122), "header_fg": (255, 255, 255),
            "day_header_bg": (68, 114, 196), "day_header_fg": (255, 255, 255),
            "weekend_bg": (240, 240, 245), "weekday_bg": (255, 255, 255),
            "today_bg": (0, 123, 255), "today_fg": (255, 255, 255),
        },
        "minimal": {
            "header_bg": (255, 255, 255), "header_fg": (33, 37, 41),
            "day_header_bg": (245, 245, 245), "day_header_fg": (33, 37, 41),
            "weekend_bg": (252, 252, 252), "weekday_bg": (255, 255, 255),
            "today_bg": (33, 37, 41), "today_fg": (255, 255, 255),
        },
        "colorful": {
            "header_bg": (0, 123, 255), "header_fg": (255, 255, 255),
            "day_header_bg": (40, 167, 69), "day_header_fg": (255, 255, 255),
            "weekend_bg": (255, 243, 205), "weekday_bg": (255, 255, 255),
            "today_bg": (220, 53, 69), "today_fg": (255, 255, 255),
        },
    }
    st = STYLES.get(style, STYLES["modern"])

    month_names = ["", "January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]

    # Title row
    t_start = _col_letter(col_start)
    t_end = _col_letter(col_start + 6)
    title_range = ws.Range(f"{t_start}{row_num}:{t_end}{row_num}")
    title_range.Merge()
    ws.Range(f"{t_start}{row_num}").Value = f"{month_names[month]} {year}"
    _apply_fill(title_range, st["header_bg"])
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=st["header_fg"])
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(row_num).RowHeight = 36

    # Day headers
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    header_row = row_num + 1
    for j, day_name in enumerate(day_names):
        cell = ws.Cells(header_row, col_start + j)
        cell.Value = day_name
    day_header_range = ws.Range(f"{t_start}{header_row}:{t_end}{header_row}")
    _apply_fill(day_header_range, st["day_header_bg"])
    _apply_font(day_header_range, bold=True, size=10, name="Segoe UI", color=st["day_header_fg"])
    _set_alignment(day_header_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(header_row).RowHeight = 24

    # Calendar days
    cal = calendar.monthcalendar(year, month)
    today = datetime.date.today()

    for week_idx, week in enumerate(cal):
        r = row_num + 2 + week_idx
        ws.Rows(r).RowHeight = 40
        for day_idx, day in enumerate(week):
            cell = ws.Cells(r, col_start + day_idx)
            if day != 0:
                cell.Value = day
                _apply_font(cell, size=12, name="Segoe UI", color=(33, 37, 41))
                # Check if today
                if year == today.year and month == today.month and day == today.day:
                    _apply_fill(cell, st["today_bg"])
                    _apply_font(cell, bold=True, color=st["today_fg"])
                elif day_idx >= 5:  # Weekend
                    _apply_fill(cell, st["weekend_bg"])
                else:
                    _apply_fill(cell, st["weekday_bg"])
            _set_alignment(cell, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)

    # Borders and column widths
    total_rows = row_num + 1 + len(cal)
    grid_range = ws.Range(f"{t_start}{header_row}:{t_end}{total_rows}")
    _apply_borders(grid_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    for j in range(7):
        ws.Columns(col_start + j).ColumnWidth = 10

    return {"sheet": ws.Name, "year": year, "month": month, "style": style}


# ---------------------------------------------------------------------------
# 8. Gantt Chart
# ---------------------------------------------------------------------------

def create_gantt_chart(sheet, tasks, start_cell="A1"):
    """Create a Gantt chart using cell colors and conditional formatting.

    Args:
        sheet: Sheet name or None for active sheet.
        tasks: List of dicts with keys: name, start_date (YYYY-MM-DD), end_date, progress (0-100), color ([R,G,B]).
        start_cell: Top-left cell.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    col_str, row_num = _parse_cell(start_cell)
    col_start = _col_num(col_str)

    if not tasks:
        return {"sheet": ws.Name, "tasks": 0}

    # Parse all dates to find range
    from datetime import datetime as dt
    all_dates = []
    parsed_tasks = []
    for task in tasks:
        s = dt.strptime(task["start_date"], "%Y-%m-%d").date()
        e = dt.strptime(task["end_date"], "%Y-%m-%d").date()
        all_dates.extend([s, e])
        color = task.get("color", [66, 133, 244])
        if isinstance(color, list):
            color = tuple(color)
        parsed_tasks.append({
            "name": task["name"],
            "start": s,
            "end": e,
            "progress": task.get("progress", 0),
            "color": color,
        })

    min_date = min(all_dates)
    max_date = max(all_dates)
    total_days = (max_date - min_date).days + 1

    # Limit columns to avoid excessive width
    if total_days > 90:
        # Use weekly granularity
        day_step = 7
    else:
        day_step = 1

    date_cols = []
    current = min_date
    while current <= max_date:
        date_cols.append(current)
        current += datetime.timedelta(days=day_step)

    # Headers: Task | Progress | date columns...
    info_cols = 3  # Task name, Duration, Progress
    ws.Cells(row_num, col_start).Value = "Task"
    ws.Cells(row_num, col_start + 1).Value = "Duration"
    ws.Cells(row_num, col_start + 2).Value = "Progress"

    # Header formatting
    h_start = _col_letter(col_start)
    h_end = _col_letter(col_start + info_cols - 1 + len(date_cols))
    header_range = ws.Range(f"{h_start}{row_num}:{h_end}{row_num}")
    _apply_fill(header_range, (33, 37, 41))
    _apply_font(header_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(header_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(row_num).RowHeight = 26

    # Date headers
    for idx, d in enumerate(date_cols):
        col = col_start + info_cols + idx
        ws.Cells(row_num, col).Value = d.strftime("%m/%d")
        ws.Columns(col).ColumnWidth = 4

    # Column widths for info columns
    ws.Columns(col_start).ColumnWidth = 25
    ws.Columns(col_start + 1).ColumnWidth = 12
    ws.Columns(col_start + 2).ColumnWidth = 10

    # Task rows
    for t_idx, task in enumerate(parsed_tasks):
        r = row_num + 1 + t_idx
        ws.Cells(r, col_start).Value = task["name"]
        duration_days = (task["end"] - task["start"]).days + 1
        ws.Cells(r, col_start + 1).Value = f"{duration_days}d"
        ws.Cells(r, col_start + 2).Value = f"{task['progress']}%"
        ws.Rows(r).RowHeight = 22

        _apply_font(ws.Cells(r, col_start), bold=False, size=10, name="Segoe UI")
        _set_alignment(ws.Cells(r, col_start + 1), h_align=XL_HALIGN_CENTER)
        _set_alignment(ws.Cells(r, col_start + 2), h_align=XL_HALIGN_CENTER)

        # Fill Gantt bars
        for idx, d in enumerate(date_cols):
            col = col_start + info_cols + idx
            d_end = d + datetime.timedelta(days=max(day_step - 1, 0))
            if task["start"] <= d_end and task["end"] >= d:
                cell = ws.Cells(r, col)
                _apply_fill(cell, task["color"])
                # Progress indicator (darker shade for completed portion)
                if task["progress"] > 0:
                    # Approximate which portion is done
                    task_total = (task["end"] - task["start"]).days + 1
                    done_days = int(task_total * task["progress"] / 100)
                    done_date = task["start"] + datetime.timedelta(days=done_days)
                    if d < done_date:
                        # Darken the color for completed part
                        darker = tuple(max(0, c - 40) for c in task["color"])
                        _apply_fill(cell, darker)

        # Alternate row background for info columns
        info_range = ws.Range(f"{_col_letter(col_start)}{r}:{_col_letter(col_start + info_cols - 1)}{r}")
        if t_idx % 2 == 0:
            _apply_fill(info_range, (248, 249, 250))
        else:
            _apply_fill(info_range, (255, 255, 255))

    # Border on info section
    info_end_row = row_num + len(parsed_tasks)
    info_range = ws.Range(f"{_col_letter(col_start)}{row_num}:{_col_letter(col_start + info_cols - 1)}{info_end_row}")
    _apply_borders(info_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 200))

    return {"sheet": ws.Name, "tasks": len(parsed_tasks), "days": total_days}


# ---------------------------------------------------------------------------
# 9. Scorecard
# ---------------------------------------------------------------------------

def create_scorecard(sheet, title, metrics, start_cell="A1", style="traffic_light"):
    """Create a performance scorecard.

    Args:
        sheet: Sheet name or None for active sheet.
        title: Scorecard title.
        metrics: List of dicts with keys: name, target, actual, unit.
        start_cell: Top-left cell.
        style: traffic_light, progress_bar, rating.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    col_str, row_num = _parse_cell(start_cell)
    col_start = _col_num(col_str)

    # Title
    sc = _col_letter(col_start)
    ec = _col_letter(col_start + 5)
    title_range = ws.Range(f"{sc}{row_num}:{ec}{row_num}")
    title_range.Merge()
    ws.Range(f"{sc}{row_num}").Value = title
    _apply_fill(title_range, (33, 37, 41))
    _apply_font(title_range, bold=True, size=14, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER)
    title_range.IndentLevel = 1
    ws.Rows(row_num).RowHeight = 36

    # Column headers
    headers_row = row_num + 1
    header_labels = ["Metric", "Target", "Actual", "Variance", "Status", ""]
    for j, lbl in enumerate(header_labels):
        ws.Cells(headers_row, col_start + j).Value = lbl
    h_range = ws.Range(f"{sc}{headers_row}:{ec}{headers_row}")
    _apply_fill(h_range, (68, 114, 196))
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(headers_row).RowHeight = 26

    # Metrics rows
    for i, m in enumerate(metrics):
        r = headers_row + 1 + i
        target = m.get("target", 0)
        actual = m.get("actual", 0)
        unit = m.get("unit", "")
        variance = actual - target
        pct = (actual / target * 100) if target != 0 else 0

        ws.Cells(r, col_start).Value = m.get("name", "")
        ws.Cells(r, col_start + 1).Value = f"{target}{unit}"
        ws.Cells(r, col_start + 2).Value = f"{actual}{unit}"
        ws.Cells(r, col_start + 3).Value = f"{variance:+}{unit}"

        # Status indicator based on style
        if style == "traffic_light":
            if pct >= 95:
                status = "●"
                status_color = (40, 167, 69)
            elif pct >= 80:
                status = "●"
                status_color = (255, 193, 7)
            else:
                status = "●"
                status_color = (220, 53, 69)
            ws.Cells(r, col_start + 4).Value = status
            _apply_font(ws.Cells(r, col_start + 4), size=16, color=status_color)
            _set_alignment(ws.Cells(r, col_start + 4), h_align=XL_HALIGN_CENTER)

        elif style == "progress_bar":
            bar_pct = min(pct, 100) / 100
            bar_len = int(bar_pct * 10)
            bar = "█" * bar_len + "░" * (10 - bar_len)
            ws.Cells(r, col_start + 4).Value = bar
            bar_color = (40, 167, 69) if pct >= 95 else (255, 193, 7) if pct >= 80 else (220, 53, 69)
            _apply_font(ws.Cells(r, col_start + 4), size=10, color=bar_color, name="Consolas")

        elif style == "rating":
            stars = min(5, int(pct / 20))
            ws.Cells(r, col_start + 4).Value = "★" * stars + "☆" * (5 - stars)
            _apply_font(ws.Cells(r, col_start + 4), size=12, color=(243, 156, 18), name="Segoe UI")
            _set_alignment(ws.Cells(r, col_start + 4), h_align=XL_HALIGN_CENTER)

        # Percentage column
        ws.Cells(r, col_start + 5).Value = f"{pct:.1f}%"
        _set_alignment(ws.Cells(r, col_start + 5), h_align=XL_HALIGN_CENTER)

        # Row formatting
        row_range = ws.Range(f"{sc}{r}:{ec}{r}")
        bg = (248, 249, 250) if i % 2 == 0 else (255, 255, 255)
        _apply_fill(row_range, bg)
        _apply_font(ws.Cells(r, col_start), size=10, name="Segoe UI")
        _set_alignment(ws.Cells(r, col_start), h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER)
        ws.Rows(r).RowHeight = 26

        # Variance color
        var_color = (40, 167, 69) if variance >= 0 else (220, 53, 69)
        _apply_font(ws.Cells(r, col_start + 3), color=var_color, bold=True)

    # Borders
    end_row = headers_row + len(metrics)
    full_range = ws.Range(f"{sc}{headers_row}:{ec}{end_row}")
    _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Column widths
    widths = [25, 12, 12, 12, 14, 10]
    for j, w in enumerate(widths):
        ws.Columns(col_start + j).ColumnWidth = w

    return {"sheet": ws.Name, "title": title, "metrics": len(metrics), "style": style}


# ---------------------------------------------------------------------------
# 10. Financial Report
# ---------------------------------------------------------------------------

def create_financial_report(sheet, title, categories, periods, values, style="standard"):
    """Create a financial statement layout.

    Args:
        sheet: Sheet name or None for active sheet.
        title: Report title.
        categories: List of category dicts: {"name": str, "level": int (0=header,1=item,2=subtotal,3=total), "bold": bool}.
        periods: List of period labels (column headers).
        values: 2D list matching categories x periods.
        style: standard, detailed, summary.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    num_periods = len(periods)

    # Title
    t_end = _col_letter(1 + num_periods)
    title_range = ws.Range(f"A1:{t_end}1")
    title_range.Merge()
    ws.Range("A1").Value = title
    _apply_fill(title_range, (33, 37, 41))
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER)
    title_range.IndentLevel = 1
    ws.Rows(1).RowHeight = 42

    # Subtitle / date
    sub_range = ws.Range(f"A2:{t_end}2")
    sub_range.Merge()
    ws.Range("A2").Value = f"Generated: {datetime.date.today().strftime('%Y/%m/%d')}"
    _apply_fill(sub_range, (52, 58, 64))
    _apply_font(sub_range, size=9, name="Segoe UI", color=(173, 181, 189))
    _set_alignment(sub_range, h_align=XL_HALIGN_RIGHT, v_align=XL_VALIGN_CENTER)
    ws.Rows(2).RowHeight = 22

    # Period headers (row 4)
    ws.Cells(4, 1).Value = "Category"
    _apply_font(ws.Cells(4, 1), bold=True, size=10, name="Segoe UI")
    for j, p in enumerate(periods):
        cell = ws.Cells(4, 2 + j)
        cell.Value = p
        _apply_font(cell, bold=True, size=10, name="Segoe UI")
        _set_alignment(cell, h_align=XL_HALIGN_RIGHT)

    h_range = ws.Range(f"A4:{t_end}4")
    _apply_fill(h_range, (41, 65, 122))
    _apply_font(h_range, color=(255, 255, 255))
    ws.Rows(4).RowHeight = 28

    # Data rows
    for i, cat in enumerate(categories):
        r = 5 + i
        name = cat.get("name", "")
        level = cat.get("level", 1)
        is_bold = cat.get("bold", False) or level in (0, 2, 3)

        ws.Cells(r, 1).Value = name
        _apply_font(ws.Cells(r, 1), bold=is_bold, size=10, name="Segoe UI")

        # Indentation based on level
        if level == 1:
            ws.Cells(r, 1).IndentLevel = 2
        elif level == 0:
            ws.Cells(r, 1).IndentLevel = 0
        elif level == 2:
            ws.Cells(r, 1).IndentLevel = 1

        # Values
        if i < len(values):
            for j, val in enumerate(values[i]):
                cell = ws.Cells(r, 2 + j)
                cell.Value = val
                _set_alignment(cell, h_align=XL_HALIGN_RIGHT)
                _apply_font(cell, bold=is_bold, size=10, name="Segoe UI")
                # Number format
                cell.NumberFormat = "#,##0;(#,##0);-"

        # Row styling based on level
        row_range = ws.Range(f"A{r}:{t_end}{r}")
        if level == 0:  # Section header
            _apply_fill(row_range, (68, 114, 196))
            _apply_font(row_range, bold=True, color=(255, 255, 255))
            ws.Rows(r).RowHeight = 26
        elif level == 2:  # Subtotal
            _apply_fill(row_range, (228, 236, 250))
            _apply_borders(row_range, edges=[XL_BORDER_TOP],
                           weight=XL_BORDER_WEIGHT_THIN, color=(41, 65, 122))
            ws.Rows(r).RowHeight = 24
        elif level == 3:  # Grand total
            _apply_fill(row_range, (33, 37, 41))
            _apply_font(row_range, bold=True, size=11, color=(255, 255, 255))
            _apply_borders(row_range, edges=[XL_BORDER_TOP],
                           weight=XL_BORDER_WEIGHT_MEDIUM, color=(33, 37, 41))
            ws.Rows(r).RowHeight = 28
        else:
            bg = (248, 249, 250) if i % 2 == 0 else (255, 255, 255)
            _apply_fill(row_range, bg)
            ws.Rows(r).RowHeight = 22

    # Column widths
    ws.Columns(1).ColumnWidth = 35
    for j in range(num_periods):
        ws.Columns(2 + j).ColumnWidth = 16

    # Outer borders
    end_row = 4 + len(categories)
    full_range = ws.Range(f"A4:{t_end}{end_row}")
    _apply_borders(full_range, edges=[XL_BORDER_LEFT, XL_BORDER_RIGHT, XL_BORDER_TOP, XL_BORDER_BOTTOM],
                   weight=XL_BORDER_WEIGHT_THIN, color=(150, 150, 150))

    return {"sheet": ws.Name, "title": title, "categories": len(categories), "periods": len(periods)}


# ---------------------------------------------------------------------------
# 11. Comparison Table
# ---------------------------------------------------------------------------

def create_comparison_table(sheet, headers, row_labels, data, highlight_best=True, start_cell="A1"):
    """Create a comparison matrix with optional best-value highlighting.

    Args:
        sheet: Sheet name or None for active sheet.
        headers: List of column headers (comparison items).
        row_labels: List of row labels (criteria).
        data: 2D list of values (rows x cols).
        highlight_best: If True, highlights the best (max numeric) value per row.
        start_cell: Top-left cell.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    col_str, row_num = _parse_cell(start_cell)
    col_start = _col_num(col_str)

    num_cols = len(headers)

    # Corner cell
    ws.Cells(row_num, col_start).Value = "Criteria"
    _apply_font(ws.Cells(row_num, col_start), bold=True, size=10, name="Segoe UI")

    # Column headers
    for j, h in enumerate(headers):
        cell = ws.Cells(row_num, col_start + 1 + j)
        cell.Value = h

    # Header formatting
    sc = _col_letter(col_start)
    ec = _col_letter(col_start + num_cols)
    h_range = ws.Range(f"{sc}{row_num}:{ec}{row_num}")
    _apply_fill(h_range, (41, 65, 122))
    _apply_font(h_range, bold=True, size=11, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(row_num).RowHeight = 30

    # Data rows
    for i, (label, row_data) in enumerate(zip(row_labels, data)):
        r = row_num + 1 + i
        ws.Cells(r, col_start).Value = label
        _apply_font(ws.Cells(r, col_start), bold=True, size=10, name="Segoe UI", color=(33, 37, 41))
        _set_alignment(ws.Cells(r, col_start), h_align=XL_HALIGN_LEFT)

        # Find best value in row (max for numeric)
        numeric_vals = []
        for v in row_data:
            try:
                numeric_vals.append(float(v))
            except (ValueError, TypeError):
                numeric_vals.append(None)

        best_idx = None
        if highlight_best:
            valid = [(idx, v) for idx, v in enumerate(numeric_vals) if v is not None]
            if valid:
                best_idx = max(valid, key=lambda x: x[1])[0]

        for j, val in enumerate(row_data):
            cell = ws.Cells(r, col_start + 1 + j)
            cell.Value = val
            _apply_font(cell, size=10, name="Segoe UI")
            _set_alignment(cell, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)

            if j == best_idx and highlight_best:
                _apply_fill(cell, (212, 237, 218))
                _apply_font(cell, bold=True, color=(21, 87, 36))

        # Row coloring
        row_range = ws.Range(f"{sc}{r}:{ec}{r}")
        if i % 2 == 0:
            if best_idx is None or not highlight_best:
                _apply_fill(row_range, (248, 249, 250))
        ws.Rows(r).RowHeight = 24

    # Borders
    end_row = row_num + len(row_labels)
    full_range = ws.Range(f"{sc}{row_num}:{ec}{end_row}")
    _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=(180, 198, 231))

    # Auto-fit
    ws.Columns(col_start).ColumnWidth = 20
    for j in range(num_cols):
        col = ws.Columns(col_start + 1 + j)
        col.AutoFit()
        if col.ColumnWidth < 14:
            col.ColumnWidth = 14

    return {"sheet": ws.Name, "rows": len(row_labels), "cols": num_cols, "highlighted": highlight_best}


# ---------------------------------------------------------------------------
# 12. Print-Ready Setup
# ---------------------------------------------------------------------------

def setup_print_ready(sheet, title=None, orientation="portrait", fit_to_pages=True):
    """Make a sheet print-ready with proper margins, headers/footers, and fit-to-page.

    Args:
        sheet: Sheet name or None for active sheet.
        title: Optional title for header.
        orientation: portrait or landscape.
        fit_to_pages: If True, fits content to one page wide.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    ps = ws.PageSetup

    # Orientation
    if orientation == "landscape":
        ps.Orientation = XL_ORIENT_LANDSCAPE
    else:
        ps.Orientation = XL_ORIENT_PORTRAIT

    # Margins (in points, 72 points = 1 inch)
    ps.LeftMargin = app.InchesToPoints(0.5)
    ps.RightMargin = app.InchesToPoints(0.5)
    ps.TopMargin = app.InchesToPoints(0.75)
    ps.BottomMargin = app.InchesToPoints(0.75)
    ps.HeaderMargin = app.InchesToPoints(0.3)
    ps.FooterMargin = app.InchesToPoints(0.3)

    # Header/Footer
    if title:
        ps.LeftHeader = f"&\"Segoe UI,Bold\"&12{title}"
    ps.RightHeader = "&D"  # Date
    ps.CenterFooter = "&P / &N"  # Page X of Y

    # Fit to page
    if fit_to_pages:
        ps.Zoom = False
        ps.FitToPagesWide = 1
        ps.FitToPagesTall = False  # As many pages tall as needed

    # Print gridlines off for clean look
    ps.PrintGridlines = False

    # Center horizontally
    ps.CenterHorizontally = True

    return {"sheet": ws.Name, "orientation": orientation, "fit_to_pages": fit_to_pages}


# ---------------------------------------------------------------------------
# 13. Heatmap
# ---------------------------------------------------------------------------

def create_heatmap(sheet, data_range, color_low=None, color_mid=None, color_high=None):
    """Apply heatmap coloring to a range of numbers.

    Args:
        sheet: Sheet name or None for active sheet.
        data_range: Range string (e.g., "B2:F10").
        color_low: RGB tuple for low values. Default: green (40,167,69).
        color_mid: RGB tuple for mid values. Default: yellow (255,193,7).
        color_high: RGB tuple for high values. Default: red (220,53,69).

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    if color_low is None:
        color_low = (40, 167, 69)
    if color_mid is None:
        color_mid = (255, 193, 7)
    if color_high is None:
        color_high = (220, 53, 69)

    rng = ws.Range(data_range)

    # Use Excel's built-in 3-color scale conditional formatting
    # FormatConditions.AddColorScale(3)
    cs = rng.FormatConditions.AddColorScale(3)

    # Low (minimum)
    cs.ColorScaleCriteria(1).Type = 1  # xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = rgb(*color_low)

    # Mid (percentile 50)
    cs.ColorScaleCriteria(2).Type = 4  # xlConditionValuePercentile
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = rgb(*color_mid)

    # High (maximum)
    cs.ColorScaleCriteria(3).Type = 2  # xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = rgb(*color_high)

    return {"sheet": ws.Name, "range": data_range}


# ---------------------------------------------------------------------------
# 14. Sheet Navigation
# ---------------------------------------------------------------------------

def add_sheet_navigation(sheets_info, nav_sheet_name="Menu"):
    """Create a navigation/menu sheet with hyperlinks to all sheets.

    Args:
        sheets_info: List of dicts with keys: name, description.
        nav_sheet_name: Name for the navigation sheet.

    Returns:
        dict with status info.
    """
    app = _get_app()
    wb = app.ActiveWorkbook

    # Create or get nav sheet
    try:
        ws = wb.Worksheets(nav_sheet_name)
        ws.Cells.Clear()
    except Exception:
        ws = wb.Worksheets.Add(Before=wb.Worksheets(1))
        ws.Name = nav_sheet_name

    # Title
    title_range = ws.Range("A1:E1")
    title_range.Merge()
    ws.Range("A1").Value = "Navigation Menu"
    _apply_fill(title_range, (33, 37, 41))
    _apply_font(title_range, bold=True, size=18, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 48

    # Subtitle
    sub_range = ws.Range("A2:E2")
    sub_range.Merge()
    ws.Range("A2").Value = f"Last updated: {datetime.date.today().strftime('%Y/%m/%d')}"
    _apply_fill(sub_range, (52, 58, 64))
    _apply_font(sub_range, size=10, name="Segoe UI", color=(173, 181, 189))
    _set_alignment(sub_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(2).RowHeight = 26

    # Headers
    ws.Cells(4, 1).Value = "#"
    ws.Cells(4, 2).Value = "Sheet Name"
    ws.Cells(4, 3).Value = "Description"
    ws.Cells(4, 4).Value = "Go To"
    h_range = ws.Range("A4:D4")
    _apply_fill(h_range, (41, 65, 122))
    _apply_font(h_range, bold=True, size=11, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(4).RowHeight = 28

    # Sheet entries
    for i, info in enumerate(sheets_info):
        r = 5 + i
        name = info.get("name", "")
        desc = info.get("description", "")

        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = name
        ws.Cells(r, 3).Value = desc

        # Hyperlink to sheet
        link_cell = ws.Cells(r, 4)
        link_cell.Value = "→ Open"
        try:
            ws.Hyperlinks.Add(
                Anchor=link_cell,
                Address="",
                SubAddress=f"'{name}'!A1",
                TextToDisplay="→ Open"
            )
        except Exception:
            pass

        _apply_font(link_cell, color=(0, 102, 204), bold=True)

        # Row formatting
        row_range = ws.Range(f"A{r}:D{r}")
        bg = (248, 249, 250) if i % 2 == 0 else (255, 255, 255)
        _apply_fill(row_range, bg)
        _apply_font(ws.Cells(r, 1), size=10, name="Segoe UI", color=(108, 117, 125))
        _set_alignment(ws.Cells(r, 1), h_align=XL_HALIGN_CENTER)
        _apply_font(ws.Cells(r, 2), bold=True, size=10, name="Segoe UI")
        _apply_font(ws.Cells(r, 3), size=10, name="Segoe UI", color=(108, 117, 125))
        ws.Rows(r).RowHeight = 26

    # Borders
    end_row = 4 + len(sheets_info)
    full_range = ws.Range(f"A4:D{end_row}")
    _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Column widths
    ws.Columns("A").ColumnWidth = 6
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 40
    ws.Columns("D").ColumnWidth = 12

    return {"nav_sheet": nav_sheet_name, "entries": len(sheets_info)}


# ---------------------------------------------------------------------------
# 15. Invoice Template
# ---------------------------------------------------------------------------

def create_invoice_template(sheet, company_name, company_address=None, logo_path=None, style="modern"):
    """Create a professional invoice template with formulas.

    Args:
        sheet: Sheet name or None for active sheet.
        company_name: Company name for the invoice.
        company_address: Optional company address.
        logo_path: Optional path to logo image.
        style: modern.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    # Company header (A1:H2)
    header_range = ws.Range("A1:H1")
    header_range.Merge()
    ws.Range("A1").Value = company_name
    _apply_font(header_range, bold=True, size=22, name="Segoe UI", color=(41, 65, 122))
    ws.Rows(1).RowHeight = 40

    if company_address:
        addr_range = ws.Range("A2:H2")
        addr_range.Merge()
        ws.Range("A2").Value = company_address
        _apply_font(addr_range, size=10, name="Segoe UI", color=(108, 117, 125))
        ws.Rows(2).RowHeight = 20

    # Logo
    if logo_path:
        try:
            from microsoft_office.com_utils import ensure_absolute_path
            ws.Shapes.AddPicture(
                ensure_absolute_path(logo_path),
                False, True, 400, 5, 80, 40
            )
        except Exception:
            pass

    # "INVOICE" label
    inv_range = ws.Range("F3:H3")
    inv_range.Merge()
    ws.Range("F3").Value = "INVOICE"
    _apply_font(inv_range, bold=True, size=28, name="Segoe UI", color=(41, 65, 122))
    _set_alignment(inv_range, h_align=XL_HALIGN_RIGHT)
    ws.Rows(3).RowHeight = 42

    # Invoice details
    details = [
        ("Invoice No:", "INV-0001", 5),
        ("Date:", datetime.date.today().strftime("%Y/%m/%d"), 6),
        ("Due Date:", "", 7),
        ("Payment Terms:", "Net 30", 8),
    ]
    for label, value, r in details:
        ws.Cells(r, 6).Value = label
        ws.Cells(r, 7).Value = value
        _apply_font(ws.Cells(r, 6), bold=True, size=10, name="Segoe UI", color=(108, 117, 125))
        _apply_font(ws.Cells(r, 7), size=10, name="Segoe UI")
        ws.Rows(r).RowHeight = 20

    # Bill To section
    ws.Cells(5, 1).Value = "Bill To:"
    _apply_font(ws.Cells(5, 1), bold=True, size=10, name="Segoe UI", color=(41, 65, 122))
    ws.Cells(6, 1).Value = "[Client Name]"
    ws.Cells(7, 1).Value = "[Address]"
    ws.Cells(8, 1).Value = "[City, State, ZIP]"
    for r in range(6, 9):
        _apply_font(ws.Cells(r, 1), size=10, name="Segoe UI", color=(108, 117, 125))

    # Separator
    sep_range = ws.Range("A10:H10")
    sep_range.Merge()
    _apply_fill(sep_range, (41, 65, 122))
    ws.Rows(10).RowHeight = 3

    # Item table headers (row 12)
    item_headers = ["#", "Description", "", "", "Qty", "Unit Price", "Tax %", "Amount"]
    for j, h in enumerate(item_headers):
        ws.Cells(12, 1 + j).Value = h
    h_range = ws.Range("A12:H12")
    _apply_fill(h_range, (41, 65, 122))
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(12).RowHeight = 28

    # Merge description columns
    ws.Range("B12:D12").Merge()

    # Item rows (13-22, 10 rows for items)
    for i in range(10):
        r = 13 + i
        ws.Cells(r, 1).Value = i + 1
        _set_alignment(ws.Cells(r, 1), h_align=XL_HALIGN_CENTER)
        ws.Range(f"B{r}:D{r}").Merge()

        # Amount formula: Qty * Unit Price * (1 + Tax%/100)
        ws.Cells(r, 8).Formula = f'=IF(E{r}="","",E{r}*F{r}*(1+G{r}/100))'
        ws.Cells(r, 8).NumberFormat = "#,##0"
        ws.Cells(r, 5).NumberFormat = "#,##0"
        ws.Cells(r, 6).NumberFormat = "#,##0"

        # Alternating rows
        row_range = ws.Range(f"A{r}:H{r}")
        bg = (248, 249, 250) if i % 2 == 0 else (255, 255, 255)
        _apply_fill(row_range, bg)
        _apply_font(row_range, size=10, name="Segoe UI")
        ws.Rows(r).RowHeight = 22
        _apply_borders(row_range, edges=[XL_BORDER_BOTTOM],
                       weight=XL_BORDER_WEIGHT_HAIRLINE, color=(220, 220, 220))

    # Summary section (rows 24-27)
    summary_items = [
        ("Subtotal", "=SUM(H13:H22)", 24),
        ("Tax", '=SUMPRODUCT((E13:E22)*(F13:F22)*(G13:G22)/100)', 25),
        ("Discount", "", 26),
        ("TOTAL", "=H24+H25-H26", 27),
    ]

    for label, formula, r in summary_items:
        ws.Range(f"F{r}:G{r}").Merge()
        ws.Cells(r, 6).Value = label
        _apply_font(ws.Cells(r, 6), bold=True, size=10 if r < 27 else 13, name="Segoe UI")
        _set_alignment(ws.Cells(r, 6), h_align=XL_HALIGN_RIGHT)
        if formula:
            ws.Cells(r, 8).Formula = formula
        ws.Cells(r, 8).NumberFormat = "#,##0"
        ws.Rows(r).RowHeight = 24

    # Total row styling
    total_range = ws.Range("F27:H27")
    _apply_fill(total_range, (41, 65, 122))
    _apply_font(total_range, bold=True, size=13, name="Segoe UI", color=(255, 255, 255))
    ws.Rows(27).RowHeight = 32

    # Notes section
    ws.Cells(29, 1).Value = "Notes:"
    _apply_font(ws.Cells(29, 1), bold=True, size=10, name="Segoe UI", color=(41, 65, 122))
    ws.Cells(30, 1).Value = "Thank you for your business."
    _apply_font(ws.Cells(30, 1), size=10, name="Segoe UI", color=(108, 117, 125))

    # Column widths
    col_widths = [5, 12, 12, 12, 8, 12, 8, 14]
    for j, w in enumerate(col_widths):
        ws.Columns(j + 1).ColumnWidth = w

    return {"sheet": ws.Name, "company": company_name, "style": style}


# ---------------------------------------------------------------------------
# 16. Timesheet
# ---------------------------------------------------------------------------

def create_timesheet(sheet, employee_name=None, month=None, year=None, style="standard"):
    """Create a weekly/monthly timesheet template with auto-calculation.

    Args:
        sheet: Sheet name or None for active sheet.
        employee_name: Optional employee name.
        month: Month number (1-12). Defaults to current month.
        year: Year. Defaults to current year.
        style: standard.

    Returns:
        dict with status info.
    """
    import calendar

    app = _get_app()
    ws = _get_ws(app, sheet)

    today = datetime.date.today()
    if year is None:
        year = today.year
    if month is None:
        month = today.month

    month_names = ["", "January", "February", "March", "April", "May", "June",
                   "July", "August", "September", "October", "November", "December"]

    # Title
    title_range = ws.Range("A1:I1")
    title_range.Merge()
    ws.Range("A1").Value = "TIMESHEET"
    _apply_fill(title_range, (33, 37, 41))
    _apply_font(title_range, bold=True, size=20, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 44

    # Info section
    info = [
        ("Employee:", employee_name or "[Name]", 3, 1),
        ("Department:", "[Department]", 3, 5),
        ("Period:", f"{month_names[month]} {year}", 4, 1),
        ("Manager:", "[Manager Name]", 4, 5),
    ]
    for label, value, r, c in info:
        ws.Cells(r, c).Value = label
        ws.Cells(r, c + 1).Value = value
        _apply_font(ws.Cells(r, c), bold=True, size=10, name="Segoe UI", color=(41, 65, 122))
        _apply_font(ws.Cells(r, c + 1), size=10, name="Segoe UI")

    # Table headers (row 6)
    headers = ["Date", "Day", "Start", "End", "Break", "Hours", "Project", "Notes", "Status"]
    for j, h in enumerate(headers):
        ws.Cells(6, 1 + j).Value = h
    h_range = ws.Range("A6:I6")
    _apply_fill(h_range, (41, 65, 122))
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(6).RowHeight = 28

    # Days in month
    days_in_month = calendar.monthrange(year, month)[1]
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    for d in range(1, days_in_month + 1):
        r = 6 + d
        date_obj = datetime.date(year, month, d)
        day_of_week = date_obj.weekday()

        ws.Cells(r, 1).Value = date_obj.strftime("%m/%d")
        ws.Cells(r, 2).Value = day_names[day_of_week]
        _set_alignment(ws.Cells(r, 1), h_align=XL_HALIGN_CENTER)
        _set_alignment(ws.Cells(r, 2), h_align=XL_HALIGN_CENTER)

        # Hours formula: End - Start - Break (if filled)
        ws.Cells(r, 6).Formula = f'=IF(OR(C{r}="",D{r}=""),"",(D{r}-C{r})*24-IF(E{r}="",0,E{r}))'
        ws.Cells(r, 6).NumberFormat = "0.00"

        # Time format for start/end
        ws.Cells(r, 3).NumberFormat = "HH:MM"
        ws.Cells(r, 4).NumberFormat = "HH:MM"
        ws.Cells(r, 5).NumberFormat = "0.00"

        # Weekend highlighting
        row_range = ws.Range(f"A{r}:I{r}")
        if day_of_week >= 5:
            _apply_fill(row_range, (255, 243, 205))
            _apply_font(row_range, size=10, name="Segoe UI", color=(130, 120, 80))
        else:
            bg = (248, 249, 250) if d % 2 == 0 else (255, 255, 255)
            _apply_fill(row_range, bg)
            _apply_font(row_range, size=10, name="Segoe UI")

        ws.Rows(r).RowHeight = 22

    # Summary row
    summary_row = 7 + days_in_month
    ws.Cells(summary_row, 1).Value = "TOTAL"
    ws.Range(f"A{summary_row}:E{summary_row}").Merge()
    _set_alignment(ws.Cells(summary_row, 1), h_align=XL_HALIGN_RIGHT)
    ws.Cells(summary_row, 6).Formula = f"=SUM(F7:F{summary_row - 1})"
    ws.Cells(summary_row, 6).NumberFormat = "0.00"

    total_range = ws.Range(f"A{summary_row}:I{summary_row}")
    _apply_fill(total_range, (33, 37, 41))
    _apply_font(total_range, bold=True, size=11, name="Segoe UI", color=(255, 255, 255))
    ws.Rows(summary_row).RowHeight = 30

    # Borders
    full_range = ws.Range(f"A6:I{summary_row}")
    _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Signature area
    sig_row = summary_row + 3
    ws.Cells(sig_row, 1).Value = "Employee Signature:"
    ws.Cells(sig_row, 5).Value = "Manager Approval:"
    _apply_font(ws.Cells(sig_row, 1), bold=True, size=10, name="Segoe UI")
    _apply_font(ws.Cells(sig_row, 5), bold=True, size=10, name="Segoe UI")
    ws.Range(f"B{sig_row + 1}:C{sig_row + 1}").Merge()
    ws.Range(f"F{sig_row + 1}:G{sig_row + 1}").Merge()
    _apply_borders(ws.Range(f"B{sig_row + 1}:C{sig_row + 1}"),
                   edges=[XL_BORDER_BOTTOM], weight=XL_BORDER_WEIGHT_THIN)
    _apply_borders(ws.Range(f"F{sig_row + 1}:G{sig_row + 1}"),
                   edges=[XL_BORDER_BOTTOM], weight=XL_BORDER_WEIGHT_THIN)

    # Column widths
    col_widths = [10, 6, 8, 8, 8, 8, 16, 20, 10]
    for j, w in enumerate(col_widths):
        ws.Columns(j + 1).ColumnWidth = w

    return {"sheet": ws.Name, "month": f"{month_names[month]} {year}", "days": days_in_month}


# ---------------------------------------------------------------------------
# 17. Format as Currency
# ---------------------------------------------------------------------------

def format_as_currency(sheet, range_str, currency_symbol="¥", decimal_places=0):
    """Format a range as currency.

    Args:
        sheet: Sheet name or None for active sheet.
        range_str: Range to format (e.g., "B2:B20").
        currency_symbol: Currency symbol (default ¥).
        decimal_places: Number of decimal places.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    rng = ws.Range(range_str)
    dec = "0" * decimal_places
    if decimal_places > 0:
        fmt = f'{currency_symbol}#,##0.{dec};{currency_symbol}-#,##0.{dec};{currency_symbol}0.{dec}'
    else:
        fmt = f'{currency_symbol}#,##0;{currency_symbol}-#,##0;{currency_symbol}0'
    rng.NumberFormat = fmt

    return {"sheet": ws.Name, "range": range_str, "symbol": currency_symbol, "decimals": decimal_places}


# ---------------------------------------------------------------------------
# 18. Format as Percentage
# ---------------------------------------------------------------------------

def format_as_percentage(sheet, range_str, decimal_places=1):
    """Format a range as percentage.

    Args:
        sheet: Sheet name or None for active sheet.
        range_str: Range to format (e.g., "C2:C20").
        decimal_places: Number of decimal places.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    rng = ws.Range(range_str)
    dec = "0" * decimal_places
    if decimal_places > 0:
        fmt = f"0.{dec}%"
    else:
        fmt = "0%"
    rng.NumberFormat = fmt

    return {"sheet": ws.Name, "range": range_str, "decimals": decimal_places}


# ---------------------------------------------------------------------------
# 19. Data Summary
# ---------------------------------------------------------------------------

def add_data_summary(sheet, data_range, summary_cell, summary_type="dashboard"):
    """Add a summary section pulling stats from a data range.

    Args:
        sheet: Sheet name or None for active sheet.
        data_range: Source data range (e.g., "B2:B100").
        summary_cell: Top-left cell for the summary output.
        summary_type: dashboard (grid of stats), simple (key numbers), detailed.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    col_str, row_num = _parse_cell(summary_cell)
    col_start = _col_num(col_str)

    if summary_type == "dashboard":
        stats = [
            ("SUM", "Total", (41, 65, 122)),
            ("AVERAGE", "Average", (68, 114, 196)),
            ("COUNT", "Count", (40, 167, 69)),
            ("MAX", "Maximum", (230, 126, 34)),
            ("MIN", "Minimum", (220, 53, 69)),
            ("STDEV", "Std Dev", (108, 117, 125)),
        ]

        # Title
        tc = _col_letter(col_start)
        te = _col_letter(col_start + 2)
        title_range = ws.Range(f"{tc}{row_num}:{te}{row_num}")
        title_range.Merge()
        ws.Range(f"{tc}{row_num}").Value = "Data Summary"
        _apply_fill(title_range, (33, 37, 41))
        _apply_font(title_range, bold=True, size=12, name="Segoe UI", color=(255, 255, 255))
        _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Rows(row_num).RowHeight = 30

        # Stats in 2x3 grid
        for i, (func, label, color) in enumerate(stats):
            grid_row = row_num + 1 + (i // 3) * 2
            grid_col = col_start + (i % 3)

            # Label
            cell = ws.Cells(grid_row, grid_col)
            cell.Value = label
            _apply_font(cell, bold=False, size=8, name="Segoe UI", color=(120, 120, 130))
            _set_alignment(cell, h_align=XL_HALIGN_CENTER)

            # Value
            val_cell = ws.Cells(grid_row + 1, grid_col)
            val_cell.Formula = f"={func}({data_range})"
            val_cell.NumberFormat = "#,##0.00"
            _apply_font(val_cell, bold=True, size=13, name="Segoe UI", color=color)
            _set_alignment(val_cell, h_align=XL_HALIGN_CENTER)

        # Set widths
        for c in range(3):
            ws.Columns(col_start + c).ColumnWidth = 14

    elif summary_type == "simple":
        stats = [("SUM", "Total"), ("AVERAGE", "Average"), ("COUNT", "Count")]
        for i, (func, label) in enumerate(stats):
            r = row_num + i
            ws.Cells(r, col_start).Value = f"{label}:"
            _apply_font(ws.Cells(r, col_start), bold=True, size=10, name="Segoe UI")
            ws.Cells(r, col_start + 1).Formula = f"={func}({data_range})"
            ws.Cells(r, col_start + 1).NumberFormat = "#,##0.00"
            _apply_font(ws.Cells(r, col_start + 1), bold=True, size=10, name="Segoe UI", color=(41, 65, 122))
            ws.Rows(r).RowHeight = 22

    elif summary_type == "detailed":
        stats = [
            ("SUM", "Total"), ("AVERAGE", "Average"), ("MEDIAN", "Median"),
            ("COUNT", "Count"), ("MAX", "Maximum"), ("MIN", "Minimum"),
            ("STDEV", "Std Dev"), ("VAR", "Variance"),
        ]

        # Title
        ws.Cells(row_num, col_start).Value = "Detailed Summary"
        _apply_font(ws.Cells(row_num, col_start), bold=True, size=12, name="Segoe UI", color=(41, 65, 122))
        ws.Rows(row_num).RowHeight = 28

        for i, (func, label) in enumerate(stats):
            r = row_num + 1 + i
            ws.Cells(r, col_start).Value = label
            _apply_font(ws.Cells(r, col_start), bold=True, size=10, name="Segoe UI")
            ws.Cells(r, col_start + 1).Formula = f"={func}({data_range})"
            ws.Cells(r, col_start + 1).NumberFormat = "#,##0.00"
            _apply_font(ws.Cells(r, col_start + 1), size=10, name="Segoe UI")

            bg = (248, 249, 250) if i % 2 == 0 else (255, 255, 255)
            row_range = ws.Range(f"{_col_letter(col_start)}{r}:{_col_letter(col_start + 1)}{r}")
            _apply_fill(row_range, bg)
            ws.Rows(r).RowHeight = 22

        ws.Columns(col_start).ColumnWidth = 14
        ws.Columns(col_start + 1).ColumnWidth = 16

    return {"sheet": ws.Name, "data_range": data_range, "summary_type": summary_type}


# ---------------------------------------------------------------------------
# 20. Checklist
# ---------------------------------------------------------------------------

def create_checklist(sheet, title, items, start_cell="A1"):
    """Create a checklist with checkbox column and conditional formatting.

    Args:
        sheet: Sheet name or None for active sheet.
        title: Checklist title.
        items: List of item strings.
        start_cell: Top-left cell.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    col_str, row_num = _parse_cell(start_cell)
    col_start = _col_num(col_str)

    sc = _col_letter(col_start)
    ec = _col_letter(col_start + 2)

    # Title
    title_range = ws.Range(f"{sc}{row_num}:{ec}{row_num}")
    title_range.Merge()
    ws.Range(f"{sc}{row_num}").Value = title
    _apply_fill(title_range, (41, 65, 122))
    _apply_font(title_range, bold=True, size=14, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_LEFT, v_align=XL_VALIGN_CENTER)
    title_range.IndentLevel = 1
    ws.Rows(row_num).RowHeight = 36

    # Headers
    h_row = row_num + 1
    ws.Cells(h_row, col_start).Value = "✓"
    ws.Cells(h_row, col_start + 1).Value = "#"
    ws.Cells(h_row, col_start + 2).Value = "Item"
    h_range = ws.Range(f"{sc}{h_row}:{ec}{h_row}")
    _apply_fill(h_range, (68, 114, 196))
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(h_row).RowHeight = 26

    # Items
    for i, item in enumerate(items):
        r = h_row + 1 + i
        # Checkbox column - user types "x" or "X" to check
        check_cell = ws.Cells(r, col_start)
        check_cell.Value = "☐"
        _apply_font(check_cell, size=14, name="Segoe UI")
        _set_alignment(check_cell, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)

        # Number
        ws.Cells(r, col_start + 1).Value = i + 1
        _apply_font(ws.Cells(r, col_start + 1), size=10, name="Segoe UI", color=(150, 150, 150))
        _set_alignment(ws.Cells(r, col_start + 1), h_align=XL_HALIGN_CENTER)

        # Item text
        ws.Cells(r, col_start + 2).Value = item
        _apply_font(ws.Cells(r, col_start + 2), size=10, name="Segoe UI")

        # Alternating background
        row_range = ws.Range(f"{sc}{r}:{ec}{r}")
        bg = (248, 249, 250) if i % 2 == 0 else (255, 255, 255)
        _apply_fill(row_range, bg)
        ws.Rows(r).RowHeight = 24

    # Add data validation for checkbox column (allow only ☐ and ☑)
    first_item_row = h_row + 1
    last_item_row = h_row + len(items)
    check_range = ws.Range(f"{sc}{first_item_row}:{sc}{last_item_row}")
    try:
        check_range.Validation.Delete()
        check_range.Validation.Add(
            Type=3,  # xlValidateList
            Formula1="☐,☑"
        )
    except Exception:
        pass

    # Conditional formatting: when cell = ☑, make the row item text gray + strikethrough
    item_range = ws.Range(f"{_col_letter(col_start + 2)}{first_item_row}:{_col_letter(col_start + 2)}{last_item_row}")
    try:
        for i in range(len(items)):
            r = first_item_row + i
            cell = ws.Cells(r, col_start + 2)
            fc = cell.FormatConditions.Add(
                Type=2,  # xlExpression
                Formula1=f'={sc}{r}="☑"'
            )
            fc.Font.Strikethrough = True
            fc.Font.Color = rgb(180, 180, 180)
    except Exception:
        pass

    # Progress counter
    prog_row = last_item_row + 2
    ws.Cells(prog_row, col_start).Value = "Progress:"
    _apply_font(ws.Cells(prog_row, col_start), bold=True, size=10, name="Segoe UI")
    ws.Cells(prog_row, col_start + 1).Formula = f'=COUNTIF({sc}{first_item_row}:{sc}{last_item_row},"☑")&"/"&{len(items)}'
    _apply_font(ws.Cells(prog_row, col_start + 1), bold=True, size=12, name="Segoe UI", color=(41, 65, 122))

    # Borders
    full_range = ws.Range(f"{sc}{h_row}:{ec}{last_item_row}")
    _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Column widths
    ws.Columns(col_start).ColumnWidth = 5
    ws.Columns(col_start + 1).ColumnWidth = 5
    ws.Columns(col_start + 2).ColumnWidth = 40

    return {"sheet": ws.Name, "title": title, "items": len(items)}
