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


# ---------------------------------------------------------------------------
# 22. Budget Template
# ---------------------------------------------------------------------------

def create_budget_template(sheet, title, categories, periods,
                           style="detailed"):
    """Create a budget template with categories, periods, and variance formulas.

    Args:
        sheet: Sheet name or None for active sheet.
        title: Budget title.
        categories: [{"name": "...", "subcategories": ["..."]}] for detailed,
                    or [{"name": "..."}] for summary/quarterly.
        periods: List of period labels (e.g., ["Jan", "Feb", "Mar"]).
        style: detailed (with subcategories), summary, quarterly.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    scheme = _get_scheme("corporate_blue")
    primary = scheme["primary"]
    light = scheme["light"]

    # Title row
    title_range = ws.Range(f"A1:{_col_letter(len(periods) + 3)}1")
    title_range.Merge()
    ws.Cells(1, 1).Value = title
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 40

    # Subtitle
    ws.Cells(2, 1).Value = f"Prepared: {datetime.date.today().strftime('%B %d, %Y')}"
    _apply_font(ws.Cells(2, 1), italic=True, size=9, name="Segoe UI", color=(120, 120, 130))
    ws.Rows(2).RowHeight = 22

    # Header row
    header_row = 4
    ws.Cells(header_row, 1).Value = "Category"
    for p_idx, period in enumerate(periods):
        ws.Cells(header_row, 2 + p_idx).Value = period
    budget_col = 2 + len(periods)
    actual_col = budget_col + 1
    variance_col = actual_col + 1
    ws.Cells(header_row, budget_col).Value = "Budget Total"
    ws.Cells(header_row, actual_col).Value = "Actual Total"
    ws.Cells(header_row, variance_col).Value = "Variance"

    last_col = variance_col
    header_range = ws.Range(f"A{header_row}:{_col_letter(last_col)}{header_row}")
    _apply_fill(header_range, primary)
    _apply_font(header_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(header_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(header_row).RowHeight = 28

    cur_row = header_row + 1

    if style == "detailed":
        for cat in categories:
            cat_name = cat.get("name", "")
            subcats = cat.get("subcategories", [])

            # Category header row
            ws.Cells(cur_row, 1).Value = cat_name
            cat_range = ws.Range(f"A{cur_row}:{_col_letter(last_col)}{cur_row}")
            _apply_fill(cat_range, (217, 226, 243))
            _apply_font(cat_range, bold=True, size=10, name="Segoe UI", color=(41, 65, 122))
            ws.Rows(cur_row).RowHeight = 24
            cur_row += 1

            sub_start = cur_row
            for sub in subcats:
                ws.Cells(cur_row, 1).Value = f"    {sub}"
                _apply_font(ws.Cells(cur_row, 1), size=10, name="Segoe UI")
                # Budget total formula = sum of periods
                period_start = _col_letter(2)
                period_end = _col_letter(1 + len(periods))
                ws.Cells(cur_row, budget_col).Formula = f"=SUM({period_start}{cur_row}:{period_end}{cur_row})"
                # Number format for all value cells
                for c in range(2, last_col + 1):
                    ws.Cells(cur_row, c).NumberFormat = "#,##0"
                # Variance
                bc = _col_letter(budget_col)
                ac = _col_letter(actual_col)
                ws.Cells(cur_row, variance_col).Formula = f"={ac}{cur_row}-{bc}{cur_row}"
                cur_row += 1
            sub_end = cur_row - 1

            # Subtotal row
            ws.Cells(cur_row, 1).Value = f"  Total {cat_name}"
            _apply_font(ws.Cells(cur_row, 1), bold=True, size=10, name="Segoe UI")
            for c in range(2, last_col + 1):
                cl = _col_letter(c)
                ws.Cells(cur_row, c).Formula = f"=SUM({cl}{sub_start}:{cl}{sub_end})"
                ws.Cells(cur_row, c).NumberFormat = "#,##0"
                _apply_font(ws.Cells(cur_row, c), bold=True, size=10, name="Segoe UI")
            subtotal_range = ws.Range(f"A{cur_row}:{_col_letter(last_col)}{cur_row}")
            _apply_borders(subtotal_range, edges=[XL_BORDER_TOP, XL_BORDER_BOTTOM],
                          weight=XL_BORDER_WEIGHT_MEDIUM)
            cur_row += 1

    else:  # summary or quarterly
        for cat in categories:
            cat_name = cat.get("name", "")
            ws.Cells(cur_row, 1).Value = cat_name
            _apply_font(ws.Cells(cur_row, 1), size=10, name="Segoe UI")
            period_start = _col_letter(2)
            period_end = _col_letter(1 + len(periods))
            ws.Cells(cur_row, budget_col).Formula = f"=SUM({period_start}{cur_row}:{period_end}{cur_row})"
            bc = _col_letter(budget_col)
            ac = _col_letter(actual_col)
            ws.Cells(cur_row, variance_col).Formula = f"={ac}{cur_row}-{bc}{cur_row}"
            for c in range(2, last_col + 1):
                ws.Cells(cur_row, c).NumberFormat = "#,##0"
            if cur_row % 2 == 0:
                row_range = ws.Range(f"A{cur_row}:{_col_letter(last_col)}{cur_row}")
                _apply_fill(row_range, (248, 249, 250))
            cur_row += 1

    # Grand total row
    _add_empty_row = cur_row
    ws.Cells(cur_row, 1).Value = "GRAND TOTAL"
    grand_range = ws.Range(f"A{cur_row}:{_col_letter(last_col)}{cur_row}")
    _apply_fill(grand_range, (33, 37, 41))
    _apply_font(grand_range, bold=True, size=11, name="Segoe UI", color=(255, 255, 255))
    ws.Rows(cur_row).RowHeight = 28

    for c in range(2, last_col + 1):
        cl = _col_letter(c)
        ws.Cells(cur_row, c).Formula = f"=SUM({cl}{header_row + 1}:{cl}{cur_row - 1})"
        ws.Cells(cur_row, c).NumberFormat = "#,##0"

    # Conditional formatting for variance column
    vc = _col_letter(variance_col)
    var_range = ws.Range(f"{vc}{header_row + 1}:{vc}{cur_row}")
    try:
        fc1 = var_range.FormatConditions.Add(Type=1, Operator=5, Formula1="0")  # xlGreater
        fc1.Font.Color = rgb(0, 128, 0)
        fc2 = var_range.FormatConditions.Add(Type=1, Operator=6, Formula1="0")  # xlLess
        fc2.Font.Color = rgb(220, 53, 69)
    except Exception:
        pass

    # Column widths
    ws.Columns(1).ColumnWidth = 28
    for c in range(2, last_col + 1):
        ws.Columns(c).ColumnWidth = 14

    # Borders
    full_range = ws.Range(f"A{header_row}:{_col_letter(last_col)}{cur_row}")
    _apply_borders(full_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    return {"sheet": ws.Name, "title": title, "categories": len(categories),
            "periods": len(periods), "style": style}


# ---------------------------------------------------------------------------
# 23. Project Tracker
# ---------------------------------------------------------------------------

def create_project_tracker(sheet, title, tasks, style="gantt_lite"):
    """Create a project tracking sheet.

    Args:
        sheet: Sheet name or None.
        title: Tracker title.
        tasks: [{"name": "...", "assignee": "...", "status": "Not Started",
                 "priority": "High", "start_date": "...", "end_date": "..."}]
        style: gantt_lite (with status colors), kanban (status-grouped), simple.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    primary = (41, 65, 122)
    accent = (68, 114, 196)

    STATUS_COLORS = {
        "Not Started": (220, 220, 220),
        "In Progress": (0, 123, 255),
        "Completed": (40, 167, 69),
        "On Hold": (255, 193, 7),
        "Cancelled": (220, 53, 69),
    }

    PRIORITY_COLORS = {
        "Critical": (220, 53, 69),
        "High": (255, 128, 0),
        "Medium": (255, 193, 7),
        "Low": (40, 167, 69),
    }

    # Title
    title_range = ws.Range("A1:H1")
    title_range.Merge()
    ws.Cells(1, 1).Value = title
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 40

    # Date
    ws.Cells(2, 1).Value = f"Last Updated: {datetime.date.today().strftime('%B %d, %Y')}"
    _apply_font(ws.Cells(2, 1), italic=True, size=9, name="Segoe UI", color=(120, 120, 130))

    if style in ("gantt_lite", "simple"):
        headers = ["#", "Task Name", "Assignee", "Status", "Priority", "Start Date", "End Date", "Notes"]
        header_row = 4

        for c_idx, h in enumerate(headers, 1):
            ws.Cells(header_row, c_idx).Value = h
        h_range = ws.Range(f"A{header_row}:H{header_row}")
        _apply_fill(h_range, primary)
        _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
        _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Rows(header_row).RowHeight = 28

        for i, task in enumerate(tasks):
            r = header_row + 1 + i
            ws.Cells(r, 1).Value = i + 1
            ws.Cells(r, 2).Value = task.get("name", "")
            ws.Cells(r, 3).Value = task.get("assignee", "")
            status = task.get("status", "Not Started")
            ws.Cells(r, 4).Value = status
            priority = task.get("priority", "Medium")
            ws.Cells(r, 5).Value = priority
            ws.Cells(r, 6).Value = task.get("start_date", "")
            ws.Cells(r, 7).Value = task.get("end_date", "")
            ws.Cells(r, 8).Value = task.get("notes", "")

            _apply_font(ws.Range(f"A{r}:H{r}"), size=10, name="Segoe UI")
            _set_alignment(ws.Cells(r, 1), h_align=XL_HALIGN_CENTER)

            # Status color
            if style == "gantt_lite" and status in STATUS_COLORS:
                sc = STATUS_COLORS[status]
                _apply_fill(ws.Cells(r, 4), sc)
                if status in ("In Progress", "Completed", "Cancelled"):
                    _apply_font(ws.Cells(r, 4), color=(255, 255, 255), bold=True)

            # Priority color
            if style == "gantt_lite" and priority in PRIORITY_COLORS:
                pc = PRIORITY_COLORS[priority]
                _apply_fill(ws.Cells(r, 5), pc)
                if priority in ("Critical", "High"):
                    _apply_font(ws.Cells(r, 5), color=(255, 255, 255), bold=True)

            # Alternating rows for simple
            if style == "simple" and i % 2 == 1:
                _apply_fill(ws.Range(f"A{r}:H{r}"), (248, 249, 250))

        last_row = header_row + len(tasks)

        # Borders
        _apply_borders(ws.Range(f"A{header_row}:H{last_row}"),
                      weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

        # Summary row
        summary_row = last_row + 2
        ws.Cells(summary_row, 1).Value = "Summary:"
        _apply_font(ws.Cells(summary_row, 1), bold=True, size=10, name="Segoe UI")
        for si, (status_name, sc) in enumerate(STATUS_COLORS.items()):
            col = 2 + si
            ws.Cells(summary_row, col).Value = status_name
            ws.Cells(summary_row + 1, col).Formula = f'=COUNTIF(D{header_row + 1}:D{last_row},"{status_name}")'
            _apply_font(ws.Cells(summary_row, col), bold=True, size=9, name="Segoe UI")
            _apply_fill(ws.Cells(summary_row, col), sc)
            if status_name in ("In Progress", "Completed", "Cancelled"):
                _apply_font(ws.Cells(summary_row, col), color=(255, 255, 255), bold=True, size=9)
            _apply_font(ws.Cells(summary_row + 1, col), bold=True, size=12, name="Segoe UI", color=primary)
            _set_alignment(ws.Cells(summary_row + 1, col), h_align=XL_HALIGN_CENTER)

    elif style == "kanban":
        # Group by status
        statuses = list(STATUS_COLORS.keys())
        col_offset = 1

        for s_idx, status_name in enumerate(statuses):
            col = col_offset + s_idx * 2
            status_tasks = [t for t in tasks if t.get("status", "Not Started") == status_name]

            # Status header
            ws.Cells(4, col).Value = status_name
            ws.Cells(4, col + 1).Value = ""
            h_rng = ws.Range(f"{_col_letter(col)}4:{_col_letter(col + 1)}4")
            h_rng.Merge()
            sc = STATUS_COLORS[status_name]
            _apply_fill(h_rng, sc)
            if status_name in ("In Progress", "Completed", "Cancelled"):
                _apply_font(h_rng, bold=True, size=11, name="Segoe UI", color=(255, 255, 255))
            else:
                _apply_font(h_rng, bold=True, size=11, name="Segoe UI", color=(51, 51, 51))
            _set_alignment(h_rng, h_align=XL_HALIGN_CENTER)
            ws.Rows(4).RowHeight = 28

            for t_idx, task in enumerate(status_tasks):
                r = 5 + t_idx
                ws.Cells(r, col).Value = task.get("name", "")
                ws.Cells(r, col + 1).Value = task.get("assignee", "")
                _apply_font(ws.Cells(r, col), size=10, name="Segoe UI", bold=True)
                _apply_font(ws.Cells(r, col + 1), size=9, name="Segoe UI", color=(100, 100, 100))

            ws.Columns(col).ColumnWidth = 22
            ws.Columns(col + 1).ColumnWidth = 14

    # Column widths for gantt_lite/simple
    if style in ("gantt_lite", "simple"):
        widths = [5, 30, 16, 14, 12, 13, 13, 22]
        for i, w in enumerate(widths, 1):
            ws.Columns(i).ColumnWidth = w

    return {"sheet": ws.Name, "title": title, "task_count": len(tasks), "style": style}


# ---------------------------------------------------------------------------
# 24. Expense Report
# ---------------------------------------------------------------------------

def create_expense_report(sheet, employee_name, department, expenses,
                          approval_chain=None):
    """Create an expense report.

    Args:
        sheet: Sheet name or None.
        employee_name: Employee name.
        department: Department name.
        expenses: [{"date": "...", "category": "...", "description": "...", "amount": 0}]
        approval_chain: Optional [{"name": "...", "title": "..."}]

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    primary = (41, 65, 122)

    # Title
    title_range = ws.Range("A1:F1")
    title_range.Merge()
    ws.Cells(1, 1).Value = "EXPENSE REPORT"
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=18, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 44

    # Employee info
    info_labels = ["Employee Name:", "Department:", "Date Submitted:", "Report Period:"]
    info_values = [employee_name, department,
                   datetime.date.today().strftime("%B %d, %Y"), ""]

    for i, (label, value) in enumerate(zip(info_labels, info_values)):
        r = 3 + i
        ws.Cells(r, 1).Value = label
        ws.Cells(r, 2).Value = value
        _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI")
        _apply_font(ws.Cells(r, 2), size=10, name="Segoe UI")
        _apply_fill(ws.Cells(r, 1), (217, 226, 243))

    # Expense table
    header_row = 8
    headers = ["Date", "Category", "Description", "Amount", "Receipt", "Notes"]
    for c, h in enumerate(headers, 1):
        ws.Cells(header_row, c).Value = h
    h_range = ws.Range(f"A{header_row}:F{header_row}")
    _apply_fill(h_range, primary)
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(header_row).RowHeight = 28

    for i, exp in enumerate(expenses):
        r = header_row + 1 + i
        ws.Cells(r, 1).Value = exp.get("date", "")
        ws.Cells(r, 2).Value = exp.get("category", "")
        ws.Cells(r, 3).Value = exp.get("description", "")
        ws.Cells(r, 4).Value = exp.get("amount", 0)
        ws.Cells(r, 5).Value = exp.get("receipt", "")
        ws.Cells(r, 6).Value = exp.get("notes", "")

        _apply_font(ws.Range(f"A{r}:F{r}"), size=10, name="Segoe UI")
        ws.Cells(r, 4).NumberFormat = "#,##0.00"

        if i % 2 == 1:
            _apply_fill(ws.Range(f"A{r}:F{r}"), (248, 249, 250))

    last_data_row = header_row + len(expenses)

    # Total row
    total_row = last_data_row + 1
    ws.Cells(total_row, 1).Value = "TOTAL"
    total_range = ws.Range(f"A{total_row}:F{total_row}")
    _apply_fill(total_range, (33, 37, 41))
    _apply_font(total_range, bold=True, size=11, name="Segoe UI", color=(255, 255, 255))
    ws.Cells(total_row, 4).Formula = f"=SUM(D{header_row + 1}:D{last_data_row})"
    ws.Cells(total_row, 4).NumberFormat = "#,##0.00"
    ws.Rows(total_row).RowHeight = 28

    # Borders
    _apply_borders(ws.Range(f"A{header_row}:F{total_row}"),
                  weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Approval chain
    if approval_chain:
        approve_row = total_row + 3
        ws.Cells(approve_row, 1).Value = "APPROVALS"
        _apply_font(ws.Cells(approve_row, 1), bold=True, size=12, name="Segoe UI", color=primary)
        ws.Rows(approve_row).RowHeight = 26

        for j, approver in enumerate(approval_chain):
            r = approve_row + 1 + j
            ws.Cells(r, 1).Value = approver.get("title", "")
            ws.Cells(r, 2).Value = approver.get("name", "")
            ws.Cells(r, 3).Value = "Signature: ________________"
            ws.Cells(r, 4).Value = "Date: ________"
            _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI")
            _apply_font(ws.Cells(r, 2), size=10, name="Segoe UI")
            _apply_font(ws.Cells(r, 3), size=9, name="Segoe UI", color=(120, 120, 130))
            _apply_font(ws.Cells(r, 4), size=9, name="Segoe UI", color=(120, 120, 130))

    # Column widths
    widths = [14, 16, 30, 14, 10, 20]
    for i, w in enumerate(widths, 1):
        ws.Columns(i).ColumnWidth = w

    # Calculate total for return
    total_amount = sum(e.get("amount", 0) for e in expenses)

    return {"sheet": ws.Name, "employee": employee_name, "expense_count": len(expenses),
            "total": total_amount}


# ---------------------------------------------------------------------------
# 25. Inventory Tracker
# ---------------------------------------------------------------------------

def create_inventory_tracker(sheet, title, items, style="standard"):
    """Create an inventory management sheet.

    Args:
        sheet: Sheet name or None.
        title: Tracker title.
        items: [{"sku": "...", "name": "...", "quantity": 0, "min_stock": 0,
                 "unit_price": 0, "location": "..."}]
        style: standard.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    primary = (41, 65, 122)

    # Title
    title_range = ws.Range("A1:I1")
    title_range.Merge()
    ws.Cells(1, 1).Value = title
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 40

    ws.Cells(2, 1).Value = f"Last Updated: {datetime.date.today().strftime('%B %d, %Y')}"
    _apply_font(ws.Cells(2, 1), italic=True, size=9, name="Segoe UI", color=(120, 120, 130))

    # Headers
    header_row = 4
    headers = ["SKU", "Item Name", "Category", "Quantity", "Min Stock",
               "Unit Price", "Total Value", "Location", "Status"]
    for c, h in enumerate(headers, 1):
        ws.Cells(header_row, c).Value = h
    h_range = ws.Range(f"A{header_row}:I{header_row}")
    _apply_fill(h_range, primary)
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(header_row).RowHeight = 28

    for i, item in enumerate(items):
        r = header_row + 1 + i
        ws.Cells(r, 1).Value = item.get("sku", "")
        ws.Cells(r, 2).Value = item.get("name", "")
        ws.Cells(r, 3).Value = item.get("category", "")
        qty = item.get("quantity", 0)
        min_stock = item.get("min_stock", 0)
        unit_price = item.get("unit_price", 0)
        ws.Cells(r, 4).Value = qty
        ws.Cells(r, 5).Value = min_stock
        ws.Cells(r, 6).Value = unit_price
        ws.Cells(r, 6).NumberFormat = "#,##0.00"
        # Total value formula
        ws.Cells(r, 7).Formula = f"=D{r}*F{r}"
        ws.Cells(r, 7).NumberFormat = "#,##0.00"
        ws.Cells(r, 8).Value = item.get("location", "")
        # Status formula
        ws.Cells(r, 9).Formula = f'=IF(D{r}<=0,"Out of Stock",IF(D{r}<=E{r},"Low Stock","In Stock"))'

        _apply_font(ws.Range(f"A{r}:I{r}"), size=10, name="Segoe UI")
        _set_alignment(ws.Cells(r, 4), h_align=XL_HALIGN_CENTER)
        _set_alignment(ws.Cells(r, 5), h_align=XL_HALIGN_CENTER)

        if i % 2 == 1:
            _apply_fill(ws.Range(f"A{r}:I{r}"), (248, 249, 250))

    last_row = header_row + len(items)

    # Conditional formatting for status column
    status_range = ws.Range(f"I{header_row + 1}:I{last_row}")
    try:
        fc1 = status_range.FormatConditions.Add(
            Type=2, Formula1=f'=I{header_row + 1}="Out of Stock"')
        fc1.Interior.Color = rgb(255, 200, 200)
        fc1.Font.Color = rgb(180, 0, 0)
        fc1.Font.Bold = True
        fc2 = status_range.FormatConditions.Add(
            Type=2, Formula1=f'=I{header_row + 1}="Low Stock"')
        fc2.Interior.Color = rgb(255, 243, 205)
        fc2.Font.Color = rgb(180, 100, 0)
        fc2.Font.Bold = True
        fc3 = status_range.FormatConditions.Add(
            Type=2, Formula1=f'=I{header_row + 1}="In Stock"')
        fc3.Interior.Color = rgb(212, 237, 218)
        fc3.Font.Color = rgb(0, 100, 0)
    except Exception:
        pass

    # Summary row
    summary_row = last_row + 2
    ws.Cells(summary_row, 1).Value = "SUMMARY"
    _apply_font(ws.Cells(summary_row, 1), bold=True, size=12, name="Segoe UI", color=primary)

    ws.Cells(summary_row + 1, 1).Value = "Total Items:"
    ws.Cells(summary_row + 1, 2).Formula = f"=COUNTA(B{header_row + 1}:B{last_row})"
    ws.Cells(summary_row + 2, 1).Value = "Total Value:"
    ws.Cells(summary_row + 2, 2).Formula = f"=SUM(G{header_row + 1}:G{last_row})"
    ws.Cells(summary_row + 2, 2).NumberFormat = "#,##0.00"
    ws.Cells(summary_row + 3, 1).Value = "Low Stock Items:"
    ws.Cells(summary_row + 3, 2).Formula = f'=COUNTIF(I{header_row + 1}:I{last_row},"Low Stock")'
    ws.Cells(summary_row + 4, 1).Value = "Out of Stock:"
    ws.Cells(summary_row + 4, 2).Formula = f'=COUNTIF(I{header_row + 1}:I{last_row},"Out of Stock")'

    for r in range(summary_row + 1, summary_row + 5):
        _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI")
        _apply_font(ws.Cells(r, 2), bold=True, size=10, name="Segoe UI", color=primary)

    # Borders
    _apply_borders(ws.Range(f"A{header_row}:I{last_row}"),
                  weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Column widths
    widths = [12, 24, 14, 10, 10, 12, 14, 14, 12]
    for i, w in enumerate(widths, 1):
        ws.Columns(i).ColumnWidth = w

    return {"sheet": ws.Name, "title": title, "item_count": len(items)}


# ---------------------------------------------------------------------------
# 26. Sales Report
# ---------------------------------------------------------------------------

def create_sales_report(sheet, title, sales_data, period="monthly",
                        style="dashboard"):
    """Create a sales report.

    Args:
        sheet: Sheet name or None.
        title: Report title.
        sales_data: [{"product": "...", "revenue": 0, "units": 0, "target": 0}]
        period: monthly, quarterly, yearly.
        style: dashboard, simple.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    primary = (41, 65, 122)
    accent = (68, 114, 196)

    # Title
    title_range = ws.Range("A1:H1")
    title_range.Merge()
    ws.Cells(1, 1).Value = title
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 40

    # Period and date
    ws.Cells(2, 1).Value = f"Period: {period.capitalize()} | Generated: {datetime.date.today().strftime('%B %d, %Y')}"
    _apply_font(ws.Cells(2, 1), italic=True, size=9, name="Segoe UI", color=(120, 120, 130))

    if style == "dashboard":
        # KPI cards row
        kpi_row = 4
        total_revenue = sum(d.get("revenue", 0) for d in sales_data)
        total_units = sum(d.get("units", 0) for d in sales_data)
        total_target = sum(d.get("target", 0) for d in sales_data)
        achievement = (total_revenue / total_target * 100) if total_target > 0 else 0

        kpis = [
            ("Total Revenue", f"${total_revenue:,.0f}", primary),
            ("Units Sold", f"{total_units:,}", accent),
            ("Target", f"${total_target:,.0f}", (40, 167, 69)),
            ("Achievement", f"{achievement:.1f}%", (230, 126, 34) if achievement < 100 else (40, 167, 69)),
        ]

        for k_idx, (kpi_label, kpi_value, kpi_color) in enumerate(kpis):
            col = 1 + k_idx * 2
            kpi_rng = ws.Range(f"{_col_letter(col)}{kpi_row}:{_col_letter(col + 1)}{kpi_row + 1}")
            kpi_rng.Merge()

            ws.Cells(kpi_row, col).Value = kpi_label
            _apply_font(ws.Cells(kpi_row, col), size=9, name="Segoe UI", color=(120, 120, 130))

            val_rng = ws.Range(f"{_col_letter(col)}{kpi_row + 1}:{_col_letter(col + 1)}{kpi_row + 1}")
            try:
                val_rng.Merge()
            except Exception:
                pass
            ws.Cells(kpi_row + 1, col).Value = kpi_value
            _apply_font(ws.Cells(kpi_row + 1, col), bold=True, size=16, name="Segoe UI", color=kpi_color)

            _set_alignment(ws.Cells(kpi_row, col), h_align=XL_HALIGN_CENTER)
            _set_alignment(ws.Cells(kpi_row + 1, col), h_align=XL_HALIGN_CENTER)

        data_start_row = kpi_row + 3
    else:
        data_start_row = 4

    # Data table
    headers = ["Product", "Revenue", "Units", "Target", "Variance", "% of Target", "Avg Price", "Rank"]
    for c, h in enumerate(headers, 1):
        ws.Cells(data_start_row, c).Value = h
    h_range = ws.Range(f"A{data_start_row}:H{data_start_row}")
    _apply_fill(h_range, primary)
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(data_start_row).RowHeight = 28

    for i, sd in enumerate(sales_data):
        r = data_start_row + 1 + i
        revenue = sd.get("revenue", 0)
        units = sd.get("units", 0)
        target = sd.get("target", 0)
        variance = revenue - target
        pct = (revenue / target * 100) if target > 0 else 0
        avg_price = (revenue / units) if units > 0 else 0

        ws.Cells(r, 1).Value = sd.get("product", "")
        ws.Cells(r, 2).Value = revenue
        ws.Cells(r, 3).Value = units
        ws.Cells(r, 4).Value = target
        ws.Cells(r, 5).Value = variance
        ws.Cells(r, 6).Value = pct / 100
        ws.Cells(r, 7).Value = avg_price

        ws.Cells(r, 2).NumberFormat = "#,##0"
        ws.Cells(r, 3).NumberFormat = "#,##0"
        ws.Cells(r, 4).NumberFormat = "#,##0"
        ws.Cells(r, 5).NumberFormat = "#,##0"
        ws.Cells(r, 6).NumberFormat = "0.0%"
        ws.Cells(r, 7).NumberFormat = "#,##0.00"

        _apply_font(ws.Range(f"A{r}:H{r}"), size=10, name="Segoe UI")

        # Variance color
        if variance >= 0:
            _apply_font(ws.Cells(r, 5), color=(40, 167, 69))
        else:
            _apply_font(ws.Cells(r, 5), color=(220, 53, 69))

        if i % 2 == 1:
            _apply_fill(ws.Range(f"A{r}:H{r}"), (248, 249, 250))

    last_data = data_start_row + len(sales_data)

    # Rank formula
    for i in range(len(sales_data)):
        r = data_start_row + 1 + i
        ws.Cells(r, 8).Formula = f"=RANK(B{r},B{data_start_row + 1}:B{last_data})"
        _set_alignment(ws.Cells(r, 8), h_align=XL_HALIGN_CENTER)

    # Total row
    total_row = last_data + 1
    ws.Cells(total_row, 1).Value = "TOTAL"
    for c in [2, 3, 4, 5]:
        cl = _col_letter(c)
        ws.Cells(total_row, c).Formula = f"=SUM({cl}{data_start_row + 1}:{cl}{last_data})"
    ws.Cells(total_row, 6).Formula = f"=IF(D{total_row}>0,B{total_row}/D{total_row},0)"
    ws.Cells(total_row, 7).Formula = f"=IF(C{total_row}>0,B{total_row}/C{total_row},0)"
    for c in [2, 3, 4, 5]:
        ws.Cells(total_row, c).NumberFormat = "#,##0"
    ws.Cells(total_row, 6).NumberFormat = "0.0%"
    ws.Cells(total_row, 7).NumberFormat = "#,##0.00"

    t_range = ws.Range(f"A{total_row}:H{total_row}")
    _apply_fill(t_range, (33, 37, 41))
    _apply_font(t_range, bold=True, size=11, name="Segoe UI", color=(255, 255, 255))
    ws.Rows(total_row).RowHeight = 28

    # Borders
    _apply_borders(ws.Range(f"A{data_start_row}:H{total_row}"),
                  weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Column widths
    widths = [20, 14, 10, 14, 14, 12, 12, 8]
    for i, w in enumerate(widths, 1):
        ws.Columns(i).ColumnWidth = w

    return {"sheet": ws.Name, "title": title, "product_count": len(sales_data),
            "total_revenue": sum(d.get("revenue", 0) for d in sales_data),
            "period": period, "style": style}


# ---------------------------------------------------------------------------
# 27. Employee Roster
# ---------------------------------------------------------------------------

def create_employee_roster(sheet, title, employees, style="detailed"):
    """Create an employee directory/roster.

    Args:
        sheet: Sheet name or None.
        title: Roster title.
        employees: [{"name": "...", "department": "...", "position": "...",
                     "email": "...", "phone": "..."}]
        style: detailed, compact.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    primary = (41, 65, 122)

    # Title
    title_range = ws.Range("A1:G1")
    title_range.Merge()
    ws.Cells(1, 1).Value = title
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 40

    ws.Cells(2, 1).Value = f"Total Employees: {len(employees)} | Updated: {datetime.date.today().strftime('%B %d, %Y')}"
    _apply_font(ws.Cells(2, 1), italic=True, size=9, name="Segoe UI", color=(120, 120, 130))

    header_row = 4
    if style == "detailed":
        headers = ["#", "Name", "Department", "Position", "Email", "Phone", "Start Date"]
    else:
        headers = ["#", "Name", "Department", "Position", "Email", "Phone", "Start Date"]

    for c, h in enumerate(headers, 1):
        ws.Cells(header_row, c).Value = h
    last_col = len(headers)
    h_range = ws.Range(f"A{header_row}:{_col_letter(last_col)}{header_row}")
    _apply_fill(h_range, primary)
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(header_row).RowHeight = 28

    for i, emp in enumerate(employees):
        r = header_row + 1 + i
        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = emp.get("name", "")
        ws.Cells(r, 3).Value = emp.get("department", "")
        ws.Cells(r, 4).Value = emp.get("position", "")
        ws.Cells(r, 5).Value = emp.get("email", "")
        ws.Cells(r, 6).Value = emp.get("phone", "")
        ws.Cells(r, 7).Value = emp.get("start_date", "")

        _apply_font(ws.Range(f"A{r}:{_col_letter(last_col)}{r}"), size=10, name="Segoe UI")
        _set_alignment(ws.Cells(r, 1), h_align=XL_HALIGN_CENTER)
        _apply_font(ws.Cells(r, 2), bold=True)

        if i % 2 == 1:
            _apply_fill(ws.Range(f"A{r}:{_col_letter(last_col)}{r}"), (248, 249, 250))

    last_row = header_row + len(employees)

    # Borders
    _apply_borders(ws.Range(f"A{header_row}:{_col_letter(last_col)}{last_row}"),
                  weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Department summary
    summary_row = last_row + 3
    ws.Cells(summary_row, 1).Value = "Department Summary"
    _apply_font(ws.Cells(summary_row, 1), bold=True, size=12, name="Segoe UI", color=primary)
    ws.Rows(summary_row).RowHeight = 26

    departments = list(set(e.get("department", "") for e in employees if e.get("department")))
    departments.sort()
    ws.Cells(summary_row + 1, 1).Value = "Department"
    ws.Cells(summary_row + 1, 2).Value = "Count"
    _apply_font(ws.Range(f"A{summary_row + 1}:B{summary_row + 1}"), bold=True, size=10, name="Segoe UI")
    _apply_fill(ws.Range(f"A{summary_row + 1}:B{summary_row + 1}"), (217, 226, 243))

    for d_idx, dept in enumerate(departments):
        r = summary_row + 2 + d_idx
        ws.Cells(r, 1).Value = dept
        ws.Cells(r, 2).Formula = f'=COUNTIF(C{header_row + 1}:C{last_row},"{dept}")'
        _apply_font(ws.Cells(r, 1), size=10, name="Segoe UI")
        _apply_font(ws.Cells(r, 2), bold=True, size=10, name="Segoe UI", color=primary)
        _set_alignment(ws.Cells(r, 2), h_align=XL_HALIGN_CENTER)

    # Column widths
    widths = [5, 22, 18, 20, 26, 16, 12]
    for i, w in enumerate(widths, 1):
        ws.Columns(i).ColumnWidth = w

    return {"sheet": ws.Name, "title": title, "employee_count": len(employees),
            "department_count": len(departments)}


# ---------------------------------------------------------------------------
# 28. Risk Matrix
# ---------------------------------------------------------------------------

def create_risk_matrix(sheet, title, risks, style="heatmap"):
    """Create a risk assessment matrix.

    Args:
        sheet: Sheet name or None.
        title: Matrix title.
        risks: [{"name": "...", "probability": 1-5, "impact": 1-5, "mitigation": "..."}]
        style: heatmap, simple.

    Returns:
        dict with status info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    primary = (41, 65, 122)

    RISK_COLORS = {
        "Critical": (220, 53, 69),
        "High": (255, 128, 0),
        "Medium": (255, 193, 7),
        "Low": (40, 167, 69),
        "Very Low": (144, 202, 249),
    }

    def _risk_level(prob, impact):
        score = prob * impact
        if score >= 20:
            return "Critical"
        elif score >= 12:
            return "High"
        elif score >= 8:
            return "Medium"
        elif score >= 4:
            return "Low"
        else:
            return "Very Low"

    # Title
    title_range = ws.Range("A1:H1")
    title_range.Merge()
    ws.Cells(1, 1).Value = title
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=16, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 40

    ws.Cells(2, 1).Value = f"Assessment Date: {datetime.date.today().strftime('%B %d, %Y')}"
    _apply_font(ws.Cells(2, 1), italic=True, size=9, name="Segoe UI", color=(120, 120, 130))

    # Risk register table
    header_row = 4
    headers = ["#", "Risk Name", "Probability (1-5)", "Impact (1-5)",
               "Risk Score", "Risk Level", "Mitigation Strategy", "Owner"]
    for c, h in enumerate(headers, 1):
        ws.Cells(header_row, c).Value = h
    h_range = ws.Range(f"A{header_row}:H{header_row}")
    _apply_fill(h_range, primary)
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER, wrap=True)
    ws.Rows(header_row).RowHeight = 32

    for i, risk in enumerate(risks):
        r = header_row + 1 + i
        prob = risk.get("probability", 1)
        impact = risk.get("impact", 1)
        score = prob * impact
        level = _risk_level(prob, impact)

        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = risk.get("name", "")
        ws.Cells(r, 3).Value = prob
        ws.Cells(r, 4).Value = impact
        ws.Cells(r, 5).Formula = f"=C{r}*D{r}"
        ws.Cells(r, 6).Value = level
        ws.Cells(r, 7).Value = risk.get("mitigation", "")
        ws.Cells(r, 8).Value = risk.get("owner", "")

        _apply_font(ws.Range(f"A{r}:H{r}"), size=10, name="Segoe UI")
        _set_alignment(ws.Cells(r, 1), h_align=XL_HALIGN_CENTER)
        _set_alignment(ws.Cells(r, 3), h_align=XL_HALIGN_CENTER)
        _set_alignment(ws.Cells(r, 4), h_align=XL_HALIGN_CENTER)
        _set_alignment(ws.Cells(r, 5), h_align=XL_HALIGN_CENTER)
        _set_alignment(ws.Cells(r, 6), h_align=XL_HALIGN_CENTER)

        # Color the risk level cell
        if style == "heatmap" and level in RISK_COLORS:
            rc = RISK_COLORS[level]
            _apply_fill(ws.Cells(r, 6), rc)
            if level in ("Critical", "High"):
                _apply_font(ws.Cells(r, 6), bold=True, color=(255, 255, 255))
            else:
                _apply_font(ws.Cells(r, 6), bold=True)

            # Also color score cell
            _apply_fill(ws.Cells(r, 5), rc)
            if level in ("Critical", "High"):
                _apply_font(ws.Cells(r, 5), bold=True, color=(255, 255, 255))

    last_row = header_row + len(risks)

    # Borders
    _apply_borders(ws.Range(f"A{header_row}:H{last_row}"),
                  weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # 5x5 heatmap matrix
    if style == "heatmap":
        matrix_start_row = last_row + 3
        ws.Cells(matrix_start_row, 1).Value = "Risk Heatmap Matrix"
        _apply_font(ws.Cells(matrix_start_row, 1), bold=True, size=12, name="Segoe UI", color=primary)

        # Matrix header: Impact 1-5 across columns
        for imp in range(1, 6):
            ws.Cells(matrix_start_row + 1, 2 + imp - 1).Value = f"Impact {imp}"
            _apply_font(ws.Cells(matrix_start_row + 1, 2 + imp - 1), bold=True, size=9, name="Segoe UI")
            _set_alignment(ws.Cells(matrix_start_row + 1, 2 + imp - 1), h_align=XL_HALIGN_CENTER)

        # Matrix rows: Probability 5 (top) to 1 (bottom)
        for prob in range(5, 0, -1):
            mr = matrix_start_row + 2 + (5 - prob)
            ws.Cells(mr, 1).Value = f"Prob {prob}"
            _apply_font(ws.Cells(mr, 1), bold=True, size=9, name="Segoe UI")
            _set_alignment(ws.Cells(mr, 1), h_align=XL_HALIGN_CENTER)

            for imp in range(1, 6):
                mc = 2 + imp - 1
                score = prob * imp
                level = _risk_level(prob, imp)
                ws.Cells(mr, mc).Value = score
                _set_alignment(ws.Cells(mr, mc), h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
                _apply_font(ws.Cells(mr, mc), bold=True, size=10, name="Segoe UI")

                rc = RISK_COLORS.get(level, (200, 200, 200))
                _apply_fill(ws.Cells(mr, mc), rc)
                if level in ("Critical", "High"):
                    _apply_font(ws.Cells(mr, mc), color=(255, 255, 255), bold=True)

                ws.Rows(mr).RowHeight = 28
                ws.Columns(mc).ColumnWidth = 12

        # Borders on matrix
        m_range = ws.Range(f"A{matrix_start_row + 1}:F{matrix_start_row + 7}")
        _apply_borders(m_range, weight=XL_BORDER_WEIGHT_THIN, color=(180, 180, 180))

    # Legend
    legend_row = (last_row + 10) if style == "heatmap" else (last_row + 3)
    ws.Cells(legend_row, 1).Value = "Legend:"
    _apply_font(ws.Cells(legend_row, 1), bold=True, size=10, name="Segoe UI")
    for l_idx, (level_name, lc) in enumerate(RISK_COLORS.items()):
        col = 2 + l_idx
        ws.Cells(legend_row, col).Value = level_name
        _apply_fill(ws.Cells(legend_row, col), lc)
        if level_name in ("Critical", "High"):
            _apply_font(ws.Cells(legend_row, col), bold=True, size=9, name="Segoe UI", color=(255, 255, 255))
        else:
            _apply_font(ws.Cells(legend_row, col), bold=True, size=9, name="Segoe UI")
        _set_alignment(ws.Cells(legend_row, col), h_align=XL_HALIGN_CENTER)

    # Column widths
    widths = [5, 26, 14, 12, 10, 12, 30, 14]
    for i, w in enumerate(widths, 1):
        ws.Columns(i).ColumnWidth = w

    return {"sheet": ws.Name, "title": title, "risk_count": len(risks), "style": style}


# ---------------------------------------------------------------------------
# 29. Attendance Tracker
# ---------------------------------------------------------------------------

def create_attendance_tracker(sheet, title, employees, month, year,
                              style="calendar"):
    """Create a monthly attendance tracking sheet.

    Args:
        sheet: Sheet name or None.
        title: Tracker title.
        employees: List of employee names.
        month: Month number (1-12).
        year: Year (e.g., 2025).
        style: calendar.

    Returns:
        dict with status info.
    """
    import calendar as cal_mod

    app = _get_app()
    ws = _get_ws(app, sheet)
    primary = (41, 65, 122)
    accent = (68, 114, 196)

    days_in_month = cal_mod.monthrange(year, month)[1]
    month_name = cal_mod.month_name[month]
    day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    # Title
    last_col = 3 + days_in_month  # cols: #, Name, then days, then totals
    title_range = ws.Range(f"A1:{_col_letter(last_col)}1")
    title_range.Merge()
    ws.Cells(1, 1).Value = f"{title} - {month_name} {year}"
    _apply_fill(title_range, primary)
    _apply_font(title_range, bold=True, size=14, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(title_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(1).RowHeight = 36

    # Legend row
    ws.Cells(2, 1).Value = "Legend: P=Present  A=Absent  L=Leave  H=Holiday  WFH=Work From Home"
    _apply_font(ws.Cells(2, 1), italic=True, size=8, name="Segoe UI", color=(100, 100, 100))

    # Header row
    header_row = 4
    ws.Cells(header_row, 1).Value = "#"
    ws.Cells(header_row, 2).Value = "Employee Name"

    # Day headers with day-of-week
    for d in range(1, days_in_month + 1):
        col = 2 + d
        dow = cal_mod.weekday(year, month, d)
        ws.Cells(header_row, col).Value = d
        ws.Cells(header_row - 1, col).Value = day_names[dow][:2]
        _apply_font(ws.Cells(header_row - 1, col), size=7, name="Segoe UI", color=(120, 120, 130))
        _set_alignment(ws.Cells(header_row - 1, col), h_align=XL_HALIGN_CENTER)

        # Weekend columns get different background
        if dow >= 5:  # Saturday=5, Sunday=6
            for er in range(header_row, header_row + len(employees) + 1):
                _apply_fill(ws.Cells(er, col), (240, 240, 245))

    total_col = 3 + days_in_month
    ws.Cells(header_row, total_col).Value = "Present"

    # Header formatting
    h_range = ws.Range(f"A{header_row}:{_col_letter(total_col)}{header_row}")
    _apply_fill(h_range, primary)
    _apply_font(h_range, bold=True, size=9, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
    ws.Rows(header_row).RowHeight = 24

    # Employee rows
    for i, emp_name in enumerate(employees):
        r = header_row + 1 + i
        ws.Cells(r, 1).Value = i + 1
        ws.Cells(r, 2).Value = emp_name
        _apply_font(ws.Cells(r, 2), size=10, name="Segoe UI", bold=True)
        _set_alignment(ws.Cells(r, 1), h_align=XL_HALIGN_CENTER)

        # Present count formula
        day_start_col = _col_letter(3)
        day_end_col = _col_letter(2 + days_in_month)
        ws.Cells(r, total_col).Formula = f'=COUNTIF({day_start_col}{r}:{day_end_col}{r},"P")+COUNTIF({day_start_col}{r}:{day_end_col}{r},"WFH")'
        _apply_font(ws.Cells(r, total_col), bold=True, size=10, name="Segoe UI", color=primary)
        _set_alignment(ws.Cells(r, total_col), h_align=XL_HALIGN_CENTER)

        # Set alignment for day cells
        for d in range(1, days_in_month + 1):
            col = 2 + d
            _set_alignment(ws.Cells(r, col), h_align=XL_HALIGN_CENTER)
            _apply_font(ws.Cells(r, col), size=9, name="Segoe UI")

        if i % 2 == 1:
            row_range = ws.Range(f"A{r}:B{r}")
            _apply_fill(row_range, (248, 249, 250))

    last_row = header_row + len(employees)

    # Data validation for attendance cells
    att_range = ws.Range(f"C{header_row + 1}:{_col_letter(2 + days_in_month)}{last_row}")
    try:
        att_range.Validation.Delete()
        att_range.Validation.Add(
            Type=3,  # xlValidateList
            Formula1="P,A,L,H,WFH"
        )
    except Exception:
        pass

    # Conditional formatting for attendance
    try:
        fc_p = att_range.FormatConditions.Add(Type=1, Operator=3, Formula1='"P"')
        fc_p.Interior.Color = rgb(212, 237, 218)
        fc_p.Font.Color = rgb(0, 100, 0)
        fc_a = att_range.FormatConditions.Add(Type=1, Operator=3, Formula1='"A"')
        fc_a.Interior.Color = rgb(255, 200, 200)
        fc_a.Font.Color = rgb(180, 0, 0)
        fc_l = att_range.FormatConditions.Add(Type=1, Operator=3, Formula1='"L"')
        fc_l.Interior.Color = rgb(255, 243, 205)
        fc_l.Font.Color = rgb(180, 100, 0)
        fc_h = att_range.FormatConditions.Add(Type=1, Operator=3, Formula1='"H"')
        fc_h.Interior.Color = rgb(217, 226, 243)
        fc_h.Font.Color = rgb(41, 65, 122)
        fc_w = att_range.FormatConditions.Add(Type=1, Operator=3, Formula1='"WFH"')
        fc_w.Interior.Color = rgb(232, 245, 233)
        fc_w.Font.Color = rgb(0, 128, 0)
    except Exception:
        pass

    # Summary row
    summary_row = last_row + 2
    ws.Cells(summary_row, 2).Value = "Daily Present Count:"
    _apply_font(ws.Cells(summary_row, 2), bold=True, size=9, name="Segoe UI")
    for d in range(1, days_in_month + 1):
        col = 2 + d
        cl = _col_letter(col)
        ws.Cells(summary_row, col).Formula = f'=COUNTIF({cl}{header_row + 1}:{cl}{last_row},"P")+COUNTIF({cl}{header_row + 1}:{cl}{last_row},"WFH")'
        _apply_font(ws.Cells(summary_row, col), bold=True, size=9, name="Segoe UI", color=primary)
        _set_alignment(ws.Cells(summary_row, col), h_align=XL_HALIGN_CENTER)

    # Borders
    _apply_borders(ws.Range(f"A{header_row}:{_col_letter(total_col)}{last_row}"),
                  weight=XL_BORDER_WEIGHT_THIN, color=(200, 200, 210))

    # Column widths
    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 20
    for d in range(1, days_in_month + 1):
        ws.Columns(2 + d).ColumnWidth = 4
    ws.Columns(total_col).ColumnWidth = 8

    return {"sheet": ws.Name, "title": title, "employee_count": len(employees),
            "month": f"{month_name} {year}", "days": days_in_month}


# ---------------------------------------------------------------------------
# 19. KPI Dashboard
# ---------------------------------------------------------------------------

def create_kpi_dashboard(
    sheet, title: str, kpis: list[dict],
    chart_config: dict | None = None,
    style: str = "executive",
) -> dict:
    """Create a full KPI dashboard with header, KPI cards, and optional chart area.

    Args:
        sheet: Target worksheet name or None for active.
        title: Dashboard title.
        kpis: List of dicts with "name", "value", "target", "unit", "status"
              (green/yellow/red).
        chart_config: Optional dict (reserved for future chart integration).
        style: "executive" (corporate dark), "modern" (clean light),
               "compact" (minimal space).

    Returns:
        dict with sheet, title, kpi_count, style.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    style_configs = {
        "executive": {"primary": (41, 65, 122), "accent": (68, 114, 196),
                      "bg": (248, 249, 252), "card_bg": (255, 255, 255),
                      "text": (33, 37, 41)},
        "modern": {"primary": (33, 37, 41), "accent": (0, 123, 255),
                   "bg": (255, 255, 255), "card_bg": (248, 249, 250),
                   "text": (33, 37, 41)},
        "compact": {"primary": (0, 105, 148), "accent": (0, 151, 167),
                    "bg": (245, 248, 252), "card_bg": (255, 255, 255),
                    "text": (44, 62, 80)},
    }
    s = style_configs.get(style, style_configs["executive"])
    primary = s["primary"]
    accent = s["accent"]

    status_colors = {
        "green": (40, 167, 69),
        "yellow": (255, 193, 7),
        "red": (220, 53, 69),
    }

    # Background
    ws.Range("A1:O40").Interior.Color = rgb(*s["bg"])

    # Header band
    header_range = ws.Range("A1:O2")
    _apply_fill(header_range, primary)
    ws.Rows(1).RowHeight = 12
    ws.Rows(2).RowHeight = 36

    ws.Cells(2, 2).Value = title
    _apply_font(ws.Cells(2, 2), bold=True, size=18, name="Segoe UI",
                color=(255, 255, 255))

    # Date
    today_str = datetime.date.today().strftime("%B %d, %Y")
    ws.Cells(2, 12).Value = today_str
    _apply_font(ws.Cells(2, 12), size=10, name="Segoe UI", color=(200, 210, 230))
    _set_alignment(ws.Cells(2, 12), h_align=XL_HALIGN_RIGHT)

    # Accent line under header
    ws.Rows(3).RowHeight = 4
    _apply_fill(ws.Range("A3:O3"), accent)

    # KPI Cards
    ws.Rows(4).RowHeight = 8  # spacer
    n = len(kpis)
    cols_per_card = max(2, 14 // n)
    card_start_row = 5

    ws.Rows(card_start_row).RowHeight = 12
    ws.Rows(card_start_row + 1).RowHeight = 36
    ws.Rows(card_start_row + 2).RowHeight = 20
    ws.Rows(card_start_row + 3).RowHeight = 16
    ws.Rows(card_start_row + 4).RowHeight = 4  # status bar

    for i, kpi in enumerate(kpis):
        start_col = 2 + i * cols_per_card
        end_col = start_col + cols_per_card - 1
        end_col_l = _col_letter(end_col)
        start_col_l = _col_letter(start_col)

        # Card background
        card_range = ws.Range(
            f"{start_col_l}{card_start_row}:{end_col_l}{card_start_row + 3}")
        _apply_fill(card_range, s["card_bg"])
        _apply_borders(card_range, weight=XL_BORDER_WEIGHT_THIN,
                      color=(220, 225, 235))

        # KPI name
        ws.Cells(card_start_row, start_col).Value = kpi.get("name", "")
        _apply_font(ws.Cells(card_start_row, start_col), bold=True, size=9,
                    name="Segoe UI", color=(100, 110, 130))
        _set_alignment(ws.Cells(card_start_row, start_col), h_align=XL_HALIGN_CENTER)
        ws.Range(f"{start_col_l}{card_start_row}:{end_col_l}{card_start_row}").MergeCells = True

        # KPI value
        value_str = str(kpi.get("value", ""))
        unit = kpi.get("unit", "")
        ws.Cells(card_start_row + 1, start_col).Value = f"{value_str}{unit}"
        _apply_font(ws.Cells(card_start_row + 1, start_col), bold=True, size=22,
                    name="Segoe UI", color=s["text"])
        _set_alignment(ws.Cells(card_start_row + 1, start_col),
                      h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Range(f"{start_col_l}{card_start_row + 1}:{end_col_l}{card_start_row + 1}").MergeCells = True

        # Target
        target = kpi.get("target", "")
        if target:
            ws.Cells(card_start_row + 2, start_col).Value = f"Target: {target}{unit}"
            _apply_font(ws.Cells(card_start_row + 2, start_col), size=9,
                        name="Segoe UI", color=(130, 140, 155))
            _set_alignment(ws.Cells(card_start_row + 2, start_col),
                          h_align=XL_HALIGN_CENTER)
            ws.Range(f"{start_col_l}{card_start_row + 2}:{end_col_l}{card_start_row + 2}").MergeCells = True

        # Status indicator bar
        status = kpi.get("status", "green")
        status_color = status_colors.get(status, (100, 100, 100))
        status_range = ws.Range(
            f"{start_col_l}{card_start_row + 3}:{end_col_l}{card_start_row + 3}")
        _apply_fill(status_range, status_color)
        ws.Rows(card_start_row + 3).RowHeight = 4

    # Column widths
    ws.Columns(1).ColumnWidth = 2
    for c in range(2, 16):
        ws.Columns(c).ColumnWidth = 10

    return {"sheet": ws.Name, "title": title, "kpi_count": n, "style": style}


# ---------------------------------------------------------------------------
# 20. Vendor Comparison
# ---------------------------------------------------------------------------

def create_vendor_comparison(
    sheet, title: str, vendors: list[str], criteria: list[str],
    scores: list[list], style: str = "weighted",
) -> dict:
    """Create a vendor scoring/comparison matrix.

    Args:
        sheet: Target worksheet name or None.
        title: Title of the comparison.
        vendors: List of vendor names.
        criteria: List of evaluation criteria.
        scores: 2D list [criteria_index][vendor_index] of scores (1-10).
        style: "weighted", "simple", "visual".

    Returns:
        dict with sheet, title, vendor_count, criteria_count, style.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    primary = (41, 65, 122)
    accent = (68, 114, 196)

    # Title
    ws.Range("A1:J1").MergeCells = True
    ws.Cells(1, 1).Value = title
    _apply_font(ws.Cells(1, 1), bold=True, size=16, name="Segoe UI", color=primary)
    ws.Rows(1).RowHeight = 30

    # Accent bar
    ws.Rows(2).RowHeight = 3
    col_end = 2 + len(vendors) + (1 if style == "weighted" else 0)
    _apply_fill(ws.Range(f"A2:{_col_letter(col_end)}2"), accent)

    header_row = 3
    ws.Rows(header_row).RowHeight = 28

    # Column headers
    ws.Cells(header_row, 1).Value = "Criteria"
    if style == "weighted":
        ws.Cells(header_row, 2).Value = "Weight"
        v_start_col = 3
    else:
        v_start_col = 2

    for j, vendor in enumerate(vendors):
        ws.Cells(header_row, v_start_col + j).Value = vendor

    # Total column
    total_col = v_start_col + len(vendors)
    if style == "weighted":
        # No total header for weighted (totals per vendor at bottom)
        pass

    # Header formatting
    h_range = ws.Range(f"A{header_row}:{_col_letter(total_col - 1)}{header_row}")
    _apply_fill(h_range, primary)
    _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
    _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)

    # Data rows
    for i, criterion in enumerate(criteria):
        r = header_row + 1 + i
        ws.Cells(r, 1).Value = criterion
        _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI", color=primary)
        ws.Rows(r).RowHeight = 24

        if style == "weighted":
            # Default equal weights
            weight = round(1.0 / len(criteria), 2)
            ws.Cells(r, 2).Value = weight
            _apply_font(ws.Cells(r, 2), size=10, name="Segoe UI", color=(100, 100, 100))
            _set_alignment(ws.Cells(r, 2), h_align=XL_HALIGN_CENTER)

        for j in range(len(vendors)):
            score_val = scores[i][j] if i < len(scores) and j < len(scores[i]) else 0
            ws.Cells(r, v_start_col + j).Value = score_val
            _set_alignment(ws.Cells(r, v_start_col + j), h_align=XL_HALIGN_CENTER)
            _apply_font(ws.Cells(r, v_start_col + j), size=11, name="Segoe UI")

            # Conditional color for scores
            if style == "visual":
                if score_val >= 8:
                    _apply_fill(ws.Cells(r, v_start_col + j), (212, 237, 218))
                elif score_val >= 5:
                    _apply_fill(ws.Cells(r, v_start_col + j), (255, 243, 205))
                else:
                    _apply_fill(ws.Cells(r, v_start_col + j), (255, 220, 220))

        # Alternating row color
        if i % 2 == 1 and style != "visual":
            row_range = ws.Range(f"A{r}:{_col_letter(v_start_col + len(vendors) - 1)}{r}")
            _apply_fill(row_range, (240, 243, 250))

    # Total row
    total_row = header_row + len(criteria) + 1
    ws.Cells(total_row, 1).Value = "TOTAL"
    _apply_font(ws.Cells(total_row, 1), bold=True, size=11, name="Segoe UI", color=primary)
    ws.Rows(total_row).RowHeight = 28

    data_start_r = header_row + 1
    data_end_r = header_row + len(criteria)
    for j in range(len(vendors)):
        col = v_start_col + j
        cl = _col_letter(col)
        if style == "weighted":
            # Weighted sum: SUMPRODUCT(weights, scores)
            w_cl = _col_letter(2)
            ws.Cells(total_row, col).Formula = (
                f"=SUMPRODUCT({w_cl}{data_start_r}:{w_cl}{data_end_r},"
                f"{cl}{data_start_r}:{cl}{data_end_r})"
            )
        else:
            ws.Cells(total_row, col).Formula = f"=SUM({cl}{data_start_r}:{cl}{data_end_r})"

        _apply_font(ws.Cells(total_row, col), bold=True, size=12, name="Segoe UI",
                    color=accent)
        _set_alignment(ws.Cells(total_row, col), h_align=XL_HALIGN_CENTER)

    _apply_fill(ws.Range(f"A{total_row}:{_col_letter(v_start_col + len(vendors) - 1)}{total_row}"),
               (230, 235, 245))

    # Borders
    data_range = ws.Range(f"A{header_row}:{_col_letter(v_start_col + len(vendors) - 1)}{total_row}")
    _apply_borders(data_range, weight=XL_BORDER_WEIGHT_THIN, color=(200, 205, 215))

    # Column widths
    ws.Columns(1).ColumnWidth = 22
    if style == "weighted":
        ws.Columns(2).ColumnWidth = 10
    for j in range(len(vendors)):
        ws.Columns(v_start_col + j).ColumnWidth = 14

    return {"sheet": ws.Name, "title": title, "vendor_count": len(vendors),
            "criteria_count": len(criteria), "style": style}


# ---------------------------------------------------------------------------
# 21. Cash Flow Statement
# ---------------------------------------------------------------------------

def create_cash_flow_statement(
    sheet, title: str,
    operating: list[dict], investing: list[dict], financing: list[dict],
    periods: list[str],
) -> dict:
    """Create a cash flow statement.

    Args:
        sheet: Target worksheet name or None.
        title: Statement title.
        operating: List of {"item": "...", "values": [...]}.
        investing: List of {"item": "...", "values": [...]}.
        financing: List of {"item": "...", "values": [...]}.
        periods: Period header labels.

    Returns:
        dict with sheet, title, period_count.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    primary = (0, 51, 102)
    accent = (0, 112, 192)

    # Title
    num_periods = len(periods)
    end_col = 2 + num_periods
    end_cl = _col_letter(end_col)

    ws.Range(f"A1:{end_cl}1").MergeCells = True
    ws.Cells(1, 1).Value = title
    _apply_font(ws.Cells(1, 1), bold=True, size=16, name="Segoe UI", color=primary)
    ws.Rows(1).RowHeight = 32

    ws.Rows(2).RowHeight = 3
    _apply_fill(ws.Range(f"A2:{end_cl}2"), accent)

    r = 3
    ws.Rows(r).RowHeight = 24

    # Period headers
    ws.Cells(r, 1).Value = "(in thousands)"
    _apply_font(ws.Cells(r, 1), italic=True, size=9, name="Segoe UI", color=(120, 120, 130))

    for j, period in enumerate(periods):
        ws.Cells(r, 2 + j).Value = period
        _apply_font(ws.Cells(r, 2 + j), bold=True, size=10, name="Segoe UI",
                    color=(255, 255, 255))
        _set_alignment(ws.Cells(r, 2 + j), h_align=XL_HALIGN_RIGHT)

    _apply_fill(ws.Range(f"A{r}:{end_cl}{r}"), primary)

    def _write_section(start_row, section_name, items, section_color):
        row = start_row
        # Section header
        ws.Cells(row, 1).Value = section_name
        _apply_font(ws.Cells(row, 1), bold=True, size=11, name="Segoe UI",
                    color=section_color)
        _apply_fill(ws.Range(f"A{row}:{end_cl}{row}"), (235, 240, 248))
        ws.Rows(row).RowHeight = 24
        row += 1

        for item in items:
            ws.Cells(row, 1).Value = f"   {item.get('item', '')}"
            _apply_font(ws.Cells(row, 1), size=10, name="Segoe UI", color=(51, 51, 51))
            ws.Rows(row).RowHeight = 20

            values = item.get("values", [])
            for j, val in enumerate(values):
                if j < num_periods:
                    ws.Cells(row, 2 + j).Value = val
                    _set_alignment(ws.Cells(row, 2 + j), h_align=XL_HALIGN_RIGHT)
                    _apply_font(ws.Cells(row, 2 + j), size=10, name="Segoe UI")
                    ws.Cells(row, 2 + j).NumberFormat = "#,##0"
            row += 1

        # Subtotal
        ws.Cells(row, 1).Value = f"   Net {section_name}"
        _apply_font(ws.Cells(row, 1), bold=True, size=10, name="Segoe UI",
                    color=section_color)
        ws.Rows(row).RowHeight = 22

        for j in range(num_periods):
            col = 2 + j
            cl = _col_letter(col)
            first_data_row = start_row + 1
            last_data_row = row - 1
            ws.Cells(row, col).Formula = f"=SUM({cl}{first_data_row}:{cl}{last_data_row})"
            _apply_font(ws.Cells(row, col), bold=True, size=10, name="Segoe UI",
                        color=section_color)
            _set_alignment(ws.Cells(row, col), h_align=XL_HALIGN_RIGHT)
            ws.Cells(row, col).NumberFormat = "#,##0"

        # Bottom border for subtotal
        _apply_borders(ws.Range(f"A{row}:{end_cl}{row}"),
                      edges=[XL_BORDER_TOP, XL_BORDER_BOTTOM],
                      weight=XL_BORDER_WEIGHT_THIN, color=(150, 160, 180))

        return row + 1

    r = 4
    r = _write_section(r, "Operating Activities", operating, primary)
    r += 1
    r = _write_section(r, "Investing Activities", investing, (0, 128, 80))
    r += 1
    r = _write_section(r, "Financing Activities", financing, (140, 80, 0))

    # Grand total - net change in cash
    r += 1
    ws.Cells(r, 1).Value = "Net Change in Cash"
    _apply_font(ws.Cells(r, 1), bold=True, size=12, name="Segoe UI", color=primary)
    _apply_fill(ws.Range(f"A{r}:{end_cl}{r}"), (220, 228, 240))
    ws.Rows(r).RowHeight = 28

    # Calculate net totals (find the subtotal rows)
    # For simplicity, sum the three section subtotals
    op_total_r = 4 + len(operating) + 1
    inv_total_r = op_total_r + 2 + len(investing) + 1
    fin_total_r = inv_total_r + 2 + len(financing) + 1

    for j in range(num_periods):
        col = 2 + j
        cl = _col_letter(col)
        ws.Cells(r, col).Formula = f"={cl}{op_total_r}+{cl}{inv_total_r}+{cl}{fin_total_r}"
        _apply_font(ws.Cells(r, col), bold=True, size=12, name="Segoe UI", color=primary)
        _set_alignment(ws.Cells(r, col), h_align=XL_HALIGN_RIGHT)
        ws.Cells(r, col).NumberFormat = "#,##0"

    _apply_borders(ws.Range(f"A{r}:{end_cl}{r}"),
                  edges=[XL_BORDER_TOP, XL_BORDER_BOTTOM],
                  weight=XL_BORDER_WEIGHT_MEDIUM, color=primary)

    # Column widths
    ws.Columns(1).ColumnWidth = 32
    for j in range(num_periods):
        ws.Columns(2 + j).ColumnWidth = 16

    return {"sheet": ws.Name, "title": title, "period_count": num_periods}


# ---------------------------------------------------------------------------
# 22. Workout Tracker
# ---------------------------------------------------------------------------

def create_workout_tracker(
    sheet, title: str, exercises: list[dict],
    style: str = "weekly",
) -> dict:
    """Create a fitness/workout tracker.

    Args:
        sheet: Target worksheet name or None.
        title: Tracker title.
        exercises: List of {"name": "...", "sets": 3, "reps": 10, "weight": "..."}.
        style: "weekly" (7 day tracker), "daily" (single day detail),
               "simple" (minimal).

    Returns:
        dict with sheet, title, exercise_count, style.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    primary = (41, 65, 122)
    accent = (0, 150, 136)

    days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    # Title header
    ws.Range("A1:L1").MergeCells = True
    ws.Cells(1, 1).Value = title
    _apply_font(ws.Cells(1, 1), bold=True, size=16, name="Segoe UI", color=primary)
    _apply_fill(ws.Range("A1:L1"), (245, 248, 255))
    ws.Rows(1).RowHeight = 32

    ws.Rows(2).RowHeight = 3
    _apply_fill(ws.Range("A2:L2"), accent)

    if style == "weekly":
        # Headers: Exercise | Sets | Reps | Weight | Mon-Sun (checkboxes)
        header_row = 3
        headers = ["Exercise", "Sets", "Reps", "Weight"] + days
        for j, h in enumerate(headers):
            ws.Cells(header_row, j + 1).Value = h

        h_range = ws.Range(f"A{header_row}:{_col_letter(len(headers))}{header_row}")
        _apply_fill(h_range, primary)
        _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
        _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Rows(header_row).RowHeight = 26

        for i, ex in enumerate(exercises):
            r = header_row + 1 + i
            ws.Cells(r, 1).Value = ex.get("name", "")
            ws.Cells(r, 2).Value = ex.get("sets", 3)
            ws.Cells(r, 3).Value = ex.get("reps", 10)
            ws.Cells(r, 4).Value = ex.get("weight", "")

            _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI")
            for c in range(2, 5):
                _set_alignment(ws.Cells(r, c), h_align=XL_HALIGN_CENTER)
                _apply_font(ws.Cells(r, c), size=10, name="Segoe UI")

            # Day columns for check marks
            for d in range(7):
                _set_alignment(ws.Cells(r, 5 + d), h_align=XL_HALIGN_CENTER)

            if i % 2 == 1:
                _apply_fill(ws.Range(f"A{r}:{_col_letter(len(headers))}{r}"),
                           (245, 248, 252))

            ws.Rows(r).RowHeight = 22

        last_row = header_row + len(exercises)

        # Summary row
        summary_r = last_row + 2
        ws.Cells(summary_r, 1).Value = "Daily Completion:"
        _apply_font(ws.Cells(summary_r, 1), bold=True, size=9, name="Segoe UI")

        for d in range(7):
            col = 5 + d
            cl = _col_letter(col)
            ws.Cells(summary_r, col).Formula = (
                f'=COUNTIF({cl}{header_row + 1}:{cl}{last_row},"*")'
            )
            _apply_font(ws.Cells(summary_r, col), bold=True, size=10,
                        name="Segoe UI", color=accent)
            _set_alignment(ws.Cells(summary_r, col), h_align=XL_HALIGN_CENTER)

        # Weekend highlight
        for d in [5, 6]:
            for r in range(header_row, last_row + 1):
                _apply_fill(ws.Cells(r, 5 + d), (240, 248, 245))

        # Borders
        _apply_borders(ws.Range(f"A{header_row}:{_col_letter(len(headers))}{last_row}"),
                      weight=XL_BORDER_WEIGHT_THIN, color=(200, 210, 220))

        # Column widths
        ws.Columns(1).ColumnWidth = 22
        ws.Columns(2).ColumnWidth = 7
        ws.Columns(3).ColumnWidth = 7
        ws.Columns(4).ColumnWidth = 10
        for d in range(7):
            ws.Columns(5 + d).ColumnWidth = 6

    elif style == "daily":
        header_row = 3
        headers = ["Exercise", "Set 1", "Set 2", "Set 3", "Set 4", "Set 5",
                    "Rest (s)", "Notes"]
        for j, h in enumerate(headers):
            ws.Cells(header_row, j + 1).Value = h

        h_range = ws.Range(f"A{header_row}:{_col_letter(len(headers))}{header_row}")
        _apply_fill(h_range, primary)
        _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
        _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Rows(header_row).RowHeight = 26

        for i, ex in enumerate(exercises):
            r = header_row + 1 + i
            ws.Cells(r, 1).Value = ex.get("name", "")
            _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI")

            n_sets = ex.get("sets", 3)
            reps = ex.get("reps", 10)
            weight = ex.get("weight", "")
            for s_idx in range(1, 6):
                if s_idx <= n_sets:
                    ws.Cells(r, 1 + s_idx).Value = f"{reps}x{weight}"
                _set_alignment(ws.Cells(r, 1 + s_idx), h_align=XL_HALIGN_CENTER)
                _apply_font(ws.Cells(r, 1 + s_idx), size=9, name="Segoe UI")

            if i % 2 == 1:
                _apply_fill(ws.Range(f"A{r}:{_col_letter(len(headers))}{r}"),
                           (245, 248, 252))
            ws.Rows(r).RowHeight = 22

        last_row = header_row + len(exercises)
        _apply_borders(ws.Range(f"A{header_row}:{_col_letter(len(headers))}{last_row}"),
                      weight=XL_BORDER_WEIGHT_THIN, color=(200, 210, 220))

        ws.Columns(1).ColumnWidth = 22
        for c in range(2, 7):
            ws.Columns(c).ColumnWidth = 10
        ws.Columns(7).ColumnWidth = 9
        ws.Columns(8).ColumnWidth = 18

    else:  # simple
        header_row = 3
        headers = ["Exercise", "Sets x Reps", "Weight", "Done"]
        for j, h in enumerate(headers):
            ws.Cells(header_row, j + 1).Value = h

        h_range = ws.Range(f"A{header_row}:D{header_row}")
        _apply_fill(h_range, primary)
        _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
        _set_alignment(h_range, h_align=XL_HALIGN_CENTER)
        ws.Rows(header_row).RowHeight = 26

        for i, ex in enumerate(exercises):
            r = header_row + 1 + i
            ws.Cells(r, 1).Value = ex.get("name", "")
            ws.Cells(r, 2).Value = f"{ex.get('sets', 3)} x {ex.get('reps', 10)}"
            ws.Cells(r, 3).Value = ex.get("weight", "")
            _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI")
            for c in range(2, 5):
                _set_alignment(ws.Cells(r, c), h_align=XL_HALIGN_CENTER)
                _apply_font(ws.Cells(r, c), size=10, name="Segoe UI")
            if i % 2 == 1:
                _apply_fill(ws.Range(f"A{r}:D{r}"), (245, 248, 252))
            ws.Rows(r).RowHeight = 22

        last_row = header_row + len(exercises)
        _apply_borders(ws.Range(f"A{header_row}:D{last_row}"),
                      weight=XL_BORDER_WEIGHT_THIN, color=(200, 210, 220))

        ws.Columns(1).ColumnWidth = 22
        ws.Columns(2).ColumnWidth = 14
        ws.Columns(3).ColumnWidth = 10
        ws.Columns(4).ColumnWidth = 8

    return {"sheet": ws.Name, "title": title,
            "exercise_count": len(exercises), "style": style}


# ---------------------------------------------------------------------------
# 23. Meal Planner
# ---------------------------------------------------------------------------

def create_meal_planner(
    sheet, title: str, days: int = 7, meals_per_day: int = 3,
    style: str = "weekly",
) -> dict:
    """Create a meal planning template.

    Args:
        sheet: Target worksheet name or None.
        title: Planner title.
        days: Number of days (default 7).
        meals_per_day: Number of meals per day (default 3).
        style: "weekly" (7-day grid), "detailed" (with calories/notes),
               "simple" (minimal).

    Returns:
        dict with sheet, title, days, meals_per_day, style.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)

    primary = (76, 175, 80)
    accent = (56, 142, 60)
    header_bg = (27, 94, 32)

    day_names = ["Monday", "Tuesday", "Wednesday", "Thursday",
                 "Friday", "Saturday", "Sunday"]
    meal_names = ["Breakfast", "Lunch", "Dinner", "Snack 1", "Snack 2"][:meals_per_day]

    # Title
    total_cols = 2 + meals_per_day if style == "detailed" else 1 + meals_per_day
    end_cl = _col_letter(total_cols)

    ws.Range(f"A1:{end_cl}1").MergeCells = True
    ws.Cells(1, 1).Value = title
    _apply_font(ws.Cells(1, 1), bold=True, size=16, name="Segoe UI", color=header_bg)
    _apply_fill(ws.Range(f"A1:{end_cl}1"), (232, 245, 233))
    ws.Rows(1).RowHeight = 32

    ws.Rows(2).RowHeight = 3
    _apply_fill(ws.Range(f"A2:{end_cl}2"), primary)

    if style == "weekly" or style == "simple":
        header_row = 3
        # Headers
        ws.Cells(header_row, 1).Value = "Day"
        for m, meal in enumerate(meal_names):
            ws.Cells(header_row, 2 + m).Value = meal

        h_range = ws.Range(f"A{header_row}:{_col_letter(1 + meals_per_day)}{header_row}")
        _apply_fill(h_range, header_bg)
        _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
        _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Rows(header_row).RowHeight = 26

        for d in range(days):
            r = header_row + 1 + d
            dn = day_names[d % 7] if d < 7 else f"Day {d + 1}"
            ws.Cells(r, 1).Value = dn
            _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI",
                        color=header_bg)

            for m in range(meals_per_day):
                _set_alignment(ws.Cells(r, 2 + m), h_align=XL_HALIGN_CENTER, wrap=True)
                _apply_font(ws.Cells(r, 2 + m), size=10, name="Segoe UI")

            if d % 2 == 1:
                _apply_fill(ws.Range(f"A{r}:{_col_letter(1 + meals_per_day)}{r}"),
                           (232, 245, 233))

            # Weekend highlight
            if d % 7 >= 5:
                _apply_fill(ws.Cells(r, 1), (200, 230, 201))

            ws.Rows(r).RowHeight = 36 if style == "weekly" else 24

        last_row = header_row + days
        _apply_borders(ws.Range(f"A{header_row}:{_col_letter(1 + meals_per_day)}{last_row}"),
                      weight=XL_BORDER_WEIGHT_THIN, color=(180, 210, 185))

        ws.Columns(1).ColumnWidth = 14
        for m in range(meals_per_day):
            ws.Columns(2 + m).ColumnWidth = 20

    else:  # detailed
        header_row = 3
        ws.Cells(header_row, 1).Value = "Day"
        for m, meal in enumerate(meal_names):
            ws.Cells(header_row, 2 + m).Value = meal
        ws.Cells(header_row, 2 + meals_per_day).Value = "Calories"

        h_range = ws.Range(f"A{header_row}:{_col_letter(2 + meals_per_day)}{header_row}")
        _apply_fill(h_range, header_bg)
        _apply_font(h_range, bold=True, size=10, name="Segoe UI", color=(255, 255, 255))
        _set_alignment(h_range, h_align=XL_HALIGN_CENTER, v_align=XL_VALIGN_CENTER)
        ws.Rows(header_row).RowHeight = 26

        for d in range(days):
            r = header_row + 1 + d
            dn = day_names[d % 7] if d < 7 else f"Day {d + 1}"
            ws.Cells(r, 1).Value = dn
            _apply_font(ws.Cells(r, 1), bold=True, size=10, name="Segoe UI",
                        color=header_bg)

            for m in range(meals_per_day):
                _set_alignment(ws.Cells(r, 2 + m), h_align=XL_HALIGN_CENTER, wrap=True)
                _apply_font(ws.Cells(r, 2 + m), size=10, name="Segoe UI")

            # Calories column
            cal_col = 2 + meals_per_day
            _set_alignment(ws.Cells(r, cal_col), h_align=XL_HALIGN_CENTER)
            _apply_font(ws.Cells(r, cal_col), size=10, name="Segoe UI", color=accent)

            if d % 2 == 1:
                _apply_fill(ws.Range(f"A{r}:{_col_letter(2 + meals_per_day)}{r}"),
                           (232, 245, 233))

            if d % 7 >= 5:
                _apply_fill(ws.Cells(r, 1), (200, 230, 201))

            ws.Rows(r).RowHeight = 36

        last_row = header_row + days

        # Total calories row
        total_r = last_row + 1
        ws.Cells(total_r, 1).Value = "Weekly Total"
        _apply_font(ws.Cells(total_r, 1), bold=True, size=10, name="Segoe UI",
                    color=header_bg)
        cal_cl = _col_letter(2 + meals_per_day)
        ws.Cells(total_r, 2 + meals_per_day).Formula = (
            f"=SUM({cal_cl}{header_row + 1}:{cal_cl}{last_row})"
        )
        _apply_font(ws.Cells(total_r, 2 + meals_per_day), bold=True, size=12,
                    name="Segoe UI", color=primary)
        _set_alignment(ws.Cells(total_r, 2 + meals_per_day), h_align=XL_HALIGN_CENTER)
        _apply_fill(ws.Range(f"A{total_r}:{_col_letter(2 + meals_per_day)}{total_r}"),
                   (200, 230, 201))

        _apply_borders(ws.Range(f"A{header_row}:{_col_letter(2 + meals_per_day)}{total_r}"),
                      weight=XL_BORDER_WEIGHT_THIN, color=(180, 210, 185))

        ws.Columns(1).ColumnWidth = 14
        for m in range(meals_per_day):
            ws.Columns(2 + m).ColumnWidth = 20
        ws.Columns(2 + meals_per_day).ColumnWidth = 12

    return {"sheet": ws.Name, "title": title, "days": days,
            "meals_per_day": meals_per_day, "style": style}
