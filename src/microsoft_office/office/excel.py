"""Excel COM automation logic for the Microsoft Office MCP server.

Provides comprehensive Excel workbook manipulation including formatting,
charts, pivot tables, data validation, conditional formatting, and more.
"""

from microsoft_office.com_utils import ensure_absolute_path, get_or_create_app, rgb

# ---------------------------------------------------------------------------
# Excel Constants
# ---------------------------------------------------------------------------

# Chart types
XL_CHART_AREA = 1
XL_CHART_BAR_CLUSTERED = 57
XL_CHART_COLUMN_CLUSTERED = 51
XL_CHART_COLUMN_STACKED = 52
XL_CHART_DOUGHNUT = -4120
XL_CHART_LINE = 4
XL_CHART_LINE_MARKERS = 65
XL_CHART_PIE = 5
XL_CHART_SCATTER = -4169

# Border edge indices
XL_BORDER_LEFT = 7
XL_BORDER_TOP = 8
XL_BORDER_BOTTOM = 9
XL_BORDER_RIGHT = 10
XL_BORDER_INSIDE_VERTICAL = 11
XL_BORDER_INSIDE_HORIZONTAL = 12
XL_BORDER_ALL_EDGES = [7, 8, 9, 10, 11, 12]

# Border line styles
XL_LINE_STYLE_CONTINUOUS = 1
XL_LINE_STYLE_DASH = -4115
XL_LINE_STYLE_DOT = -4118
XL_LINE_STYLE_NONE = -4142

# Border weights
XL_BORDER_WEIGHT_HAIRLINE = 1
XL_BORDER_WEIGHT_THIN = 2
XL_BORDER_WEIGHT_MEDIUM = -4138
XL_BORDER_WEIGHT_THICK = 4

# Horizontal alignment
XL_HALIGN_GENERAL = 1
XL_HALIGN_LEFT = -4131
XL_HALIGN_CENTER = -4108
XL_HALIGN_RIGHT = -4152

# Vertical alignment
XL_VALIGN_TOP = -4160
XL_VALIGN_CENTER = -4108
XL_VALIGN_BOTTOM = -4107

# Sort order
XL_SORT_ASCENDING = 1
XL_SORT_DESCENDING = 2

# Orientation
XL_ORIENT_PORTRAIT = 1
XL_ORIENT_LANDSCAPE = 2

# Fill pattern
XL_PATTERN_SOLID = 1
XL_PATTERN_NONE = -4142

# Pivot field orientations
XL_PIVOT_ROW = 1
XL_PIVOT_COLUMN = 2
XL_PIVOT_PAGE = 3
XL_PIVOT_DATA = 4

# Aggregate functions for pivot tables
XL_FUNC_SUM = -4157
XL_FUNC_COUNT = -4112
XL_FUNC_AVERAGE = -4106
XL_FUNC_MAX = -4136
XL_FUNC_MIN = -4139

# Conditional formatting operators
XL_CF_BETWEEN = 1
XL_CF_EQUAL = 3
XL_CF_GREATER = 5
XL_CF_LESS = 6
XL_CF_GREATER_EQUAL = 7
XL_CF_LESS_EQUAL = 8

# Data validation types
XL_DV_LIST = 3
XL_DV_WHOLE_NUMBER = 1
XL_DV_DECIMAL = 2
XL_DV_DATE = 4
XL_DV_TEXT_LENGTH = 6

# Export formats
XL_EXPORT_PDF = 0

# Pivot cache source types
XL_PIVOT_SOURCE_WORKSHEET = 1

# AutoFilter operator constants
XL_FILTER_AND = 1
XL_FILTER_OR = 2


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _get_app():
    """Get or create the Excel application instance."""
    return get_or_create_app("Excel.Application")


def _get_ws(app, sheet: str | None = None):
    """Return the target worksheet from the active workbook."""
    wb = app.ActiveWorkbook
    return wb.Worksheets(sheet) if sheet else wb.ActiveSheet


# ---------------------------------------------------------------------------
# Basic workbook operations
# ---------------------------------------------------------------------------

def create_workbook(file_path: str | None = None) -> dict:
    """Create a new Excel workbook, optionally saving it immediately.

    Args:
        file_path: Optional path to save the new workbook.

    Returns:
        dict with workbook name and sheet count.
    """
    app = _get_app()
    wb = app.Workbooks.Add()
    if file_path:
        wb.SaveAs(ensure_absolute_path(file_path))
    return {"name": wb.Name, "sheet_count": wb.Worksheets.Count}


def open_workbook(file_path: str) -> dict:
    """Open an existing workbook.

    Args:
        file_path: Path to the workbook file.

    Returns:
        dict with workbook name and list of sheet names.
    """
    app = _get_app()
    wb = app.Workbooks.Open(ensure_absolute_path(file_path))
    sheets = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    return {"name": wb.Name, "sheets": sheets}


def save_workbook(file_path: str | None = None) -> dict:
    """Save the active workbook, optionally to a new path.

    Args:
        file_path: Optional new path (SaveAs). If omitted, saves in place.

    Returns:
        dict with workbook name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    if file_path:
        wb.SaveAs(ensure_absolute_path(file_path))
    else:
        wb.Save()
    return {"name": wb.Name}


def close_workbook() -> dict:
    """Close the active workbook without saving.

    Returns:
        dict with the closed workbook name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    name = wb.Name
    wb.Close(SaveChanges=False)
    return {"name": name}


# ---------------------------------------------------------------------------
# Sheet operations
# ---------------------------------------------------------------------------

def get_sheets() -> dict:
    """List all sheet names in the active workbook.

    Returns:
        dict with workbook name and list of sheet names.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    sheets = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    return {"name": wb.Name, "sheets": sheets}


def add_sheet(name: str) -> dict:
    """Add a new worksheet with the given name.

    Args:
        name: Name for the new sheet.

    Returns:
        dict with the new sheet name and total sheet count.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    ws = wb.Worksheets.Add()
    ws.Name = name
    return {"sheet_name": name, "sheet_count": wb.Worksheets.Count}


def delete_sheet(sheet: str) -> dict:
    """Delete a worksheet by name.

    Args:
        sheet: Name of the sheet to delete.

    Returns:
        dict with deleted sheet name and remaining sheet count.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    app.DisplayAlerts = False
    try:
        wb.Worksheets(sheet).Delete()
    finally:
        app.DisplayAlerts = True
    return {"deleted": sheet, "sheet_count": wb.Worksheets.Count}


def rename_sheet(old_name: str, new_name: str) -> dict:
    """Rename a worksheet.

    Args:
        old_name: Current name of the sheet.
        new_name: New name for the sheet.

    Returns:
        dict with old and new names.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    wb.Worksheets(old_name).Name = new_name
    return {"old_name": old_name, "new_name": new_name}


# ---------------------------------------------------------------------------
# Cell and range read/write
# ---------------------------------------------------------------------------

def read_cell(cell: str, sheet: str | None = None) -> dict:
    """Read the value of a single cell.

    Args:
        cell: Cell address (e.g. "A1").
        sheet: Optional sheet name.

    Returns:
        dict with cell address, value, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    value = ws.Range(cell).Value
    return {"cell": cell, "value": value, "sheet": ws.Name}


def write_cell(cell: str, value, sheet: str | None = None) -> dict:
    """Write a value to a single cell.

    Args:
        cell: Cell address (e.g. "A1").
        value: Value to write.
        sheet: Optional sheet name.

    Returns:
        dict with cell address, value, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(cell).Value = value
    return {"cell": cell, "value": value, "sheet": ws.Name}


def read_range(range_address: str, sheet: str | None = None) -> dict:
    """Read values from a cell range.

    Handles COM returning a single value, a flat tuple, or a tuple of tuples.

    Args:
        range_address: Range address (e.g. "A1:C3").
        sheet: Optional sheet name.

    Returns:
        dict with range address, 2D data list, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_address)
    data = rng.Value
    if data is None:
        result = []
    elif isinstance(data, tuple):
        result = [list(row) if isinstance(row, tuple) else [row] for row in data]
    else:
        result = [[data]]
    return {"range": range_address, "data": result, "sheet": ws.Name}


def write_range(start_cell: str, data: list[list], sheet: str | None = None) -> dict:
    """Write a 2D list of values starting at the given cell.

    Data rows are padded with None to equal length before writing.

    Args:
        start_cell: Top-left cell address (e.g. "A1").
        data: 2D list of values to write.
        sheet: Optional sheet name.

    Returns:
        dict with start cell, row/col counts, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rows = len(data)
    cols = max(len(row) for row in data) if data else 0
    if rows > 0 and cols > 0:
        start = ws.Range(start_cell)
        end = ws.Cells(start.Row + rows - 1, start.Column + cols - 1)
        target_range = ws.Range(start, end)
        padded = [row + [None] * (cols - len(row)) for row in data]
        target_range.Value = padded
    return {"start_cell": start_cell, "rows": rows, "cols": cols, "sheet": ws.Name}


def get_sheet_data(sheet: str | None = None) -> dict:
    """Read all data from the used range of a worksheet.

    Args:
        sheet: Optional sheet name.

    Returns:
        dict with sheet name, 2D data list, and row/col counts.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    used = ws.UsedRange
    if used.Rows.Count == 1 and used.Columns.Count == 1 and used.Value is None:
        return {"sheet": ws.Name, "data": [], "rows": 0, "cols": 0}
    data = used.Value
    if data is None:
        result = []
    elif isinstance(data, tuple):
        result = [list(row) if isinstance(row, tuple) else [row] for row in data]
    else:
        result = [[data]]
    return {
        "sheet": ws.Name,
        "data": result,
        "rows": len(result),
        "cols": len(result[0]) if result else 0,
    }


# ---------------------------------------------------------------------------
# Formatting
# ---------------------------------------------------------------------------

def set_cell_format(
    cell: str,
    bold: bool | None = None,
    italic: bool | None = None,
    font_size: float | None = None,
    font_name: str | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    underline: bool | None = None,
    number_format: str | None = None,
    sheet: str | None = None,
) -> dict:
    """Apply font and number formatting to a cell or range.

    Args:
        cell: Cell or range address.
        bold: Set bold.
        italic: Set italic.
        font_size: Font size in points.
        font_name: Font family name.
        font_color_rgb: Font color as (R, G, B) tuple.
        underline: Set underline.
        number_format: Excel number format string (e.g. "#,##0.00").
        sheet: Optional sheet name.

    Returns:
        dict with cell address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    if bold is not None:
        rng.Font.Bold = bold
    if italic is not None:
        rng.Font.Italic = italic
    if font_size is not None:
        rng.Font.Size = font_size
    if font_name is not None:
        rng.Font.Name = font_name
    if font_color_rgb is not None:
        rng.Font.Color = rgb(*font_color_rgb)
    if underline is not None:
        rng.Font.Underline = underline
    if number_format is not None:
        rng.NumberFormat = number_format
    return {"cell": cell, "sheet": ws.Name}


def set_cell_fill(
    range_address: str,
    color_rgb: tuple[int, int, int] | None = None,
    pattern: int | None = None,
    pattern_color_rgb: tuple[int, int, int] | None = None,
    sheet: str | None = None,
) -> dict:
    """Set the fill (background) color and pattern for a range.

    Args:
        range_address: Target range address.
        color_rgb: Fill color as (R, G, B) tuple.
        pattern: Fill pattern constant (1=Solid, -4142=None).
        pattern_color_rgb: Pattern foreground color as (R, G, B) tuple.
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_address)
    if color_rgb is not None:
        rng.Interior.Color = rgb(*color_rgb)
    if pattern is not None:
        rng.Interior.Pattern = pattern
    if pattern_color_rgb is not None:
        rng.Interior.PatternColor = rgb(*pattern_color_rgb)
    return {"range": range_address, "sheet": ws.Name}


def set_cell_borders(
    range_address: str,
    border_style: int = XL_LINE_STYLE_CONTINUOUS,
    border_weight: int = XL_BORDER_WEIGHT_THIN,
    color_rgb: tuple[int, int, int] | None = None,
    edges: list[int] | None = None,
    sheet: str | None = None,
) -> dict:
    """Apply borders to a range.

    Args:
        range_address: Target range address.
        border_style: Line style (1=Continuous, -4115=Dash, -4118=Dot, -4142=None).
        border_weight: Line weight (2=Thin, -4138=Medium, 4=Thick).
        color_rgb: Border color as (R, G, B) tuple.
        edges: List of border edge indices (default all: 7=Left, 8=Top,
               9=Bottom, 10=Right, 11=InsideVertical, 12=InsideHorizontal).
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_address)
    if edges is None:
        edges = XL_BORDER_ALL_EDGES
    for edge in edges:
        border = rng.Borders(edge)
        border.LineStyle = border_style
        border.Weight = border_weight
        if color_rgb is not None:
            border.Color = rgb(*color_rgb)
    return {"range": range_address, "sheet": ws.Name}


def set_cell_alignment(
    range_address: str,
    horizontal: int | None = None,
    vertical: int | None = None,
    wrap_text: bool | None = None,
    text_rotation: int | None = None,
    sheet: str | None = None,
) -> dict:
    """Set text alignment and wrapping for a range.

    Args:
        range_address: Target range address.
        horizontal: Horizontal alignment (-4131=Left, -4108=Center,
                    -4152=Right, 1=General).
        vertical: Vertical alignment (-4160=Top, -4108=Center, -4107=Bottom).
        wrap_text: Enable or disable text wrapping.
        text_rotation: Text rotation angle in degrees.
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_address)
    if horizontal is not None:
        rng.HorizontalAlignment = horizontal
    if vertical is not None:
        rng.VerticalAlignment = vertical
    if wrap_text is not None:
        rng.WrapText = wrap_text
    if text_rotation is not None:
        rng.Orientation = text_rotation
    return {"range": range_address, "sheet": ws.Name}


def merge_cells(range_address: str, sheet: str | None = None) -> dict:
    """Merge cells in the given range.

    Args:
        range_address: Range to merge (e.g. "A1:C1").
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_address).Merge()
    return {"range": range_address, "sheet": ws.Name}


def unmerge_cells(range_address: str, sheet: str | None = None) -> dict:
    """Unmerge previously merged cells.

    Args:
        range_address: Range to unmerge.
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_address).UnMerge()
    return {"range": range_address, "sheet": ws.Name}


def set_column_width(
    columns: str,
    width: float | None = None,
    auto_fit: bool = False,
    sheet: str | None = None,
) -> dict:
    """Set column width or auto-fit columns.

    Args:
        columns: Column specifier (e.g. "A", "A:C", "1:3").
        width: Explicit column width in character units.
        auto_fit: If True, auto-fit column width to content.
        sheet: Optional sheet name.

    Returns:
        dict with column specifier and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    col_range = ws.Columns(columns)
    if auto_fit:
        col_range.AutoFit()
    elif width is not None:
        col_range.ColumnWidth = width
    return {"columns": columns, "sheet": ws.Name}


def set_row_height(rows: str, height: float, sheet: str | None = None) -> dict:
    """Set the height of one or more rows.

    Args:
        rows: Row specifier (e.g. "1", "1:5").
        height: Row height in points.
        sheet: Optional sheet name.

    Returns:
        dict with row specifier and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Rows(rows).RowHeight = height
    return {"rows": rows, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Formulas & Data
# ---------------------------------------------------------------------------

def set_formula(cell: str, formula: str, sheet: str | None = None) -> dict:
    """Set a formula on a cell.

    Args:
        cell: Cell address (e.g. "B2").
        formula: Excel formula string (e.g. "=SUM(A1:A10)").
        sheet: Optional sheet name.

    Returns:
        dict with cell address, formula, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(cell).Formula = formula
    return {"cell": cell, "formula": formula, "sheet": ws.Name}


def set_number_format(
    range_address: str, format_string: str, sheet: str | None = None
) -> dict:
    """Apply a number format to a range.

    Args:
        range_address: Target range address.
        format_string: Excel number format (e.g. "#,##0.00", "yyyy-mm-dd").
        sheet: Optional sheet name.

    Returns:
        dict with range address, format string, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_address).NumberFormat = format_string
    return {"range": range_address, "format": format_string, "sheet": ws.Name}


def add_named_range(
    name: str, range_address: str, sheet: str | None = None
) -> dict:
    """Create a named range in the active workbook.

    Args:
        name: Name for the range.
        range_address: Cell range address (e.g. "A1:D10").
        sheet: Optional sheet name for scoping the reference.

    Returns:
        dict with the named range name and its reference.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    ws = wb.Worksheets(sheet) if sheet else wb.ActiveSheet
    refers_to = f"={ws.Name}!{range_address}"
    wb.Names.Add(Name=name, RefersTo=refers_to)
    return {"name": name, "refers_to": refers_to}


def delete_named_range(name: str) -> dict:
    """Delete a named range from the active workbook.

    Args:
        name: Name of the named range to delete.

    Returns:
        dict with the deleted name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    wb.Names(name).Delete()
    return {"deleted": name}


def get_named_ranges() -> dict:
    """List all named ranges in the active workbook.

    Returns:
        dict with a list of named range details.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    ranges = []
    for i in range(1, wb.Names.Count + 1):
        n = wb.Names(i)
        ranges.append({"name": n.Name, "refers_to": n.RefersTo, "visible": n.Visible})
    return {"named_ranges": ranges, "count": len(ranges)}


def sort_range(
    range_address: str,
    sort_field: str,
    order: int = XL_SORT_ASCENDING,
    has_header: bool = True,
    sheet: str | None = None,
) -> dict:
    """Sort a range by a specified key column.

    Args:
        range_address: Range to sort (e.g. "A1:D20").
        sort_field: Key column range (e.g. "B1:B20").
        order: Sort order (1=Ascending, 2=Descending).
        has_header: Whether the first row is a header.
        sheet: Optional sheet name.

    Returns:
        dict with range address, sort field, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_address)
    key = ws.Range(sort_field)
    ws.Sort.SortFields.Clear()
    ws.Sort.SortFields.Add2(Key=key, Order=order)
    ws.Sort.SetRange(rng)
    ws.Sort.Header = 1 if has_header else 2
    ws.Sort.Apply()
    return {"range": range_address, "sort_field": sort_field, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Data Features
# ---------------------------------------------------------------------------

def set_auto_filter(
    range_address: str | None = None,
    field: int | None = None,
    criteria1=None,
    criteria2=None,
    operator: int | None = None,
    sheet: str | None = None,
) -> dict:
    """Apply or configure AutoFilter on a range.

    If called with only range_address, toggles AutoFilter on the range.
    If field and criteria are provided, filters a specific column.

    Args:
        range_address: Range to apply the filter to. If None, uses UsedRange.
        field: Column number within the range to filter (1-based).
        criteria1: First filter criteria value.
        criteria2: Second filter criteria value (for AND/OR operations).
        operator: Filter operator constant (1=And, 2=Or).
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if range_address:
        rng = ws.Range(range_address)
    else:
        rng = ws.UsedRange

    kwargs = {}
    if field is not None:
        kwargs["Field"] = field
    if criteria1 is not None:
        kwargs["Criteria1"] = criteria1
    if criteria2 is not None:
        kwargs["Criteria2"] = criteria2
    if operator is not None:
        kwargs["Operator"] = operator

    rng.AutoFilter(**kwargs)
    return {"range": rng.Address, "sheet": ws.Name}


def freeze_panes(cell: str, sheet: str | None = None) -> dict:
    """Freeze panes at the specified cell.

    Rows above and columns to the left of the cell will be frozen.

    Args:
        cell: Cell address at which to freeze (e.g. "B2" freezes row 1 and column A).
        sheet: Optional sheet name.

    Returns:
        dict with the freeze cell and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Activate()
    app.ActiveWindow.FreezePanes = False
    ws.Range(cell).Select()
    app.ActiveWindow.FreezePanes = True
    return {"cell": cell, "sheet": ws.Name}


def set_conditional_formatting(
    range_address: str,
    operator: int,
    formula1: str,
    formula2: str | None = None,
    font_color_rgb: tuple[int, int, int] | None = None,
    fill_color_rgb: tuple[int, int, int] | None = None,
    sheet: str | None = None,
) -> dict:
    """Add a conditional formatting rule to a range.

    Args:
        range_address: Target range address.
        operator: Comparison operator (1=Between, 3=Equal, 5=Greater,
                  6=Less, 7=GreaterEqual, 8=LessEqual).
        formula1: First comparison value or formula.
        formula2: Second value (required for Between operator).
        font_color_rgb: Font color for matching cells as (R, G, B).
        fill_color_rgb: Fill color for matching cells as (R, G, B).
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_address)
    kwargs = {"Type": 1, "Operator": operator, "Formula1": formula1}
    if formula2 is not None:
        kwargs["Formula2"] = formula2
    fc = rng.FormatConditions.Add(**kwargs)
    if font_color_rgb is not None:
        fc.Font.Color = rgb(*font_color_rgb)
    if fill_color_rgb is not None:
        fc.Interior.Color = rgb(*fill_color_rgb)
    return {"range": range_address, "sheet": ws.Name}


def add_data_validation(
    range_address: str,
    validation_type: int,
    operator: int | None = None,
    formula1: str | None = None,
    formula2: str | None = None,
    input_title: str | None = None,
    input_message: str | None = None,
    error_title: str | None = None,
    error_message: str | None = None,
    sheet: str | None = None,
) -> dict:
    """Add data validation to a range.

    Args:
        range_address: Target range address.
        validation_type: Validation type (3=List, 1=WholeNumber, 2=Decimal,
                         4=Date, 6=TextLength).
        operator: Comparison operator for numeric/date/text validations.
        formula1: First validation value or list source (e.g. "A,B,C" for lists).
        formula2: Second validation value (for Between operator).
        input_title: Title for the input prompt.
        input_message: Body text for the input prompt.
        error_title: Title for the error alert.
        error_message: Body text for the error alert.
        sheet: Optional sheet name.

    Returns:
        dict with range address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_address)
    # Remove any existing validation first
    rng.Validation.Delete()
    kwargs = {"Type": validation_type}
    if operator is not None:
        kwargs["Operator"] = operator
    if formula1 is not None:
        kwargs["Formula1"] = formula1
    if formula2 is not None:
        kwargs["Formula2"] = formula2
    rng.Validation.Add(**kwargs)
    if input_title is not None:
        rng.Validation.InputTitle = input_title
    if input_message is not None:
        rng.Validation.InputMessage = input_message
    if error_title is not None:
        rng.Validation.ErrorTitle = error_title
    if error_message is not None:
        rng.Validation.ErrorMessage = error_message
    return {"range": range_address, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Charts
# ---------------------------------------------------------------------------

def add_chart(
    chart_type: int,
    data_range: str,
    sheet: str | None = None,
    title: str | None = None,
    left: float | None = None,
    top: float | None = None,
    width: float = 400,
    height: float = 300,
    has_legend: bool = True,
) -> dict:
    """Add an embedded chart to a worksheet.

    Args:
        chart_type: Chart type constant (51=ColumnClustered, 57=BarClustered,
                    4=Line, 5=Pie, -4169=Scatter, 1=Area, 52=ColumnStacked,
                    65=LineMarkers, -4120=Doughnut).
        data_range: Source data range address (e.g. "A1:B10").
        sheet: Optional sheet name.
        title: Optional chart title.
        left: Left position in points (defaults to 100).
        top: Top position in points (defaults to 100).
        width: Chart width in points.
        height: Chart height in points.
        has_legend: Whether to show the legend.

    Returns:
        dict with chart name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    source = ws.Range(data_range)
    chart_left = left if left is not None else 100
    chart_top = top if top is not None else 100
    chart_obj = ws.ChartObjects().Add(chart_left, chart_top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=source)
    chart.ChartType = chart_type
    if title is not None:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    chart.HasLegend = has_legend
    return {"chart_name": chart_obj.Name, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Pivot Tables
# ---------------------------------------------------------------------------

def create_pivot_table(
    source_range: str,
    dest_cell: str,
    table_name: str,
    row_fields: list[str],
    column_fields: list[str] | None = None,
    data_fields: list[dict] | None = None,
    page_fields: list[str] | None = None,
    source_sheet: str | None = None,
    dest_sheet: str | None = None,
) -> dict:
    """Create a pivot table from worksheet data.

    Args:
        source_range: Source data range address (e.g. "A1:E100").
        dest_cell: Destination cell for the pivot table.
        table_name: Name for the pivot table.
        row_fields: List of field names for rows.
        column_fields: Optional list of field names for columns.
        data_fields: Optional list of dicts with "name" (field name) and
                     "function" (aggregate: -4157=Sum, -4112=Count,
                     -4106=Average, -4136=Max, -4139=Min).
        page_fields: Optional list of field names for page/filter area.
        source_sheet: Optional source sheet name.
        dest_sheet: Optional destination sheet name.

    Returns:
        dict with pivot table name and destination sheet name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    src_ws = wb.Worksheets(source_sheet) if source_sheet else wb.ActiveSheet
    dst_ws = wb.Worksheets(dest_sheet) if dest_sheet else wb.ActiveSheet

    src_range = src_ws.Range(source_range)
    pivot_cache = wb.PivotCaches().Create(
        SourceType=XL_PIVOT_SOURCE_WORKSHEET, SourceData=src_range
    )
    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=dst_ws.Range(dest_cell), TableName=table_name
    )

    # Configure row fields
    for field_name in row_fields:
        pf = pivot_table.PivotFields(field_name)
        pf.Orientation = XL_PIVOT_ROW

    # Configure column fields
    if column_fields:
        for field_name in column_fields:
            pf = pivot_table.PivotFields(field_name)
            pf.Orientation = XL_PIVOT_COLUMN

    # Configure page (filter) fields
    if page_fields:
        for field_name in page_fields:
            pf = pivot_table.PivotFields(field_name)
            pf.Orientation = XL_PIVOT_PAGE

    # Configure data fields
    if data_fields:
        for df in data_fields:
            pf = pivot_table.PivotFields(df["name"])
            pf.Orientation = XL_PIVOT_DATA
            if "function" in df:
                pf.Function = df["function"]

    return {"table_name": table_name, "sheet": dst_ws.Name}


# ---------------------------------------------------------------------------
# Comments
# ---------------------------------------------------------------------------

def add_comment(cell: str, text: str, sheet: str | None = None) -> dict:
    """Add a comment to a cell.

    Args:
        cell: Cell address (e.g. "A1").
        text: Comment text.
        sheet: Optional sheet name.

    Returns:
        dict with cell address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    # Remove existing comment if present
    if rng.Comment is not None:
        rng.Comment.Delete()
    rng.AddComment(text)
    return {"cell": cell, "text": text, "sheet": ws.Name}


def get_comments(sheet: str | None = None) -> dict:
    """Retrieve all comments from a worksheet.

    Args:
        sheet: Optional sheet name.

    Returns:
        dict with a list of comments (cell address, author, text).
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    comments = []
    for comment in ws.Comments:
        comments.append({
            "cell": comment.Parent.Address,
            "author": comment.Author,
            "text": comment.Text(),
        })
    return {"comments": comments, "count": len(comments), "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Print & Export
# ---------------------------------------------------------------------------

def set_print_setup(
    sheet: str | None = None,
    orientation: int | None = None,
    paper_size: int | None = None,
    fit_to_pages_wide: int | None = None,
    fit_to_pages_tall: int | None = None,
    print_title_rows: str | None = None,
    print_area: str | None = None,
    center_horizontally: bool | None = None,
    center_vertically: bool | None = None,
) -> dict:
    """Configure print/page setup for a worksheet.

    Args:
        sheet: Optional sheet name.
        orientation: Page orientation (1=Portrait, 2=Landscape).
        paper_size: Paper size constant.
        fit_to_pages_wide: Number of pages wide for fit-to-page scaling.
        fit_to_pages_tall: Number of pages tall for fit-to-page scaling.
        print_title_rows: Rows to repeat at top (e.g. "$1:$2").
        print_area: Print area range address.
        center_horizontally: Center content horizontally on the page.
        center_vertically: Center content vertically on the page.

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ps = ws.PageSetup
    if orientation is not None:
        ps.Orientation = orientation
    if paper_size is not None:
        ps.PaperSize = paper_size
    if fit_to_pages_wide is not None:
        ps.FitToPagesWide = fit_to_pages_wide
    if fit_to_pages_tall is not None:
        ps.FitToPagesTall = fit_to_pages_tall
    if print_title_rows is not None:
        ps.PrintTitleRows = print_title_rows
    if print_area is not None:
        ps.PrintArea = print_area
    if center_horizontally is not None:
        ps.CenterHorizontally = center_horizontally
    if center_vertically is not None:
        ps.CenterVertically = center_vertically
    return {"sheet": ws.Name}


def export_to_pdf(output_path: str, sheet: str | None = None) -> dict:
    """Export a worksheet or the active workbook to PDF.

    Args:
        output_path: Destination file path for the PDF.
        sheet: Optional sheet name. If provided, exports only that sheet;
               otherwise exports the active sheet.

    Returns:
        dict with output path and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    abs_path = ensure_absolute_path(output_path)
    ws.ExportAsFixedFormat(Type=XL_EXPORT_PDF, Filename=abs_path)
    return {"output_path": abs_path, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Protection
# ---------------------------------------------------------------------------

def protect_sheet(
    password: str | None = None,
    allow_formatting_cells: bool = False,
    allow_sorting: bool = False,
    allow_filtering: bool = False,
    sheet: str | None = None,
) -> dict:
    """Protect a worksheet with optional permissions.

    Args:
        password: Optional protection password.
        allow_formatting_cells: Allow users to format cells.
        allow_sorting: Allow users to sort data.
        allow_filtering: Allow users to use AutoFilter.
        sheet: Optional sheet name.

    Returns:
        dict with sheet name and protection status.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    kwargs = {
        "AllowFormattingCells": allow_formatting_cells,
        "AllowSorting": allow_sorting,
        "AllowFiltering": allow_filtering,
    }
    if password is not None:
        kwargs["Password"] = password
    ws.Protect(**kwargs)
    return {"sheet": ws.Name, "protected": True}


def unprotect_sheet(password: str | None = None, sheet: str | None = None) -> dict:
    """Remove protection from a worksheet.

    Args:
        password: Password used to protect the sheet (if any).
        sheet: Optional sheet name.

    Returns:
        dict with sheet name and protection status.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if password is not None:
        ws.Unprotect(Password=password)
    else:
        ws.Unprotect()
    return {"sheet": ws.Name, "protected": False}
