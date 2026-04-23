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


# ---------------------------------------------------------------------------
# Sparklines
# ---------------------------------------------------------------------------

# Sparkline type constants
XL_SPARKLINE_LINE = 1
XL_SPARKLINE_COLUMN = 2
XL_SPARKLINE_STACKED = 3  # win_loss


def add_sparkline(
    sheet: str | None = None,
    data_range: str = "",
    location_cell: str = "",
    sparkline_type: str = "line",
    color: tuple[int, int, int] | None = None,
) -> dict:
    """Add a sparkline to a cell.

    Args:
        sheet: Optional sheet name.
        data_range: Source data range (e.g. "A1:A10").
        location_cell: Cell where the sparkline is placed.
        sparkline_type: Type of sparkline ("line", "column", "win_loss").
        color: Optional sparkline color as (R, G, B).

    Returns:
        dict with location cell and sheet name.
    """
    type_map = {
        "line": XL_SPARKLINE_LINE,
        "column": XL_SPARKLINE_COLUMN,
        "win_loss": XL_SPARKLINE_STACKED,
    }
    app = _get_app()
    ws = _get_ws(app, sheet)
    sl_type = type_map.get(sparkline_type, XL_SPARKLINE_LINE)
    loc = ws.Range(location_cell)
    sg = loc.SparklineGroups.Add(Type=sl_type, SourceData=data_range)
    if color is not None:
        sg.SeriesColor.Color = rgb(*color)
    return {"location": location_cell, "sheet": ws.Name}


def format_sparkline(
    sheet: str | None = None,
    location_cell: str = "",
    high_point: bool = False,
    low_point: bool = False,
    first_point: bool = False,
    last_point: bool = False,
    negative_points: bool = False,
    markers: bool = False,
    line_weight: float | None = None,
) -> dict:
    """Format sparkline display options.

    Args:
        sheet: Optional sheet name.
        location_cell: Cell containing the sparkline.
        high_point: Show high point marker.
        low_point: Show low point marker.
        first_point: Show first point marker.
        last_point: Show last point marker.
        negative_points: Show negative point markers.
        markers: Show all markers (line sparklines).
        line_weight: Line weight in points.

    Returns:
        dict with location cell and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(location_cell)
    sg = rng.SparklineGroups(1)
    sg.Points.Highpoint.Visible = high_point
    sg.Points.Lowpoint.Visible = low_point
    sg.Points.Firstpoint.Visible = first_point
    sg.Points.Lastpoint.Visible = last_point
    sg.Points.Negative.Visible = negative_points
    sg.Points.Markers.Visible = markers
    if line_weight is not None:
        sg.LineWeight = line_weight
    return {"location": location_cell, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Advanced Conditional Formatting
# ---------------------------------------------------------------------------

# Conditional format types
XL_CF_COLOR_SCALE = 3
XL_CF_DATA_BAR = 4
XL_CF_ICON_SET = 6

# Condition value types
XL_COND_VALUE_LOWEST = 1
XL_COND_VALUE_HIGHEST = 2
XL_COND_VALUE_NUMBER = 0
XL_COND_VALUE_PERCENT = 3
XL_COND_VALUE_PERCENTILE = 5
XL_COND_VALUE_AUTOMATIC = 7

# Icon set types
XL_ICON_3_ARROWS = 1
XL_ICON_3_TRAFFIC_LIGHTS = 4
XL_ICON_3_STARS = 18
XL_ICON_4_ARROWS = 8
XL_ICON_5_ARROWS = 13
XL_ICON_3_FLAGS = 3
XL_ICON_3_SYMBOLS = 6


def set_conditional_formatting_color_scale(
    sheet: str | None = None,
    range_str: str = "",
    min_color: tuple[int, int, int] = (255, 255, 255),
    mid_color: tuple[int, int, int] | None = None,
    max_color: tuple[int, int, int] | None = None,
) -> dict:
    """Add a 2-color or 3-color scale conditional formatting rule.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        min_color: Color for the minimum value as (R, G, B).
        mid_color: Color for the midpoint (if provided, uses 3-color scale).
        max_color: Color for the maximum value as (R, G, B).

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    if mid_color is not None:
        # 3-color scale
        cs = rng.FormatConditions.AddColorScale(ColorScaleType=3)
        cs.ColorScaleCriteria(1).Type = XL_COND_VALUE_LOWEST
        cs.ColorScaleCriteria(1).FormatColor.Color = rgb(*min_color)
        cs.ColorScaleCriteria(2).Type = XL_COND_VALUE_PERCENTILE
        cs.ColorScaleCriteria(2).Value = 50
        cs.ColorScaleCriteria(2).FormatColor.Color = rgb(*mid_color)
        cs.ColorScaleCriteria(3).Type = XL_COND_VALUE_HIGHEST
        cs.ColorScaleCriteria(3).FormatColor.Color = rgb(*max_color) if max_color else rgb(255, 255, 255)
    else:
        # 2-color scale
        cs = rng.FormatConditions.AddColorScale(ColorScaleType=2)
        cs.ColorScaleCriteria(1).Type = XL_COND_VALUE_LOWEST
        cs.ColorScaleCriteria(1).FormatColor.Color = rgb(*min_color)
        cs.ColorScaleCriteria(2).Type = XL_COND_VALUE_HIGHEST
        cs.ColorScaleCriteria(2).FormatColor.Color = rgb(*max_color) if max_color else rgb(255, 255, 255)
    return {"range": range_str, "sheet": ws.Name}


def set_conditional_formatting_data_bar(
    sheet: str | None = None,
    range_str: str = "",
    bar_color: tuple[int, int, int] = (0, 0, 255),
    show_value: bool = True,
    min_type: int | None = None,
    max_type: int | None = None,
) -> dict:
    """Add data bar conditional formatting to a range.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        bar_color: Bar color as (R, G, B).
        show_value: Whether to show the cell value alongside the bar.
        min_type: Minimum value type (0=Number, 1=Lowest, 3=Percent, 5=Percentile, 7=Automatic).
        max_type: Maximum value type.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    db = rng.FormatConditions.AddDatabar()
    db.BarColor.Color = rgb(*bar_color)
    db.ShowValue = show_value
    if min_type is not None:
        db.MinPoint.Modify(newtype=min_type)
    if max_type is not None:
        db.MaxPoint.Modify(newtype=max_type)
    return {"range": range_str, "sheet": ws.Name}


def set_conditional_formatting_icon_set(
    sheet: str | None = None,
    range_str: str = "",
    icon_style: str = "3_arrows",
    reverse: bool = False,
    show_icon_only: bool = False,
) -> dict:
    """Add icon set conditional formatting to a range.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        icon_style: Icon set style name (e.g. "3_arrows", "3_traffic_lights",
                    "3_stars", "4_arrows", "5_arrows", "3_flags", "3_symbols").
        reverse: Reverse the icon order.
        show_icon_only: Show only icons without cell values.

    Returns:
        dict with range and sheet name.
    """
    icon_map = {
        "3_arrows": XL_ICON_3_ARROWS,
        "3_traffic_lights": XL_ICON_3_TRAFFIC_LIGHTS,
        "3_stars": XL_ICON_3_STARS,
        "4_arrows": XL_ICON_4_ARROWS,
        "5_arrows": XL_ICON_5_ARROWS,
        "3_flags": XL_ICON_3_FLAGS,
        "3_symbols": XL_ICON_3_SYMBOLS,
    }
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    icon_id = icon_map.get(icon_style, XL_ICON_3_ARROWS)
    iset = rng.FormatConditions.AddIconSetCondition()
    iset.IconSet = app.ActiveWorkbook.IconSets(icon_id)
    iset.ReverseOrder = reverse
    iset.ShowIconOnly = show_icon_only
    return {"range": range_str, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Grouping & Outline
# ---------------------------------------------------------------------------

def group_rows(
    sheet: str | None = None,
    start_row: int = 1,
    end_row: int = 1,
) -> dict:
    """Group rows (collapsible).

    Args:
        sheet: Optional sheet name.
        start_row: First row number to group.
        end_row: Last row number to group.

    Returns:
        dict with row range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Rows(f"{start_row}:{end_row}").Group()
    return {"rows": f"{start_row}:{end_row}", "sheet": ws.Name}


def ungroup_rows(
    sheet: str | None = None,
    start_row: int = 1,
    end_row: int = 1,
) -> dict:
    """Ungroup rows.

    Args:
        sheet: Optional sheet name.
        start_row: First row number to ungroup.
        end_row: Last row number to ungroup.

    Returns:
        dict with row range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Rows(f"{start_row}:{end_row}").Ungroup()
    return {"rows": f"{start_row}:{end_row}", "sheet": ws.Name}


def group_columns(
    sheet: str | None = None,
    start_col: str = "A",
    end_col: str = "A",
) -> dict:
    """Group columns (collapsible).

    Args:
        sheet: Optional sheet name.
        start_col: First column letter to group.
        end_col: Last column letter to group.

    Returns:
        dict with column range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Columns(f"{start_col}:{end_col}").Group()
    return {"columns": f"{start_col}:{end_col}", "sheet": ws.Name}


def ungroup_columns(
    sheet: str | None = None,
    start_col: str = "A",
    end_col: str = "A",
) -> dict:
    """Ungroup columns.

    Args:
        sheet: Optional sheet name.
        start_col: First column letter to ungroup.
        end_col: Last column letter to ungroup.

    Returns:
        dict with column range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Columns(f"{start_col}:{end_col}").Ungroup()
    return {"columns": f"{start_col}:{end_col}", "sheet": ws.Name}


def set_outline_level(
    sheet: str | None = None,
    show_detail: bool = True,
    summary_below: bool = True,
    summary_right: bool = True,
) -> dict:
    """Configure outline settings for a worksheet.

    Args:
        sheet: Optional sheet name.
        show_detail: Whether to show detail rows/columns.
        summary_below: Place summary rows below detail.
        summary_right: Place summary columns to the right of detail.

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Outline.SummaryRow = 1 if summary_below else 0  # xlAbove=0, xlBelow=1
    ws.Outline.SummaryColumn = 1 if summary_right else 0  # xlLeft=0, xlRight=1
    if not show_detail:
        ws.Outline.ShowLevels(RowLevels=1, ColumnLevels=1)
    return {"sheet": ws.Name}


def collapse_group(
    sheet: str | None = None,
    level: int = 1,
) -> dict:
    """Collapse outline to a specific level.

    Args:
        sheet: Optional sheet name.
        level: Outline level to show (1-8).

    Returns:
        dict with level and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Outline.ShowLevels(RowLevels=level, ColumnLevels=level)
    return {"level": level, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Hyperlinks
# ---------------------------------------------------------------------------

def add_hyperlink(
    sheet: str | None = None,
    cell: str = "",
    url: str = "",
    display_text: str | None = None,
    tooltip: str | None = None,
) -> dict:
    """Add a hyperlink to a cell.

    Args:
        sheet: Optional sheet name.
        cell: Cell address for the hyperlink.
        url: URL target.
        display_text: Optional display text.
        tooltip: Optional tooltip text.

    Returns:
        dict with cell and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    anchor = ws.Range(cell)
    kwargs = {"Anchor": anchor, "Address": url}
    if display_text is not None:
        kwargs["TextToDisplay"] = display_text
    if tooltip is not None:
        kwargs["ScreenTip"] = tooltip
    ws.Hyperlinks.Add(**kwargs)
    return {"cell": cell, "url": url, "sheet": ws.Name}


def add_internal_link(
    sheet: str | None = None,
    cell: str = "",
    target_sheet: str = "",
    target_cell: str = "",
    display_text: str | None = None,
) -> dict:
    """Add a hyperlink to another cell/sheet within the workbook.

    Args:
        sheet: Optional sheet name (source sheet).
        cell: Source cell address.
        target_sheet: Name of the target sheet.
        target_cell: Target cell address.
        display_text: Optional display text.

    Returns:
        dict with cell, target info, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    anchor = ws.Range(cell)
    sub_address = f"'{target_sheet}'!{target_cell}"
    kwargs = {"Anchor": anchor, "Address": "", "SubAddress": sub_address}
    if display_text is not None:
        kwargs["TextToDisplay"] = display_text
    ws.Hyperlinks.Add(**kwargs)
    return {"cell": cell, "target_sheet": target_sheet, "target_cell": target_cell, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Images & Shapes
# ---------------------------------------------------------------------------

# Shape type constants (msoAutoShapeType)
MSO_SHAPE_RECTANGLE = 1
MSO_SHAPE_ROUNDED_RECTANGLE = 5
MSO_SHAPE_OVAL = 9
MSO_SHAPE_DIAMOND = 4
MSO_SHAPE_RIGHT_TRIANGLE = 8
MSO_SHAPE_ISOSCELES_TRIANGLE = 7


def insert_image(
    sheet: str | None = None,
    image_path: str = "",
    cell: str = "A1",
    width: float | None = None,
    height: float | None = None,
) -> dict:
    """Insert an image near a cell on a worksheet.

    Args:
        sheet: Optional sheet name.
        image_path: Path to the image file.
        cell: Cell near which the image is placed.
        width: Optional width in points.
        height: Optional height in points.

    Returns:
        dict with image info and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    abs_path = ensure_absolute_path(image_path)
    anchor = ws.Range(cell)
    left = anchor.Left
    top = anchor.Top
    pic = ws.Shapes.AddPicture(
        Filename=abs_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=left,
        Top=top,
        Width=width if width is not None else -1,
        Height=height if height is not None else -1,
    )
    # If width/height not specified, keep original size
    if width is None and height is None:
        pic.ScaleWidth(1, True)
        pic.ScaleHeight(1, True)
    return {"name": pic.Name, "cell": cell, "sheet": ws.Name}


def add_shape(
    sheet: str | None = None,
    shape_type: int = MSO_SHAPE_RECTANGLE,
    left: float = 100,
    top: float = 100,
    width: float = 100,
    height: float = 50,
    fill_color: tuple[int, int, int] | None = None,
    line_color: tuple[int, int, int] | None = None,
    text: str | None = None,
) -> dict:
    """Add an auto shape to a worksheet.

    Args:
        sheet: Optional sheet name.
        shape_type: AutoShape type constant (1=Rectangle, 5=RoundedRect,
                    9=Oval, 4=Diamond, 7=Triangle, 8=RightTriangle).
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.
        fill_color: Fill color as (R, G, B).
        line_color: Line color as (R, G, B).
        text: Optional text inside the shape.

    Returns:
        dict with shape name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    shp = ws.Shapes.AddShape(shape_type, left, top, width, height)
    if fill_color is not None:
        shp.Fill.ForeColor.RGB = rgb(*fill_color)
    if line_color is not None:
        shp.Line.ForeColor.RGB = rgb(*line_color)
    if text is not None:
        shp.TextFrame.Characters().Text = text
    return {"name": shp.Name, "sheet": ws.Name}


def add_textbox(
    sheet: str | None = None,
    left: float = 100,
    top: float = 100,
    width: float = 200,
    height: float = 50,
    text: str = "",
    font_size: float | None = None,
    font_color: tuple[int, int, int] | None = None,
    bold: bool = False,
) -> dict:
    """Add a textbox to a worksheet.

    Args:
        sheet: Optional sheet name.
        left: Left position in points.
        top: Top position in points.
        width: Width in points.
        height: Height in points.
        text: Text content.
        font_size: Font size in points.
        font_color: Font color as (R, G, B).
        bold: Whether text is bold.

    Returns:
        dict with textbox name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    # msoTextOrientationHorizontal = 1
    tb = ws.Shapes.AddTextbox(1, left, top, width, height)
    tb.TextFrame.Characters().Text = text
    if font_size is not None:
        tb.TextFrame.Characters().Font.Size = font_size
    if font_color is not None:
        tb.TextFrame.Characters().Font.Color = rgb(*font_color)
    if bold:
        tb.TextFrame.Characters().Font.Bold = True
    return {"name": tb.Name, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Advanced Cell Formatting
# ---------------------------------------------------------------------------

def set_cell_style(
    sheet: str | None = None,
    range_str: str = "",
    style_name: str = "Normal",
) -> dict:
    """Apply a built-in Excel style to a range.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        style_name: Built-in style name (e.g. "Heading 1", "Total", "Accent1").

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_str).Style = style_name
    return {"range": range_str, "style": style_name, "sheet": ws.Name}


def auto_fit_columns(
    sheet: str | None = None,
    range_str: str | None = None,
) -> dict:
    """Auto-fit column widths to content.

    Args:
        sheet: Optional sheet name.
        range_str: Optional range to auto-fit. If None, fits all used columns.

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if range_str:
        ws.Range(range_str).Columns.AutoFit()
    else:
        ws.UsedRange.Columns.AutoFit()
    return {"range": range_str or "UsedRange", "sheet": ws.Name}


def auto_fit_rows(
    sheet: str | None = None,
    range_str: str | None = None,
) -> dict:
    """Auto-fit row heights to content.

    Args:
        sheet: Optional sheet name.
        range_str: Optional range to auto-fit. If None, fits all used rows.

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if range_str:
        ws.Range(range_str).Rows.AutoFit()
    else:
        ws.UsedRange.Rows.AutoFit()
    return {"range": range_str or "UsedRange", "sheet": ws.Name}


def set_cell_indent(
    sheet: str | None = None,
    range_str: str = "",
    indent_level: int = 0,
) -> dict:
    """Set the indent level for cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        indent_level: Indent level (0 or greater).

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_str).IndentLevel = indent_level
    return {"range": range_str, "indent": indent_level, "sheet": ws.Name}


def set_text_rotation(
    sheet: str | None = None,
    range_str: str = "",
    angle: int = 0,
) -> dict:
    """Set text rotation angle for cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        angle: Rotation angle (-90 to 90 degrees).

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_str).Orientation = angle
    return {"range": range_str, "angle": angle, "sheet": ws.Name}


def set_wrap_text(
    sheet: str | None = None,
    range_str: str = "",
    wrap: bool = True,
) -> dict:
    """Enable or disable text wrapping for cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        wrap: True to wrap text, False to disable.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_str).WrapText = wrap
    return {"range": range_str, "wrap": wrap, "sheet": ws.Name}


def set_cell_pattern(
    sheet: str | None = None,
    range_str: str = "",
    pattern_type: int = 1,
    fore_color: tuple[int, int, int] | None = None,
    back_color: tuple[int, int, int] | None = None,
) -> dict:
    """Set pattern fill for cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        pattern_type: Pattern type constant (1-18).
        fore_color: Pattern foreground color as (R, G, B).
        back_color: Pattern background color as (R, G, B).

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    rng.Interior.Pattern = pattern_type
    if fore_color is not None:
        rng.Interior.PatternColor = rgb(*fore_color)
    if back_color is not None:
        rng.Interior.Color = rgb(*back_color)
    return {"range": range_str, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Data Features (Advanced)
# ---------------------------------------------------------------------------

def copy_range(
    sheet: str | None = None,
    source_range: str = "",
    target_cell: str = "",
    target_sheet: str | None = None,
) -> dict:
    """Copy a range to another location.

    Args:
        sheet: Optional source sheet name.
        source_range: Source range address.
        target_cell: Target cell address (top-left of paste area).
        target_sheet: Optional target sheet name.

    Returns:
        dict with source, target, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    wb = app.ActiveWorkbook
    dst_ws = wb.Worksheets(target_sheet) if target_sheet else ws
    ws.Range(source_range).Copy(Destination=dst_ws.Range(target_cell))
    return {"source": source_range, "target": target_cell, "sheet": ws.Name, "target_sheet": dst_ws.Name}


def clear_range(
    sheet: str | None = None,
    range_str: str = "",
    clear_type: str = "all",
) -> dict:
    """Clear a range of cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        clear_type: What to clear ("all", "contents", "formats", "comments").

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    if clear_type == "contents":
        rng.ClearContents()
    elif clear_type == "formats":
        rng.ClearFormats()
    elif clear_type == "comments":
        rng.ClearComments()
    else:
        rng.Clear()
    return {"range": range_str, "clear_type": clear_type, "sheet": ws.Name}


def find_value(
    sheet: str | None = None,
    value: str = "",
    match_case: bool = False,
    match_entire: bool = False,
) -> dict:
    """Find a value in a worksheet.

    Args:
        sheet: Optional sheet name.
        value: Value to search for.
        match_case: Case-sensitive search.
        match_entire: Match entire cell contents only.

    Returns:
        dict with found cell address (or None) and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    # xlPart = 2, xlWhole = 1
    look_at = 1 if match_entire else 2
    found = ws.UsedRange.Find(
        What=value,
        MatchCase=match_case,
        LookAt=look_at,
    )
    address = found.Address if found else None
    return {"value": value, "found": address, "sheet": ws.Name}


def replace_value(
    sheet: str | None = None,
    find_text: str = "",
    replace_text: str = "",
    match_case: bool = False,
    match_entire: bool = False,
) -> dict:
    """Find and replace values in a worksheet.

    Args:
        sheet: Optional sheet name.
        find_text: Text to find.
        replace_text: Replacement text.
        match_case: Case-sensitive search.
        match_entire: Match entire cell contents only.

    Returns:
        dict with find/replace info and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    look_at = 1 if match_entire else 2
    ws.UsedRange.Replace(
        What=find_text,
        Replacement=replace_text,
        MatchCase=match_case,
        LookAt=look_at,
    )
    return {"find": find_text, "replace": replace_text, "sheet": ws.Name}


def remove_duplicates(
    sheet: str | None = None,
    range_str: str = "",
    columns: list[int] | None = None,
) -> dict:
    """Remove duplicate rows from a range.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        columns: List of column numbers (1-based) to check for duplicates.
                 If None, checks all columns.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    if columns is None:
        col_count = rng.Columns.Count
        columns = list(range(1, col_count + 1))
    rng.RemoveDuplicates(Columns=columns, Header=1)  # xlYes=1
    return {"range": range_str, "sheet": ws.Name}


def text_to_columns(
    sheet: str | None = None,
    range_str: str = "",
    delimiter: str = ",",
) -> dict:
    """Split text in cells into multiple columns.

    Args:
        sheet: Optional sheet name.
        range_str: Source range address (typically a single column).
        delimiter: Delimiter character.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    # xlDelimited = 1
    delimiter_map = {
        ",": {"Comma": True},
        "\t": {"Tab": True},
        ";": {"Semicolon": True},
        " ": {"Space": True},
    }
    kwargs = delimiter_map.get(delimiter, {"Other": True, "OtherChar": delimiter})
    rng.TextToColumns(
        Destination=rng.Cells(1, 1),
        DataType=1,
        **kwargs,
    )
    return {"range": range_str, "delimiter": delimiter, "sheet": ws.Name}


def transpose_range(
    sheet: str | None = None,
    source_range: str = "",
    target_cell: str = "",
) -> dict:
    """Transpose data from source range to target location.

    Args:
        sheet: Optional sheet name.
        source_range: Source range address.
        target_cell: Target cell address (top-left of paste area).

    Returns:
        dict with source, target, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    src = ws.Range(source_range)
    src.Copy()
    dst = ws.Range(target_cell)
    # xlPasteAll = -4104, xlPasteSpecialOperationNone = -4142
    dst.PasteSpecial(Paste=-4104, Operation=-4142, Transpose=True)
    app.CutCopyMode = False
    return {"source": source_range, "target": target_cell, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Advanced Chart
# ---------------------------------------------------------------------------

# Legend position constants
XL_LEGEND_BOTTOM = -4107
XL_LEGEND_TOP = -4160
XL_LEGEND_LEFT = -4131
XL_LEGEND_RIGHT = -4152

# Trendline types
XL_TREND_LINEAR = -4132
XL_TREND_EXPONENTIAL = 5
XL_TREND_LOGARITHMIC = -4133
XL_TREND_POLYNOMIAL = 3
XL_TREND_POWER = 4
XL_TREND_MOVING_AVERAGE = 6

# Marker styles
XL_MARKER_NONE = -4142
XL_MARKER_CIRCLE = 8
XL_MARKER_SQUARE = 1
XL_MARKER_TRIANGLE = 3
XL_MARKER_DIAMOND = 2


def format_chart(
    sheet: str | None = None,
    chart_index: int = 1,
    title: str | None = None,
    x_axis_title: str | None = None,
    y_axis_title: str | None = None,
    legend_position: str | None = None,
    style: int | None = None,
) -> dict:
    """Format an existing chart.

    Args:
        sheet: Optional sheet name.
        chart_index: Index of the chart object (1-based).
        title: Chart title text.
        x_axis_title: X-axis title.
        y_axis_title: Y-axis title.
        legend_position: Legend position ("bottom", "top", "left", "right", "none").
        style: Chart style number.

    Returns:
        dict with chart name and sheet name.
    """
    legend_map = {
        "bottom": XL_LEGEND_BOTTOM,
        "top": XL_LEGEND_TOP,
        "left": XL_LEGEND_LEFT,
        "right": XL_LEGEND_RIGHT,
    }
    app = _get_app()
    ws = _get_ws(app, sheet)
    chart_obj = ws.ChartObjects(chart_index)
    chart = chart_obj.Chart
    if title is not None:
        chart.HasTitle = True
        chart.ChartTitle.Text = title
    if x_axis_title is not None:
        # xlCategory = 1
        axis = chart.Axes(1)
        axis.HasTitle = True
        axis.AxisTitle.Text = x_axis_title
    if y_axis_title is not None:
        # xlValue = 2
        axis = chart.Axes(2)
        axis.HasTitle = True
        axis.AxisTitle.Text = y_axis_title
    if legend_position is not None:
        if legend_position == "none":
            chart.HasLegend = False
        else:
            chart.HasLegend = True
            chart.Legend.Position = legend_map.get(legend_position, XL_LEGEND_BOTTOM)
    if style is not None:
        chart.ChartStyle = style
    return {"chart": chart_obj.Name, "sheet": ws.Name}


def format_chart_series(
    sheet: str | None = None,
    chart_index: int = 1,
    series_index: int = 1,
    color: tuple[int, int, int] | None = None,
    line_weight: float | None = None,
    marker_style: int | None = None,
    marker_size: int | None = None,
) -> dict:
    """Format an individual data series in a chart.

    Args:
        sheet: Optional sheet name.
        chart_index: Index of the chart object (1-based).
        series_index: Index of the data series (1-based).
        color: Series color as (R, G, B).
        line_weight: Line weight in points.
        marker_style: Marker style constant (-4142=None, 8=Circle, 1=Square,
                      3=Triangle, 2=Diamond).
        marker_size: Marker size in points.

    Returns:
        dict with chart and series info.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    chart = ws.ChartObjects(chart_index).Chart
    series = chart.SeriesCollection(series_index)
    if color is not None:
        series.Format.Fill.ForeColor.RGB = rgb(*color)
        series.Format.Line.ForeColor.RGB = rgb(*color)
    if line_weight is not None:
        series.Format.Line.Weight = line_weight
    if marker_style is not None:
        series.MarkerStyle = marker_style
    if marker_size is not None:
        series.MarkerSize = marker_size
    return {"chart_index": chart_index, "series_index": series_index, "sheet": ws.Name}


def add_chart_trendline(
    sheet: str | None = None,
    chart_index: int = 1,
    series_index: int = 1,
    trend_type: str = "linear",
    display_equation: bool = False,
    display_r_squared: bool = False,
) -> dict:
    """Add a trendline to a chart series.

    Args:
        sheet: Optional sheet name.
        chart_index: Index of the chart object (1-based).
        series_index: Index of the data series (1-based).
        trend_type: Trendline type ("linear", "exponential", "logarithmic",
                    "polynomial", "power", "moving_average").
        display_equation: Show the trendline equation.
        display_r_squared: Show R-squared value.

    Returns:
        dict with chart and series info.
    """
    trend_map = {
        "linear": XL_TREND_LINEAR,
        "exponential": XL_TREND_EXPONENTIAL,
        "logarithmic": XL_TREND_LOGARITHMIC,
        "polynomial": XL_TREND_POLYNOMIAL,
        "power": XL_TREND_POWER,
        "moving_average": XL_TREND_MOVING_AVERAGE,
    }
    app = _get_app()
    ws = _get_ws(app, sheet)
    chart = ws.ChartObjects(chart_index).Chart
    series = chart.SeriesCollection(series_index)
    xl_type = trend_map.get(trend_type, XL_TREND_LINEAR)
    tl = series.Trendlines().Add(Type=xl_type)
    tl.DisplayEquation = display_equation
    tl.DisplayRSquared = display_r_squared
    return {"chart_index": chart_index, "series_index": series_index, "trend_type": trend_type, "sheet": ws.Name}


def set_chart_area_format(
    sheet: str | None = None,
    chart_index: int = 1,
    fill_color: tuple[int, int, int] | None = None,
    border_color: tuple[int, int, int] | None = None,
    border_weight: float | None = None,
) -> dict:
    """Format the chart area (background and border).

    Args:
        sheet: Optional sheet name.
        chart_index: Index of the chart object (1-based).
        fill_color: Chart area fill color as (R, G, B).
        border_color: Chart area border color as (R, G, B).
        border_weight: Border line weight in points.

    Returns:
        dict with chart info and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    chart = ws.ChartObjects(chart_index).Chart
    if fill_color is not None:
        chart.ChartArea.Format.Fill.ForeColor.RGB = rgb(*fill_color)
    if border_color is not None:
        chart.ChartArea.Format.Line.ForeColor.RGB = rgb(*border_color)
    if border_weight is not None:
        chart.ChartArea.Format.Line.Weight = border_weight
    return {"chart_index": chart_index, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Pivot Table Advanced
# ---------------------------------------------------------------------------

def format_pivot_table(
    sheet: str | None = None,
    pivot_name: str = "",
    style: str | None = None,
    show_grand_total_rows: bool | None = None,
    show_grand_total_cols: bool | None = None,
    repeat_item_labels: bool | None = None,
) -> dict:
    """Format a pivot table.

    Args:
        sheet: Optional sheet name.
        pivot_name: Name of the pivot table.
        style: Pivot table style name (e.g. "PivotStyleMedium9").
        show_grand_total_rows: Show grand total for rows.
        show_grand_total_cols: Show grand total for columns.
        repeat_item_labels: Repeat all item labels.

    Returns:
        dict with pivot table name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    pt = ws.PivotTables(pivot_name)
    if style is not None:
        pt.TableStyle2 = style
    if show_grand_total_rows is not None:
        pt.RowGrand = show_grand_total_rows
    if show_grand_total_cols is not None:
        pt.ColumnGrand = show_grand_total_cols
    if repeat_item_labels is not None:
        pt.RepeatAllLabels(1 if repeat_item_labels else 2)  # xlRepeatLabels=1, xlDoNotRepeatLabels=2
    return {"pivot_name": pivot_name, "sheet": ws.Name}


def add_pivot_field(
    sheet: str | None = None,
    pivot_name: str = "",
    field_name: str = "",
    area: str = "row",
    position: int | None = None,
) -> dict:
    """Add a field to a pivot table.

    Args:
        sheet: Optional sheet name.
        pivot_name: Name of the pivot table.
        field_name: Name of the field to add.
        area: Target area ("row", "column", "data", "filter").
        position: Optional position within the area.

    Returns:
        dict with pivot table and field info.
    """
    area_map = {
        "row": XL_PIVOT_ROW,
        "column": XL_PIVOT_COLUMN,
        "data": XL_PIVOT_DATA,
        "filter": XL_PIVOT_PAGE,
    }
    app = _get_app()
    ws = _get_ws(app, sheet)
    pt = ws.PivotTables(pivot_name)
    pf = pt.PivotFields(field_name)
    pf.Orientation = area_map.get(area, XL_PIVOT_ROW)
    if position is not None:
        pf.Position = position
    return {"pivot_name": pivot_name, "field": field_name, "area": area, "sheet": ws.Name}


def refresh_pivot_table(
    sheet: str | None = None,
    pivot_name: str | None = None,
) -> dict:
    """Refresh pivot table(s).

    Args:
        sheet: Optional sheet name.
        pivot_name: Name of specific pivot table to refresh.
                    If None, refreshes all pivot tables on the sheet.

    Returns:
        dict with refresh info and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if pivot_name:
        ws.PivotTables(pivot_name).RefreshTable()
        return {"pivot_name": pivot_name, "sheet": ws.Name}
    else:
        count = ws.PivotTables().Count
        for i in range(1, count + 1):
            ws.PivotTables(i).RefreshTable()
        return {"refreshed_count": count, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Workbook Features
# ---------------------------------------------------------------------------

def set_workbook_properties(
    title: str | None = None,
    author: str | None = None,
    subject: str | None = None,
    keywords: str | None = None,
    comments: str | None = None,
) -> dict:
    """Set document metadata properties.

    Args:
        title: Document title.
        author: Document author.
        subject: Document subject.
        keywords: Document keywords.
        comments: Document comments.

    Returns:
        dict with workbook name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    props = wb.BuiltinDocumentProperties
    if title is not None:
        props("Title").Value = title
    if author is not None:
        props("Author").Value = author
    if subject is not None:
        props("Subject").Value = subject
    if keywords is not None:
        props("Keywords").Value = keywords
    if comments is not None:
        props("Comments").Value = comments
    return {"name": wb.Name}


def get_workbook_statistics() -> dict:
    """Get workbook statistics including sheet count and used ranges.

    Returns:
        dict with workbook name, sheet count, and per-sheet info.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    sheets_info = []
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        used = ws.UsedRange
        sheets_info.append({
            "name": ws.Name,
            "used_range": used.Address,
            "rows": used.Rows.Count,
            "cols": used.Columns.Count,
        })
    return {
        "name": wb.Name,
        "sheet_count": wb.Worksheets.Count,
        "sheets": sheets_info,
    }


def set_tab_color(
    sheet: str | None = None,
    color: tuple[int, int, int] = (0, 0, 0),
) -> dict:
    """Set the worksheet tab color.

    Args:
        sheet: Optional sheet name.
        color: Tab color as (R, G, B).

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Tab.Color = rgb(*color)
    return {"sheet": ws.Name}


def hide_sheet(sheet: str) -> dict:
    """Hide a worksheet.

    Args:
        sheet: Name of the sheet to hide.

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Visible = False
    return {"sheet": ws.Name, "visible": False}


def unhide_sheet(sheet: str) -> dict:
    """Unhide a worksheet.

    Args:
        sheet: Name of the sheet to unhide.

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Visible = True
    return {"sheet": ws.Name, "visible": True}


def copy_sheet(
    source_sheet: str,
    target_name: str | None = None,
    before: str | None = None,
    after: str | None = None,
) -> dict:
    """Copy a worksheet.

    Args:
        source_sheet: Name of the sheet to copy.
        target_name: New name for the copied sheet.
        before: Name of sheet to place copy before.
        after: Name of sheet to place copy after.

    Returns:
        dict with source and new sheet name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    src = wb.Worksheets(source_sheet)
    kwargs = {}
    if before:
        kwargs["Before"] = wb.Worksheets(before)
    elif after:
        kwargs["After"] = wb.Worksheets(after)
    else:
        kwargs["After"] = wb.Worksheets(wb.Worksheets.Count)
    src.Copy(**kwargs)
    new_ws = app.ActiveSheet
    if target_name:
        new_ws.Name = target_name
    return {"source": source_sheet, "new_sheet": new_ws.Name}


def move_sheet(
    sheet: str,
    before: str | None = None,
    after: str | None = None,
) -> dict:
    """Move a worksheet to a new position.

    Args:
        sheet: Name of the sheet to move.
        before: Name of sheet to place before.
        after: Name of sheet to place after.

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    ws = wb.Worksheets(sheet)
    kwargs = {}
    if before:
        kwargs["Before"] = wb.Worksheets(before)
    elif after:
        kwargs["After"] = wb.Worksheets(after)
    ws.Move(**kwargs)
    return {"sheet": sheet}


# ---------------------------------------------------------------------------
# Print & Page Setup Advanced
# ---------------------------------------------------------------------------

def set_print_area(
    sheet: str | None = None,
    range_str: str = "",
) -> dict:
    """Set the print area for a worksheet.

    Args:
        sheet: Optional sheet name.
        range_str: Range address to set as print area.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.PageSetup.PrintArea = range_str
    return {"print_area": range_str, "sheet": ws.Name}


def set_print_titles(
    sheet: str | None = None,
    rows: str | None = None,
    columns: str | None = None,
) -> dict:
    """Set rows/columns to repeat on each printed page.

    Args:
        sheet: Optional sheet name.
        rows: Row range to repeat (e.g. "1:2").
        columns: Column range to repeat (e.g. "A:B").

    Returns:
        dict with sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if rows is not None:
        ws.PageSetup.PrintTitleRows = f"${rows.replace(':', ':$')}" if not rows.startswith("$") else rows
    if columns is not None:
        ws.PageSetup.PrintTitleColumns = f"${columns.replace(':', ':$')}" if not columns.startswith("$") else columns
    return {"rows": rows, "columns": columns, "sheet": ws.Name}


def add_page_break(
    sheet: str | None = None,
    row: int | None = None,
    col: int | None = None,
) -> dict:
    """Insert a page break.

    Args:
        sheet: Optional sheet name.
        row: Row number for horizontal page break.
        col: Column number for vertical page break.

    Returns:
        dict with break info and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if row is not None:
        ws.HPageBreaks.Add(Before=ws.Rows(row))
    if col is not None:
        ws.VPageBreaks.Add(Before=ws.Columns(col))
    return {"row": row, "col": col, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Protection Advanced
# ---------------------------------------------------------------------------

def protect_workbook(
    password: str | None = None,
    structure: bool = True,
    windows: bool = False,
) -> dict:
    """Protect the workbook structure and/or windows.

    Args:
        password: Optional protection password.
        structure: Protect workbook structure (prevent adding/deleting sheets).
        windows: Protect workbook windows.

    Returns:
        dict with workbook name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    kwargs = {"Structure": structure, "Windows": windows}
    if password is not None:
        kwargs["Password"] = password
    wb.Protect(**kwargs)
    return {"name": wb.Name, "protected": True}


def unprotect_workbook(password: str | None = None) -> dict:
    """Unprotect the workbook.

    Args:
        password: Password used to protect the workbook.

    Returns:
        dict with workbook name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    if password is not None:
        wb.Unprotect(Password=password)
    else:
        wb.Unprotect()
    return {"name": wb.Name, "protected": False}


def lock_cells(
    sheet: str | None = None,
    range_str: str = "",
    locked: bool = True,
) -> dict:
    """Lock or unlock specific cells (works with sheet protection).

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        locked: True to lock cells, False to unlock.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Range(range_str).Locked = locked
    return {"range": range_str, "locked": locked, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Advanced Formula Functions
# ---------------------------------------------------------------------------

def set_array_formula(sheet: str | None, range_str: str, formula: str) -> dict:
    """Set an array formula (CSE or dynamic) on a range.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address (e.g. "A1:A10").
        formula: Array formula string (e.g. "=TRANSPOSE(B1:D1)").

    Returns:
        dict with range, formula, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    rng.FormulaArray = formula
    return {"range": range_str, "formula": formula, "sheet": ws.Name}


def evaluate_formula(sheet: str | None, formula: str) -> dict:
    """Evaluate a formula and return the result without placing it in a cell.

    Args:
        sheet: Optional sheet name (for context).
        formula: Formula to evaluate (e.g. "=SUM(1,2,3)").

    Returns:
        dict with formula and result.
    """
    app = _get_app()
    _get_ws(app, sheet)  # ensure sheet context
    result = app.Evaluate(formula)
    return {"formula": formula, "result": result}


def set_formula_range(sheet: str | None, start_cell: str, formulas: list[list[str]]) -> dict:
    """Set multiple formulas at once from a 2D list.

    Args:
        sheet: Optional sheet name.
        start_cell: Top-left cell address (e.g. "A1").
        formulas: 2D list of formula strings.

    Returns:
        dict with start cell, dimensions, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    start = ws.Range(start_cell)
    rows = len(formulas)
    cols = len(formulas[0]) if rows > 0 else 0
    for r_idx, row in enumerate(formulas):
        for c_idx, formula in enumerate(row):
            ws.Cells(start.Row + r_idx, start.Column + c_idx).Formula = formula
    return {"start_cell": start_cell, "rows": rows, "cols": cols, "sheet": ws.Name}


def create_named_formula(name: str, formula: str, sheet: str | None = None) -> dict:
    """Create a named formula (not a named range).

    Args:
        name: Name for the formula.
        formula: Formula string (e.g. "=Sheet1!A1*2").
        sheet: Optional sheet name for scope. If None, workbook-level.

    Returns:
        dict with name and formula.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    if sheet is not None:
        ws = wb.Worksheets(sheet)
        # Sheet-level named formula
        ws.Names.Add(Name=name, RefersTo=formula)
    else:
        wb.Names.Add(Name=name, RefersTo=formula)
    return {"name": name, "formula": formula, "scope": sheet or "Workbook"}


# ---------------------------------------------------------------------------
# Data Analysis
# ---------------------------------------------------------------------------

def calculate_statistics(sheet: str | None, range_str: str) -> dict:
    """Calculate descriptive statistics for a range.

    Args:
        sheet: Optional sheet name.
        range_str: Data range address (e.g. "A1:A100").

    Returns:
        dict with sum, avg, min, max, count, stdev, median.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng_ref = f"'{ws.Name}'!{range_str}"
    stats = {
        "sum": app.Evaluate(f"=SUM({rng_ref})"),
        "avg": app.Evaluate(f"=AVERAGE({rng_ref})"),
        "min": app.Evaluate(f"=MIN({rng_ref})"),
        "max": app.Evaluate(f"=MAX({rng_ref})"),
        "count": app.Evaluate(f"=COUNT({rng_ref})"),
        "stdev": app.Evaluate(f"=STDEV({rng_ref})"),
        "median": app.Evaluate(f"=MEDIAN({rng_ref})"),
        "sheet": ws.Name,
        "range": range_str,
    }
    return stats


def create_frequency_distribution(
    sheet: str | None,
    data_range: str,
    bins_range: str,
    output_cell: str,
) -> dict:
    """Create a frequency distribution.

    Args:
        sheet: Optional sheet name.
        data_range: Range containing data values.
        bins_range: Range containing bin boundaries.
        output_cell: Top-left cell for output.

    Returns:
        dict with output location and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    data_ref = f"'{ws.Name}'!{data_range}"
    bins_ref = f"'{ws.Name}'!{bins_range}"
    # Count the number of bins to determine output size
    bins_count = ws.Range(bins_range).Rows.Count
    output_rows = bins_count + 1  # FREQUENCY returns one more than bins
    out_start = ws.Range(output_cell)
    end_row = out_start.Row + output_rows - 1
    end_col_letter = out_start.Columns(1).Address.split("$")[1]
    out_range_str = f"{output_cell}:{end_col_letter}{end_row}"
    out_rng = ws.Range(out_range_str)
    out_rng.FormulaArray = f"=FREQUENCY({data_ref},{bins_ref})"
    return {"output_range": out_range_str, "bins_count": bins_count, "sheet": ws.Name}


def goal_seek(
    sheet: str | None,
    target_cell: str,
    target_value: float,
    changing_cell: str,
) -> dict:
    """Perform Goal Seek analysis.

    Args:
        sheet: Optional sheet name.
        target_cell: Cell containing the formula to reach target value.
        target_value: Desired value for the target cell.
        changing_cell: Cell to adjust to reach the target.

    Returns:
        dict with target cell, achieved value, changing cell value, and sheet.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    target_rng = ws.Range(target_cell)
    changing_rng = ws.Range(changing_cell)
    target_rng.GoalSeek(Goal=target_value, ChangingCell=changing_rng)
    return {
        "target_cell": target_cell,
        "target_value": target_value,
        "achieved_value": target_rng.Value,
        "changing_cell": changing_cell,
        "changing_value": changing_rng.Value,
        "sheet": ws.Name,
    }


def create_data_table_analysis(
    sheet: str | None,
    row_input_cell: str | None,
    col_input_cell: str | None,
    formula_cell: str,
    row_values_range: str | None = None,
    col_values_range: str | None = None,
) -> dict:
    """Create a What-If data table for sensitivity analysis.

    Args:
        sheet: Optional sheet name.
        row_input_cell: Row input cell reference (for one/two-variable table).
        col_input_cell: Column input cell reference (for one/two-variable table).
        formula_cell: Cell containing the formula.
        row_values_range: Range with row input values (top row of table).
        col_values_range: Range with column input values (left column of table).

    Returns:
        dict with table info and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    # Build the data table range
    # For a one-variable (column) table: formula in top-left, values in left column
    # For a two-variable table: formula in top-left, row values on top, col values on left
    row_input = ws.Range(row_input_cell) if row_input_cell else None
    col_input = ws.Range(col_input_cell) if col_input_cell else None

    # Determine table range from the values ranges and formula cell
    formula_rng = ws.Range(formula_cell)
    if row_values_range and col_values_range:
        # Two-variable table
        row_vals = ws.Range(row_values_range)
        col_vals = ws.Range(col_values_range)
        last_col = row_vals.Columns(row_vals.Columns.Count).Column
        last_row = col_vals.Rows(col_vals.Rows.Count).Row
        table_range = ws.Range(
            ws.Cells(formula_rng.Row, formula_rng.Column),
            ws.Cells(last_row, last_col),
        )
    elif col_values_range:
        # One-variable column table
        col_vals = ws.Range(col_values_range)
        last_row = col_vals.Rows(col_vals.Rows.Count).Row
        table_range = ws.Range(
            ws.Cells(formula_rng.Row, formula_rng.Column),
            ws.Cells(last_row, formula_rng.Column),
        )
    elif row_values_range:
        # One-variable row table
        row_vals = ws.Range(row_values_range)
        last_col = row_vals.Columns(row_vals.Columns.Count).Column
        table_range = ws.Range(
            ws.Cells(formula_rng.Row, formula_rng.Column),
            ws.Cells(formula_rng.Row, last_col),
        )
    else:
        raise ValueError("At least one of row_values_range or col_values_range is required")

    if row_input and col_input:
        table_range.Table(RowInput=row_input, ColumnInput=col_input)
    elif col_input:
        table_range.Table(ColumnInput=col_input)
    elif row_input:
        table_range.Table(RowInput=row_input)

    return {
        "formula_cell": formula_cell,
        "table_address": table_range.Address,
        "sheet": ws.Name,
    }


# ---------------------------------------------------------------------------
# Advanced Formatting
# ---------------------------------------------------------------------------

def set_rich_text_cell(sheet: str | None, cell: str, runs: list[dict]) -> dict:
    """Set rich text (multiple formatted runs) in a single cell.

    Args:
        sheet: Optional sheet name.
        cell: Cell address (e.g. "A1").
        runs: List of dicts with keys: text, bold, italic, color, size, underline.
              Example: [{"text": "Hello ", "bold": True, "color": [255,0,0]},
                        {"text": "World", "italic": True}]

    Returns:
        dict with cell address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    # Build the full text first
    full_text = "".join(run["text"] for run in runs)
    rng.Value = full_text
    # Apply formatting to each run
    pos = 1  # COM uses 1-based character positions
    for run in runs:
        length = len(run["text"])
        chars = rng.Characters(pos, length)
        if run.get("bold") is not None:
            chars.Font.Bold = run["bold"]
        if run.get("italic") is not None:
            chars.Font.Italic = run["italic"]
        if run.get("color") is not None:
            c = run["color"]
            chars.Font.Color = rgb(c[0], c[1], c[2])
        if run.get("size") is not None:
            chars.Font.Size = run["size"]
        if run.get("underline") is not None:
            chars.Font.Underline = run["underline"]
        pos += length
    return {"cell": cell, "runs_count": len(runs), "sheet": ws.Name}


def add_cell_dropdown(
    sheet: str | None,
    cell: str,
    items: list[str],
    show_error: bool = True,
) -> dict:
    """Add a simple dropdown list to a cell.

    Args:
        sheet: Optional sheet name.
        cell: Cell address (e.g. "B2").
        items: List of dropdown items.
        show_error: Show error if invalid value entered.

    Returns:
        dict with cell address and item count.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    rng.Validation.Delete()
    formula_list = ",".join(items)
    rng.Validation.Add(Type=XL_DV_LIST, Formula1=formula_list)
    rng.Validation.ShowError = show_error
    return {"cell": cell, "items_count": len(items), "sheet": ws.Name}


def set_cell_hyperlink_format(
    sheet: str | None,
    cell: str,
    display_text: str | None = None,
    color: list[int] | None = None,
    underline: bool = True,
) -> dict:
    """Format the appearance of a hyperlink cell.

    Args:
        sheet: Optional sheet name.
        cell: Cell address (e.g. "A1").
        display_text: Optional display text override.
        color: Optional RGB color list [r, g, b].
        underline: Whether to underline the text.

    Returns:
        dict with cell address and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    if display_text is not None:
        rng.Value = display_text
    if color is not None:
        rng.Font.Color = rgb(color[0], color[1], color[2])
    # 2 = xlUnderlineStyleSingle, -4142 = xlUnderlineStyleNone
    rng.Font.Underline = 2 if underline else -4142
    return {"cell": cell, "sheet": ws.Name}


def clear_all_formatting(sheet: str | None, range_str: str | None = None) -> dict:
    """Clear all formatting from a range or the entire sheet.

    Args:
        sheet: Optional sheet name.
        range_str: Optional range address. If None, clears entire sheet.

    Returns:
        dict with range (or "All") and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if range_str:
        ws.Range(range_str).ClearFormats()
    else:
        ws.Cells.ClearFormats()
    return {"range": range_str or "All", "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Chart Types Extended
# ---------------------------------------------------------------------------

# Additional chart type constants
XL_CHART_PIE_EXPLODED = 69
XL_CHART_SCATTER_LINES = -4169
XL_CHART_SCATTER_SMOOTH = 72
XL_CHART_BUBBLE = 15
XL_CHART_RADAR = -4151
XL_CHART_RADAR_FILLED = 82
XL_CHART_STOCK_HLC = 88
XL_CHART_STOCK_OHLC = 89
XL_CHART_AREA_STACKED = 76
XL_CHART_LINE_STACKED = 63
XL_CHART_COLUMN_CLUSTERED_2 = 51


def _cell_to_position(ws, position_cell: str) -> tuple[float, float]:
    """Convert a cell address to (left, top) in points."""
    cell = ws.Range(position_cell)
    return cell.Left, cell.Top


def add_combo_chart(
    sheet: str | None,
    data_range: str,
    chart_types: list[str],
    series_on_secondary: list[int] | None = None,
    position_cell: str = "E1",
    width: float = 480,
    height: float = 300,
) -> dict:
    """Add a combo chart with multiple chart types per series.

    Args:
        sheet: Optional sheet name.
        data_range: Source data range address.
        chart_types: List of chart type strings per series ("column", "line", "area").
        series_on_secondary: List of 1-based series indices on secondary axis.
        position_cell: Cell for chart position.
        width: Chart width in points.
        height: Chart height in points.

    Returns:
        dict with chart name and sheet name.
    """
    type_map = {
        "column": XL_CHART_COLUMN_CLUSTERED,
        "line": XL_CHART_LINE,
        "area": XL_CHART_AREA,
    }
    app = _get_app()
    ws = _get_ws(app, sheet)
    left, top = _cell_to_position(ws, position_cell)
    chart_obj = ws.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=ws.Range(data_range))
    # Default to column clustered initially
    chart.ChartType = XL_CHART_COLUMN_CLUSTERED
    # Set chart type per series
    for i, ct in enumerate(chart_types):
        if i < chart.SeriesCollection().Count:
            series = chart.SeriesCollection(i + 1)
            series.ChartType = type_map.get(ct, XL_CHART_COLUMN_CLUSTERED)
    # Set secondary axis for specified series
    if series_on_secondary:
        for idx in series_on_secondary:
            if idx <= chart.SeriesCollection().Count:
                chart.SeriesCollection(idx).AxisGroup = 2  # xlSecondary
    return {"chart_name": chart_obj.Name, "sheet": ws.Name}


def add_pie_chart(
    sheet: str | None,
    data_range: str,
    position_cell: str = "E1",
    width: float = 400,
    height: float = 300,
    explode: list[int] | None = None,
    show_percentage: bool = True,
) -> dict:
    """Add a pie chart.

    Args:
        sheet: Optional sheet name.
        data_range: Source data range address.
        position_cell: Cell for chart position.
        width: Chart width in points.
        height: Chart height in points.
        explode: List of 1-based point indices to explode (separate from pie).
        show_percentage: Show percentage data labels.

    Returns:
        dict with chart name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    left, top = _cell_to_position(ws, position_cell)
    chart_obj = ws.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=ws.Range(data_range))
    chart.ChartType = XL_CHART_PIE
    if show_percentage:
        chart.SeriesCollection(1).HasDataLabels = True
        chart.SeriesCollection(1).DataLabels().ShowPercentage = True
        chart.SeriesCollection(1).DataLabels().ShowValue = False
    if explode:
        for pt_idx in explode:
            chart.SeriesCollection(1).Points(pt_idx).Explosion = 25
    return {"chart_name": chart_obj.Name, "sheet": ws.Name}


def add_scatter_chart(
    sheet: str | None,
    data_range: str,
    position_cell: str = "E1",
    width: float = 480,
    height: float = 300,
    show_trendline: bool = False,
    bubble: bool = False,
) -> dict:
    """Add a scatter or bubble chart.

    Args:
        sheet: Optional sheet name.
        data_range: Source data range address.
        position_cell: Cell for chart position.
        width: Chart width in points.
        height: Chart height in points.
        show_trendline: Add linear trendline.
        bubble: Use bubble chart instead of scatter.

    Returns:
        dict with chart name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    left, top = _cell_to_position(ws, position_cell)
    chart_obj = ws.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=ws.Range(data_range))
    chart.ChartType = XL_CHART_BUBBLE if bubble else XL_CHART_SCATTER
    if show_trendline and chart.SeriesCollection().Count > 0:
        chart.SeriesCollection(1).Trendlines().Add()
    return {"chart_name": chart_obj.Name, "sheet": ws.Name}


def add_stock_chart(
    sheet: str | None,
    data_range: str,
    position_cell: str = "E1",
    width: float = 480,
    height: float = 300,
    chart_subtype: str = "hlc",
) -> dict:
    """Add a stock chart.

    Args:
        sheet: Optional sheet name.
        data_range: Source data range address.
        position_cell: Cell for chart position.
        width: Chart width in points.
        height: Chart height in points.
        chart_subtype: "hlc" for High-Low-Close, "ohlc" for Open-High-Low-Close.

    Returns:
        dict with chart name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    left, top = _cell_to_position(ws, position_cell)
    chart_obj = ws.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=ws.Range(data_range))
    chart.ChartType = XL_CHART_STOCK_OHLC if chart_subtype == "ohlc" else XL_CHART_STOCK_HLC
    return {"chart_name": chart_obj.Name, "sheet": ws.Name}


def add_radar_chart(
    sheet: str | None,
    data_range: str,
    position_cell: str = "E1",
    width: float = 400,
    height: float = 300,
    filled: bool = False,
) -> dict:
    """Add a radar (spider) chart.

    Args:
        sheet: Optional sheet name.
        data_range: Source data range address.
        position_cell: Cell for chart position.
        width: Chart width in points.
        height: Chart height in points.
        filled: Use filled radar chart.

    Returns:
        dict with chart name and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    left, top = _cell_to_position(ws, position_cell)
    chart_obj = ws.ChartObjects().Add(left, top, width, height)
    chart = chart_obj.Chart
    chart.SetSourceData(Source=ws.Range(data_range))
    chart.ChartType = XL_CHART_RADAR_FILLED if filled else XL_CHART_RADAR
    return {"chart_name": chart_obj.Name, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Workbook Navigation
# ---------------------------------------------------------------------------

def activate_sheet(sheet: str) -> dict:
    """Activate (switch to) a specific worksheet.

    Args:
        sheet: Sheet name to activate.

    Returns:
        dict with activated sheet name.
    """
    app = _get_app()
    wb = app.ActiveWorkbook
    ws = wb.Worksheets(sheet)
    ws.Activate()
    return {"sheet": ws.Name}


def get_active_sheet() -> dict:
    """Get the name of the currently active sheet.

    Returns:
        dict with active sheet name.
    """
    app = _get_app()
    ws = app.ActiveWorkbook.ActiveSheet
    return {"sheet": ws.Name}


def get_used_range(sheet: str | None = None) -> dict:
    """Get the used range address and dimensions.

    Args:
        sheet: Optional sheet name.

    Returns:
        dict with address, rows, columns, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    used = ws.UsedRange
    return {
        "address": used.Address,
        "rows": used.Rows.Count,
        "columns": used.Columns.Count,
        "sheet": ws.Name,
    }


def get_last_row(sheet: str | None = None, column: str = "A") -> dict:
    """Get the last used row number in a specific column.

    Args:
        sheet: Optional sheet name.
        column: Column letter (default "A").

    Returns:
        dict with last row number, column, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    # Find last used row by going up from the bottom
    last_row = ws.Cells(ws.Rows.Count, ws.Range(f"{column}1").Column).End(-4162).Row  # xlUp = -4162
    return {"last_row": last_row, "column": column, "sheet": ws.Name}


def get_last_column(sheet: str | None = None, row: int = 1) -> dict:
    """Get the last used column number in a specific row.

    Args:
        sheet: Optional sheet name.
        row: Row number (default 1).

    Returns:
        dict with last column number, column letter, row, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    # Find last used column by going left from the right
    last_col = ws.Cells(row, ws.Columns.Count).End(-4159).Column  # xlToLeft = -4159
    col_letter = ws.Cells(1, last_col).Address.split("$")[1]
    return {"last_column": last_col, "column_letter": col_letter, "row": row, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Advanced Data Operations
# ---------------------------------------------------------------------------

def fill_series(
    sheet: str | None,
    start_cell: str,
    end_cell: str,
    fill_type: str = "linear",
    step: float = 1,
) -> dict:
    """Auto-fill a series of values.

    Args:
        sheet: Optional sheet name.
        start_cell: Starting cell with seed value.
        end_cell: Ending cell for the fill.
        fill_type: "linear", "growth", "date", or "auto".
        step: Step value for the series.

    Returns:
        dict with start, end, fill type, and sheet name.
    """
    fill_type_map = {
        "linear": 0,    # xlFillDefault -> use DataSeries with xlLinear
        "growth": 1,     # xlGrowth
        "date": 3,       # xlChronological
        "auto": -1,      # use AutoFill instead
    }
    app = _get_app()
    ws = _get_ws(app, sheet)
    start_rng = ws.Range(start_cell)
    end_rng = ws.Range(end_cell)
    fill_rng = ws.Range(start_cell, end_cell)

    if fill_type == "auto":
        # Use AutoFill
        start_rng.AutoFill(Destination=fill_rng)
    else:
        # Determine direction
        if start_rng.Column == end_rng.Column:
            # Vertical - fill down rows
            row_size = end_rng.Row - start_rng.Row + 1
            type_const = 0 if fill_type == "linear" else fill_type_map.get(fill_type, 0)
            fill_rng.DataSeries(Rowcol=1, Type=type_const, Step=step)  # 1=xlColumns
        else:
            # Horizontal - fill across columns
            type_const = 0 if fill_type == "linear" else fill_type_map.get(fill_type, 0)
            fill_rng.DataSeries(Rowcol=2, Type=type_const, Step=step)  # 2=xlRows

    return {"start": start_cell, "end": end_cell, "fill_type": fill_type, "sheet": ws.Name}


def concatenate_range(
    sheet: str | None,
    range_str: str,
    separator: str = ", ",
) -> dict:
    """Concatenate all values in a range into a single string.

    Args:
        sheet: Optional sheet name.
        range_str: Range address (e.g. "A1:A10").
        separator: Separator between values.

    Returns:
        dict with the concatenated result and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    values = rng.Value
    # Handle different return types from COM
    parts = []
    if values is None:
        pass
    elif isinstance(values, tuple):
        for row in values:
            if isinstance(row, tuple):
                for v in row:
                    if v is not None:
                        parts.append(str(v))
            else:
                if row is not None:
                    parts.append(str(row))
    else:
        parts.append(str(values))
    result = separator.join(parts)
    return {"result": result, "count": len(parts), "sheet": ws.Name}


def split_text_by_rows(
    sheet: str | None,
    cell: str,
    separator: str | None = None,
    target_cell: str | None = None,
) -> dict:
    """Split cell text into multiple rows.

    Args:
        sheet: Optional sheet name.
        cell: Source cell address.
        separator: Delimiter to split on. If None, splits by newline.
        target_cell: Starting cell for output. If None, uses cell below source.

    Returns:
        dict with parts count and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    value = ws.Range(cell).Value
    if value is None:
        return {"parts": 0, "sheet": ws.Name}
    text = str(value)
    sep = separator if separator is not None else "\n"
    parts = text.split(sep)
    # Determine output location
    if target_cell:
        out = ws.Range(target_cell)
    else:
        src = ws.Range(cell)
        out = ws.Cells(src.Row + 1, src.Column)
    for i, part in enumerate(parts):
        ws.Cells(out.Row + i, out.Column).Value = part.strip()
    return {"parts": len(parts), "target_start": out.Address, "sheet": ws.Name}


def apply_formula_to_range(
    sheet: str | None,
    range_str: str,
    formula_template: str,
) -> dict:
    """Apply a formula pattern to each cell in a range.

    The formula_template uses {row} and {col} as placeholders for the
    current cell's row number and column letter.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        formula_template: Formula with {row} and {col} placeholders.
                         Example: "=A{row}*B{row}"

    Returns:
        dict with range, formula template, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    for r in range(1, rng.Rows.Count + 1):
        for c in range(1, rng.Columns.Count + 1):
            cell = rng.Cells(r, c)
            row_num = cell.Row
            col_letter = ws.Cells(1, cell.Column).Address.split("$")[1]
            formula = formula_template.replace("{row}", str(row_num)).replace("{col}", col_letter)
            cell.Formula = formula
    return {"range": range_str, "formula_template": formula_template, "sheet": ws.Name}


def create_sequence(
    sheet: str | None,
    start_cell: str,
    count: int,
    start_value: float = 1,
    step: float = 1,
    direction: str = "down",
) -> dict:
    """Generate a number sequence in cells.

    Args:
        sheet: Optional sheet name.
        start_cell: Starting cell address.
        count: Number of values to generate.
        start_value: First value in the sequence.
        step: Increment between values.
        direction: "down" for vertical, "right" for horizontal.

    Returns:
        dict with start cell, count, direction, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    cell = ws.Range(start_cell)
    for i in range(count):
        value = start_value + i * step
        if direction == "right":
            ws.Cells(cell.Row, cell.Column + i).Value = value
        else:
            ws.Cells(cell.Row + i, cell.Column).Value = value
    return {
        "start_cell": start_cell,
        "count": count,
        "start_value": start_value,
        "step": step,
        "direction": direction,
        "sheet": ws.Name,
    }


# ---------------------------------------------------------------------------
# Conditional Operations
# ---------------------------------------------------------------------------

def highlight_cells(
    sheet: str | None,
    range_str: str,
    condition: str,
    color: list[int],
) -> dict:
    """Highlight cells meeting a condition.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        condition: Condition string: ">50", "<0", "=100", ">=10", "<=5",
                   "contains:text", "empty", "not_empty".
        color: RGB color list [r, g, b] for the fill.

    Returns:
        dict with range, condition, highlighted count, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    fill_color = rgb(color[0], color[1], color[2])
    count = 0
    for r in range(1, rng.Rows.Count + 1):
        for c in range(1, rng.Columns.Count + 1):
            cell = rng.Cells(r, c)
            val = cell.Value
            match = False
            if condition == "empty":
                match = val is None or str(val).strip() == ""
            elif condition == "not_empty":
                match = val is not None and str(val).strip() != ""
            elif condition.startswith("contains:"):
                search_text = condition[9:]
                match = val is not None and search_text in str(val)
            else:
                # Numeric comparison
                if val is not None:
                    try:
                        num_val = float(val)
                        if condition.startswith(">="):
                            match = num_val >= float(condition[2:])
                        elif condition.startswith("<="):
                            match = num_val <= float(condition[2:])
                        elif condition.startswith(">"):
                            match = num_val > float(condition[1:])
                        elif condition.startswith("<"):
                            match = num_val < float(condition[1:])
                        elif condition.startswith("="):
                            match = num_val == float(condition[1:])
                    except (ValueError, TypeError):
                        pass
            if match:
                cell.Interior.Color = fill_color
                count += 1
    return {"range": range_str, "condition": condition, "highlighted": count, "sheet": ws.Name}


def count_if(sheet: str | None, range_str: str, criteria: str) -> dict:
    """Count cells matching criteria (COUNTIF).

    Args:
        sheet: Optional sheet name.
        range_str: Range to count in.
        criteria: Criteria string (e.g. ">10", "Apple", "<>0").

    Returns:
        dict with count result, range, criteria, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng_ref = f"'{ws.Name}'!{range_str}"
    result = app.Evaluate(f'=COUNTIF({rng_ref},"{criteria}")')
    return {"count": result, "range": range_str, "criteria": criteria, "sheet": ws.Name}


def sum_if(
    sheet: str | None,
    range_str: str,
    criteria: str,
    sum_range: str | None = None,
) -> dict:
    """Sum cells matching criteria (SUMIF).

    Args:
        sheet: Optional sheet name.
        range_str: Range to evaluate criteria against.
        criteria: Criteria string (e.g. ">10", "Apple").
        sum_range: Optional range to sum. If None, sums the criteria range.

    Returns:
        dict with sum result, range, criteria, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng_ref = f"'{ws.Name}'!{range_str}"
    if sum_range:
        sum_ref = f"'{ws.Name}'!{sum_range}"
        result = app.Evaluate(f'=SUMIF({rng_ref},"{criteria}",{sum_ref})')
    else:
        result = app.Evaluate(f'=SUMIF({rng_ref},"{criteria}")')
    return {"sum": result, "range": range_str, "criteria": criteria, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Sheet Appearance
# ---------------------------------------------------------------------------

def set_gridlines_visible(sheet: str | None, visible: bool = True) -> dict:
    """Show or hide gridlines on a worksheet.

    Args:
        sheet: Optional sheet name.
        visible: True to show gridlines, False to hide.

    Returns:
        dict with visibility state and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Activate()
    app.ActiveWindow.DisplayGridlines = visible
    return {"visible": visible, "sheet": ws.Name}


def set_headings_visible(sheet: str | None, visible: bool = True) -> dict:
    """Show or hide row and column headings on a worksheet.

    Args:
        sheet: Optional sheet name.
        visible: True to show headings, False to hide.

    Returns:
        dict with visibility state and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Activate()
    app.ActiveWindow.DisplayHeadings = visible
    return {"visible": visible, "sheet": ws.Name}


def set_zoom_level(sheet: str | None, zoom_percent: int) -> dict:
    """Set the zoom level of a worksheet.

    Args:
        sheet: Optional sheet name.
        zoom_percent: Zoom percentage (10-400).

    Returns:
        dict with zoom level and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    ws.Activate()
    zoom = max(10, min(400, zoom_percent))
    app.ActiveWindow.Zoom = zoom
    return {"zoom": zoom, "sheet": ws.Name}


def set_sheet_direction(sheet: str | None, direction: str = "ltr") -> dict:
    """Set the sheet reading direction.

    Args:
        sheet: Optional sheet name.
        direction: "ltr" for left-to-right, "rtl" for right-to-left.

    Returns:
        dict with direction and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    # xlLTR = -5003, xlRTL = -5004
    ws.DisplayRightToLeft = (direction == "rtl")
    return {"direction": direction, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Validation Extended
# ---------------------------------------------------------------------------

def add_number_validation(
    sheet: str | None,
    range_str: str,
    min_value: float | None = None,
    max_value: float | None = None,
    input_message: str | None = None,
    error_message: str | None = None,
) -> dict:
    """Add number range validation to cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        min_value: Minimum allowed value.
        max_value: Maximum allowed value.
        input_message: Optional input prompt message.
        error_message: Optional error alert message.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    rng.Validation.Delete()
    kwargs = {"Type": XL_DV_DECIMAL}
    if min_value is not None and max_value is not None:
        kwargs["Operator"] = XL_CF_BETWEEN  # 1 = xlBetween
        kwargs["Formula1"] = str(min_value)
        kwargs["Formula2"] = str(max_value)
    elif min_value is not None:
        kwargs["Operator"] = XL_CF_GREATER_EQUAL  # 7
        kwargs["Formula1"] = str(min_value)
    elif max_value is not None:
        kwargs["Operator"] = XL_CF_LESS_EQUAL  # 8
        kwargs["Formula1"] = str(max_value)
    rng.Validation.Add(**kwargs)
    if input_message:
        rng.Validation.InputMessage = input_message
    if error_message:
        rng.Validation.ErrorMessage = error_message
    return {"range": range_str, "sheet": ws.Name}


def add_date_validation(
    sheet: str | None,
    range_str: str,
    min_date: str | None = None,
    max_date: str | None = None,
    input_message: str | None = None,
) -> dict:
    """Add date range validation to cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        min_date: Minimum date string (e.g. "2024-01-01").
        max_date: Maximum date string (e.g. "2024-12-31").
        input_message: Optional input prompt message.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    rng.Validation.Delete()
    kwargs = {"Type": XL_DV_DATE}
    if min_date is not None and max_date is not None:
        kwargs["Operator"] = XL_CF_BETWEEN
        kwargs["Formula1"] = min_date
        kwargs["Formula2"] = max_date
    elif min_date is not None:
        kwargs["Operator"] = XL_CF_GREATER_EQUAL
        kwargs["Formula1"] = min_date
    elif max_date is not None:
        kwargs["Operator"] = XL_CF_LESS_EQUAL
        kwargs["Formula1"] = max_date
    rng.Validation.Add(**kwargs)
    if input_message:
        rng.Validation.InputMessage = input_message
    return {"range": range_str, "sheet": ws.Name}


def add_text_length_validation(
    sheet: str | None,
    range_str: str,
    min_length: int | None = None,
    max_length: int | None = None,
    input_message: str | None = None,
) -> dict:
    """Add text length validation to cells.

    Args:
        sheet: Optional sheet name.
        range_str: Target range address.
        min_length: Minimum text length.
        max_length: Maximum text length.
        input_message: Optional input prompt message.

    Returns:
        dict with range and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(range_str)
    rng.Validation.Delete()
    kwargs = {"Type": XL_DV_TEXT_LENGTH}
    if min_length is not None and max_length is not None:
        kwargs["Operator"] = XL_CF_BETWEEN
        kwargs["Formula1"] = str(min_length)
        kwargs["Formula2"] = str(max_length)
    elif min_length is not None:
        kwargs["Operator"] = XL_CF_GREATER_EQUAL
        kwargs["Formula1"] = str(min_length)
    elif max_length is not None:
        kwargs["Operator"] = XL_CF_LESS_EQUAL
        kwargs["Formula1"] = str(max_length)
    rng.Validation.Add(**kwargs)
    if input_message:
        rng.Validation.InputMessage = input_message
    return {"range": range_str, "sheet": ws.Name}


# ---------------------------------------------------------------------------
# Error Handling & Audit
# ---------------------------------------------------------------------------

def trace_precedents(sheet: str | None, cell: str) -> dict:
    """Get cells that a formula references (precedents).

    Args:
        sheet: Optional sheet name.
        cell: Cell address containing a formula.

    Returns:
        dict with list of precedent cell addresses and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    formula = rng.Formula
    if not formula or not str(formula).startswith("="):
        return {"cell": cell, "precedents": [], "sheet": ws.Name}
    # Use NavigateArrow after ShowPrecedents
    rng.ShowPrecedents()
    precedents = []
    try:
        prec = rng.Precedents
        for i in range(1, prec.Count + 1):
            precedents.append(prec.Item(i).Address)
    except Exception:
        pass
    rng.Parent.ClearArrows()
    return {"cell": cell, "precedents": precedents, "sheet": ws.Name}


def trace_dependents(sheet: str | None, cell: str) -> dict:
    """Get cells that reference this cell (dependents).

    Args:
        sheet: Optional sheet name.
        cell: Cell address.

    Returns:
        dict with list of dependent cell addresses and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    rng.ShowDependents()
    dependents = []
    try:
        deps = rng.Dependents
        for i in range(1, deps.Count + 1):
            dependents.append(deps.Item(i).Address)
    except Exception:
        pass
    rng.Parent.ClearArrows()
    return {"cell": cell, "dependents": dependents, "sheet": ws.Name}


def check_errors(sheet: str | None, range_str: str | None = None) -> dict:
    """Find all error cells in a range or the entire sheet.

    Args:
        sheet: Optional sheet name.
        range_str: Optional range address. If None, checks entire used range.

    Returns:
        dict with list of error cells (address and error type) and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    if range_str:
        rng = ws.Range(range_str)
    else:
        rng = ws.UsedRange
    errors = []
    # xlCellTypeFormulas with error value
    try:
        error_cells = rng.SpecialCells(1, 16)  # xlCellTypeFormulas=1, xlErrors=16
        for i in range(1, error_cells.Areas.Count + 1):
            area = error_cells.Areas(i)
            for r in range(1, area.Rows.Count + 1):
                for c in range(1, area.Columns.Count + 1):
                    cell = area.Cells(r, c)
                    errors.append({
                        "cell": cell.Address,
                        "error": str(cell.Value),
                    })
    except Exception:
        # No error cells found (SpecialCells raises if none match)
        pass
    return {"errors": errors, "count": len(errors), "sheet": ws.Name}


def get_cell_formula(sheet: str | None, cell: str) -> dict:
    """Get the formula in a cell (not the calculated value).

    Args:
        sheet: Optional sheet name.
        cell: Cell address.

    Returns:
        dict with cell address, formula, has_formula flag, and sheet name.
    """
    app = _get_app()
    ws = _get_ws(app, sheet)
    rng = ws.Range(cell)
    formula = rng.Formula
    has_formula = bool(formula and str(formula).startswith("="))
    return {
        "cell": cell,
        "formula": str(formula) if formula else "",
        "has_formula": has_formula,
        "sheet": ws.Name,
    }
