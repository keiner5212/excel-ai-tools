"""MCP server for spreadsheet-tools.

Exposes every spreadsheet-tools command as a native MCP tool so AI agents can
call them directly — no shell scripting, no openpyxl imports, no helper files.

Run via:
    uv run spreadsheet-tools-mcp

Or from Claude Desktop / Cursor mcp.json:
    {
      "command": "uv",
      "args": ["run", "--directory", "/path/to/repo", "spreadsheet-tools-mcp"]
    }

IMPORTANT FOR AI AGENTS
-----------------------
- NEVER create Python scripts or import openpyxl yourself.
- NEVER call `uv run python` for spreadsheet operations.
- ALL Excel reads and writes go through these MCP tools ONLY.
- New workbooks: pass a bare filename (no directory) to `create_empty_workbook`.
  The workspace/ directory is created automatically.
- Row numbering: `from_row`/`to_row` are ZERO-BASED (row 0 = Excel row 1).
  `row` in `set_row_height` / `batch_set_dimensions` is 1-BASED (Excel row number).
- Columns: always use Excel letters (A, B, C...), never numeric indices.
- Merged cells: master cell (top-left) holds value. Slave cells return null.
  Use `sheet_info` to discover merged ranges before editing.
- Prefer `batch_edit` over repeated `edit_cell` calls (single save, safer).
- Prefer `batch_set_dimensions` over repeated width/height calls.
"""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from spreadsheet_tools.reader import (
    audit_range as _audit_range,
    describe_section as _describe_section,
    find_values as _find_values,
    list_sheets as _list_sheets,
    read_cell as _read_cell,
    read_range as _read_range,
    section_map as _section_map,
    sheet_info as _sheet_info,
    validate_rules as _validate_rules,
    workbook_info as _workbook_info,
)
from spreadsheet_tools.writer import (
    add_sheet as _add_sheet,
    batch_edit as _batch_edit,
    batch_set_dimensions as _batch_set_dimensions,
    copy_sheet_structure as _copy_sheet_structure,
    create_empty_workbook as _create_empty_workbook,
    edit_cell as _edit_cell,
    find_replace as _find_replace,
    format_range as _format_range,
    freeze_panes as _freeze_panes,
    get_cell_style_info as _get_cell_style_info,
    merge_cells as _merge_cells,
    rename_sheet as _rename_sheet,
    set_column_width as _set_column_width,
    set_row_height as _set_row_height,
    set_tab_color as _set_tab_color,
    toggle_auto_filter as _toggle_auto_filter,
    unfreeze_panes as _unfreeze_panes,
    unmerge_cells as _unmerge_cells,
)
from spreadsheet_tools.snapshot import (
    create_snapshot as _create_snapshot,
    diff_snapshots as _diff_snapshots,
    list_snapshots as _list_snapshots,
)

mcp = FastMCP(
    "spreadsheet-tools",
    instructions=(
        "Excel (.xlsx/.xlsm) read/write tools. "
        "NEVER use Python scripts or import openpyxl — use these tools exclusively. "
        "Pass any file path to operate on existing files. "
        "For new workbooks call create_empty_workbook; bare filenames auto-go to workspace/ "
        "but you may pass any absolute or relative path. "
        "Row indices for read_range/audit_range/section_map are ZERO-BASED (Excel row 1 = 0). "
        "Row numbers for set_row_height/batch_set_dimensions are 1-BASED (Excel row number). "
        "Columns are always Excel letters (A, B, ...), never numeric indices."
    ),
)


# ---------------------------------------------------------------------------
# Discovery / inspection
# ---------------------------------------------------------------------------


@mcp.tool()
def list_sheets(file: str) -> dict[str, Any]:
    """List all sheet names in an Excel workbook.

    Args:
        file: Path to .xlsx or .xlsm file.
    """
    return _list_sheets(file)


@mcp.tool()
def workbook_info(file: str) -> dict[str, Any]:
    """Return workbook-level metadata: sheet names, active sheet, defined names.

    Args:
        file: Path to .xlsx or .xlsm file.
    """
    return _workbook_info(file)


@mcp.tool()
def sheet_info(file: str, sheet: str | None = None) -> dict[str, Any]:
    """Return sheet structure: dimensions, all merged cell ranges, freeze panes, filters.

    Always call this before editing a sheet with merged cells.

    Args:
        file: Path to .xlsx or .xlsm file.
        sheet: Sheet name. Defaults to the active (first) sheet.
    """
    return _sheet_info(file, sheet)


# ---------------------------------------------------------------------------
# Reads
# ---------------------------------------------------------------------------


@mcp.tool()
def read_range(
    file: str,
    from_col: str,
    to_col: str,
    from_row: int,
    to_row: int,
    sheet: str | None = None,
    include_empty_rows: bool = False,
    keep_trailing_empty: bool = False,
    include_formulas: bool = False,
) -> dict[str, Any]:
    """Read a rectangular range of cells, returning values only (no formatting).

    Row indices are ZERO-BASED: Excel row 1 = from_row=0.
    Columns use Excel letters: A, B, C, ...

    Args:
        file: Path to .xlsx or .xlsm file.
        from_col: Start column letter (e.g. "A").
        to_col: End column letter (e.g. "L").
        from_row: Zero-based start row (Excel row 1 → 0).
        to_row: Zero-based end row (inclusive).
        sheet: Sheet name. Defaults to active sheet.
        include_empty_rows: Keep fully blank rows in output.
        keep_trailing_empty: Do not trim trailing empty rows/columns.
        include_formulas: Return formula text instead of computed values.
    """
    return _read_range(
        file,
        sheet_name=sheet,
        from_col=from_col,
        to_col=to_col,
        from_row=from_row,
        to_row=to_row,
        include_empty_rows=include_empty_rows,
        trim_empty=not keep_trailing_empty,
        include_formulas=include_formulas,
    )


@mcp.tool()
def read_cell(file: str, cell: str, sheet: str | None = None) -> dict[str, Any]:
    """Read a single cell value.

    Args:
        file: Path to .xlsx or .xlsm file.
        cell: Cell address in Excel notation (e.g. "B3").
        sheet: Sheet name. Defaults to active sheet.
    """
    return _read_cell(file, sheet_name=sheet, address=cell)


@mcp.tool()
def cell_style(file: str, cell: str, sheet: str | None = None) -> dict[str, Any]:
    """Read style metadata for a single cell: font, fill, alignment, border, number format.

    Use sparingly — heavier than value reads.

    Args:
        file: Path to .xlsx or .xlsm file.
        cell: Cell address (e.g. "B3").
        sheet: Sheet name. Defaults to active sheet.
    """
    return _get_cell_style_info(file, sheet_name=sheet, address=cell)


# ---------------------------------------------------------------------------
# Search
# ---------------------------------------------------------------------------


@mcp.tool()
def find(
    file: str,
    query: str,
    sheet: str | None = None,
    case_sensitive: bool = False,
    max_results: int = 50,
) -> dict[str, Any]:
    """Search for text across all cell values in the workbook (or a specific sheet).

    Args:
        file: Path to .xlsx or .xlsm file.
        query: Text to search for.
        sheet: Limit search to this sheet. Omit to search all sheets.
        case_sensitive: Match case exactly.
        max_results: Maximum number of results to return.
    """
    return _find_values(
        file,
        query=query,
        sheet_name=sheet,
        case_sensitive=case_sensitive,
        max_results=max_results,
    )


@mcp.tool()
def find_replace(
    file: str,
    query: str,
    replace_with: str | int | float | None = None,
    sheet: str | None = None,
    case_sensitive: bool = False,
    use_regex: bool = False,
    dry_run: bool = False,
    max_results: int = 200,
) -> dict[str, Any]:
    """Search cell values and optionally replace matches.

    Without replace_with: find-only mode, no file write.
    Without use_regex: entire cell value must equal query.
    With use_regex: substring match via re.sub.

    Args:
        file: Path to .xlsx or .xlsm file.
        query: Search string or regex pattern.
        replace_with: Replacement value. Omit for find-only.
        sheet: Limit to this sheet. Omit for all sheets.
        case_sensitive: Match case exactly.
        use_regex: Treat query as a regex pattern.
        dry_run: Preview matches without writing the file.
        max_results: Maximum matches to return.
    """
    return _find_replace(
        file,
        query=query,
        replace_with=replace_with,
        sheet_name=sheet,
        case_sensitive=case_sensitive,
        use_regex=use_regex,
        dry_run=dry_run,
        max_results=max_results,
    )


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------


@mcp.tool()
def section_map(
    file: str,
    sheet: str | None = None,
    min_row: int = 0,
    max_row: int | None = None,
) -> dict[str, Any]:
    """Auto-discover numbered section headers (e.g. "1.2.3 Title") and their row ranges.

    Returns header text, prefix, title, depth, header_row, and row_range for each section.
    Use before bulk-editing to understand the sheet layout.

    Args:
        file: Path to .xlsx or .xlsm file.
        sheet: Sheet name. Defaults to active sheet.
        min_row: Zero-based start row for scan.
        max_row: Zero-based end row for scan (default: sheet max).
    """
    return _section_map(file, sheet_name=sheet, min_row=min_row, max_row=max_row)


@mcp.tool()
def audit_range(
    file: str,
    from_col: str,
    to_col: str,
    from_row: int,
    to_row: int,
    sheet: str | None = None,
    show_slaves: bool = False,
) -> dict[str, Any]:
    """List all master cells in a range with their is_empty flag.

    Slave cells (non-master side of a merge) are skipped by default.
    Returns per-cell status and summary counts (total_master, empty_master, filled_master).

    Row indices are ZERO-BASED.

    Args:
        file: Path to .xlsx or .xlsm file.
        from_col: Start column letter.
        to_col: End column letter.
        from_row: Zero-based start row.
        to_row: Zero-based end row (inclusive).
        sheet: Sheet name. Defaults to active sheet.
        show_slaves: Include slave (non-writable) cells in output.
    """
    return _audit_range(
        file,
        sheet_name=sheet,
        from_col=from_col,
        to_col=to_col,
        from_row=from_row,
        to_row=to_row,
        show_slaves=show_slaves,
    )


@mcp.tool()
def describe_section(
    file: str,
    data_rows: str,
    sheet: str | None = None,
    name_col: str = "B",
    desc_col: str = "D",
    cost_col: str = "L",
    header_row: int | None = None,
) -> dict[str, Any]:
    """Audit a strategy table for name/description consistency issues.

    Issues detected per row:
    - generic_name (warning): name matches "Estrategia de X N" pattern
    - missing_description (error): name present but description empty
    - name_desc_mismatch (error): keyword overlap < 30%
    - missing_cost (warning): cost cell is empty

    Args:
        file: Path to .xlsx or .xlsm file.
        data_rows: Zero-based row range as "START-END", e.g. "325-328".
        sheet: Sheet name. Defaults to active sheet.
        name_col: Column letter with strategy names.
        desc_col: Column letter with descriptions.
        cost_col: Column letter with annual costs. Pass "" to skip cost check.
        header_row: Optional zero-based header row index.
    """
    start_str, end_str = data_rows.split("-", 1)
    return _describe_section(
        file,
        sheet_name=sheet,
        from_data_row=int(start_str),
        to_data_row=int(end_str),
        name_col=name_col,
        desc_col=desc_col,
        cost_col=cost_col if cost_col.strip() else None,
        header_row=header_row,
    )


@mcp.tool()
def validate(
    file: str,
    rules: list[str],
    sheet: str | None = None,
) -> dict[str, Any]:
    """Apply one or more validation rules against a sheet.

    Rule syntax examples:
    - "not-empty:D325"
    - "not-empty-range:D325:D328"
    - "price-min-max:F308:H308"
    - "name-matches-desc:B325:D325"
    - "numeric-range:L317:50000:2000000"
    - "string-contains:D317:Descuento"
    - "no-generic-name:B325"

    Args:
        file: Path to .xlsx or .xlsm file.
        rules: List of rule strings in "rule-name:ARGS" format.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _validate_rules(file, sheet_name=sheet, rules=rules)


# ---------------------------------------------------------------------------
# Writes — single cell
# ---------------------------------------------------------------------------


@mcp.tool()
def edit_cell(
    file: str,
    cell: str,
    value: str | int | float | bool | None = None,
    clear: bool = False,
    style: dict[str, Any] | None = None,
    sheet: str | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Edit a single cell's value and/or style.

    At least one of value, clear, or style must be provided.

    Style keys (all optional, partial updates supported):
    {
      "font": {"name": "Arial", "size": 12, "bold": true, "italic": false, "color": "FF0000"},
      "fill": {"fill_type": "solid", "start_color": "FFFF00"},
      "alignment": {"horizontal": "center", "vertical": "center", "wrap_text": true},
      "number_format": "0.00",
      "comment": "Review this",
      "hyperlink": "https://example.com"
    }
    Colors are 6-char hex RGB without # (e.g. "FF0000" for red).

    Args:
        file: Path to .xlsx or .xlsm file.
        cell: Cell address (e.g. "B3").
        value: New cell value. Omit to keep existing value.
        clear: Set true to clear the cell value.
        style: Style overrides dict. Omit to keep existing style.
        sheet: Sheet name. Defaults to active sheet.
        dry_run: Apply in memory without saving.
    """
    return _edit_cell(
        file,
        sheet_name=sheet,
        address=cell,
        value=value,
        clear_value=clear,
        style=style,
        save=not dry_run,
    )


# ---------------------------------------------------------------------------
# Writes — bulk
# ---------------------------------------------------------------------------


@mcp.tool()
def batch_edit(
    file: str,
    edits: list[dict[str, Any]],
    sheet: str | None = None,
    dry_run: bool = False,
) -> dict[str, Any]:
    """Apply multiple cell edits atomically in a single save. Preferred over repeated edit_cell.

    All addresses are validated upfront — if any fail, the file is never modified.
    Slave cells auto-redirect to their master cell (with a warning).

    Each edit dict format:
    {
      "cell": "D325",           # required
      "value": "Description",  # optional: new value
      "clear": true,           # optional: clear the cell
      "style": {"alignment": {"wrap_text": true}}  # optional: style overrides
    }

    Args:
        file: Path to .xlsx or .xlsm file.
        edits: List of edit dicts (see format above).
        sheet: Sheet name. Defaults to active sheet.
        dry_run: Validate and preview without saving.
    """
    return _batch_edit(file, sheet_name=sheet, edits=edits, dry_run=dry_run)


# ---------------------------------------------------------------------------
# Structure / formatting
# ---------------------------------------------------------------------------


@mcp.tool()
def merge_cells(
    file: str,
    cell_range: str,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Merge a rectangular cell range. Master cell (top-left) keeps value and style.

    Args:
        file: Path to .xlsx or .xlsm file.
        cell_range: Range to merge, e.g. "A1:E1".
        sheet: Sheet name. Defaults to active sheet.
    """
    return _merge_cells(file, sheet_name=sheet, cell_range=cell_range)


@mcp.tool()
def unmerge_cells(
    file: str,
    cell_range: str,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Remove merge from a cell range. Cells become independent again.

    Args:
        file: Path to .xlsx or .xlsm file.
        cell_range: Range to unmerge, e.g. "A1:E1".
        sheet: Sheet name. Defaults to active sheet.
    """
    return _unmerge_cells(file, sheet_name=sheet, cell_range=cell_range)


@mcp.tool()
def set_column_width(
    file: str,
    col: str,
    width: float,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Set width of a single column in Excel character units.

    For 3+ columns, prefer batch_set_dimensions (single save).

    Args:
        file: Path to .xlsx or .xlsm file.
        col: Column letter (e.g. "A").
        width: Width in Excel character units.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _set_column_width(file, sheet_name=sheet, col=col, width=width)


@mcp.tool()
def set_row_height(
    file: str,
    row: int,
    height: float,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Set height of a single row in points. Row number is 1-BASED (Excel convention).

    For 3+ rows, prefer batch_set_dimensions (single save).

    Args:
        file: Path to .xlsx or .xlsm file.
        row: Row number, 1-based (Excel row 1 = row=1).
        height: Height in points.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _set_row_height(file, sheet_name=sheet, row=row, height=height)


@mcp.tool()
def batch_set_dimensions(
    file: str,
    column_widths: list[dict[str, Any]] | None = None,
    row_heights: list[dict[str, Any]] | None = None,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Set multiple column widths and/or row heights in a single save.

    Preferred over repeated set_column_width / set_row_height calls.

    column_widths format: [{"col": "A", "width": 8}, {"col": "B", "width": 20}]
    row_heights format:   [{"row": 1, "height": 40}, {"row": 8, "height": 25}]
    Row numbers in row_heights are 1-BASED (Excel convention).

    Args:
        file: Path to .xlsx or .xlsm file.
        column_widths: List of column width dicts. At least one of column_widths/row_heights required.
        row_heights: List of row height dicts. Row numbers are 1-based.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _batch_set_dimensions(
        file,
        sheet_name=sheet,
        column_widths=column_widths,
        row_heights=row_heights,
    )


@mcp.tool()
def format_range(
    file: str,
    cell_range: str,
    style: dict[str, Any],
    sheet: str | None = None,
) -> dict[str, Any]:
    """Apply the same style to every master cell in a rectangular range.

    Slave cells are skipped automatically. Partial style updates supported —
    omitted keys leave existing style unchanged.

    Style keys (see edit_cell for full schema):
    {"font": {"bold": true, "color": "FFFFFF"}, "fill": {"fill_type": "solid", "start_color": "1F4E79"}}

    Args:
        file: Path to .xlsx or .xlsm file.
        cell_range: Range to format, e.g. "A8:E8".
        style: Style overrides dict.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _format_range(file, sheet_name=sheet, cell_range=cell_range, style=style)


@mcp.tool()
def freeze_panes(
    file: str,
    cell: str,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Freeze rows above and columns to the left of a given cell.

    Examples:
    - cell="B2" → freeze row 1 and column A (most common for tables)
    - cell="A2" → freeze only row 1
    - cell="B1" → freeze only column A

    Args:
        file: Path to .xlsx or .xlsm file.
        cell: Freeze point cell address (e.g. "B2").
        sheet: Sheet name. Defaults to active sheet.
    """
    return _freeze_panes(file, sheet_name=sheet, cell=cell)


@mcp.tool()
def unfreeze_panes(file: str, sheet: str | None = None) -> dict[str, Any]:
    """Remove all freeze panes from a sheet.

    Args:
        file: Path to .xlsx or .xlsm file.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _unfreeze_panes(file, sheet_name=sheet)


@mcp.tool()
def set_tab_color(
    file: str,
    color: str,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Set the sheet tab color.

    Args:
        file: Path to .xlsx or .xlsm file.
        color: 6-char RGB hex string without # (e.g. "1F4E79" for dark blue, "70AD47" for green).
        sheet: Sheet name. Defaults to active sheet.
    """
    return _set_tab_color(file, sheet_name=sheet, color=color)


@mcp.tool()
def auto_filter(
    file: str,
    cell_range: str | None = None,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Enable auto-filter on a header range, or clear it when cell_range is omitted.

    Args:
        file: Path to .xlsx or .xlsm file.
        cell_range: Header row range for auto-filter (e.g. "A8:E8"). Omit to clear.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _toggle_auto_filter(file, sheet_name=sheet, cell_range=cell_range)


# ---------------------------------------------------------------------------
# Sheet management
# ---------------------------------------------------------------------------


@mcp.tool()
def add_sheet(
    file: str,
    sheet: str,
    position: int | None = None,
) -> dict[str, Any]:
    """Add a new empty sheet to an existing workbook.

    Args:
        file: Path to .xlsx or .xlsm file.
        sheet: Name for the new sheet.
        position: Zero-based insertion index. Omit to append at end.
    """
    return _add_sheet(file, sheet_name=sheet, position=position)


@mcp.tool()
def rename_sheet(
    file: str,
    old_name: str,
    new_name: str,
) -> dict[str, Any]:
    """Rename a sheet. Fails if old name not found or new name already taken.

    Args:
        file: Path to .xlsx or .xlsm file.
        old_name: Current sheet name.
        new_name: New sheet name.
    """
    return _rename_sheet(file, old_name=old_name, new_name=new_name)


@mcp.tool()
def copy_sheet(
    file: str,
    source_sheet: str,
    target_sheet: str,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Duplicate a sheet within the same workbook.

    Args:
        file: Path to .xlsx or .xlsm file.
        source_sheet: Name of the sheet to copy.
        target_sheet: Name for the new copy.
        overwrite: Overwrite target_sheet if it already exists.
    """
    return _copy_sheet_structure(
        file,
        source_sheet=source_sheet,
        target_sheet=target_sheet,
        overwrite=overwrite,
    )


# ---------------------------------------------------------------------------
# Workbook creation
# ---------------------------------------------------------------------------


@mcp.tool()
def create_empty_workbook(filename: str, sheet: str = "Sheet1") -> dict[str, Any]:
    """Create a new empty Excel workbook.

    Pass any file path. Bare filenames (no directory separator) are automatically
    placed in workspace/ as a convenience convention, but any path works:
        create_empty_workbook("report.xlsx")           → workspace/report.xlsx
        create_empty_workbook("/tmp/report.xlsx")      → /tmp/report.xlsx
        create_empty_workbook("output/report.xlsx")    → output/report.xlsx (relative to cwd)

    Args:
        filename: Output path. Bare name → workspace/ automatically.
        sheet: Name of the first sheet.
    """
    return _create_empty_workbook(filename, sheet_name=sheet)


# ---------------------------------------------------------------------------
# Snapshots
# ---------------------------------------------------------------------------


@mcp.tool()
def snapshot(
    file: str,
    tag: str,
    sheet: str | None = None,
    description: str | None = None,
) -> dict[str, Any]:
    """Capture all non-empty cell values as a named snapshot for later comparison.

    Take a snapshot with tag="before" before any destructive batch operation.

    Args:
        file: Path to .xlsx or .xlsm file.
        tag: Snapshot name (e.g. "before-edit", "after-fill").
        sheet: Sheet name. Defaults to active sheet.
        description: Optional human-readable note stored with the snapshot.
    """
    return _create_snapshot(file, sheet_name=sheet, tag=tag, description=description)


@mcp.tool()
def snapshot_diff(
    file: str,
    tag_a: str,
    tag_b: str,
    sheet: str | None = None,
) -> dict[str, Any]:
    """Compare two snapshots and list per-cell changes.

    Args:
        file: Path to .xlsx or .xlsm file.
        tag_a: First snapshot tag (e.g. "before-edit").
        tag_b: Second snapshot tag (e.g. "after-edit").
        sheet: Sheet name. Defaults to active sheet.
    """
    return _diff_snapshots(file, sheet_name=sheet, tag_a=tag_a, tag_b=tag_b)


@mcp.tool()
def list_snapshots(file: str, sheet: str | None = None) -> dict[str, Any]:
    """List all available snapshots for a workbook/sheet.

    Args:
        file: Path to .xlsx or .xlsm file.
        sheet: Sheet name. Defaults to active sheet.
    """
    return _list_snapshots(file, sheet_name=sheet)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
