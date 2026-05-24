from __future__ import annotations

from datetime import date, datetime, time
from typing import Any

from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from spreadsheet_tools.cleaner import (
    build_clean_row,
    build_merge_lookup,
    drop_fully_empty_rows,
    trim_trailing_empty_columns,
    trim_trailing_empty_rows,
)
from spreadsheet_tools.utils import (
    column_to_index,
    get_sheet,
    index_to_column,
    normalize_scalar,
    open_workbook,
)


def _serialize_value(value: object | None) -> object | None:
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.isoformat()
    return normalize_scalar(value)


def list_sheets(path: str) -> dict[str, Any]:
    workbook = open_workbook(path, read_only=True)
    try:
        active = workbook.active.title if workbook.active else None
        sheets = []
        for name in workbook.sheetnames:
            sheet = workbook[name]
            sheets.append(
                {
                    "name": name,
                    "index": workbook.sheetnames.index(name),
                    "state": sheet.sheet_state,
                }
            )
        return {
            "file": path,
            "active_sheet": active,
            "sheet_count": len(sheets),
            "sheets": sheets,
        }
    finally:
        workbook.close()


def workbook_info(path: str) -> dict[str, Any]:
    workbook = open_workbook(path, read_only=True)
    try:
        properties = workbook.properties
        return {
            "file": path,
            "format": "xlsm" if path.lower().endswith(".xlsm") else "xlsx",
            "sheet_count": len(workbook.sheetnames),
            "sheet_names": list(workbook.sheetnames),
            "active_sheet": workbook.active.title if workbook.active else None,
            "properties": {
                "title": properties.title,
                "creator": properties.creator,
                "created": properties.created.isoformat()
                if properties.created
                else None,
                "modified": properties.modified.isoformat()
                if properties.modified
                else None,
                "subject": properties.subject,
                "description": properties.description,
            },
        }
    finally:
        workbook.close()


def sheet_info(path: str, sheet_name: str | None = None) -> dict[str, Any]:
    workbook = open_workbook(path, read_only=False)
    try:
        sheet = get_sheet(workbook, sheet_name)
        merged = [str(item) for item in sheet.merged_cells.ranges]
        return {
            "file": path,
            "sheet": sheet.title,
            "dimensions": sheet.dimensions,
            "max_row": sheet.max_row,
            "max_column": sheet.max_column,
            "max_column_letter": get_column_letter(sheet.max_column)
            if sheet.max_column
            else None,
            "merged_ranges": merged,
            "freeze_panes": sheet.freeze_panes,
            "auto_filter": sheet.auto_filter.ref if sheet.auto_filter else None,
        }
    finally:
        workbook.close()


def read_range(
    path: str,
    *,
    sheet_name: str | None = None,
    from_col: str,
    to_col: str,
    from_row: int,
    to_row: int,
    include_empty_rows: bool = False,
    trim_empty: bool = True,
    include_formulas: bool = False,
) -> dict[str, Any]:
    if from_row < 0 or to_row < 0:
        raise ValueError("Row indices are zero-based and must be >= 0")
    if from_row > to_row:
        raise ValueError("from_row must be <= to_row")

    from_col_idx = column_to_index(from_col)
    to_col_idx = column_to_index(to_col)
    if from_col_idx > to_col_idx:
        raise ValueError("from_col must be <= to_col")

    data_only = not include_formulas
    workbook = open_workbook(path, read_only=False, data_only=data_only)
    try:
        sheet = get_sheet(workbook, sheet_name)
        merge_lookup = build_merge_lookup(sheet)
        rows: list[dict[str, Any]] = []
        for excel_row in range(from_row + 1, to_row + 2):
            row_payload = build_clean_row(
                sheet, excel_row, from_col_idx, to_col_idx, merge_lookup
            )
            for cell in row_payload["cells"]:
                cell["value"] = _serialize_value(cell["value"])
            rows.append(row_payload)

        if trim_empty:
            if not include_empty_rows:
                rows = drop_fully_empty_rows(rows)
            rows = trim_trailing_empty_rows(rows)
            rows = trim_trailing_empty_columns(rows)

        merged_in_range = _merged_ranges_in_area(
            sheet,
            from_col_idx,
            to_col_idx,
            from_row + 1,
            to_row + 1,
        )

        return {
            "file": path,
            "sheet": sheet.title,
            "range": {
                "from_col": index_to_column(from_col_idx),
                "to_col": index_to_column(to_col_idx),
                "from_row": from_row,
                "to_row": to_row,
            },
            "row_count": len(rows),
            "rows": rows,
            "merged_ranges": merged_in_range,
            "notes": [
                "Values only; formatting metadata excluded.",
                "Rows are zero-based in output.",
                "Empty rows/columns trimmed unless include_empty_rows is enabled.",
                "Merged cells: 'master' is the top-left cell of the range (write target). 'slave' cells are non-writable; their value is always null.",
            ],
        }
    finally:
        workbook.close()


def _merged_ranges_in_area(
    sheet: Worksheet,
    from_col_idx: int,
    to_col_idx: int,
    from_row: int,
    to_row: int,
) -> list[str]:
    merged: list[str] = []
    for item in sheet.merged_cells.ranges:
        if (
            item.min_col <= to_col_idx
            and item.max_col >= from_col_idx
            and item.min_row <= to_row
            and item.max_row >= from_row
        ):
            merged.append(str(item))
    return merged


def find_values(
    path: str,
    *,
    query: str,
    sheet_name: str | None = None,
    case_sensitive: bool = False,
    max_results: int = 50,
) -> dict[str, Any]:
    workbook = open_workbook(path, read_only=True)
    try:
        sheets = (
            [get_sheet(workbook, sheet_name)]
            if sheet_name
            else [workbook[name] for name in workbook.sheetnames]
        )
        needle = query if case_sensitive else query.casefold()
        matches: list[dict[str, Any]] = []

        for sheet in sheets:
            for row in sheet.iter_rows():
                for cell in row:
                    value = _serialize_value(cell.value)
                    if value is None:
                        continue
                    haystack = str(value)
                    compare = haystack if case_sensitive else haystack.casefold()
                    if needle in compare:
                        matches.append(
                            {
                                "sheet": sheet.title,
                                "address": cell.coordinate,
                                "value": value,
                            }
                        )
                        if len(matches) >= max_results:
                            return {
                                "file": path,
                                "query": query,
                                "match_count": len(matches),
                                "truncated": True,
                                "matches": matches,
                            }

        return {
            "file": path,
            "query": query,
            "match_count": len(matches),
            "truncated": False,
            "matches": matches,
        }
    finally:
        workbook.close()


def read_cell(path: str, *, sheet_name: str | None, address: str) -> dict[str, Any]:
    workbook = open_workbook(path, read_only=True)
    try:
        sheet = get_sheet(workbook, sheet_name)
        cell = sheet[address.upper()]
        return {
            "file": path,
            "sheet": sheet.title,
            "address": cell.coordinate,
            "value": _serialize_value(cell.value),
        }
    finally:
        workbook.close()
