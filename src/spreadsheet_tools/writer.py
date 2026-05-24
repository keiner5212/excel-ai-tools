from __future__ import annotations

import re
from typing import Any

from openpyxl import Workbook

from spreadsheet_tools.cleaner import build_merge_lookup
from spreadsheet_tools.styles import apply_style_updates, get_cell_style
from spreadsheet_tools.utils import (
    get_sheet,
    open_workbook,
    open_workbook_for_write,
    safe_save_workbook,
    validate_cell_address,
)


def _serialize_value(value: object | None) -> object | None:
    from datetime import date, datetime, time

    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.isoformat()
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return value


def _resolve_write_address(sheet: Any, address: str) -> tuple[str, str | None]:
    """Return (effective_address, warning).

    If address is a slave cell of a merge, redirect to master and warn.
    Writing to slave cells is silently ignored by Excel; only master shows data.
    """
    merge_lookup = build_merge_lookup(sheet)
    if address not in merge_lookup:
        return address, None
    range_str, master = merge_lookup[address]
    if address == master:
        return address, None
    warning = (
        f"Address {address} is a slave cell of merged range {range_str}. "
        f"Redirected to master cell {master}."
    )
    return master, warning


def edit_cell(
    path: str,
    *,
    sheet_name: str | None,
    address: str,
    value: object | None = None,
    clear_value: bool = False,
    style: dict[str, Any] | None = None,
    save: bool = True,
) -> dict[str, Any]:
    normalized_address = validate_cell_address(address)
    workbook = open_workbook_for_write(path)
    try:
        sheet = get_sheet(workbook, sheet_name)
        effective_address, warning = _resolve_write_address(sheet, normalized_address)
        cell = sheet[effective_address]

        if clear_value:
            cell.value = None
        elif value is not None:
            cell.value = value

        if style:
            apply_style_updates(cell, style)

        result: dict[str, Any] = {
            "file": path,
            "sheet": sheet.title,
            "address": effective_address,
            "updated": {
                "value": cell.value,
                "style_changed": bool(style),
            },
        }

        if warning:
            result["warning"] = warning

        if save:
            safe_save_workbook(workbook, path)
            result["saved"] = True
        else:
            result["saved"] = False

        return result
    finally:
        workbook.close()


def get_cell_style_info(
    path: str, *, sheet_name: str | None, address: str
) -> dict[str, Any]:
    normalized_address = validate_cell_address(address)
    # data_only=False needed to read number formats and formula-based cells correctly.
    # read_only=False required because read-only mode doesn't expose full cell style objects.
    workbook = open_workbook(path, read_only=False, data_only=False)
    try:
        sheet = get_sheet(workbook, sheet_name)
        style = get_cell_style(sheet, normalized_address)
        if "value" in style:
            style["value"] = _serialize_value(style["value"])
        return {"file": path, "sheet": sheet.title, **style}
    finally:
        workbook.close()


def copy_sheet_structure(
    path: str,
    *,
    source_sheet: str,
    target_sheet: str,
    overwrite: bool = False,
) -> dict[str, Any]:
    workbook = open_workbook_for_write(path)
    try:
        if source_sheet not in workbook.sheetnames:
            raise ValueError(f"Source sheet {source_sheet!r} not found")

        if target_sheet in workbook.sheetnames:
            if not overwrite:
                raise ValueError(f"Target sheet {target_sheet!r} already exists")
            del workbook[target_sheet]

        source = workbook[source_sheet]
        target = workbook.copy_worksheet(source)
        target.title = target_sheet
        safe_save_workbook(workbook, path)
        return {
            "file": path,
            "source_sheet": source_sheet,
            "target_sheet": target_sheet,
            "saved": True,
        }
    finally:
        workbook.close()


def batch_edit(
    path: str,
    *,
    sheet_name: str | None = None,
    edits: list[dict[str, Any]],
    dry_run: bool = False,
) -> dict[str, Any]:
    """Apply multiple cell edits atomically in one workbook load/save cycle.

    Each edit dict must contain ``cell`` (address string) and at least one of:
    ``value``, ``clear`` (bool), or ``style`` (dict).

    All addresses are validated before any edit is applied.  If validation
    fails the workbook is never opened for writing.
    """
    if not edits:
        raise ValueError("edits list is empty")

    # Validate all addresses upfront so we fail early before touching the file
    for i, edit in enumerate(edits):
        if "cell" not in edit:
            raise ValueError(f"edits[{i}] missing required 'cell' field")
        validate_cell_address(edit["cell"])

    workbook = open_workbook_for_write(path)
    try:
        sheet = get_sheet(workbook, sheet_name)
        merge_lookup = build_merge_lookup(sheet)
        applied: list[dict[str, Any]] = []

        for edit in edits:
            address = validate_cell_address(edit["cell"])
            effective_addr = address
            warning: str | None = None

            if address in merge_lookup:
                range_str, master = merge_lookup[address]
                if address != master:
                    warning = (
                        f"Address {address} is a slave of merged range {range_str}. "
                        f"Redirected to master cell {master}."
                    )
                    effective_addr = master

            cell = sheet[effective_addr]
            value = edit.get("value")
            clear = edit.get("clear", False)
            style = edit.get("style")

            if clear:
                cell.value = None
            elif value is not None:
                cell.value = value

            if style:
                apply_style_updates(cell, style)

            entry: dict[str, Any] = {
                "cell": edit["cell"],
                "effective_address": effective_addr,
                "new_value": _serialize_value(cell.value),
                "style_changed": bool(style),
            }
            if warning:
                entry["warning"] = warning
            applied.append(entry)

        result: dict[str, Any] = {
            "file": path,
            "sheet": sheet.title,
            "applied": len(applied),
            "dry_run": dry_run,
            "cells": applied,
        }

        if not dry_run:
            safe_save_workbook(workbook, path)
            result["saved"] = True
        else:
            result["saved"] = False

        return result
    finally:
        workbook.close()


def create_empty_workbook(path: str, sheet_name: str = "Sheet1") -> dict[str, Any]:
    """Create new empty workbook with one sheet, overwriting existing file."""
    workbook = Workbook()
    ws = workbook.active
    ws.title = sheet_name
    workbook.save(path)
    return {"file": path, "sheet": sheet_name, "created": True}


def find_replace(
    path: str,
    *,
    query: str,
    replace_with: Any = None,
    sheet_name: str | None = None,
    case_sensitive: bool = False,
    use_regex: bool = False,
    dry_run: bool = False,
    max_results: int = 200,
) -> dict[str, Any]:
    """Search cell values and optionally replace matches.

    Default matching is exact-value: the entire string representation of the
    cell value must equal the query.  Pass ``use_regex=True`` for
    substring/pattern matching (``re.search``); replacement is then applied
    via ``re.sub`` on the string representation.

    ``replace_with`` accepts any Python value (int, float, str).  Numeric
    values are stored as numbers in Excel, not text.  When ``use_regex=True``
    the replacement is forced to str (required by ``re.sub``).

    Two-pass approach: read pass collects matches with data_only=True (computed
    values); write pass applies replacements using a separate workbook load.
    """
    from spreadsheet_tools.reader import _serialize_value

    flags = 0 if case_sensitive else re.IGNORECASE
    pattern: re.Pattern[str] | None = re.compile(query, flags) if use_regex else None
    needle = query if case_sensitive else query.lower()

    # --- Pass 1: find matches (read-only, computed values) ---
    matches: list[dict[str, Any]] = []
    read_wb = open_workbook(path, read_only=False, data_only=True)
    try:
        target_sheets = (
            [get_sheet(read_wb, sheet_name)]
            if sheet_name
            else [read_wb[name] for name in read_wb.sheetnames]
        )
        for sheet in target_sheets:
            if len(matches) >= max_results:
                break
            for row in sheet.iter_rows():
                if len(matches) >= max_results:
                    break
                for cell in row:
                    if cell.value is None:
                        continue
                    val_str = str(cell.value)
                    if use_regex:
                        matched = bool(pattern.search(val_str))  # type: ignore[union-attr]
                    else:
                        compare = val_str if case_sensitive else val_str.lower()
                        matched = needle == compare
                    if matched:
                        matches.append(
                            {
                                "sheet": sheet.title,
                                "address": cell.coordinate,
                                "value": _serialize_value(cell.value),
                            }
                        )
    finally:
        read_wb.close()

    result: dict[str, Any] = {
        "file": path,
        "query": query,
        "total": len(matches),
        "truncated": len(matches) >= max_results,
        "matches": matches,
    }

    if replace_with is None:
        return result

    result["replace_with"] = replace_with

    # Compute new values for each match
    replacements: list[dict[str, Any]] = []
    for m in matches:
        if use_regex:
            # re.sub requires string operands; result is always str
            new_val: Any = re.sub(
                query, str(replace_with), str(m["value"]), flags=flags
            )
        else:
            # Preserve the caller's type (int, float, str, …)
            new_val = replace_with
        replacements.append(
            {
                "sheet": m["sheet"],
                "address": m["address"],
                "before": m["value"],
                "after": new_val,
            }
        )

    if dry_run:
        result["dry_run"] = True
        result["replacements_preview"] = replacements
        result["saved"] = False
        return result

    # --- Pass 2: apply replacements ---
    write_wb = open_workbook_for_write(path)
    try:
        for rep in replacements:
            ws = get_sheet(write_wb, rep["sheet"])
            ws[rep["address"]].value = rep["after"]  # type: ignore[union-attr]
        safe_save_workbook(write_wb, path)
    finally:
        write_wb.close()

    result["dry_run"] = False
    result["replacements"] = replacements
    result["saved"] = True
    return result
