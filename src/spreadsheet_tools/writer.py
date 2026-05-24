from __future__ import annotations

from typing import Any

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
        cell = sheet[normalized_address]

        if clear_value:
            cell.value = None
        elif value is not None:
            cell.value = value

        if style:
            apply_style_updates(cell, style)

        result = {
            "file": path,
            "sheet": sheet.title,
            "address": normalized_address,
            "updated": {
                "value": cell.value,
                "style_changed": bool(style),
            },
        }

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
