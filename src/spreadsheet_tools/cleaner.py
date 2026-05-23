from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from spreadsheet_tools.utils import index_to_column, normalize_scalar


def is_empty_value(value: object | None) -> bool:
    return normalize_scalar(value) is None


def row_has_content(values: list[object | None]) -> bool:
    return any(not is_empty_value(value) for value in values)


def trim_trailing_empty_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    trimmed = list(rows)
    while trimmed and not row_has_content([cell.get("value") for cell in trimmed[-1]["cells"]]):
        trimmed.pop()
    return trimmed


def trim_trailing_empty_columns(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    if not rows:
        return rows

    max_cols = max(len(row["cells"]) for row in rows)
    last_content_col = -1
    for col_idx in range(max_cols):
        if any(
            col_idx < len(row["cells"]) and not is_empty_value(row["cells"][col_idx].get("value"))
            for row in rows
        ):
            last_content_col = col_idx

    if last_content_col == -1:
        return []

    trimmed_rows: list[dict[str, Any]] = []
    for row in rows:
        cells = row["cells"][: last_content_col + 1]
        trimmed_rows.append({**row, "cells": cells})
    return trimmed_rows


def drop_fully_empty_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        row
        for row in rows
        if row_has_content([cell.get("value") for cell in row["cells"]])
    ]


def build_clean_row(
    worksheet: Worksheet,
    row_index: int,
    from_col_idx: int,
    to_col_idx: int,
) -> dict[str, Any]:
    cells: list[dict[str, Any]] = []
    for col_idx in range(from_col_idx, to_col_idx + 1):
        column = index_to_column(col_idx)
        address = f"{column}{row_index}"
        value = normalize_scalar(worksheet.cell(row=row_index, column=col_idx).value)
        cell_payload: dict[str, Any] = {
            "column": column,
            "address": address,
            "value": value,
        }
        cells.append(cell_payload)
    return {"row": row_index - 1, "cells": cells}
