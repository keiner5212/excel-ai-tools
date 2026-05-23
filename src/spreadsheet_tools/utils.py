from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

COL_RE = re.compile(r"^[A-Z]+$")
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def resolve_path(path: str) -> Path:
    resolved = Path(path).expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"File not found: {resolved}")
    return resolved


def open_workbook(path: str, *, read_only: bool = False, data_only: bool = True) -> Workbook:
    return load_workbook(resolve_path(path), read_only=read_only, data_only=data_only, keep_vba=True)


def open_workbook_for_write(path: str) -> Workbook:
    return load_workbook(resolve_path(path), data_only=False, keep_vba=True)


def validate_column(column: str) -> str:
    normalized = column.strip().upper()
    if not COL_RE.match(normalized):
        raise ValueError(f"Invalid column letter: {column!r}")
    return normalized


def validate_cell_address(address: str) -> str:
    normalized = address.strip().upper()
    if not CELL_RE.match(normalized):
        raise ValueError(f"Invalid cell address: {address!r}")
    return normalized


def column_to_index(column: str) -> int:
    return column_index_from_string(validate_column(column))


def index_to_column(index: int) -> str:
    if index < 1:
        raise ValueError(f"Column index must be >= 1, got {index}")
    return get_column_letter(index)


def parse_cell_address(address: str) -> tuple[str, int]:
    match = CELL_RE.match(validate_cell_address(address))
    assert match is not None
    return match.group(1), int(match.group(2))


def get_sheet(workbook: Workbook, sheet_name: str | None) -> Worksheet:
    if sheet_name is None:
        return workbook.active
    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(f"Sheet {sheet_name!r} not found. Available: {available}")
    return workbook[sheet_name]


def normalize_scalar(value: object) -> object | None:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return value
