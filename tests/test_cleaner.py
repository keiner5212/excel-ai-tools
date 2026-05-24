from __future__ import annotations

from spreadsheet_tools.cleaner import (
    drop_fully_empty_rows,
    trim_trailing_empty_columns,
    trim_trailing_empty_rows,
)


def _row(values: list[object | None]) -> dict:
    return {
        "row": 0,
        "cells": [
            {"column": "A", "address": f"A{i + 1}", "value": v}
            for i, v in enumerate(values)
        ],
    }


def test_trim_trailing_empty_rows() -> None:
    rows = [_row(["a"]), _row([None]), _row([None])]
    trimmed = trim_trailing_empty_rows(rows)
    assert len(trimmed) == 1


def test_trim_trailing_empty_columns() -> None:
    rows = [_row(["a", None, None]), _row(["b", None, None])]
    trimmed = trim_trailing_empty_columns(rows)
    assert len(trimmed[0]["cells"]) == 1


def test_drop_fully_empty_rows() -> None:
    rows = [_row([None]), _row(["x"])]
    kept = drop_fully_empty_rows(rows)
    assert len(kept) == 1
    assert kept[0]["cells"][0]["value"] == "x"
