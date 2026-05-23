from __future__ import annotations

import argparse
import json
import sys
from typing import Any

from spreadsheet_tools.reader import (
    find_values,
    list_sheets,
    read_cell,
    read_range,
    sheet_info,
    workbook_info,
)
from spreadsheet_tools.writer import copy_sheet_structure, edit_cell, get_cell_style_info


def _print_json(payload: dict[str, Any]) -> None:
    json.dump(payload, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")


def _load_json_arg(value: str | None) -> dict[str, Any] | None:
    if value is None:
        return None
    parsed = json.loads(value)
    if not isinstance(parsed, dict):
        raise ValueError("JSON argument must be an object")
    return parsed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="spreadsheet-tools",
        description="AI-friendly Excel (.xlsx/.xlsm) tools with cleaned reads and explicit style edits.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    list_cmd = subparsers.add_parser("list-sheets", help="List workbook sheets")
    list_cmd.add_argument("file", help="Path to .xlsx or .xlsm file")

    info_cmd = subparsers.add_parser("workbook-info", help="Workbook metadata")
    info_cmd.add_argument("file")

    sheet_cmd = subparsers.add_parser("sheet-info", help="Sheet dimensions and structure")
    sheet_cmd.add_argument("file")
    sheet_cmd.add_argument("--sheet", help="Sheet name (defaults to active sheet)")

    read_cmd = subparsers.add_parser("read-range", help="Read a cleaned value-only range")
    read_cmd.add_argument("file")
    read_cmd.add_argument("--sheet")
    read_cmd.add_argument("--from-col", required=True, help="Start column letter, e.g. A")
    read_cmd.add_argument("--to-col", required=True, help="End column letter, e.g. L")
    read_cmd.add_argument("--from-row", type=int, required=True, help="Zero-based start row")
    read_cmd.add_argument("--to-row", type=int, required=True, help="Zero-based end row")
    read_cmd.add_argument(
        "--include-empty-rows",
        action="store_true",
        help="Keep fully empty rows inside the requested range",
    )
    read_cmd.add_argument(
        "--keep-trailing-empty",
        action="store_true",
        help="Do not trim trailing empty rows/columns",
    )
    read_cmd.add_argument(
        "--include-formulas",
        action="store_true",
        help="Return formula text instead of computed values",
    )

    cell_cmd = subparsers.add_parser("read-cell", help="Read one cell value")
    cell_cmd.add_argument("file")
    cell_cmd.add_argument("--sheet")
    cell_cmd.add_argument("--cell", required=True)

    style_cmd = subparsers.add_parser("cell-style", help="Read one cell style metadata")
    style_cmd.add_argument("file")
    style_cmd.add_argument("--sheet")
    style_cmd.add_argument("--cell", required=True)

    edit_cmd = subparsers.add_parser("edit-cell", help="Edit cell value and/or style")
    edit_cmd.add_argument("file")
    edit_cmd.add_argument("--sheet")
    edit_cmd.add_argument("--cell", required=True)
    edit_cmd.add_argument("--value", help="New cell value")
    edit_cmd.add_argument("--clear", action="store_true", help="Clear the cell value")
    edit_cmd.add_argument(
        "--style-json",
        help='Style overrides as JSON, e.g. {"font":{"bold":true},"number_format":"0.00"}',
    )
    edit_cmd.add_argument(
        "--dry-run",
        action="store_true",
        help="Apply changes in memory without saving the workbook",
    )

    find_cmd = subparsers.add_parser("find", help="Search for text in workbook values")
    find_cmd.add_argument("file")
    find_cmd.add_argument("--query", required=True)
    find_cmd.add_argument("--sheet")
    find_cmd.add_argument("--case-sensitive", action="store_true")
    find_cmd.add_argument("--max-results", type=int, default=50)

    copy_cmd = subparsers.add_parser("copy-sheet", help="Duplicate a sheet inside the workbook")
    copy_cmd.add_argument("file")
    copy_cmd.add_argument("--source-sheet", required=True)
    copy_cmd.add_argument("--target-sheet", required=True)
    copy_cmd.add_argument("--overwrite", action="store_true")

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        if args.command == "list-sheets":
            _print_json(list_sheets(args.file))
        elif args.command == "workbook-info":
            _print_json(workbook_info(args.file))
        elif args.command == "sheet-info":
            _print_json(sheet_info(args.file, args.sheet))
        elif args.command == "read-range":
            _print_json(
                read_range(
                    args.file,
                    sheet_name=args.sheet,
                    from_col=args.from_col,
                    to_col=args.to_col,
                    from_row=args.from_row,
                    to_row=args.to_row,
                    include_empty_rows=args.include_empty_rows,
                    trim_empty=not args.keep_trailing_empty,
                    include_formulas=args.include_formulas,
                )
            )
        elif args.command == "read-cell":
            _print_json(read_cell(args.file, sheet_name=args.sheet, address=args.cell))
        elif args.command == "cell-style":
            _print_json(get_cell_style_info(args.file, sheet_name=args.sheet, address=args.cell))
        elif args.command == "edit-cell":
            if args.value is None and not args.clear and not args.style_json:
                parser.error("edit-cell requires at least one of --value, --clear, or --style-json")
            _print_json(
                edit_cell(
                    args.file,
                    sheet_name=args.sheet,
                    address=args.cell,
                    value=args.value,
                    clear_value=args.clear,
                    style=_load_json_arg(args.style_json),
                    save=not args.dry_run,
                )
            )
        elif args.command == "find":
            _print_json(
                find_values(
                    args.file,
                    query=args.query,
                    sheet_name=args.sheet,
                    case_sensitive=args.case_sensitive,
                    max_results=args.max_results,
                )
            )
        elif args.command == "copy-sheet":
            _print_json(
                copy_sheet_structure(
                    args.file,
                    source_sheet=args.source_sheet,
                    target_sheet=args.target_sheet,
                    overwrite=args.overwrite,
                )
            )
        else:
            parser.error(f"Unknown command: {args.command}")
            return 2
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
