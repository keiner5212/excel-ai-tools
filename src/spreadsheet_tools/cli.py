from __future__ import annotations

import argparse
import json
import math
import sys
import zipfile
from typing import Any

from spreadsheet_tools.reader import (
    audit_range,
    describe_section,
    find_values,
    list_sheets,
    read_cell,
    read_range,
    section_map,
    sheet_info,
    validate_rules,
    workbook_info,
)
from spreadsheet_tools.writer import (
    batch_edit,
    copy_sheet_structure,
    edit_cell,
    find_replace,
    get_cell_style_info,
)
from spreadsheet_tools.snapshot import (
    create_snapshot,
    diff_snapshots,
    list_snapshots,
)


def _coerce_value(raw: str) -> object:
    """Coerce a CLI string value to int, float, or str.

    This prevents numeric data from being stored as text in Excel,
    which would break formulas and sorting. Booleans are kept as strings
    because "TRUE"/"FALSE" are ambiguous across locales.

    nan and inf are rejected as strings to avoid writing XML-invalid float
    values that corrupt the workbook.
    """
    try:
        return int(raw)
    except ValueError:
        pass
    try:
        v = float(raw)
        if math.isnan(v) or math.isinf(v):
            # nan/inf are not valid xs:double in OOXML; store as plain string.
            return raw
        return v
    except ValueError:
        pass
    return raw


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

    sheet_cmd = subparsers.add_parser(
        "sheet-info", help="Sheet dimensions and structure"
    )
    sheet_cmd.add_argument("file")
    sheet_cmd.add_argument("--sheet", help="Sheet name (defaults to active sheet)")

    read_cmd = subparsers.add_parser(
        "read-range", help="Read a cleaned value-only range"
    )
    read_cmd.add_argument("file")
    read_cmd.add_argument("--sheet")
    read_cmd.add_argument(
        "--from-col", required=True, help="Start column letter, e.g. A"
    )
    read_cmd.add_argument("--to-col", required=True, help="End column letter, e.g. L")
    read_cmd.add_argument(
        "--from-row", type=int, required=True, help="Zero-based start row"
    )
    read_cmd.add_argument(
        "--to-row", type=int, required=True, help="Zero-based end row"
    )
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

    copy_cmd = subparsers.add_parser(
        "copy-sheet", help="Duplicate a sheet inside the workbook"
    )
    copy_cmd.add_argument("file")
    copy_cmd.add_argument("--source-sheet", required=True)
    copy_cmd.add_argument("--target-sheet", required=True)
    copy_cmd.add_argument("--overwrite", action="store_true")

    # --- section-map ---
    smap_cmd = subparsers.add_parser(
        "section-map", help="Discover numbered section headers and their row ranges"
    )
    smap_cmd.add_argument("file")
    smap_cmd.add_argument("--sheet")
    smap_cmd.add_argument("--min-row", type=int, default=0, help="Zero-based start row")
    smap_cmd.add_argument(
        "--max-row",
        type=int,
        default=None,
        help="Zero-based end row (default: sheet max)",
    )

    # --- audit-range ---
    audit_cmd = subparsers.add_parser(
        "audit-range", help="Audit master cells in a range, flagging empty ones"
    )
    audit_cmd.add_argument("file")
    audit_cmd.add_argument("--sheet")
    audit_cmd.add_argument("--from-col", required=True)
    audit_cmd.add_argument("--to-col", required=True)
    audit_cmd.add_argument("--from-row", type=int, required=True)
    audit_cmd.add_argument("--to-row", type=int, required=True)
    audit_cmd.add_argument(
        "--show-slaves", action="store_true", help="Include slave (non-writable) cells"
    )

    # --- describe-section ---
    desc_cmd = subparsers.add_parser(
        "describe-section",
        help="Audit a strategy table for name/description consistency",
    )
    desc_cmd.add_argument("file")
    desc_cmd.add_argument("--sheet")
    desc_cmd.add_argument(
        "--data-rows",
        required=True,
        help="Zero-based row range, e.g. 325-328",
    )
    desc_cmd.add_argument("--header-row", type=int, default=None)
    desc_cmd.add_argument("--name-col", default="B", help="Column with strategy names")
    desc_cmd.add_argument("--desc-col", default="D", help="Column with descriptions")
    desc_cmd.add_argument(
        "--cost-col", default="L", help="Column with annual costs (pass '' to skip)"
    )

    # --- find-replace ---
    fr_cmd = subparsers.add_parser(
        "find-replace", help="Search cell values and optionally replace matches"
    )
    fr_cmd.add_argument("file")
    fr_cmd.add_argument("--query", required=True)
    fr_cmd.add_argument("--replace-with", default=None, help="Replacement value")
    fr_cmd.add_argument("--sheet")
    fr_cmd.add_argument("--case-sensitive", action="store_true")
    fr_cmd.add_argument("--regex", action="store_true", help="Treat --query as regex")
    fr_cmd.add_argument("--dry-run", action="store_true")
    fr_cmd.add_argument("--max-results", type=int, default=200)

    # --- validate ---
    val_cmd = subparsers.add_parser(
        "validate", help="Apply validation rules against a sheet"
    )
    val_cmd.add_argument("file")
    val_cmd.add_argument("--sheet")
    val_cmd.add_argument(
        "--rule",
        dest="rules",
        action="append",
        required=True,
        metavar="RULE",
        help=(
            "Validation rule, repeatable. Examples: "
            "'not-empty:D317', "
            "'price-min-max:F308:H308', "
            "'name-matches-desc:B325:D325', "
            "'numeric-range:L317:50000:2000000', "
            "'string-contains:D317:Descuento', "
            "'no-generic-name:B325'"
        ),
    )

    # --- batch-edit ---
    batch_cmd = subparsers.add_parser(
        "batch-edit", help="Apply multiple cell edits atomically in one save"
    )
    batch_cmd.add_argument("file")
    batch_cmd.add_argument("--sheet")
    _batch_src = batch_cmd.add_mutually_exclusive_group(required=True)
    _batch_src.add_argument("--edits-file", help="Path to JSON file with edits array")
    _batch_src.add_argument("--edits-json", help="Inline JSON edits array")
    batch_cmd.add_argument("--dry-run", action="store_true")

    # --- snapshot ---
    snap_cmd = subparsers.add_parser(
        "snapshot", help="Capture sheet cell values to a named snapshot"
    )
    snap_cmd.add_argument("file")
    snap_cmd.add_argument("--sheet")
    snap_cmd.add_argument("--tag", required=True, help="Snapshot tag name")
    snap_cmd.add_argument("--description", default=None)

    # --- snapshot-diff ---
    diff_cmd = subparsers.add_parser(
        "snapshot-diff", help="Compare two snapshots and list changed cells"
    )
    diff_cmd.add_argument("file")
    diff_cmd.add_argument("--sheet")
    diff_cmd.add_argument("--tag-a", required=True)
    diff_cmd.add_argument("--tag-b", required=True)

    # --- list-snapshots ---
    ls_snap_cmd = subparsers.add_parser(
        "list-snapshots", help="List available snapshots for a sheet"
    )
    ls_snap_cmd.add_argument("file")
    ls_snap_cmd.add_argument("--sheet")

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
            _print_json(
                get_cell_style_info(args.file, sheet_name=args.sheet, address=args.cell)
            )
        elif args.command == "edit-cell":
            if args.value is None and not args.clear and not args.style_json:
                parser.error(
                    "edit-cell requires at least one of --value, --clear, or --style-json"
                )
            _print_json(
                edit_cell(
                    args.file,
                    sheet_name=args.sheet,
                    address=args.cell,
                    value=_coerce_value(args.value) if args.value is not None else None,
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
        elif args.command == "section-map":
            _print_json(
                section_map(
                    args.file,
                    sheet_name=args.sheet,
                    min_row=args.min_row,
                    max_row=args.max_row,
                )
            )
        elif args.command == "audit-range":
            _print_json(
                audit_range(
                    args.file,
                    sheet_name=args.sheet,
                    from_col=args.from_col,
                    to_col=args.to_col,
                    from_row=args.from_row,
                    to_row=args.to_row,
                    show_slaves=args.show_slaves,
                )
            )
        elif args.command == "describe-section":
            raw = args.data_rows
            try:
                start_str, end_str = raw.split("-", 1)
                from_data = int(start_str)
                to_data = int(end_str)
            except ValueError:
                parser.error("--data-rows must be in format START-END, e.g. 325-328")
            cost_col_arg = args.cost_col if args.cost_col.strip() else None
            _print_json(
                describe_section(
                    args.file,
                    sheet_name=args.sheet,
                    from_data_row=from_data,
                    to_data_row=to_data,
                    name_col=args.name_col,
                    desc_col=args.desc_col,
                    cost_col=cost_col_arg,
                    header_row=args.header_row,
                )
            )
        elif args.command == "find-replace":
            _print_json(
                find_replace(
                    args.file,
                    query=args.query,
                    replace_with=_coerce_value(args.replace_with)
                    if args.replace_with is not None
                    else None,
                    sheet_name=args.sheet,
                    case_sensitive=args.case_sensitive,
                    use_regex=args.regex,
                    dry_run=args.dry_run,
                    max_results=args.max_results,
                )
            )
        elif args.command == "validate":
            _print_json(
                validate_rules(
                    args.file,
                    sheet_name=args.sheet,
                    rules=args.rules,
                )
            )
        elif args.command == "batch-edit":
            if args.edits_file:
                with open(args.edits_file, encoding="utf-8") as fh:
                    raw_edits = json.load(fh)
            else:
                raw_edits = json.loads(args.edits_json)
            if not isinstance(raw_edits, list):
                parser.error("edits must be a JSON array")
            _print_json(
                batch_edit(
                    args.file,
                    sheet_name=args.sheet,
                    edits=raw_edits,
                    dry_run=args.dry_run,
                )
            )
        elif args.command == "snapshot":
            _print_json(
                create_snapshot(
                    args.file,
                    sheet_name=args.sheet,
                    tag=args.tag,
                    description=args.description,
                )
            )
        elif args.command == "snapshot-diff":
            _print_json(
                diff_snapshots(
                    args.file,
                    sheet_name=args.sheet,
                    tag_a=args.tag_a,
                    tag_b=args.tag_b,
                )
            )
        elif args.command == "list-snapshots":
            _print_json(list_snapshots(args.file, sheet_name=args.sheet))
        else:
            parser.error(f"Unknown command: {args.command}")
    except (
        ValueError,
        KeyError,
        FileNotFoundError,
        PermissionError,
        AttributeError,
        zipfile.BadZipFile,
    ) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
