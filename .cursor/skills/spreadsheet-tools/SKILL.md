---
name: spreadsheet-tools
description: >-
  Read, search, and edit Excel (.xlsx/.xlsm) workbooks via local Python CLI tools
  with cleaned value-only reads and explicit style metadata commands. Use when
  working with spreadsheets, Excel files, .xlsx, .xlsm, sheet ranges, cell edits,
  or workbook inspection.
---

# Spreadsheet Tools

Local CLI for AI-friendly Excel interaction. Always run commands from the repo root.

## Setup

```bash
export PATH="$HOME/.local/bin:$PATH"
uv sync
```

Run tools with:

```bash
uv run spreadsheet-tools <command> ...
```

## Workflow

1. `workbook-info` or `list-sheets` to discover structure.
2. `sheet-info` for dimensions, merged ranges, and filters.
3. `read-range` for value-only data (formatting stripped automatically).
4. `cell-style` only when formatting matters.
5. `edit-cell` to change values and/or style metadata.
6. `find` to locate text quickly.

## Commands

### List sheets

```bash
uv run spreadsheet-tools list-sheets "file.xlsm"
```

### Workbook metadata

```bash
uv run spreadsheet-tools workbook-info "file.xlsm"
```

### Sheet structure

```bash
uv run spreadsheet-tools sheet-info "file.xlsm" --sheet "Sheet1"
```

### Read cleaned range

Zero-based rows. Columns use Excel letters.

```bash
uv run spreadsheet-tools read-range "file.xlsm" \
  --sheet "Sheet1" \
  --from-col A --to-col L \
  --from-row 0 --to-row 10
```

Options:
- `--include-empty-rows`: keep blank rows inside the requested window
- `--keep-trailing-empty`: disable trailing empty row/column trimming
- `--include-formulas`: return formula text instead of computed values

Output contains values only. Empty rows/columns are trimmed by default.

### Read one cell

```bash
uv run spreadsheet-tools read-cell "file.xlsm" --sheet "Sheet1" --cell B3
```

### Cell style metadata

Use when the task depends on formatting, not just content.

```bash
uv run spreadsheet-tools cell-style "file.xlsm" --sheet "Sheet1" --cell B3
```

Returns font, fill, alignment, border, number format, comment, and hyperlink when present.

### Edit cell

Content-only edit (style unchanged):

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 --value "New text"
```

Style edit via JSON (unspecified style fields stay as-is):

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 \
  --style-json '{"font":{"bold":true,"color":"FF0000"},"number_format":"0.00"}'
```

Clear value:

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 --clear
```

Preview without saving:

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 --value "Test" --dry-run
```

### Search values

```bash
uv run spreadsheet-tools find "file.xlsm" --query "empresa" --max-results 20
```

### Copy sheet

```bash
uv run spreadsheet-tools copy-sheet "file.xlsm" --source-sheet "Sheet1" --target-sheet "Sheet1_copy"
```

## Style JSON reference

`--style-json` accepts partial updates:

```json
{
  "font": {"name": "Arial", "size": 12, "bold": true, "italic": false, "color": "FF0000"},
  "fill": {"fill_type": "solid", "start_color": "FFFF00"},
  "alignment": {"horizontal": "center", "vertical": "center", "wrap_text": true},
  "number_format": "0.00",
  "comment": "Review this value",
  "hyperlink": "https://example.com"
}
```

Omitted keys are left unchanged on the cell.

## Rules for the agent

- Prefer `read-range` over opening raw Excel files in context.
- Read small ranges first; expand only when needed.
- Do not assume 1-based rows in CLI flags: `--from-row` and `--to-row` are zero-based.
- Use `cell-style` sparingly; it is heavier than value reads.
- Before bulk edits, inspect with `sheet-info` and `find`.
- After edits, re-read the affected range to verify.
- Keep `.xlsx/.xlsm` files out of git; they are gitignored by design.

## Additional resources

- Command details and JSON shapes: [reference.md](reference.md)
