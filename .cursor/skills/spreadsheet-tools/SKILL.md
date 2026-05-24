---
name: spreadsheet-tools
description: >-
  Read, search, and edit Excel (.xlsx/.xlsm) workbooks via local Python CLI tools
  with cleaned value-only reads and explicit style metadata commands. Use when
  working with spreadsheets, Excel files, .xlsx, .xlsm, sheet ranges, cell edits,
  or workbook inspection.
---

# Spreadsheet Tools

Local CLI for AI-friendly Excel interaction. Always run from the repo root.

## ⚠ CRITICAL — AI AGENT RULES (NO EXCEPTIONS)

> **NEVER create Python scripts or call openpyxl directly.**
> Every operation — reads, writes, styles, merges, dimensions — MUST go through
> `uv run spreadsheet-tools <command>`. No exceptions. Do NOT write helper `.py`
> files. Do NOT use `uv run python -c`. Do NOT import openpyxl yourself.
>
> **ALL writes use ZIP-merge save internally.** This preserves images, drawings,
> VBA macros, charts, and all embedded objects in the file. If you bypass the CLI
> you WILL corrupt those assets.
>
> **Workspace**: always create new workbooks in `workspace/` by passing a bare
> filename (no directory) to `create-empty-workbook`. The directory is created
> automatically. Reference the file as `workspace/file.xlsx` in all subsequent
> commands.

---

## Setup

```bash
export PATH="$HOME/.local/bin:$PATH"
uv sync
```

Run any command:

```bash
uv run spreadsheet-tools <command> [args...]
```

---

## Critical Architecture Facts

**Before touching any cell, understand these:**

1. **Row numbering**: CLI flags `--from-row` and `--to-row` are **zero-based**. Row 0 in CLI = Excel row 1. Write commands (`--row`, `sheet.row_dimensions`) are 1-based. Never confuse the two.

2. **Merged cells**: The master cell (top-left) holds value and style. Slave cells always return `null`. Writing to a slave is auto-redirected to master (with a warning). Use `sheet-info` to see all merged ranges before editing.

3. **Columns**: Always use Excel letters (`A`, `B`, `D`, `L`). Never use column indices.

4. **Save safety**: Every write command does: backup → in-memory buffer → ZIP merge (preserves drawings/VBA/images) → atomic rename. Original is **never** corrupted. The `.bak` file is auto-deleted after successful save.

5. **Data types**: `--value` for `edit-cell` and `batch-edit` is coerced: numeric strings become numbers, booleans stay bools. Pass `300000` not `"300000"` for cost cells.

6. **Sheets**: If `--sheet` is omitted, the first sheet is used. Always pass `--sheet` to be explicit.

7. **Format preservation**: `merge-cells`, `format-range`, `set-column-width`, `set-row-height`, `freeze-panes` all go through the same ZIP-merge save. Images and drawings are never lost.

---

## Workspace Convention

New workbooks should always be created in `workspace/`. Pass a bare filename
(no path separator) to `create-empty-workbook` — the `workspace/` directory is
auto-created:

```bash
# Creates workspace/report.xlsx
uv run spreadsheet-tools create-empty-workbook "report.xlsx" --sheet "Data"

# All subsequent commands reference the full path
uv run spreadsheet-tools edit-cell "workspace/report.xlsx" --sheet "Data" --cell A1 --value "Hello"
```

---

## Recommended Workflow

When exploring and filling a structured Excel form:

1. `list-sheets` → discover sheet names
2. `sheet-info` → see dimensions + all merged ranges + freeze panes
3. `section-map` → find section headers (`N.N.N Title`) + row ranges
4. `read-range` → inspect cell values for a specific section
5. `audit-range` → find all empty master cells before bulk filling
6. `snapshot` with `--tag before` → safety snapshot before edits
7. `batch-edit` → fill multiple cells in one save (preferred over repeated `edit-cell`)
8. `describe-section` → verify strategy table consistency (name ↔ description)
9. `validate` → assert data quality rules
10. `snapshot` with `--tag after` + `snapshot-diff` → verify what changed

When building a new workbook from scratch:

1. `create-empty-workbook` → creates `workspace/file.xlsx`
2. `batch-edit` → populate cells with values and styles
3. `merge-cells` → merge header/title cells
4. `format-range` → apply uniform style to header rows or sections
5. `batch-set-dimensions` → set all column widths and row heights in one save
6. `freeze-panes` → freeze header row for scrolling UX
7. `set-tab-color` → color-code the sheet tab

For targeted reads/writes:
- `read-cell` → single cell value
- `edit-cell` → single cell value or style
- `find` → locate text anywhere in workbook
- `find-replace` → search + replace across sheets

---

## All Commands

### list-sheets

```bash
uv run spreadsheet-tools list-sheets "file.xlsm"
```

### workbook-info

```bash
uv run spreadsheet-tools workbook-info "file.xlsm"
```

### sheet-info

```bash
uv run spreadsheet-tools sheet-info "file.xlsm" --sheet "Sheet1"
```

Returns dimensions, merged ranges, freeze_panes, and active filters.

### read-range

Zero-based rows. Columns use Excel letters. Returns values only (no formatting).

```bash
uv run spreadsheet-tools read-range "file.xlsm" \
  --sheet "Sheet1" \
  --from-col A --to-col L \
  --from-row 0 --to-row 10
```

Options:
- `--include-empty-rows` — keep blank rows in output
- `--keep-trailing-empty` — disable trailing empty row/column trimming
- `--include-formulas` — return formula text instead of computed values

### read-cell

```bash
uv run spreadsheet-tools read-cell "file.xlsm" --sheet "Sheet1" --cell B3
```

### cell-style

Returns font, fill, alignment, border, number format, comment, hyperlink. Use sparingly — heavier than value reads.

```bash
uv run spreadsheet-tools cell-style "file.xlsm" --sheet "Sheet1" --cell B3
```

### edit-cell

Content-only (style unchanged):

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 --value "New text"
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell L3 --value 300000
```

Style only (value unchanged):

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 \
  --style-json '{"alignment":{"wrap_text":true}}'
```

Both at once:

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 \
  --value "New text" \
  --style-json '{"font":{"bold":true},"alignment":{"wrap_text":true}}'
```

Clear value:

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 --clear
```

Preview without saving:

```bash
uv run spreadsheet-tools edit-cell "file.xlsm" --sheet "Sheet1" --cell B3 --value "Test" --dry-run
```

### find

```bash
uv run spreadsheet-tools find "file.xlsm" --query "empresa" --max-results 20
```

### copy-sheet

```bash
uv run spreadsheet-tools copy-sheet "file.xlsm" \
  --source-sheet "Sheet1" --target-sheet "Sheet1_copy"
```

---

## Analysis & Bulk Commands

### section-map

Auto-discover numbered section headers (`N.N.N Title`) and their row ranges.

```bash
uv run spreadsheet-tools section-map "file.xlsm" \
  --sheet "Sheet1" \
  [--min-row 0] [--max-row 400]
```

Returns: header text, prefix, title, depth, header_row, row_range.

### audit-range

List all master cells in a range with `is_empty` flag. Skips slave cells by default.

```bash
uv run spreadsheet-tools audit-range "file.xlsm" \
  --sheet "Sheet1" \
  --from-col B --to-col L \
  --from-row 324 --to-row 328 \
  [--show-slaves]
```

Returns: per-cell is_empty, summary counts (total_master, empty_master, filled_master).

### describe-section

Audit a strategy table (name col, description col, cost col) for consistency issues.

```bash
uv run spreadsheet-tools describe-section "file.xlsm" \
  --sheet "Sheet1" \
  --data-rows 325-328 \
  [--name-col B] [--desc-col D] [--cost-col L] \
  [--header-row 324]
```

Issues detected per row:
- `generic_name` (warning) — name matches `Estrategia de X N` pattern
- `missing_description` (error) — name present but description empty
- `name_desc_mismatch` (error) — keyword overlap between name and description < 30%
- `missing_cost` (warning) — cost cell is empty

### find-replace

Search cell values and optionally replace. Works across all sheets unless `--sheet` is specified.

```bash
# Find only (no --replace-with)
uv run spreadsheet-tools find-replace "file.xlsm" --query "85000"

# Find + replace (exact whole-value match)
uv run spreadsheet-tools find-replace "file.xlsm" \
  --query "85000" --replace-with 90000 \
  [--sheet "Sheet1"] [--dry-run]

# Regex pattern (substring match)
uv run spreadsheet-tools find-replace "file.xlsm" \
  --query "\\b85000\\b" --replace-with "90000" --regex \
  [--case-sensitive] [--max-results 200]
```

Without `--regex`: entire cell value must equal query.
With `--regex`: substring match; replacement via `re.sub`.

### validate

Apply one or more validation rules against a sheet. Rules are colon-separated.

```bash
uv run spreadsheet-tools validate "file.xlsm" --sheet "Sheet1" \
  --rule "not-empty:D325" \
  --rule "not-empty-range:D325:D328" \
  --rule "price-min-max:F308:H308" \
  --rule "name-matches-desc:B325:D325" \
  --rule "numeric-range:L317:50000:2000000" \
  --rule "string-contains:D317:Descuento" \
  --rule "no-generic-name:B325"
```

| Rule | Syntax | Check |
|------|--------|-------|
| `not-empty` | `not-empty:CELL` | Non-blank value |
| `not-empty-range` | `not-empty-range:FROM:TO` | All cells non-blank |
| `price-min-max` | `price-min-max:MIN_CELL:MAX_CELL` | numeric(MIN) < numeric(MAX) |
| `name-matches-desc` | `name-matches-desc:NAME_CELL:DESC_CELL` | Keyword overlap ≥ 30% |
| `numeric-range` | `numeric-range:CELL:MIN:MAX` | MIN ≤ value ≤ MAX |
| `string-contains` | `string-contains:CELL:SUBSTRING` | Substring present (case-insensitive) |
| `no-generic-name` | `no-generic-name:CELL` | Not `Estrategia de X N` pattern |

### batch-edit

Apply multiple cell edits atomically — one workbook load, one save. Preferred over repeated `edit-cell`.

```bash
# Inline JSON
uv run spreadsheet-tools batch-edit "file.xlsm" --sheet "Sheet1" \
  --edits-json '[{"cell":"D325","value":"Text"},{"cell":"L325","value":300000}]'

# From JSON file
uv run spreadsheet-tools batch-edit "file.xlsm" \
  --edits-file edits.json [--dry-run]
```

Edits JSON format:
```json
[
  {"cell": "D325", "value": "Description text"},
  {"cell": "L325", "value": 300000},
  {"cell": "B326", "clear": true},
  {"cell": "D326", "value": "Text", "style": {"alignment": {"wrap_text": true}}}
]
```

All addresses validated upfront — if any fail, the file is never modified. Slave cells auto-redirect to master.

### snapshot

Capture all non-empty cell values as a named snapshot.

```bash
uv run spreadsheet-tools snapshot "file.xlsm" \
  --sheet "Sheet1" --tag "before-edit" \
  [--description "Optional note"]
```

Stored at: `~/.cache/spreadsheet-tools/snapshots/<file_hash>/<sheet>/<tag>.json`

### snapshot-diff

Compare two snapshots and list per-cell changes.

```bash
uv run spreadsheet-tools snapshot-diff "file.xlsm" \
  --sheet "Sheet1" \
  --tag-a "before-edit" --tag-b "after-edit"
```

### list-snapshots

```bash
uv run spreadsheet-tools list-snapshots "file.xlsm" [--sheet "Sheet1"]
```

### create-empty-workbook

Create new empty workbook from scratch. Bare filenames auto-go to `workspace/`.

```bash
# Creates workspace/report.xlsx
uv run spreadsheet-tools create-empty-workbook "report.xlsx" [--sheet "Sheet1"]

# Explicit path (no workspace routing)
uv run spreadsheet-tools create-empty-workbook "some/dir/report.xlsx"
```

---

## Structure & Formatting Commands

### merge-cells

Merge a rectangular cell range. Master cell (top-left) keeps its value and style.

```bash
uv run spreadsheet-tools merge-cells "file.xlsx" --sheet "Sheet1" --range A1:E1
```

### unmerge-cells

Remove merge from a range. Cells become independent again.

```bash
uv run spreadsheet-tools unmerge-cells "file.xlsx" --sheet "Sheet1" --range A1:E1
```

### set-column-width

Set width of a single column (Excel character units).

```bash
uv run spreadsheet-tools set-column-width "file.xlsx" --sheet "Sheet1" --col A --width 8
uv run spreadsheet-tools set-column-width "file.xlsx" --sheet "Sheet1" --col B --width 20
```

### set-row-height

Set height of a single row (points). Row number is 1-based (Excel convention).

```bash
uv run spreadsheet-tools set-row-height "file.xlsx" --sheet "Sheet1" --row 1 --height 40
```

### batch-set-dimensions

Set multiple column widths and/or row heights in a **single save**. Preferred when setting 3+ dimensions.

```bash
uv run spreadsheet-tools batch-set-dimensions "file.xlsx" \
  --sheet "Sheet1" \
  --columns-json '[{"col":"A","width":8},{"col":"B","width":15},{"col":"C","width":15}]' \
  --rows-json '[{"row":1,"height":40},{"row":8,"height":25}]'
```

Both `--columns-json` and `--rows-json` are optional; at least one is required.

### freeze-panes

Freeze rows above and columns to the left of a given cell.

```bash
# Freeze row 1 and column A (most common for tables)
uv run spreadsheet-tools freeze-panes "file.xlsx" --sheet "Sheet1" --cell B2

# Freeze only row 1
uv run spreadsheet-tools freeze-panes "file.xlsx" --sheet "Sheet1" --cell A2

# Freeze only column A
uv run spreadsheet-tools freeze-panes "file.xlsx" --sheet "Sheet1" --cell B1
```

### unfreeze-panes

Remove all freeze panes from a sheet.

```bash
uv run spreadsheet-tools unfreeze-panes "file.xlsx" --sheet "Sheet1"
```

### add-sheet

Add a new empty sheet to an existing workbook.

```bash
uv run spreadsheet-tools add-sheet "file.xlsx" --sheet "Summary"
uv run spreadsheet-tools add-sheet "file.xlsx" --sheet "Summary" --position 0
```

`--position` is zero-based sheet index. Omit to append at end.

### rename-sheet

Rename a sheet. Fails if old name not found or new name already taken.

```bash
uv run spreadsheet-tools rename-sheet "file.xlsx" --old-name "Sheet1" --new-name "Amortización"
```

### format-range

Apply the same style to every **master** cell in a rectangular range. Slave cells are skipped.

```bash
# Bold + dark blue header row
uv run spreadsheet-tools format-range "file.xlsx" --sheet "Sheet1" \
  --range A8:E8 \
  --style-json '{"font":{"bold":true,"color":"FFFFFF"},"fill":{"fill_type":"solid","start_color":"1F4E79"},"alignment":{"horizontal":"center"}}'

# Number format for a column range
uv run spreadsheet-tools format-range "file.xlsx" --sheet "Sheet1" \
  --range C9:C20 \
  --style-json '{"number_format":"#,##0.00"}'
```

### set-tab-color

Set the sheet tab color. Color is a 6-char RGB hex string (no `#`).

```bash
uv run spreadsheet-tools set-tab-color "file.xlsx" --sheet "Sheet1" --color "1F4E79"
uv run spreadsheet-tools set-tab-color "file.xlsx" --sheet "Summary" --color "70AD47"
```

### auto-filter

Enable auto-filter on a header row range, or clear it.

```bash
# Enable on row 8 headers
uv run spreadsheet-tools auto-filter "file.xlsx" --sheet "Sheet1" --range A8:E8

# Clear auto-filter
uv run spreadsheet-tools auto-filter "file.xlsx" --sheet "Sheet1"
```

---

## Style JSON Reference

`--style-json` and `"style"` in batch-edit / format-range accept partial updates. Omitted keys stay unchanged.

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

Colors are hex RGB without `#` (e.g., `"FF0000"` for red, `"1F4E79"` for dark blue).

---

## Agent Rules (follow always)

- **NEVER create Python scripts.** NEVER call `uv run python` for spreadsheet ops. Use the CLI only.
- Use `section-map` + `audit-range` before bulk filling to find empty cells.
- Use `batch-edit` when filling 3+ cells — one save, no corruption risk.
- Use `batch-set-dimensions` when setting 3+ column widths or row heights.
- Never assume a row number — verify with `read-range` or `section-map` first.
- Merged cells: if a cell shows `null` on read, it's a slave. Check `sheet-info` merged ranges.
- Use `describe-section` after filling strategy/product tables to verify consistency.
- Use `validate` before marking a section complete.
- Take a `snapshot` with `--tag before` before any destructive batch operation.
- After edits, re-read the affected range to verify values landed correctly.
- `--from-row`/`--to-row` are zero-based. Excel row 325 = `--from-row 324`.
- `--row` in `set-row-height` and `batch-set-dimensions` is **1-based** (Excel row number).
- Keep `.xlsx/.xlsm` files out of git; they are gitignored by design.
- Create new workbooks with bare filenames → they go to `workspace/` automatically.

## Additional resources

- Full JSON output shapes: [reference.md](reference.md)