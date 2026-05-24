---
name: spreadsheet-tools
description: >-
  Read, search, and edit Excel (.xlsx/.xlsm) workbooks via local Python CLI tools
  with cleaned value-only reads and explicit style metadata commands. Use when
  working with spreadsheets, Excel files, .xlsx, .xlsm, sheet ranges, cell edits,
  or workbook inspection.
allowed-tools: Bash(export *) Bash(uv *) Bash(uv run *) Bash(uv sync *)
---

# Spreadsheet Tools

Local CLI for AI-friendly Excel interaction. Always run from the repo root.

Invoke manually with `/spreadsheet-tools` or let Claude load this skill when spreadsheet work is detected.

## Setup

```bash
export PATH="$HOME/.local/bin:$PATH"
uv sync
```

Run any command:

```bash
uv run spreadsheet-tools <command> [args...]
```

## Critical Architecture Facts

**Before touching any cell, understand these:**

1. **Row numbering**: CLI flags `--from-row` and `--to-row` are **zero-based**. Row 0 in CLI = Excel row 1. Internally Excel uses 1-based rows (add 1 when calling openpyxl directly).

2. **Merged cells**: Many cells in structured forms are merged across columns/rows. The merge master cell holds the value; slave cells always return `null`. Writing to a slave is auto-redirected to its master (with a warning). Use `sheet-info` to see all merged ranges before editing.

3. **Columns**: Always use Excel letters (`A`, `B`, `D`, `L`). Never use column indices.

4. **Save safety**: Every write command does: backup → in-memory buffer → ZIP merge (preserves drawings/VBA/images) → atomic rename. Original is never corrupted.

5. **Data types**: `--value` for `edit-cell` and `batch-edit` is coerced: numeric strings become numbers, booleans stay bools. Pass `300000` not `"300000"` for cost cells.

6. **Sheets**: If `--sheet` is omitted, the first sheet is used. Always pass `--sheet` to be explicit.

---

## Recommended Workflow

When exploring and filling a structured Excel form:

1. `list-sheets` → discover sheet names
2. `sheet-info` → see dimensions + all merged ranges
3. `section-map` → find section headers (`N.N.N Title`) + row ranges
4. `read-range` → inspect cell values for a specific section
5. `audit-range` → find all empty master cells before bulk filling
6. `snapshot` with `--tag before` → safety snapshot before edits
7. `batch-edit` → fill multiple cells in one save (preferred over repeated `edit-cell`)
8. `describe-section` → verify strategy table consistency (name ↔ description)
9. `validate` → assert data quality rules
10. `snapshot` with `--tag after` + `snapshot-diff` → verify what changed

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

Returns dimensions, merged ranges, and active filters.

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

## New Analysis & Bulk Commands

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

### create-empty-workbook

Create new empty workbook from scratch (overwrites existing file).

```bash
uv run spreadsheet-tools create-empty-workbook "new_file.xlsx" [--sheet "Sheet1"]
```

### list-snapshots

```bash
uv run spreadsheet-tools list-snapshots "file.xlsm" [--sheet "Sheet1"]
```

---

## Style JSON Reference

`--style-json` and `"style"` in batch-edit accept partial updates. Omitted keys stay unchanged.

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

Colors are hex RGB without `#` (e.g., `"FF0000"` for red, `"000000"` for black).

---

## Agent Rules (follow always)

- Use `section-map` + `audit-range` before bulk filling to find empty cells.
- Use `batch-edit` when filling 3+ cells — avoids repeated save overhead.
- Never assume a row number — verify with `read-range` or `section-map` first.
- Merged cells: if a cell shows `null` on read, it's a slave. Check `sheet-info` merged ranges.
- Use `describe-section` after filling strategy/product tables to verify consistency.
- Use `validate` before marking a section complete.
- Take a `snapshot` with `--tag before` before destructive batch operations.
- After edits, re-read the affected range to verify values landed correctly.
- `--from-row`/`--to-row` are zero-based. Excel row 325 = `--from-row 324`.
- Keep `.xlsx/.xlsm` files out of git; they are gitignored by design.

## Additional resources

- Full JSON output shapes: [reference.md](reference.md)
