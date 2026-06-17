# Spreadsheet Tools Reference

> **AGENT CONSTRAINT**: All output shapes below come from `uv run spreadsheet-tools`.
> Never call openpyxl directly. Never create Python scripts. Use the CLI only.

All commands output JSON to stdout. Errors print `error: <message>` to stderr and exit code 1.

---

## Existing Commands

### read-range output

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "range": {"from_col": "A", "to_col": "L", "from_row": 0, "to_row": 10},
  "row_count": 3,
  "rows": [
    {
      "row": 0,
      "cells": [
        {"column": "A", "address": "A1", "value": "Header"},
        {"column": "B", "address": "B1", "value": null},
        {"column": "C", "address": "C1", "value": 12345}
      ]
    }
  ],
  "merged_ranges": ["A1:C1"],
  "notes": ["Values only; formatting metadata excluded."]
}
```

**Notes:**
- Slave cells of merged ranges return `null`. The master cell holds the value.
- `row` is zero-based (row 0 = Excel row 1).
- Empty rows and trailing empty columns are trimmed by default.

### read-cell output

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "address": "B3",
  "resolved_address": "B3",
  "value": "Conjunto deportivo adaptativo",
  "merge_info": null
}
```

If a slave cell is read:
```json
{
  "address": "C3",
  "resolved_address": "B3",
  "value": "Conjunto deportivo adaptativo",
  "merge_info": {"range": "B3:E3", "master": "B3"}
}
```

### cell-style output

Only non-default fields are returned.

```json
{
  "address": "B3",
  "value": 1234.5,
  "data_type": "n",
  "number_format": "#,##0.00",
  "font": {"bold": true, "color": "FF000000"},
  "fill": {"fill_type": "solid", "start_color": "FFFFFF00"},
  "alignment": {"horizontal": "center", "wrap_text": true},
  "comment": "Check source",
  "hyperlink": "https://example.com"
}
```

### edit-cell result

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "address": "B3",
  "resolved_address": "B3",
  "updated": {"value": "New text", "style_changed": true},
  "saved": true,
  "warnings": []
}
```

If redirected from slave to master:
```json
{
  "address": "C3",
  "resolved_address": "B3",
  "warnings": ["Address C3 is a slave cell of merged range B3:E3. Redirected to master cell B3."]
}
```

### find output

```json
{
  "file": "workbook.xlsm",
  "query": "empresa",
  "matches": [
    {"sheet": "Sheet1", "address": "B5", "value": "La empresa Luqui Creaciones"},
    {"sheet": "Sheet2", "address": "A12", "value": "empresa"}
  ],
  "match_count": 2
}
```

### list-sheets output

```json
{
  "file": "workbook.xlsm",
  "sheets": ["Sheet1", "Estudio de Mercados", "Finanzas"],
  "sheet_count": 3
}
```

### sheet-info output

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "dimensions": {
    "min_row": 1,
    "max_row": 400,
    "min_col": "A",
    "max_col": "L"
  },
  "max_row": 400,
  "max_column": 12,
  "max_column_letter": "L",
  "merged_ranges": ["B3:E3", "D271:L271", "F308:H308"],
  "freeze_panes": null,
  "auto_filter": null
}
```

### workbook-info output

```json
{
  "file": "workbook.xlsm",
  "sheet_count": 5,
  "sheets": ["Sheet1", "Estudio de Mercados"],
  "has_vba": true,
  "file_size_bytes": 245760
}
```

### copy-sheet output

```json
{
  "file": "workbook.xlsm",
  "source_sheet": "Sheet1",
  "target_sheet": "Sheet1_copy",
  "saved": true
}
```

---

## New Commands

### section-map output

Discovers numbered headers (`N.N.N Title`) in columns B or A and maps each to its row range.

```bash
uv run spreadsheet-tools section-map "file.xlsm" \
  --sheet "Sheet1" \
  [--min-row 0] [--max-row 400]
```

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "sections": [
    {
      "header": "2.9.1. Producto",
      "prefix": "2.9.1",
      "title": "Producto",
      "depth": 3,
      "header_row": 271,
      "column": "B",
      "row_range": {"from": 271, "to": 301}
    },
    {
      "header": "2.9.2. Precio",
      "prefix": "2.9.2",
      "title": "Precio",
      "depth": 3,
      "header_row": 302,
      "column": "B",
      "row_range": {"from": 302, "to": 329}
    }
  ],
  "section_count": 2
}
```

**Notes:**
- `header_row` is zero-based (matches CLI `--from-row` convention).
- `row_range.from` and `row_range.to` are both zero-based.
- `depth` = number of numeric components (e.g., `2.9.1` → depth 3).
- Use `row_range` directly as `--from-row`/`--to-row` in subsequent `audit-range` calls.

---

### audit-range output

Lists all master cells in a range, each with an `is_empty` flag.

```bash
uv run spreadsheet-tools audit-range "file.xlsm" \
  --sheet "Sheet1" \
  --from-col B --to-col L \
  --from-row 324 --to-row 328 \
  [--show-slaves]
```

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "range": {"from_col": "B", "to_col": "L", "from_row": 324, "to_row": 328},
  "cells": [
    {"address": "B325", "is_master": true, "is_empty": false, "value": "Diseño del logo"},
    {"address": "D325", "is_master": true, "is_empty": false, "value": "Diseño del logotipo..."},
    {"address": "L325", "is_master": true, "is_empty": false, "value": 300000},
    {"address": "B326", "is_master": true, "is_empty": true,  "value": null},
    {"address": "D326", "is_master": true, "is_empty": true,  "value": null},
    {"address": "L326", "is_master": true, "is_empty": false, "value": 250000}
  ],
  "summary": {
    "total_master": 15,
    "filled_master": 12,
    "empty_master": 3
  }
}
```

**Notes:**
- Only master cells are shown unless `--show-slaves` is passed.
- `is_empty` is `true` for `null`, empty string, and whitespace-only values.
- Use `empty_master` cells as the fill target list for `batch-edit`.

---

### describe-section output

Audits a strategy table (name col, description col, cost col) for consistency.

```bash
uv run spreadsheet-tools describe-section "file.xlsm" \
  --sheet "Sheet1" \
  --data-rows 325-328 \
  [--name-col B] [--desc-col D] [--cost-col L] \
  [--header-row 324]
```

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "data_rows": "325-328",
  "name_col": "B",
  "desc_col": "D",
  "cost_col": "L",
  "rows": [
    {
      "row": 325,
      "name": "Diseño del logo",
      "description": "Diseño del logotipo de la marca Luqui Creaciones...",
      "cost": 300000,
      "issues": []
    },
    {
      "row": 326,
      "name": "Estrategia de Producto 2",
      "description": "Diversificación de línea de productos...",
      "cost": 250000,
      "issues": [
        {"type": "generic_name", "severity": "warning", "message": "Name 'Estrategia de Producto 2' matches generic pattern"}
      ]
    },
    {
      "row": 327,
      "name": "Investigación continua",
      "description": null,
      "cost": 400000,
      "issues": [
        {"type": "missing_description", "severity": "error", "message": "Name present but description is empty"}
      ]
    },
    {
      "row": 328,
      "name": "Marketing digital",
      "description": "Gestión de redes sociales empresariales...",
      "cost": null,
      "issues": [
        {"type": "missing_cost", "severity": "warning", "message": "Cost cell is empty"},
        {"type": "name_desc_mismatch", "severity": "error", "message": "Keyword overlap 0.10 < 0.30 threshold"}
      ]
    }
  ],
  "summary": {
    "total": 4,
    "ok": 1,
    "with_errors": 2,
    "with_warnings": 2
  }
}
```

**Issue types:**
| Type | Severity | Meaning |
|------|----------|---------|
| `generic_name` | warning | Name matches `Estrategia de X N` pattern |
| `missing_description` | error | Name present, description null/empty |
| `name_desc_mismatch` | error | Keyword overlap between name and description < 30% |
| `missing_cost` | warning | Cost cell null/empty |

---

### find-replace output

#### Find only (no `--replace-with`)

```bash
uv run spreadsheet-tools find-replace "file.xlsm" --query "85000"
```

```json
{
  "file": "workbook.xlsm",
  "query": "85000",
  "mode": "find",
  "regex": false,
  "case_sensitive": false,
  "matches": [
    {"sheet": "Estudio de Mercados", "address": "F308", "value": 85000},
    {"sheet": "Estudio de Mercados", "address": "I308", "value": "Precio desde $85,000 COP..."}
  ],
  "match_count": 2
}
```

#### Find + replace (dry-run)

```bash
uv run spreadsheet-tools find-replace "file.xlsm" \
  --query "85000" --replace-with 90000 --dry-run
```

```json
{
  "file": "workbook.xlsm",
  "query": "85000",
  "replace_with": 90000,
  "mode": "replace",
  "dry_run": true,
  "regex": false,
  "replacements_preview": [
    {
      "sheet": "Estudio de Mercados",
      "address": "F308",
      "before": 85000,
      "after": 90000
    }
  ],
  "replacement_count": 1,
  "saved": false
}
```

#### Find + replace (applied)

```json
{
  "file": "workbook.xlsm",
  "query": "85000",
  "replace_with": 90000,
  "mode": "replace",
  "dry_run": false,
  "replacements": [
    {"sheet": "Estudio de Mercados", "address": "F308", "before": 85000, "after": 90000}
  ],
  "replacement_count": 1,
  "saved": true
}
```

**Notes:**
- Without `--regex`: full cell value equality check (works on numeric cells too).
- With `--regex`: substring match on string representation; `re.sub` applied to string; result coerced back.
- `replace_with` preserves numeric type when passed as a number (e.g., `--replace-with 90000` sets an `int`, not `"90000"`).

---

### validate output

```bash
uv run spreadsheet-tools validate "file.xlsm" --sheet "Sheet1" \
  --rule "not-empty:D325" \
  --rule "price-min-max:F308:H308" \
  --rule "name-matches-desc:B325:D325" \
  --rule "numeric-range:L317:50000:2000000" \
  --rule "string-contains:D317:Descuento" \
  --rule "no-generic-name:B325"
```

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "results": [
    {
      "rule": "not-empty:D325",
      "type": "not-empty",
      "addresses": ["D325"],
      "passed": true,
      "message": "D325 is not empty"
    },
    {
      "rule": "price-min-max:F308:H308",
      "type": "price-min-max",
      "addresses": ["F308", "H308"],
      "passed": true,
      "message": "F308 (85000) < H308 (130000)"
    },
    {
      "rule": "name-matches-desc:B325:D325",
      "type": "name-matches-desc",
      "addresses": ["B325", "D325"],
      "passed": true,
      "message": "Keyword overlap 0.67 >= 0.30"
    },
    {
      "rule": "numeric-range:L317:50000:2000000",
      "type": "numeric-range",
      "addresses": ["L317"],
      "passed": false,
      "message": "L317 value 100 is below minimum 50000"
    },
    {
      "rule": "string-contains:D317:Descuento",
      "type": "string-contains",
      "addresses": ["D317"],
      "passed": true,
      "message": "D317 contains 'Descuento'"
    },
    {
      "rule": "no-generic-name:B325",
      "type": "no-generic-name",
      "addresses": ["B325"],
      "passed": true,
      "message": "B325 value 'Diseño del logo' is not a generic name"
    }
  ],
  "summary": {
    "total": 6,
    "passed": 5,
    "failed": 1
  }
}
```

**Rule syntax reference:**

| Rule | Syntax | Passes when |
|------|--------|-------------|
| `not-empty` | `not-empty:CELL` | Value is non-null, non-blank |
| `not-empty-range` | `not-empty-range:FROM_CELL:TO_CELL` | All master cells in rectangular range non-empty |
| `price-min-max` | `price-min-max:MIN_CELL:MAX_CELL` | numeric(MIN_CELL) < numeric(MAX_CELL) |
| `name-matches-desc` | `name-matches-desc:NAME_CELL:DESC_CELL` | Keyword overlap ratio ≥ 0.30 |
| `numeric-range` | `numeric-range:CELL:MIN:MAX` | MIN ≤ numeric(CELL) ≤ MAX |
| `string-contains` | `string-contains:CELL:SUBSTRING` | Case-insensitive substring present |
| `no-generic-name` | `no-generic-name:CELL` | Value doesn't match `Estrategia de X N` or similar |

---

### batch-edit output

```bash
uv run spreadsheet-tools batch-edit "file.xlsm" --sheet "Sheet1" \
  --edits-json '[{"cell":"D325","value":"Diseño del logo"},{"cell":"L325","value":300000}]'
```

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "dry_run": false,
  "applied": [
    {
      "cell": "D325",
      "resolved_address": "D325",
      "value": "Diseño del logo",
      "cleared": false,
      "style_changed": false,
      "warnings": []
    },
    {
      "cell": "L325",
      "resolved_address": "L325",
      "value": 300000,
      "cleared": false,
      "style_changed": false,
      "warnings": []
    }
  ],
  "skipped": [],
  "total_applied": 2,
  "total_skipped": 0,
  "saved": true
}
```

Dry-run returns the same shape with `"dry_run": true` and `"saved": false`.

**Edits JSON format:**
```json
[
  {"cell": "D325", "value": "Description text"},
  {"cell": "L325", "value": 300000},
  {"cell": "B326", "clear": true},
  {"cell": "D326", "value": "Text", "style": {"alignment": {"wrap_text": true}}}
]
```

**Notes:**
- All addresses are validated upfront. If any is invalid, no edits are applied.
- Slave cells are silently redirected to their merge master (warning added to that edit's `warnings`).
- Use `batch-edit` instead of repeated `edit-cell` calls when filling 3+ cells.

---

### snapshot output

```bash
uv run spreadsheet-tools snapshot "file.xlsm" \
  --sheet "Sheet1" --tag "before-edit"
```

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "tag": "before-edit",
  "description": null,
  "cell_count": 247,
  "snapshot_path": "/home/user/.cache/spreadsheet-tools/snapshots/8e780928b4726080/Sheet1/before-edit.json",
  "created_at": "2026-05-24T14:30:00"
}
```

Stored internally as:
```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "tag": "before-edit",
  "created_at": "2026-05-24T14:30:00",
  "cells": {
    "B3": "Conjunto deportivo adaptativo",
    "F308": 85000,
    "H308": 130000
  }
}
```

---

### snapshot-diff output

```bash
uv run spreadsheet-tools snapshot-diff "file.xlsm" \
  --sheet "Sheet1" --tag-a "before-edit" --tag-b "after-edit"
```

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "tag_a": "before-edit",
  "tag_b": "after-edit",
  "changes": [
    {"address": "D325", "before": null,    "after": "Diseño del logo"},
    {"address": "L325", "before": null,    "after": 300000},
    {"address": "F308", "before": 85000,   "after": 90000}
  ],
  "added_cells": ["D325", "L325"],
  "removed_cells": [],
  "modified_cells": ["F308"],
  "total_changes": 3
}
```

**Notes:**
- `added_cells`: present in `tag_b` but absent in `tag_a`
- `removed_cells`: present in `tag_a` but absent in `tag_b`
- `modified_cells`: present in both but with different values

---

### list-snapshots output

```bash
uv run spreadsheet-tools list-snapshots "file.xlsm" [--sheet "Sheet1"]
```

```json
{
  "file": "workbook.xlsm",
  "snapshots": [
    {
      "sheet": "Sheet1",
      "tag": "before-edit",
      "created_at": "2026-05-24T14:30:00",
      "description": null,
      "cell_count": 247,
      "path": "/home/user/.cache/spreadsheet-tools/snapshots/8e780928b4726080/Sheet1/before-edit.json"
    },
    {
      "sheet": "Sheet1",
      "tag": "after-edit",
      "created_at": "2026-05-24T15:45:00",
      "description": "After filling section 2.9",
      "cell_count": 291,
      "path": "/home/user/.cache/spreadsheet-tools/snapshots/8e780928b4726080/Sheet1/after-edit.json"
    }
  ],
  "snapshot_count": 2
}
```

---

## Error Handling

All commands exit code `1` and print to stderr for:
- Missing file: `error: No such file: file.xlsm`
- Invalid sheet: `error: Sheet 'BadSheet' not found. Available: Sheet1, Sheet2`
- Invalid cell address: `error: Invalid cell address: 'ZZZ'`
- Invalid JSON in `--style-json` or `--edits-json`: `error: Invalid JSON: ...`
- Rule parse failure: `error: Cannot parse rule 'bad-rule': ...`
- Snapshot not found: `error: Snapshot 'nonexistent' not found for sheet 'Sheet1'`

## Supported Formats

- `.xlsx`
- `.xlsm` (VBA macros preserved via `keep_vba=True`)

Legacy `.xls` is not supported.

---

## Structure & Formatting Commands

### create-empty-workbook output

```bash
uv run spreadsheet-tools create-empty-workbook "report.xlsx" --sheet "Data"
```

```json
{
  "file": "workspace/report.xlsx",
  "sheet": "Data",
  "created": true
}
```

**Notes:**
- Bare filenames (no directory) are auto-placed in `workspace/` (created if absent).
- Explicit paths like `"subdir/file.xlsx"` are used as-is.

---

### merge-cells output

```bash
uv run spreadsheet-tools merge-cells "file.xlsx" --sheet "Sheet1" --range A1:E1
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "merged": "A1:E1",
  "master": "A1",
  "saved": true
}
```

**Notes:**
- Master cell (top-left) retains its value and style.
- Slave cells (B1:E1) are cleared by Excel on open.
- Verify result with `sheet-info` → `merged_ranges`.

---

### unmerge-cells output

```bash
uv run spreadsheet-tools unmerge-cells "file.xlsx" --sheet "Sheet1" --range A1:E1
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "unmerged": "A1:E1",
  "saved": true
}
```

---

### set-column-width output

```bash
uv run spreadsheet-tools set-column-width "file.xlsx" --sheet "Sheet1" --col B --width 20
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "column": "B",
  "width": 20.0,
  "saved": true
}
```

---

### set-row-height output

```bash
uv run spreadsheet-tools set-row-height "file.xlsx" --sheet "Sheet1" --row 1 --height 40
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "row": 1,
  "height": 40.0,
  "saved": true
}
```

**Notes:**
- `--row` is **1-based** (Excel row number, NOT zero-based).
- Height is in points (Excel default is 15).

---

### batch-set-dimensions output

```bash
uv run spreadsheet-tools batch-set-dimensions "file.xlsx" \
  --sheet "Sheet1" \
  --columns-json '[{"col":"A","width":8},{"col":"B","width":15}]' \
  --rows-json '[{"row":1,"height":40},{"row":8,"height":25}]'
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "column_widths": [
    {"col": "A", "width": 8.0},
    {"col": "B", "width": 15.0}
  ],
  "row_heights": [
    {"row": 1, "height": 40.0},
    {"row": 8, "height": 25.0}
  ],
  "saved": true
}
```

**Notes:**
- `row` in `--rows-json` is **1-based**.
- At least one of `--columns-json` or `--rows-json` is required.
- One save for all changes — use this instead of repeated `set-column-width` calls.

---

### freeze-panes output

```bash
uv run spreadsheet-tools freeze-panes "file.xlsx" --sheet "Sheet1" --cell B2
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "freeze_panes": "B2",
  "saved": true
}
```

**Notes:**
- `B2` freezes row 1 (above B2) and column A (left of B2).
- `A2` freezes only row 1.
- `B1` freezes only column A.

---

### unfreeze-panes output

```bash
uv run spreadsheet-tools unfreeze-panes "file.xlsx" --sheet "Sheet1"
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "freeze_panes": null,
  "saved": true
}
```

---

### add-sheet output

```bash
uv run spreadsheet-tools add-sheet "file.xlsx" --sheet "Summary" --position 0
```

```json
{
  "file": "file.xlsx",
  "sheet": "Summary",
  "index": 0,
  "added": true,
  "saved": true
}
```

**Notes:**
- Fails with error if sheet name already exists.
- `--position` is zero-based; omit to append at end.

---

### rename-sheet output

```bash
uv run spreadsheet-tools rename-sheet "file.xlsx" --old-name "Sheet1" --new-name "Amortización"
```

```json
{
  "file": "file.xlsx",
  "old_name": "Sheet1",
  "new_name": "Amortización",
  "saved": true
}
```

---

### format-range output

```bash
uv run spreadsheet-tools format-range "file.xlsx" --sheet "Sheet1" \
  --range A8:E8 \
  --style-json '{"font":{"bold":true,"color":"FFFFFF"},"fill":{"fill_type":"solid","start_color":"1F4E79"}}'
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "range": "A8:E8",
  "cells_styled": 5,
  "saved": true
}
```

**Notes:**
- Only master cells are styled; slave cells are silently skipped.
- `cells_styled` reports exactly how many cells received the style.
- `--style-json` uses the same format as `edit-cell --style-json`.

---

### set-tab-color output

```bash
uv run spreadsheet-tools set-tab-color "file.xlsx" --sheet "Sheet1" --color "1F4E79"
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "tab_color": "FF1F4E79",
  "saved": true
}
```

**Notes:**
- Color stored as ARGB (8 chars): 6-char RGB input is auto-prefixed with `FF`.
- Leading `#` is stripped automatically.

---

### auto-filter output

#### Enable

```bash
uv run spreadsheet-tools auto-filter "file.xlsx" --sheet "Sheet1" --range A8:E8
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "auto_filter": "A8:E8",
  "saved": true
}
```

#### Clear (omit --range)

```bash
uv run spreadsheet-tools auto-filter "file.xlsx" --sheet "Sheet1"
```

```json
{
  "file": "file.xlsx",
  "sheet": "Sheet1",
  "auto_filter": null,
  "saved": true
}
```

---

## Error Handling (new commands)

All new commands follow the same exit-code-1 + stderr pattern:
- Invalid range: `error: Invalid cell range: 'A1'. Expected A1:B2 format.`
- Invalid column: `error: Invalid column letter: '1A'`
- Row < 1: `error: Row number must be >= 1, got 0`
- Width/height ≤ 0: `error: Column width must be > 0, got 0`
- Sheet exists on add: `error: Sheet 'Name' already exists`
- Sheet not found on rename: `error: Sheet 'X' not found. Available: Sheet1, Sheet2`
- Color invalid: `error: Color must be a 6-char RGB hex (e.g. '1F4E79'), got 'ZZZ'`
