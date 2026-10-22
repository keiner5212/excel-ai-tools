# Spreadsheet Tools Reference

Result shapes from the MCP tools, encoded as TOON. Call the tools. Do not parse the result as JSON.

`sheet` omitted means the active sheet. Rows noted as zero-based use Excel row 1 = 0. Cell addresses and `row_heights` use Excel's 1-based numbers.

A tool error is a failed MCP result. The message names the missing file, unknown sheet, bad address, or bad rule. The workbook on disk is unchanged when a save fails validation.

## list_sheets

```text
file: /path/book.xlsx
active_sheet: Datos
sheet_count: 2
sheets[2]{name,index,state}:
  Datos,0,visible
  Resumen,1,visible
```

`sheets` is a list of objects. Use `name`, not the object itself, as the `sheet` argument of later calls.

## workbook_info

```text
file: /path/book.xlsx
format: xlsx
sheet_count: 2
sheet_names[2]: Datos,Resumen
active_sheet: Datos
properties:
  title: null
  creator: null
  created: null
  modified: null
  subject: null
  description: null
```

`format` is `xlsm` when the path ends in `.xlsm`.

## sheet_info

`dimensions.min_row` and `dimensions.max_row` are 1-based Excel rows. `max_row` at the top level is the same 1-based number.

```text
file: /path/book.xlsx
sheet: Datos
dimensions:
  min_row: 1
  max_row: 11
  min_col: A
  max_col: G
max_row: 11
max_column: 7
max_column_letter: G
merged_ranges[1]: "A1:C1"
freeze_panes: A2
auto_filter: "A1:G1"
```

`freeze_panes` and `auto_filter` are null when unset.

## read_range

Arguments: `file`, `from_col`, `to_col`, `from_row`, `to_row`, optional `sheet`, `include_empty_rows`, `keep_trailing_empty`, `include_formulas`.

`from_row` and `to_row` are zero-based and inclusive.

```text
file: /path/book.xlsx
sheet: Datos
range:
  from_col: A
  to_col: G
  from_row: 0
  to_row: 1
row_count: 2
rows[1]:
  - row: 0
    cells[2]:
      - column: A
        address: A1
        value: ID
      - column: B
        address: B1
        value: Producto
        merge:
          range: "B1:C1"
          role: master
          master: B1
merged_ranges[1]: "B1:C1"
notes: []
```

`row` matches the zero-based index. A slave cell has `"role": "slave"` and `"value": null`. `include_formulas: true` puts formula text in `value`. Otherwise `value` is the cached result, or null if Excel has not calculated it.

## read_cell

Arguments: `file`, `cell`, optional `sheet`. No formula flag.

```text
file: /path/book.xlsx
sheet: Datos
address: G2
value: null
```

## cell_style

Only non-default fields are present.

```text
address: F2
value: 12.5
data_type: n
number_format: "#,##0.00"
font:
  bold: true
  color: FF000000
fill:
  fill_type: solid
  start_color: FFFFFF00
alignment:
  horizontal: center
  wrap_text: true
comment: Check source
hyperlink: "https://example.com"
```

## find

```text
file: /path/book.xlsx
query: teclado
match_count: 1
truncated: false
matches[1]{sheet,address,value}:
  Datos,B4,Teclado
```

## find_replace

Find-only (`replace_with` omitted) does not save. The count field is `total`.

```text
file: /path/book.xlsx
query: "85000"
total: 1
truncated: false
matches[1]{sheet,address,value}:
  Datos,F8,85000
```

A replace adds `saved`. `dry_run: true` sets `saved` false.

## audit_range

The row list is `audit`. Summary keys end in `_cells`.

```text
file: /path/book.xlsx
sheet: Datos
range:
  from_col: A
  to_col: C
  from_row: 1
  to_row: 2
audit[1]:
  - row: 1
    cells[1]{address,value,is_master,is_empty}:
      A2,1,true,false
summary:
  total_rows: 2
  total_master_cells: 6
  empty_master_cells: 1
  filled_master_cells: 5
```

## section_map

Only headers matching `N.N.N Title` in column B, then A. `header_row` and `row_range` are zero-based.

```text
file: /path/book.xlsx
sheet: Plan
scanned_rows:
  min: 0
  max: 400
section_count: 1
sections[1]{header,prefix,title,depth,header_row,column,row_range{from,to}}:
  2.9.1. Producto,2.9.1,Producto,3,270,B,270,300
```

## describe_section

Spanish strategy table only. `data_rows` is `"start-end"`, zero-based.

`row` is zero-based. `name`, `description`, and `cost` are `{address, value}` objects, not bare strings. The column names sit under `section`, not at the top.

```text
file: /path/book.xlsx
sheet: Plan
section:
  header_row: null
  data_rows: 10-12
  name_col: B
  desc_col: D
  cost_col: L
rows[1]:
  - row: 10
    name:
      address: B11
      value: Logo
    description:
      address: D11
      value: Logo de la marca
    cost:
      address: L11
      value: 300000
    issues: []
summary:
  total: 1
  errors: 0
  warnings: 0
  ok: 1
```

Each issue has `severity`, `type`, and `detail`. `type` is `generic_name`, `missing_description`, `name_desc_mismatch`, or `missing_cost`. Summary counts are `errors` and `warnings`, not `with_errors`.

## validate

`rules` is a list of strings. Each result has `rule`, `passed`, `message`, and `addresses`. There is no separate `type` field. `summary` has `total`, `passed`, `failed`.

## edit_cell and batch_edit

`batch_edit` returns `applied`, `total_applied`, `dry_run`, and `saved`. Each applied item has `address` (what you sent), `resolved_address`, `value`, `cleared`, and `style_changed`. A slave write adds a singular `warning` string on that item. Duplicate addresses add a top-level `warnings` list. There is no `skipped` list. If any address is invalid, nothing is saved.

`dry_run: true` returns the same shape with `saved: false`.

## format_range

```text
file: /path/book.xlsx
sheet: Datos
range: "A1:G1"
cells_styled: 7
saved: true
```

`cells_styled` counts master cells that received the style, empty or not.

## create_empty_workbook

```text
file: /path/book.xlsx
sheet: Datos
created: true
```

`file` is the resolved path. A bare name becomes `workspace/<name>` next to the client working directory.

## Structure tools

Each returns `saved: true` plus the field that changed:

| Tool | Extra fields |
| --- | --- |
| `merge_cells` | `merged`, `master` |
| `unmerge_cells` | `unmerged` |
| `set_column_width` | `column`, `width` |
| `set_row_height` | `row` (1-based), `height` |
| `batch_set_dimensions` | `column_widths`, `row_heights` |
| `freeze_panes` | `freeze_panes` |
| `unfreeze_panes` | `freeze_panes: null` |
| `add_sheet` | `sheet`, `index`, `added` |
| `rename_sheet` | `old_name`, `new_name` |
| `copy_sheet` | `source_sheet`, `target_sheet` |
| `set_tab_color` | `tab_color` (8-char ARGB, `FF` plus the RGB you passed) |
| `auto_filter` | `auto_filter` (the range, or null when cleared) |

## snapshot

Stored at `~/.cache/spreadsheet-tools/snapshots/<file-hash>/<sheet>/<tag>.json`. Not inside the workbook. Values are cached results, so a new formula may be null.

```text
file: /path/book.xlsx
sheet: Datos
tag: before
description: null
created_at: "2026-10-22T06:00:00+00:00"
row_count: 11
snapshot_path: ~/.cache/spreadsheet-tools/snapshots/<hash>/Datos/before.json
```

## snapshot_diff

```text
file: /path/book.xlsx
sheet: Datos
tag_a: before
tag_b: after
created_a: "2026-10-22T06:00:00+00:00"
created_b: "2026-10-22T06:05:00+00:00"
changes[1]{address,before,after}:
  D2,null,Teclado
added_rows: []
removed_rows: []
total_changes: 1
```

`added_rows` and `removed_rows` are Excel row numbers as strings. There is no `added_cells` list.

## list_snapshots

```text
file: /path/book.xlsx
sheet: Datos
count: 1
snapshots[1]{tag,description,created_at,row_count}:
  before,null,"2026-10-22T06:00:00+00:00",11
```

The count field is `count`. An unreadable file is included with `"unreadable": true`.

## Formats

`.xlsx` and `.xlsm`. `.xls` is not supported.
