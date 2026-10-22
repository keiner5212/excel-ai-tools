---
name: spreadsheet-tools
description: >-
  Read, search, and edit Excel (.xlsx/.xlsm) workbooks through the
  spreadsheet-tools MCP tools. Use for spreadsheets, Excel files, .xlsx,
  .xlsm, sheet ranges, cell edits, formulas, formatting, or workbook inspection.
---

# Spreadsheet Tools

Call the `spreadsheet-tools` MCP tools. Do not write Python, shell, or scripts to read or edit a workbook. Do not import a spreadsheet library. If the MCP server is not available, stop and say so. Do not look for another way to edit the file.

## Results are TOON

Every tool returns TOON text, not JSON. Do not `json.loads` the result. Read the keys and the rows as written.

A key on its own line is `name: value`. `null` is an empty value. A list of objects with the same keys is a table: the header names the columns once, and each following line is one row. A comma inside a value is quoted.

```text
file: /path/book.xlsx
sheet: Datos
row_count: 1
rows[1]:
  - row: 0
    cells[3]{column,address,value}:
      A,A1,"Name, last"
      B,B1,=E2*F2
      C,C1,null
```

That row is Excel row 1 (`row: 0`). Column B holds the formula text `=E2*F2`. Column C is empty. Field names are the same ones listed in [reference.md](reference.md).

## Paths

Every tool that takes `file` accepts any path.

- Absolute path or `~/...`: that file, wherever it is.
- Relative path with a directory (`reports/q1.xlsx`): relative to the client working directory.
- Bare filename (`report.xlsx`, no `/`): `workspace/` next to the client working directory.

`create_empty_workbook` uses the argument name `filename`, not `file`. It creates missing parent directories. Same path rules.

`sheet` omitted means the workbook's active sheet. Read `active_sheet` from `list_sheets` or `workbook_info` before assuming a name. `Sheet1` is only the default on a workbook this server just created.

## Indexing

Two different row conventions. Mixing them writes the wrong row.

| Call | Rows | Example |
| --- | --- | --- |
| `read_range`, `audit_range`, `section_map`, `describe_section` `data_rows` | zero-based, inclusive | Excel row 1 is `0`. Excel rows 2-11 are `from_row=1`, `to_row=10`. |
| `set_row_height`, `batch_set_dimensions` `row_heights` | 1-based Excel row | Excel row 1 is `row: 1`. |
| Cell addresses (`A1`, `B2`) | 1-based, as in Excel | Header is `A1`, first data row is `A2`. |
| `freeze_panes` `cell` | 1-based address | `A2` freezes row 1. `B2` freezes row 1 and column A. |

Columns are always letters (`A`, `B`, `AA`). Never a number.

`read_range` drops fully blank rows and trailing empty columns unless `include_empty_rows` or `keep_trailing_empty` is true. `row_count` is the number of rows returned, not the Excel span.

## Values and formulas

Pass numbers as numbers and text as strings. `300` is a number. `"300"` is text and will not sum.

A string that starts with `=` is stored as a formula: `"=E2*F2"`. Use the real Excel row in the formula, not the zero-based index.

Reads return the cached calculated value. A formula written in this session has no cache yet, so a later read returns null until Excel opens the file and calculates it. `read_range(..., include_formulas=true)` returns the formula text (`=E2*F2`) instead. `read_cell` has no formula flag. It always returns the cached value.

Dates come back as ISO strings.

## Writes

Prefer one `batch_edit` over many `edit_cell` calls. One save, and a bad address fails the whole batch before anything is written.

`batch_edit` edit object:

```json
{"cell": "D5", "value": "Hello"}
{"cell": "E2", "value": "=C2*D2"}
{"cell": "F2", "value": 12.5, "style": {"number_format": "#,##0.00"}}
{"cell": "G2", "clear": true}
```

`value`, `clear`, and `style` are optional. At least one is required on `edit_cell`. `dry_run: true` does not save.

Prefer one `batch_set_dimensions` over repeated width or height calls.

```json
{
  "column_widths": [{"col": "A", "width": 18}, {"col": "B", "width": 32}],
  "row_heights": [{"row": 1, "height": 22}]
}
```

Width is Excel character units. Height is points. Both must be greater than 0.

Saves replace the file only after the package validates. A failed save leaves the previous file bytes in place.

Merged cells: the top-left cell holds the value. Other cells in the merge read as null and are not write targets. A write to a slave cell is redirected to the master, and the result includes a warning. Call `sheet_info` and use `merged_ranges` before editing a sheet that may contain merges.

`format_range` styles master cells only. `cells_styled` is how many cells received the style, including cells that were empty. It does not mean those cells have values.

## Style

Same object for `edit_cell`, each `batch_edit` item, and `format_range`. Omitted keys stay as they are.

```json
{
  "font": {"name": "Calibri", "size": 12, "bold": true, "italic": false, "color": "FFFFFF"},
  "fill": {"fill_type": "solid", "start_color": "1F4E79"},
  "alignment": {"horizontal": "center", "vertical": "center", "wrap_text": true},
  "number_format": "#,##0.00",
  "comment": "Check source",
  "hyperlink": "https://example.com"
}
```

Colors are hex RGB. `FF0000` is red. A leading `#` is stripped. 3-digit shorthand is expanded. `set_tab_color` takes the same kind of hex (`1F4E79`). The stored tab color is 8 characters with `FF` in front.

Useful `number_format` values: `0`, `0.00`, `#,##0`, `#,##0.00`, `yyyy-mm-dd`.

## Tools that are not general

Do not use these on a normal table:

- `section_map` only finds headers shaped like `1.2.3 Title` in column B, then column A. `header_row` and `row_range` are zero-based. A sales table with a plain header row returns no sections. Use `sheet_info` and `read_range` instead.
- `describe_section` only audits a Spanish strategy table: a name column, a description column, and an optional cost column. It flags generic names like `Estrategia de Producto 2`, empty descriptions, name/description keyword overlap under 30%, and empty costs. `data_rows` is a zero-based `"start-end"` string. Default columns are name `B`, description `D`, cost `L`. Pass `cost_col: ""` to skip cost.
- `validate` rules are the same strategy-table checks, plus generic emptiness and numeric bounds. Rule strings:

| Rule | Meaning |
| --- | --- |
| `not-empty:D5` | That cell is non-blank. |
| `not-empty-range:D5:D14` | Every master cell in the rectangle is non-blank. |
| `numeric-range:L10:0:1000000` | Number is between the two bounds, inclusive. |
| `string-contains:D5:Descuento` | Case-insensitive substring. |
| `price-min-max:F10:H10` | First cell's number is less than the second. |
| `name-matches-desc:B10:D10` | Keyword overlap at least 0.30. |
| `no-generic-name:B10` | Value is not an `Estrategia de X N` placeholder. |

## Other tools

`list_sheets` returns `active_sheet`, `sheet_count`, and `sheets`: a list of `{name, index, state}`. `state` is `visible`, `hidden`, or `veryHidden`. It is not a list of name strings.

`workbook_info` returns `format` (`xlsx` or `xlsm`), `sheet_names` (strings), `active_sheet`, and `properties` (`title`, `creator`, `created`, `modified`, `subject`, `description`).

`sheet_info` returns `dimensions` (Excel `min_row` / `max_row` are 1-based), `max_column_letter`, `merged_ranges`, `freeze_panes`, `auto_filter`.

`find` searches cell text. `sheet` omitted searches every sheet. Default `max_results` is 50. Match is a substring of the cell's text form. Result key is `match_count`.

`find_replace` without `replace_with` does not write. Without `use_regex`, the whole cell text must equal `query` (numbers match their text form, so `85000` matches a numeric cell). With `use_regex`, it is a substring substitution. `replace_with` as a number stays a number. Result key is `total`, not `match_count`. `dry_run: true` previews and does not save.

`audit_range` lists master cells and whether each is empty. Rows are zero-based. Result list is `audit`, not `cells`. Summary keys are `total_master_cells`, `empty_master_cells`, `filled_master_cells`. `show_slaves: true` includes non-writable cells.

`read_cell` returns `address` and `value` only.

`add_sheet` fails if the name exists. `position` is a zero-based index. Omit it to append.

`rename_sheet` takes `old_name` and `new_name`.

`copy_sheet` copies inside the same workbook. `overwrite` defaults to false.

`auto_filter` with `cell_range` (the header row, such as `A1:G1`) turns the filter on. Omit `cell_range` to clear it.

`merge_cells` / `unmerge_cells` take `cell_range` like `A1:E1`.

## Snapshots

`snapshot` writes JSON under `~/.cache/spreadsheet-tools/snapshots/`, keyed by file path and sheet. Nothing is stored inside the xlsx. The snapshot stores cached values (`data_only`), so a formula just written may be null there. `list_snapshots` returns `count` and `snapshots` with `tag`, `description`, `created_at`, `row_count`. `snapshot_diff` compares two tags on one sheet. Changes are per address. `added_rows` and `removed_rows` are Excel row numbers as strings. It does not return `added_cells`.

Take `tag="before"` before a large edit if you need a diff afterward.

## Usual order

New workbook:

1. `create_empty_workbook` with the full path and the first sheet name.
2. `add_sheet` for any extra sheet.
3. `batch_edit` for headers, values, and formulas.
4. `format_range` for header fill, font, and number formats.
5. `batch_set_dimensions` for widths and the header height.
6. `freeze_panes` and `auto_filter` when the sheet is a table.
7. `set_tab_color` if a tab color was requested.
8. `read_range` with `include_formulas=true` on the formula columns to confirm the text. Do not expect calculated numbers yet.

Existing workbook:

1. `list_sheets`.
2. `sheet_info` on the sheet you will change.
3. `read_range` on the rectangle you need. Widen it if `row_count` looks short because blank rows were trimmed.
4. `batch_edit` for the changes.
5. `read_range` again on the edited addresses.

Result fields, encoded as TOON, are in [reference.md](reference.md).
