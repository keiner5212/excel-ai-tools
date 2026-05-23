# Spreadsheet Tools Reference

## read-range output

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
        {"column": "A", "address": "A1", "value": "Header"}
      ]
    }
  ],
  "merged_ranges": ["A1:C1"],
  "notes": ["Values only; formatting metadata excluded."]
}
```

## cell-style output

Only non-default style fields are returned when possible.

```json
{
  "address": "B3",
  "value": 1234.5,
  "data_type": "n",
  "number_format": "#,##0.00",
  "font": {"bold": true, "color": "FF000000"},
  "fill": {"fill_type": "solid", "start_color": "FFFFFF00"},
  "alignment": {"horizontal": "center"},
  "comment": "Check source",
  "hyperlink": "https://example.com"
}
```

## edit-cell result

```json
{
  "file": "workbook.xlsm",
  "sheet": "Sheet1",
  "address": "B3",
  "updated": {"value": "New text", "style_changed": false},
  "saved": true
}
```

## Error handling

Commands print `error: <message>` to stderr and exit with code `1` for:
- missing files
- invalid sheet/column/cell names
- invalid JSON in `--style-json`

## Supported formats

- `.xlsx`
- `.xlsm` (VBA macros preserved on save via `keep_vba=True`)

Legacy `.xls` is not supported.
