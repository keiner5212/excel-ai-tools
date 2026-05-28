---
name: spreadsheet-tools
description: >-
  Read, search, and edit Excel (.xlsx/.xlsm) workbooks. Use when working with
  spreadsheets, Excel files, .xlsx, .xlsm, sheet ranges, cell edits, or workbook
  inspection.
---

# Spreadsheet Tools

## How to Use — MCP (Primary)

If the `spreadsheet-tools` MCP server is active in your environment, call its tools
directly. **No shell commands. No Python scripts. No openpyxl imports.**

```
# Example: create a workbook at any path and populate it
result = create_empty_workbook(filename="/tmp/report.xlsx", sheet="Data")
batch_edit(file="/tmp/report.xlsx", sheet="Data", edits=[
  {"cell": "A1", "value": "Name"},
  {"cell": "B1", "value": "Amount"},
  {"cell": "A2", "value": "Acme", "style": {"font": {"bold": true}}}
])
```

Pass any absolute or relative path to `file`. Bare filenames (no `/`) go to
`workspace/` automatically, but that's just a convention — use whatever path works
for your context.

## How to Use — CLI (Fallback)

When MCP is not available, run from the repo root:

```bash
export PATH="$HOME/.local/bin:$PATH"
uv sync
uv run spreadsheet-tools <command> [args...]
```

---

## CRITICAL RULES (NO EXCEPTIONS)

1. **NEVER create Python scripts or import openpyxl yourself.**
2. **NEVER use `uv run python -c` for spreadsheet operations.**
3. All reads and writes go through MCP tools or `uv run spreadsheet-tools` only.
4. **Row indices** for `read_range` / `audit_range` / `section_map` are **zero-based**
   (Excel row 1 = index 0). Row numbers for `set_row_height` / `batch_set_dimensions`
   are **1-based** (Excel convention).
5. **Columns** always use Excel letters (`A`, `B`, `C`...). Never numeric indices.
6. **Merged cells**: master cell (top-left) holds value. Slave cells return `null`.
   Call `sheet_info` to see all merged ranges before editing.
7. **New workbooks**: pass any path to `create_empty_workbook`. Bare filenames
   (no `/`) go to `workspace/` automatically, but any absolute or relative path works.
8. **Formula cells**: by default, reads return cached computed values (`None` if no cache).
   Use `--include-formulas` (CLI) or `include_formulas=True` (MCP) to get formula text
   like `=SUM(A1,A2)` instead of the computed result. This is essential when the
   workbook has never been opened in Excel or was saved with recalculation disabled.

---

## All MCP Tools

### Discovery
| Tool | What it does |
|------|-------------|
| `list_sheets(file)` | List all sheet names |
| `workbook_info(file)` | Sheet names, active sheet, defined names |
| `sheet_info(file, sheet?)` | Dimensions, merged ranges, freeze panes, filters |

### Reads
| Tool | What it does |
|------|-------------|
| `read_range(file, from_col, to_col, from_row, to_row, sheet?, include_formulas?)` | Read rectangular range — values only (or formula text with `include_formulas`), zero-based rows |
| `read_cell(file, cell, sheet?, include_formulas?)` | Read single cell value (or formula text with `include_formulas`) |
| `cell_style(file, cell, sheet?)` | Read cell style metadata (font, fill, alignment, etc.) |

### Search
| Tool | What it does |
|------|-------------|
| `find(file, query, sheet?, case_sensitive?, max_results?)` | Search all cell values |
| `find_replace(file, query, replace_with?, sheet?, use_regex?, dry_run?, ...)` | Find and optionally replace |

### Analysis
| Tool | What it does |
|------|-------------|
| `section_map(file, sheet?, min_row?, max_row?)` | Discover numbered section headers and row ranges |
| `audit_range(file, from_col, to_col, from_row, to_row, sheet?, show_slaves?)` | Flag empty master cells in range |
| `describe_section(file, data_rows, sheet?, name_col?, desc_col?, cost_col?)` | Check strategy table name/description consistency |
| `validate(file, rules, sheet?)` | Apply validation rules (not-empty, numeric-range, etc.) |

### Writes — Single Cell
| Tool | What it does |
|------|-------------|
| `edit_cell(file, cell, value?, clear?, style?, sheet?, dry_run?)` | Edit one cell's value and/or style |

### Writes — Bulk (Preferred)
| Tool | What it does |
|------|-------------|
| `batch_edit(file, edits, sheet?, dry_run?)` | Edit multiple cells atomically — single save |
| `batch_set_dimensions(file, column_widths?, row_heights?, sheet?)` | Set multiple widths/heights in one save |

### Structure / Formatting
| Tool | What it does |
|------|-------------|
| `merge_cells(file, cell_range, sheet?)` | Merge a range |
| `unmerge_cells(file, cell_range, sheet?)` | Remove merge |
| `set_column_width(file, col, width, sheet?)` | Set one column width |
| `set_row_height(file, row, height, sheet?)` | Set one row height (1-based row) |
| `format_range(file, cell_range, style, sheet?)` | Apply same style to every master cell in range |
| `freeze_panes(file, cell, sheet?)` | Freeze panes at cell (e.g. "B2" = row 1 + col A) |
| `unfreeze_panes(file, sheet?)` | Remove freeze panes |
| `set_tab_color(file, color, sheet?)` | Set tab color (6-char hex, no #) |
| `auto_filter(file, cell_range?, sheet?)` | Enable or clear auto-filter |

### Sheet Management
| Tool | What it does |
|------|-------------|
| `add_sheet(file, sheet, position?)` | Add empty sheet |
| `rename_sheet(file, old_name, new_name)` | Rename sheet |
| `copy_sheet(file, source_sheet, target_sheet, overwrite?)` | Duplicate sheet |

### Workbook Creation
| Tool | What it does |
|------|-------------|
| `create_empty_workbook(filename, sheet?)` | Create new .xlsx. Bare filename → workspace/. |

### Snapshots
| Tool | What it does |
|------|-------------|
| `snapshot(file, tag, sheet?, description?)` | Capture cell values to named snapshot |
| `snapshot_diff(file, tag_a, tag_b, sheet?)` | Compare two snapshots |
| `list_snapshots(file, sheet?)` | List saved snapshots |

---

## Style Dict Reference

Used in `edit_cell(style=...)`, `batch_edit` edits `style` key, `format_range(style=...)`:

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

Colors are 6-char hex RGB without `#` (`"FF0000"` = red, `"1F4E79"` = dark blue).

---

## Recommended Workflows

### Exploring and filling an existing form

```
1. list_sheets(file)                          # discover sheet names
2. sheet_info(file, sheet)                    # dimensions + merged ranges
3. section_map(file, sheet)                   # find section headers + row ranges
4. read_range(file, ...)                      # inspect values for a section
5. audit_range(file, ...)                     # find empty master cells
6. snapshot(file, tag="before", sheet)        # safety snapshot
7. batch_edit(file, edits=[...], sheet)       # fill cells
8. validate(file, rules=[...], sheet)         # assert data quality
9. snapshot(file, tag="after", sheet)         # post-edit snapshot
10. snapshot_diff(file, "before", "after")    # verify changes
```

### Building a new workbook from scratch

```
1. create_empty_workbook("/path/to/report.xlsx")   # any path; bare name → workspace/
2. batch_edit("/path/to/report.xlsx", edits)       # populate cells
3. merge_cells(..., cell_range="A1:E1")            # merge header
4. format_range(..., style={...})                  # style header row
5. batch_set_dimensions(..., column_widths)        # set column widths
6. freeze_panes(..., cell="B2")                    # freeze header
7. set_tab_color(..., color="1F4E79")              # color-code tab
```

---

## CLI Equivalent (when MCP unavailable)

Every MCP tool maps 1:1 to a CLI command with the same name (underscores → hyphens):

```bash
uv run spreadsheet-tools list-sheets "file.xlsm"
uv run spreadsheet-tools read-range "file.xlsm" --sheet "Sheet1" --from-col A --to-col L --from-row 0 --to-row 20
uv run spreadsheet-tools batch-edit "file.xlsm" --sheet "Sheet1" --edits-json '[{"cell":"D5","value":"Hello"}]'
uv run spreadsheet-tools create-empty-workbook "report.xlsx" --sheet "Data"
```

Full CLI reference: run `uv run spreadsheet-tools --help` or `uv run spreadsheet-tools <command> --help`.

## Additional Resources

- MCP setup for Cursor, Claude Code, Claude Desktop, OpenCode: [docs/mcp-setup.md](../../docs/mcp-setup.md)
- JSON output shapes for all commands: [reference.md](reference.md)
