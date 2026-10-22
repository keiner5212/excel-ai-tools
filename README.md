# spreadsheet-tools

Version 0.2.0.

[![Python](https://img.shields.io/badge/python-%3E%3D3.13-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Excel (`.xlsx` / `.xlsm`) tools for an agent. 31 operations, exposed as MCP tools. Reads return values, not a dump of the file. Writes go through one save, and a failed save leaves the previous file in place.

## Install

Download the archive for your OS. No Python install. Steps, skill folder, and the MCP `command` path: [setup/INSTALL.md](setup/INSTALL.md).

The download is a folder. Point the agent at the executable inside it. The program does not unpack into `/tmp`.

Copy `setup/spreadsheet-tools/` into the agent's skills directory. The agent does not read it from this repo.

Building from source is in that same install page.

## MCP

`setup/mcp.json` is a template. The path in it is a placeholder, not a file on any one computer. Copy the `spreadsheet-tools` entry into the agent and set `command` to the executable on your machine. The install page lists where that config file goes for each agent.

Paths are normal filesystem paths.

- An absolute path or `~/...` edits that workbook.
- A path with a directory is relative to the client working directory.
- A bare filename (`report.xlsx`) goes to `workspace/` next to the client working directory.
- `create_empty_workbook` creates missing parent directories. Its path argument is `filename`.

`read_range`, `audit_range`, and `section_map` use zero-based rows (Excel row 1 is `0`). Cell addresses, `set_row_height`, and `batch_set_dimensions` row heights use Excel's 1-based numbers. Columns are letters.

A string that starts with `=` is a formula (`"=E2*F2"`). Pass numbers as numbers. `read_range` returns the cached value. A formula just written is null until Excel calculates it. `include_formulas=true` returns the formula text.

`section_map`, `describe_section`, and `validate` are for Spanish strategy tables (numbered headers, name, description, cost). Do not use them on a normal sheet.

Tool list and TOON result shapes: `setup/spreadsheet-tools/SKILL.md` and `setup/spreadsheet-tools/reference.md`. Tool results are TOON text, not JSON.

Example:

```text
create_empty_workbook(filename="/tmp/report.xlsx", sheet="Data")
batch_edit(file="/tmp/report.xlsx", sheet="Data", edits=[
  {"cell": "A1", "value": "Name"},
  {"cell": "B2", "value": 150000},
  {"cell": "C2", "value": "=B2*2"}
])
read_range(file="/tmp/report.xlsx", sheet="Data", from_col="A", to_col="C", from_row=0, to_row=1, include_formulas=true)
```

## CLI

The same executable is the CLI. No arguments in a terminal prints help. An agent that launches it with a pipe starts the MCP server.

```bash
./spreadsheet-tools-linux-x86_64/spreadsheet-tools-linux-x86_64 list-sheets "/path/book.xlsx"
./spreadsheet-tools-linux-x86_64/spreadsheet-tools-linux-x86_64 read-range "/path/book.xlsx" --sheet Data --from-col A --to-col C --from-row 0 --to-row 10
./spreadsheet-tools-linux-x86_64/spreadsheet-tools-linux-x86_64 --help
```

Use the folder name for your OS. Mac and Windows names are in the install page.

## Tools

| Category | Tools |
| --- | --- |
| Discovery | `list_sheets`, `workbook_info`, `sheet_info` |
| Reads | `read_range`, `read_cell`, `cell_style` |
| Search | `find`, `find_replace` |
| Analysis | `section_map`, `audit_range`, `describe_section`, `validate` |
| Write | `edit_cell`, `batch_edit`, `batch_set_dimensions` |
| Formatting | `merge_cells`, `unmerge_cells`, `set_column_width`, `set_row_height`, `format_range`, `freeze_panes`, `unfreeze_panes`, `set_tab_color`, `auto_filter` |
| Sheets | `add_sheet`, `rename_sheet`, `copy_sheet` |
| Create | `create_empty_workbook` |
| Snapshots | `snapshot`, `snapshot_diff`, `list_snapshots` |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT (c) [Keiner Alvarado](https://github.com/keinerdeveloper)
