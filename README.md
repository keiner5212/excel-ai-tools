# spreadsheet-tools

[![Python](https://img.shields.io/badge/python-%3E%3D3.13-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![uv](https://img.shields.io/badge/managed%20by-uv-purple)](https://github.com/astral-sh/uv)

AI-friendly Excel (`.xlsx` / `.xlsm`) tools. Exposes 31 operations as **MCP tools** (primary) and as a **CLI** (fallback). Designed to give language models clean, token-efficient access to spreadsheet data — value-only reads, explicit style metadata, and safe cell edits with ZIP-merge save (preserves images, drawings, VBA macros).

## For AI Agents

**Use the MCP server — no shell scripting needed.**

When the `spreadsheet-tools` MCP server is connected, call tools directly:

```python
# Create a workbook at any path
create_empty_workbook(filename="/tmp/report.xlsx", sheet="Data")

# Populate cells atomically (preferred over repeated single edits)
batch_edit(
    file="/tmp/report.xlsx",
    sheet="Data",
    edits=[
        {"cell": "A1", "value": "Name", "style": {"font": {"bold": True}}},
        {"cell": "B1", "value": "Amount"},
        {"cell": "A2", "value": "Acme Corp"},
        {"cell": "B2", "value": 150000},
    ]
)

# Style header row
format_range(file="/tmp/report.xlsx", sheet="Data", cell_range="A1:B1",
             style={"fill": {"fill_type": "solid", "start_color": "1F4E79"},
                    "font": {"color": "FFFFFF", "bold": True}})

# Freeze header and read back to verify
freeze_panes(file="/tmp/report.xlsx", sheet="Data", cell="A2")
read_range(file="/tmp/report.xlsx", sheet="Data",
           from_col="A", to_col="B", from_row=0, to_row=5)
```

See `SKILL.md` files for the complete tool reference.

## Requirements

- Python ≥ 3.13
- [`uv`](https://github.com/astral-sh/uv)

## Setup

```bash
git clone https://github.com/keinerdeveloper/spreadsheet-tools.git
cd spreadsheet-tools
uv sync
```

## MCP Integration

Client-specific config files are included in the repo and work out of the box
when you open this project:

| Client | Config file |
|--------|-------------|
| **Cursor** | `.cursor/mcp.json` |
| **Claude Code** | `.mcp.json` |
| **OpenCode** | `opencode.json` → `mcp` key |

For setup in other projects or global configuration, see **[docs/mcp-setup.md](docs/mcp-setup.md)**.

That guide covers: Cursor, Claude Code, Claude Desktop, OpenCode, and any generic MCP client.

## CLI Usage (Fallback)

When MCP is not available, every operation is accessible via CLI. Each tool maps
1:1 to a command (underscores → hyphens):

```bash
# List sheets
uv run spreadsheet-tools list-sheets "workbook.xlsm"

# Read a range (rows are zero-based)
uv run spreadsheet-tools read-range "workbook.xlsm" \
  --sheet "Sheet1" --from-col A --to-col L --from-row 0 --to-row 20

# Edit a single cell
uv run spreadsheet-tools edit-cell "workbook.xlsm" --sheet "Sheet1" --cell B5 --value "Hello"

# Bulk edit (atomic, single save)
uv run spreadsheet-tools batch-edit "workbook.xlsm" --sheet "Sheet1" \
  --edits-json '[{"cell":"D5","value":"Hello"},{"cell":"E5","value":42}]'

# Create new workbook (bare filename → workspace/)
uv run spreadsheet-tools create-empty-workbook "report.xlsx" --sheet "Data"

# Search
uv run spreadsheet-tools find "workbook.xlsm" --query "total" --max-results 20

# Workbook / sheet metadata
uv run spreadsheet-tools workbook-info "workbook.xlsm"
uv run spreadsheet-tools sheet-info "workbook.xlsm" --sheet "Sheet1"
```

Full command list: `uv run spreadsheet-tools --help`

## AI Agent Integration

Skill files for major AI coding tools:

| Tool | Skill path |
|------|-----------|
| Cursor | `.cursor/skills/spreadsheet-tools/SKILL.md` |
| Claude Code | `.claude/skills/spreadsheet-tools/SKILL.md` |
| OpenCode | `.opencode/skills/spreadsheet-tools/SKILL.md` |

## Available Tools (31 total)

| Category | Tools |
|----------|-------|
| Discovery | `list_sheets`, `workbook_info`, `sheet_info` |
| Reads | `read_range`, `read_cell`, `cell_style` |
| Search | `find`, `find_replace` |
| Analysis | `section_map`, `audit_range`, `describe_section`, `validate` |
| Write (single) | `edit_cell` |
| Write (bulk) | `batch_edit`, `batch_set_dimensions` |
| Formatting | `merge_cells`, `unmerge_cells`, `set_column_width`, `set_row_height`, `format_range`, `freeze_panes`, `unfreeze_panes`, `set_tab_color`, `auto_filter` |
| Sheet mgmt | `add_sheet`, `rename_sheet`, `copy_sheet` |
| Create | `create_empty_workbook` |
| Snapshots | `snapshot`, `snapshot_diff`, `list_snapshots` |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT © [Keiner Alvarado](https://github.com/keinerdeveloper)
