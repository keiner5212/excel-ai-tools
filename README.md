# spreadsheet-tools

[![Python](https://img.shields.io/badge/python-%3E%3D3.13-blue)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![uv](https://img.shields.io/badge/managed%20by-uv-purple)](https://github.com/astral-sh/uv)

AI-friendly CLI for reading and editing Excel (`.xlsx` / `.xlsm`) workbooks. Designed to give language models clean, token-efficient access to spreadsheet data — value-only reads, explicit style metadata, and safe cell edits.

## Features

- **Read ranges** — cleaned, value-only output with optional formula passthrough
- **Inspect styles** — font, fill, border, number format per cell
- **Edit cells** — value and/or style, with `--dry-run` support
- **Find values** — full-text search across the entire workbook
- **Copy sheets** — duplicate sheet structure within a workbook
- **JSON output** — every command returns structured JSON, ideal for AI pipelines

## Requirements

- Python ≥ 3.13
- [`uv`](https://github.com/astral-sh/uv)

## Installation

```bash
# Clone
git clone https://github.com/keinerdeveloper/spreadsheet-tools.git
cd spreadsheet-tools

# Sync dependencies
uv sync
```

## Usage

### List sheets

```bash
uv run spreadsheet-tools list-sheets "workbook.xlsm"
```

### Read a range

```bash
uv run spreadsheet-tools read-range "workbook.xlsm" \
  --sheet "Sheet1" \
  --from-col A --to-col L \
  --from-row 0 --to-row 20
```

### Read a single cell

```bash
uv run spreadsheet-tools read-cell "workbook.xlsm" --sheet "Sheet1" --cell B5
```

### Inspect cell style

```bash
uv run spreadsheet-tools cell-style "workbook.xlsm" --sheet "Sheet1" --cell B5
```

### Edit a cell

```bash
# Value only
uv run spreadsheet-tools edit-cell "workbook.xlsm" --sheet "Sheet1" --cell B5 --value "Hello"

# Value + style
uv run spreadsheet-tools edit-cell "workbook.xlsm" --sheet "Sheet1" --cell B5 \
  --value "Hello" \
  --style-json '{"font":{"bold":true},"fill":{"fgColor":"FFFF00"}}'

# Dry run (no file write)
uv run spreadsheet-tools edit-cell "workbook.xlsm" --cell B5 --value "Test" --dry-run
```

### Search workbook

```bash
uv run spreadsheet-tools find "workbook.xlsm" --query "total" --max-results 20
```

### Copy a sheet

```bash
uv run spreadsheet-tools copy-sheet "workbook.xlsm" \
  --source-sheet "Template" \
  --target-sheet "January 2026"
```

### Workbook / sheet metadata

```bash
uv run spreadsheet-tools workbook-info "workbook.xlsm"
uv run spreadsheet-tools sheet-info   "workbook.xlsm" --sheet "Sheet1"
```

## AI Agent Integration

Agent skill files are included for major AI coding tools:

| Tool | Skill path |
|------|-----------|
| Cursor | `.cursor/skills/spreadsheet-tools/SKILL.md` |
| Claude Code | `.claude/skills/spreadsheet-tools/SKILL.md` |
| OpenCode | `.opencode/skills/spreadsheet-tools/SKILL.md` |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

MIT © [Keiner Alvarado](https://github.com/keinerdeveloper)
