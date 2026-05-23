# Spreadsheet Tools

Local Python CLI for AI-friendly Excel (`.xlsx` / `.xlsm`) interaction.

## Setup

```bash
export PATH="$HOME/.local/bin:$PATH"
uv sync
```

## Usage

```bash
uv run spreadsheet-tools list-sheets "workbook.xlsm"
uv run spreadsheet-tools read-range "workbook.xlsm" --sheet "Sheet1" --from-col A --to-col L --from-row 0 --to-row 10
```

See `.cursor/skills/spreadsheet-tools/SKILL.md` for the full command reference.
