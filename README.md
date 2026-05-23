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

See skill docs in:
- `.cursor/skills/spreadsheet-tools/SKILL.md` (Cursor)
- `.claude/skills/spreadsheet-tools/SKILL.md` (Claude Code)
- `.opencode/skills/spreadsheet-tools/SKILL.md` (OpenCode)
