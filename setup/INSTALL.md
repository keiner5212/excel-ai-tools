# Install

The agent needs two things:

- The program, so it can edit workbooks.
- The skill folder, so it knows to call that program instead of writing Python.

## 1. Pick a binary for your system

Each file already contains Python and the libraries. Nothing is installed globally.

A push to `main` updates the draft release for the version in `pyproject.toml` (now `v0.1.0`): [GitHub releases](https://github.com/keinerdeveloper/spreadsheet-tools/releases). Download the file for your machine:

| Your machine | File |
| --- | --- |
| Linux, 64-bit PC | `spreadsheet-tools-linux-x86_64` |
| Mac, Apple Silicon | `spreadsheet-tools-macos-arm64` |
| Windows, 64-bit | `spreadsheet-tools-windows-x86_64.exe` |

On Linux and Mac, mark it runnable:

```bash
chmod +x spreadsheet-tools-linux-x86_64
```

Use your filename. Put the file anywhere. An absolute path in the config below is enough.

The same file is the CLI:

```bash
./spreadsheet-tools-linux-x86_64 list-sheets "/path/to/book.xlsx"
```

No arguments from a terminal prints help. An agent launching it with a pipe starts the MCP server.

Intel Macs are not in the table. Use the source install below.

## 2. Point the agent at that file

If the client already has an MCP file, add the `spreadsheet-tools` entry inside `mcpServers`. Do not wipe the other servers.

```json
{
  "mcpServers": {
    "spreadsheet-tools": {
      "type": "stdio",
      "command": "/absolute/path/to/spreadsheet-tools-linux-x86_64"
    }
  }
}
```

| Agent | MCP file |
| --- | --- |
| Cursor, every project | `~/.cursor/mcp.json` |
| Cursor, one project | `<project>/.cursor/mcp.json` |
| Claude Code, every project | `~/.claude.json` |
| Claude Code, one project | `<project>/.mcp.json` |
| Claude Desktop (Linux) | `~/.config/Claude/claude_desktop_config.json` |
| OpenCode | `opencode.json` in the project, or `~/.config/opencode/opencode.json` |

OpenCode wants the program as a command list:

```json
{
  "mcp": {
    "spreadsheet-tools": {
      "type": "local",
      "command": ["/absolute/path/to/spreadsheet-tools-linux-x86_64"],
      "enabled": true
    }
  }
}
```

Restart the client after saving.

## 3. Move the skill so the agent can see it

The release also has `spreadsheet-tools-skill.zip`. Unzip it. You get a folder named `spreadsheet-tools` with `SKILL.md` and `reference.md` inside.

That folder is not detected inside this repo or inside the zip. Copy it into the skills directory the agent scans. Keep the folder name `spreadsheet-tools`.

| Agent | Copy the folder here |
| --- | --- |
| Cursor, every project | `~/.cursor/skills/spreadsheet-tools/` |
| Cursor, one project | `<project>/.cursor/skills/spreadsheet-tools/` |
| Claude Code, every project | `~/.claude/skills/spreadsheet-tools/` |
| Claude Code, one project | `<project>/.claude/skills/spreadsheet-tools/` |
| OpenCode, every project | `~/.config/opencode/skills/spreadsheet-tools/` |
| OpenCode, one project | `<project>/.opencode/skills/spreadsheet-tools/` |

From a git checkout the same folder is `setup/spreadsheet-tools`:

```bash
mkdir -p ~/.cursor/skills
cp -a setup/spreadsheet-tools ~/.cursor/skills/
```

Start a new chat after the copy. An already open chat will not pick the skill up.

## 4. Check

In a new chat, ask the agent to list sheets in a real `.xlsx` or `.xlsm`. It should call `list_sheets` instead of writing a Python script.

If the tools are missing:

1. `command` is the absolute path of the binary you downloaded, and that file is executable.
2. The skill folder is in the table above.
3. The chat was started after both of those.

## From source

Use this when you are changing the code, or when no binary matches your machine. You need Python 3.13+ and [uv](https://github.com/astral-sh/uv).

```bash
git clone https://github.com/keinerdeveloper/spreadsheet-tools.git
cd spreadsheet-tools
uv sync
```

`setup/mcp.json` in the repo points at a built binary. For a source checkout, the MCP command is `uv` again:

```json
{
  "mcpServers": {
    "spreadsheet-tools": {
      "type": "stdio",
      "command": "uv",
      "args": ["run", "--project", "/absolute/path/to/spreadsheet-tools", "spreadsheet-tools-mcp"]
    }
  }
}
```

To build a binary for the OS you are on (not for other systems):

```bash
uv run --with pyinstaller python scripts/build_binary.py
```

The file lands in `dist/`. Linux x86_64, Mac Apple Silicon, and Windows x64 are built on every push to `main`. PyInstaller cannot cross-compile the way a Rust build can, so each runner produces only its own file. Linux ARM and Intel Mac are not in that set.
