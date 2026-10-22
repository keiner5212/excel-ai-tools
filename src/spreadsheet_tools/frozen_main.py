"""Entry point for the standalone executable.

No arguments and a piped stdin start the MCP server.
Any other invocation is the CLI.
A terminal with no arguments prints CLI help.
"""

from __future__ import annotations

import sys

from spreadsheet_tools.cli import main as cli_main
from spreadsheet_tools.mcp_server import main as mcp_main


def main() -> None:
    if len(sys.argv) == 1 and not sys.stdin.isatty():
        mcp_main()
        return
    argv = None if len(sys.argv) > 1 else ["--help"]
    raise SystemExit(cli_main(argv))


if __name__ == "__main__":
    main()
