"""Entry point for the standalone executable.

No arguments and a piped stdin start the MCP server.
Any other invocation is the CLI.
A terminal with no arguments prints CLI help.
"""

from __future__ import annotations

import os
import sys

from spreadsheet_tools.cli import main as cli_main
from spreadsheet_tools.mcp_server import main as mcp_main, package_version


def _log_startup(mode: str) -> None:
    print(
        f"spreadsheet-tools {package_version()} python {sys.version.split()[0]} "
        f"pid {os.getpid()} cwd {os.getcwd()} mode {mode}",
        file=sys.stderr,
    )


def main() -> None:
    mcp_mode = len(sys.argv) == 1 and not sys.stdin.isatty()
    if getattr(sys, "frozen", False) or mcp_mode:
        _log_startup("mcp" if mcp_mode else "cli")
    if mcp_mode:
        mcp_main()
        return
    argv = None if len(sys.argv) > 1 else ["--help"]
    raise SystemExit(cli_main(argv))


if __name__ == "__main__":
    main()
