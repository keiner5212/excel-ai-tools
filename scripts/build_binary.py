"""Build one standalone executable for the machine running this script.

The file contains Python and the libraries. The user does not install them.
Cross-compiling is not supported. CI builds the other operating systems.
"""

from __future__ import annotations

import platform
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ENTRY = ROOT / "src" / "spreadsheet_tools" / "frozen_main.py"

_SYSTEMS = {"Linux": "linux", "Darwin": "macos", "Windows": "windows"}
_ARCHS = {
    "x86_64": "x86_64",
    "amd64": "x86_64",
    "arm64": "arm64",
    "aarch64": "arm64",
}


def binary_stem() -> str:
    system = _SYSTEMS.get(platform.system())
    if system is None:
        raise SystemExit(f"Unsupported OS: {platform.system()}")
    arch = _ARCHS.get(platform.machine().lower(), platform.machine().lower())
    return f"spreadsheet-tools-{system}-{arch}"


def main() -> None:
    stem = binary_stem()
    dist = ROOT / "dist"
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--name",
        stem,
        "--distpath",
        str(dist),
        "--workpath",
        str(ROOT / "build" / "pyinstaller"),
        "--specpath",
        str(ROOT / "build"),
        "--collect-submodules",
        "spreadsheet_tools",
        "--collect-all",
        "openpyxl",
        "--exclude-module",
        "mcp.cli",
        str(ENTRY),
    ]
    subprocess.run(cmd, cwd=ROOT, check=True)
    suffix = ".exe" if platform.system() == "Windows" else ""
    print(dist / f"{stem}{suffix}")


if __name__ == "__main__":
    main()
