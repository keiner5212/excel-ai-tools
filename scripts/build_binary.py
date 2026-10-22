"""Build a standalone program for the machine running this script.

Default layout is a directory. It does not unpack into /tmp on launch.
Pass --onefile only when a single file is required. That mode extracts a
runtime on every start and leaks it if the process is killed with SIGKILL.
Cross-compiling is not supported. CI builds the other operating systems.
"""

from __future__ import annotations

import platform
import shutil
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
    onefile = "--onefile" in sys.argv[1:]
    stem = binary_stem()
    dist = ROOT / "dist"
    suffix = ".exe" if platform.system() == "Windows" else ""
    # A previous onefile build leaves a file where onedir needs a directory.
    stale = dist / f"{stem}{suffix}"
    if not onefile and stale.is_file():
        stale.unlink()
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--console",
        "--onedir" if not onefile else "--onefile",
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
        "--copy-metadata",
        "spreadsheet-tools",
        "--exclude-module",
        "mcp.cli",
        str(ENTRY),
    ]
    subprocess.run(cmd, cwd=ROOT, check=True)
    if onefile:
        print(dist / f"{stem}{suffix}")
        return
    exe = dist / stem / f"{stem}{suffix}"
    archive = shutil.make_archive(str(dist / stem), "zip", root_dir=dist, base_dir=stem)
    print(exe)
    print(archive)


if __name__ == "__main__":
    main()
