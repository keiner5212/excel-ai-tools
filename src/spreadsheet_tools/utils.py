from __future__ import annotations

import io
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

COL_RE = re.compile(r"^[A-Z]+$")
CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def resolve_path(path: str) -> Path:
    resolved = Path(path).expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"File not found: {resolved}")
    return resolved


def open_workbook(path: str, *, read_only: bool = False, data_only: bool = True) -> Workbook:
    return load_workbook(resolve_path(path), read_only=read_only, data_only=data_only, keep_vba=True)


def open_workbook_for_write(path: str) -> Workbook:
    return load_workbook(resolve_path(path), data_only=False, keep_vba=True)


def validate_column(column: str) -> str:
    normalized = column.strip().upper()
    if not COL_RE.match(normalized):
        raise ValueError(f"Invalid column letter: {column!r}")
    return normalized


def validate_cell_address(address: str) -> str:
    normalized = address.strip().upper()
    if not CELL_RE.match(normalized):
        raise ValueError(f"Invalid cell address: {address!r}")
    return normalized


def column_to_index(column: str) -> int:
    return column_index_from_string(validate_column(column))


def index_to_column(index: int) -> str:
    if index < 1:
        raise ValueError(f"Column index must be >= 1, got {index}")
    return get_column_letter(index)


def parse_cell_address(address: str) -> tuple[str, int]:
    match = CELL_RE.match(validate_cell_address(address))
    assert match is not None
    return match.group(1), int(match.group(2))


def get_sheet(workbook: Workbook, sheet_name: str | None) -> Worksheet:
    if sheet_name is None:
        return workbook.active
    if sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(f"Sheet {sheet_name!r} not found. Available: {available}")
    return workbook[sheet_name]


def normalize_scalar(value: object) -> object | None:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return value


def _merge_rels(orig_data: bytes, new_data: bytes, *, prefer_new: bool = False) -> bytes:
    """Merge two .rels XML files.

    Default (prefer_new=False): starts from original, appends any new entries
    from openpyxl that have a different Id. Used for worksheet rels to preserve
    drawing/hyperlink references openpyxl drops.

    prefer_new=True: starts from openpyxl's version, appends any original entries
    not already present. Used for workbook.xml.rels because openpyxl renumbers
    rIds in workbook.xml, so the rels must match openpyxl's numbering scheme.
    """
    RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
    ET.register_namespace("", RELS_NS)
    tag = f"{{{RELS_NS}}}Relationship"

    orig_root = ET.fromstring(orig_data)
    new_root = ET.fromstring(new_data)

    if prefer_new:
        base_root = new_root
        extra_root = orig_root
    else:
        base_root = orig_root
        extra_root = new_root

    base_ids = {el.get("Id") for el in base_root.iter(tag)}
    for el in extra_root.iter(tag):
        if el.get("Id") not in base_ids:
            base_root.append(el)

    xml_str = ET.tostring(base_root, encoding="unicode")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str).encode("utf-8")


def _merge_content_types(orig_data: bytes, new_data: bytes) -> bytes:
    """Merge [Content_Types].xml.

    Starts from the original (which registers types for all original parts:
    images, drawings, charts, …) and adds any new Override/Default entries
    from openpyxl (e.g. a newly added sheet). This ensures that files
    restored from the original ZIP still have registered content types.
    """
    CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
    ET.register_namespace("", CT_NS)
    default_tag = f"{{{CT_NS}}}Default"
    override_tag = f"{{{CT_NS}}}Override"

    orig_root = ET.fromstring(orig_data)
    new_root = ET.fromstring(new_data)

    orig_extensions = {el.get("Extension") for el in orig_root.iter(default_tag)}
    orig_partnames = {el.get("PartName") for el in orig_root.iter(override_tag)}

    for el in new_root.iter(default_tag):
        if el.get("Extension") not in orig_extensions:
            orig_root.append(el)
    for el in new_root.iter(override_tag):
        if el.get("PartName") not in orig_partnames:
            orig_root.append(el)

    xml_str = ET.tostring(orig_root, encoding="unicode")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str).encode("utf-8")


def _merge_workbook_zips(original_path: Path, new_buf: io.BytesIO, out_path: Path) -> None:
    """Build a merged ZIP that preserves ALL original content.

    Algorithm:
    - Take every file from openpyxl's output (updated cells, styles, workbook XML).
    - For .rels files present in both ZIPs: merge relationship entries so
      openpyxl never silently drops drawing/image/chart references.
    - For [Content_Types].xml: merge from original + add new openpyxl entries.
    - For any file present in the original but absent from openpyxl's output
      (images in xl/media/, drawings in xl/drawings/, charts, pivot tables,
      ActiveX, …): restore it verbatim from the original.
    """
    new_buf.seek(0)
    with (
        zipfile.ZipFile(original_path, "r") as orig_zip,
        zipfile.ZipFile(new_buf, "r") as new_zip,
        zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as out_zip,
    ):
        orig_names = set(orig_zip.namelist())
        new_names = set(new_zip.namelist())

        for name in new_zip.namelist():
            data = new_zip.read(name)
            if name in orig_names:
                try:
                    if name == "[Content_Types].xml":
                        data = _merge_content_types(orig_zip.read(name), data)
                    elif name.endswith(".rels"):
                        # workbook.xml.rels must use openpyxl's rId numbering as
                        # the base because openpyxl rewrites workbook.xml with new
                        # rIds. Using the original as base would map those new rIds
                        # to wrong targets (theme, styles, vbaProject instead of
                        # the actual worksheet files).
                        prefer_new = name == "xl/_rels/workbook.xml.rels"
                        data = _merge_rels(orig_zip.read(name), data, prefer_new=prefer_new)
                except Exception:
                    pass  # fall back to openpyxl's version on parse error
            out_zip.writestr(new_zip.getinfo(name), data)

        # Restore entries openpyxl dropped entirely
        for name in orig_names - new_names:
            out_zip.writestr(orig_zip.getinfo(name), orig_zip.read(name))


def safe_save_workbook(workbook: Workbook, path: str | Path) -> None:
    """Save workbook with zero data loss.

    Strategy:
    1. Back up original to <file>.bak.
    2. Save openpyxl's output to an in-memory buffer.
    3. Merge original ZIP and openpyxl buffer:
       - .rels files are merged (all original relationships preserved).
       - [Content_Types].xml is merged (original types preserved).
       - Files openpyxl dropped (images, drawings, charts, VBA, …) are
         restored verbatim from the original.
    4. Write merged output to a temp file in the same directory.
    5. Atomically rename temp → original (os.replace; POSIX atomic).

    If step 5 fails the original is untouched. The .bak copy provides an
    additional recovery path in case openpyxl's serialisation itself is lossy.
    """
    target = Path(path) if isinstance(path, str) else path
    backup = target.with_suffix(target.suffix + ".bak")
    shutil.copy2(target, backup)

    buf = io.BytesIO()
    workbook.save(buf)

    fd, tmp_path_str = tempfile.mkstemp(dir=target.parent, suffix=".tmp")
    try:
        os.close(fd)
        tmp_path = Path(tmp_path_str)
        _merge_workbook_zips(target, buf, tmp_path)
        os.replace(tmp_path, target)
    except Exception:
        try:
            os.unlink(tmp_path_str)
        except OSError:
            pass
        raise
