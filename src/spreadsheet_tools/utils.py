from __future__ import annotations

import io
import os
import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet

COL_RE = re.compile(r"^[A-Z]+$")
# Row 0 is invalid in Excel (1-based); [1-9]\d* rejects it at validation time.
CELL_RE = re.compile(r"^([A-Z]+)([1-9]\d*)$")
RANGE_RE = re.compile(r"^([A-Z]+[1-9]\d*):([A-Z]+[1-9]\d*)$")

# Default workspace directory for new workbooks created without an explicit path.
WORKSPACE_DIR_NAME = "workspace"

# Shared text-analysis constants used by reader.py and rules.py
SECTION_HEADER_RE = re.compile(r"^(\d+(?:\.\d+)*\.?)\s+(.+)$")
GENERIC_NAME_RE = re.compile(r"^Estrategia\s+de\s+\w+\s+\d+$", re.IGNORECASE)
STOPWORDS: frozenset[str] = frozenset(
    {
        "de",
        "la",
        "el",
        "en",
        "y",
        "para",
        "con",
        "del",
        "los",
        "las",
        "que",
        "un",
        "una",
        "a",
        "e",
        "o",
        "al",
        "se",
        "su",
        "sus",
        "and",
        "the",
        "of",
        "in",
        "to",
        "for",
        "with",
        "an",
    }
)


def keyword_overlap_ratio(name: str, desc: str) -> float:
    """Fraction of meaningful name tokens (len > 2, not stopwords) found in desc.

    Returns 1.0 when the name has no meaningful tokens (cannot detect mismatch).
    Used by describe-section and validate name-matches-desc rule.
    """
    name_tokens = {w.lower() for w in re.split(r"\W+", name) if len(w) > 2}
    name_kw = name_tokens - STOPWORDS
    if not name_kw:
        return 1.0
    desc_lower = desc.lower()
    matched = sum(1 for kw in name_kw if kw in desc_lower)
    return matched / len(name_kw)


# Register XML namespaces once at import time (module-level side effect is safe;
# repeated calls from multiple threads would corrupt the global namespace registry).
_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
ET.register_namespace("", _RELS_NS)


def validate_cell_range(range_str: str) -> tuple[str, str]:
    """Validate A1:B2 range notation. Returns (top_left, bottom_right) normalized.

    Both endpoints must be valid cell addresses. Top-left must not be below or
    to the right of bottom-right (Excel convention).
    """
    normalized = range_str.strip().upper()
    m = RANGE_RE.match(normalized)
    if not m:
        raise ValueError(f"Invalid cell range: {range_str!r}. Expected A1:B2 format.")
    return m.group(1), m.group(2)


def resolve_output_path(path: str) -> Path:
    """Route bare filenames (no directory component) to ./workspace/.

    A bare filename is one whose parent is the current directory (e.g. 'file.xlsx').
    Absolute paths, '~' paths, and paths with a directory component are kept.
    '~' is expanded so a home path is not treated as a relative name.

    The workspace directory is created if it does not already exist.
    """
    p = Path(path).expanduser()
    if not p.is_absolute() and p.parent == Path("."):
        workspace = Path(WORKSPACE_DIR_NAME)
        workspace.mkdir(exist_ok=True)
        return workspace / p.name
    return p


def resolve_path(path: str) -> Path:
    resolved = Path(path).expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"File not found: {resolved}")
    return resolved


def _close_vba_archive(workbook: Workbook) -> None:
    """Close vba_archive if open.

    openpyxl stores VBA content in an in-memory ZipFile (vba_archive). It is never
    closed by openpyxl itself, which causes ZipFile.__del__ to fail with
    "I/O operation on closed file" when the workbook is garbage-collected as part of
    a reference cycle (the save-patch closure captures the workbook).
    """
    if workbook.vba_archive is not None:
        try:
            workbook.vba_archive.close()
        except (OSError, ValueError) as exc:
            # ZipFile.close raises if the archive handle is already gone.
            if "closed" not in str(exc).lower():
                raise


def open_workbook(
    path: str, *, read_only: bool = False, data_only: bool = True
) -> Workbook:
    workbook = load_workbook(
        resolve_path(path), read_only=read_only, data_only=data_only, keep_vba=True
    )
    original_close = workbook.close

    def _safe_close() -> None:  # type: ignore[misc]
        _close_vba_archive(workbook)
        original_close()

    workbook.close = _safe_close  # type: ignore[method-assign]
    return workbook


def _patch_workbook_save(workbook: Workbook) -> None:
    """Replace workbook.save with a safe version that preserves all embedded content.

    openpyxl's native save strips xl/media/, drawings, ActiveX, and VML objects.
    This patch intercepts file-path saves and routes them through safe_save_workbook,
    which merges the openpyxl output with the original ZIP at the binary level.

    BytesIO saves (used internally by safe_save_workbook itself) are left untouched
    to avoid infinite recursion.

    Also wraps close() to eagerly release vba_archive before the workbook is GC'd.
    See _close_vba_archive for the rationale.
    """
    original_save = workbook.save
    original_close = workbook.close

    def _safe_save(path: str | Path | io.BytesIO) -> None:  # type: ignore[misc]
        # BytesIO path: safe_save_workbook calls this internally — pass through
        if isinstance(path, io.BytesIO):
            original_save(path)
            return
        safe_save_workbook(workbook, path)

    def _safe_close() -> None:  # type: ignore[misc]
        _close_vba_archive(workbook)
        original_close()

    workbook.save = _safe_save  # type: ignore[method-assign]
    workbook.close = _safe_close  # type: ignore[method-assign]


def open_workbook_for_write(path: str) -> Workbook:
    workbook = load_workbook(resolve_path(path), data_only=False, keep_vba=True)
    _patch_workbook_save(workbook)
    return workbook


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
    # Coerce the string literal "null" that some AI agents send when they mean
    # to omit the parameter (the MCP SDK may or may not coerce it automatically).
    if sheet_name == "null":
        sheet_name = None
    # Check membership before access to avoid a raw KeyError from workbook[name].
    if sheet_name is not None and sheet_name not in workbook.sheetnames:
        available = ", ".join(workbook.sheetnames)
        raise ValueError(f"Sheet {sheet_name!r} not found. Available: {available}")
    sheet = workbook.active if sheet_name is None else workbook[sheet_name]
    if sheet is None:
        raise ValueError("No active sheet found in workbook")
    if not isinstance(sheet, (Worksheet, ReadOnlyWorksheet)):
        raise ValueError(
            f"Sheet {sheet_name!r} is not a worksheet (may be a chart sheet)"
        )
    return sheet  # type: ignore[return-value]


def normalize_scalar(value: object) -> object | None:
    if value is None:
        return None
    if isinstance(value, str):
        stripped = value.strip()
        return stripped if stripped else None
    return value


def _merge_rels(
    orig_data: bytes, new_data: bytes, *, prefer_new: bool = False
) -> bytes:
    """Merge two .rels XML files.

    Default (prefer_new=False): starts from original, appends any new entries
    from openpyxl that have a different Id. Used for worksheet rels to preserve
    drawing/hyperlink references openpyxl drops.

    prefer_new=True: starts from openpyxl's version, appends any original entries
    not already present. Used for workbook.xml.rels because openpyxl renumbers
    rIds in workbook.xml, so the rels must match openpyxl's numbering scheme.
    """
    tag = f"{{{_RELS_NS}}}Relationship"

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
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str
    ).encode("utf-8")


def _merge_content_types(orig_data: bytes, new_data: bytes) -> bytes:
    """Merge [Content_Types].xml.

    Starts from the original (which registers types for all original parts:
    images, drawings, charts, …) and adds any new Override/Default entries
    from openpyxl (e.g. a newly added sheet). This ensures that files
    restored from the original ZIP still have registered content types.
    """
    default_tag = f"{{{_CT_NS}}}Default"
    override_tag = f"{{{_CT_NS}}}Override"

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
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str
    ).encode("utf-8")


_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_R_NS_DECL = f'xmlns:r="{_R_NS}"'


def _ensure_r_namespace(element_snippet: str) -> str:
    """Add xmlns:r declaration if the element uses r: prefix but lacks the declaration.

    openpyxl does not declare r: on the worksheet root element; it inlines
    xmlns:r on each child element that needs it (legacyDrawing, hyperlink).
    Drawing elements restored from the original must follow the same pattern.
    """
    if "r:" not in element_snippet or _R_NS_DECL in element_snippet:
        return element_snippet
    # Insert xmlns:r before the first attribute or before the closing />
    return re.sub(r"(<\w+)", rf"\1 {_R_NS_DECL}", element_snippet, count=1)


# Elements openpyxl silently drops from worksheet XML.
# They must be restored verbatim so Excel can locate drawings/images.
_DROPPED_ELEMENT_TAGS = ("drawing", "legacyDrawing", "picture")


def _extract_drawing_elements(sheet_xml: bytes) -> list[str]:
    """Return raw XML snippets of drawing/legacyDrawing/picture elements."""
    tag_pattern = "|".join(_DROPPED_ELEMENT_TAGS)
    return re.findall(
        rf"<(?:{tag_pattern})(?:\s[^>]*)?\s*/>",
        sheet_xml.decode("utf-8", errors="replace"),
    )


def _inject_drawing_elements(new_xml: bytes, elements: list[str]) -> bytes:
    """Inject missing drawing/picture elements into openpyxl's worksheet XML.

    Inserts them right before </worksheet> so they appear in the correct
    position per OOXML spec (after sheetData, pageMargins, etc.).
    Drawing elements reference relationships by rId; the companion .rels
    merge ensures those rIds still resolve to the correct targets.

    legacyDrawing and picture are limited to one per sheet; openpyxl already
    writes its own version of these, so we never duplicate them.
    """
    if not elements:
        return new_xml

    text = new_xml.decode("utf-8", errors="replace")
    tag_pattern = "|".join(_DROPPED_ELEMENT_TAGS)

    # OOXML allows one drawing, one legacyDrawing, and one picture per sheet.
    _SINGLE_INSTANCE_TAGS = {"drawing", "legacyDrawing", "picture"}
    existing_tags = set(re.findall(rf"<({tag_pattern})[\s/>]", text))
    # Support both single- and double-quoted r:id attributes to avoid
    # duplicate injection when the original XML used single quotes.
    existing_rids = set(
        re.findall(rf"""(?:{tag_pattern})\s+[^>]*r:id=["']([^"']+)["']""", text)
    )

    to_inject = []
    for snippet in elements:
        tag_match = re.match(r"<(\w+)", snippet)
        if not tag_match:
            continue
        tag = tag_match.group(1)
        # Skip if this single-instance tag is already present in new XML
        if tag in _SINGLE_INSTANCE_TAGS:
            if tag in existing_tags:
                continue
            existing_tags.add(tag)
        rid_match = re.search(r"""r:id=["']([^"']+)["']""", snippet)
        rid = rid_match.group(1) if rid_match else None
        if rid is None or rid not in existing_rids:
            to_inject.append(snippet)

    if not to_inject:
        return new_xml

    injected_block = "\n".join(_ensure_r_namespace(s) for s in to_inject)
    # OOXML sequence puts drawing/picture before tableParts and extLst.
    # Appending after extLst makes Excel repair the sheet.
    anchors = [text.find(token) for token in ("<tableParts", "<extLst", "</worksheet>")]
    anchors = [index for index in anchors if index != -1]
    if not anchors:
        raise ValueError("Worksheet XML has no closing tag")
    at = min(anchors)
    text = text[:at] + injected_block + text[at:]
    return text.encode("utf-8")


# Parts openpyxl drops. Structural parts (sheets, styles, strings) stay with
# openpyxl so a deleted or reordered sheet is not copied back into the package.
_RESTORE_PREFIXES = (
    "xl/media/",
    "xl/drawings/",
    "xl/charts/",
    "xl/embeddings/",
    "xl/activeX/",
    "xl/ctrlProps/",
    "xl/printerSettings/",
    "xl/diagrams/",
    "xl/customXml/",
    "xl/comments/",
    "xl/threadedComments/",
    "xl/persons/",
    "xl/ink/",
    "customUI/",
    "xl/customUI/",
)
_RESTORE_FILES = frozenset({"xl/vbaProject.bin", "xl/vbaProjectSignature.bin"})
_WORKBOOK_PRESERVE_TYPES = ("vbaProject", "customXml", "attachedTool", "customUI")
_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _restorable(name: str) -> bool:
    return name in _RESTORE_FILES or name.startswith(_RESTORE_PREFIXES)


def _read_zip(archive: zipfile.ZipFile) -> dict[str, bytes]:
    return {name: archive.read(name) for name in archive.namelist()}


def _parse_rels(data: bytes) -> list[dict[str, str]]:
    tag = f"{{{_RELS_NS}}}Relationship"
    items: list[dict[str, str]] = []
    for el in ET.fromstring(data).iter(tag):
        item = {
            "Id": el.get("Id") or "",
            "Type": el.get("Type") or "",
            "Target": el.get("Target") or "",
        }
        mode = el.get("TargetMode")
        if mode:
            item["TargetMode"] = mode
        items.append(item)
    return items


def _write_rels(items: list[dict[str, str]]) -> bytes:
    root = ET.Element(f"{{{_RELS_NS}}}Relationships")
    for item in items:
        attrib = {"Id": item["Id"], "Type": item["Type"], "Target": item["Target"]}
        if item.get("TargetMode"):
            attrib["TargetMode"] = item["TargetMode"]
        ET.SubElement(root, f"{{{_RELS_NS}}}Relationship", attrib)
    return _xml_bytes(root)


def _xml_bytes(root: ET.Element) -> bytes:
    xml = ET.tostring(root, encoding="unicode")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml).encode()


def _rels_path(part: str) -> str:
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def _source_part_for_rels(rels_name: str) -> str:
    folder, name = posixpath.split(rels_name)
    if posixpath.basename(folder) != "_rels":
        return rels_name
    parent = posixpath.dirname(folder)
    return posixpath.join(parent, name[: -len(".rels")])


def _is_external_rel(rel: dict[str, str]) -> bool:
    target = rel.get("Target") or ""
    mode = (rel.get("TargetMode") or "").lower()
    return mode == "external" or "://" in target or target.startswith(("#", "mailto:"))


def _resolve_rel_target(source_part: str, target: str) -> str | None:
    if not target or _is_external_rel({"Target": target}):
        return None
    if target.startswith("/"):
        resolved = posixpath.normpath(target.lstrip("/"))
    else:
        resolved = posixpath.normpath(
            posixpath.join(posixpath.dirname(source_part), target)
        )
    if resolved.startswith("..") or not resolved:
        return None
    return resolved


def _next_rid(used: set[str]) -> str:
    number = 1
    while f"rId{number}" in used:
        number += 1
    return f"rId{number}"


def _rewrite_rid(snippet: str, old: str, new: str) -> str:
    return re.sub(
        rf"""(r:id=["']){re.escape(old)}(["'])""",
        rf"\1{new}\2",
        snippet,
        count=1,
    )


def _sheet_parts_by_name(parts: dict[str, bytes]) -> dict[str, str]:
    """Map sheet name -> worksheet part path. Match by name, not sheetN.xml."""
    workbook = parts.get("xl/workbook.xml")
    rels_blob = parts.get("xl/_rels/workbook.xml.rels")
    if workbook is None or rels_blob is None:
        return {}
    rel_by_id = {item["Id"]: item["Target"] for item in _parse_rels(rels_blob)}
    found: dict[str, str] = {}
    for el in ET.fromstring(workbook).iter(f"{{{_MAIN_NS}}}sheet"):
        name = el.get("name")
        rid = el.get(f"{{{_R_NS}}}id")
        if not name or not rid or rid not in rel_by_id:
            continue
        target = rel_by_id[rid]
        if target.startswith("/"):
            part = posixpath.normpath(target.lstrip("/"))
        else:
            part = posixpath.normpath(posixpath.join("xl", target))
        if part.startswith("xl/worksheets/"):
            found[name] = part
    return found


def _pair_sheets(
    original: dict[str, str], rewritten: dict[str, str]
) -> list[tuple[str, str]]:
    pairs: list[tuple[str, str]] = []
    unmatched_new: list[str] = []
    for name, part in rewritten.items():
        if name in original:
            pairs.append((original[name], part))
        else:
            unmatched_new.append(part)
    unmatched_orig = [part for name, part in original.items() if name not in rewritten]
    # A single rename keeps one sheet on each side with no shared name.
    if len(unmatched_new) == 1 and len(unmatched_orig) == 1:
        pairs.append((unmatched_orig[0], unmatched_new[0]))
    return pairs


def _copy_closure(
    origin: dict[str, bytes],
    final: dict[str, bytes],
    part: str,
    seen: set[str],
) -> None:
    if part in seen or part in final or part not in origin or not _restorable(part):
        return
    seen.add(part)
    final[part] = origin[part]
    rels_name = _rels_path(part)
    if rels_name not in origin or rels_name in final:
        return
    final[rels_name] = origin[rels_name]
    for rel in _parse_rels(origin[rels_name]):
        if _is_external_rel(rel):
            continue
        child = _resolve_rel_target(part, rel["Target"])
        if child:
            _copy_closure(origin, final, child, seen)


def _graft_sheet_drawings(
    origin: dict[str, bytes],
    final: dict[str, bytes],
    orig_part: str,
    new_part: str,
) -> None:
    """Put original drawing refs back onto the matching openpyxl sheet.

    Relationship ids are remapped when openpyxl already used that rId for
    something else (a hyperlink, for example). Targets are copied only when
    they are embedded parts openpyxl dropped.
    """
    if orig_part not in origin or new_part not in final:
        return
    elements = _extract_drawing_elements(origin[orig_part])
    if not elements:
        return
    orig_rels_name = _rels_path(orig_part)
    orig_rels = _parse_rels(origin[orig_rels_name]) if orig_rels_name in origin else []
    orig_by_id = {item["Id"]: item for item in orig_rels}
    new_rels_name = _rels_path(new_part)
    new_rels = _parse_rels(final[new_rels_name]) if new_rels_name in final else []
    used_ids = {item["Id"] for item in new_rels}
    rewritten: list[str] = []
    for snippet in elements:
        rid_match = re.search(r"""r:id=["']([^"']+)["']""", snippet)
        if not rid_match:
            continue
        old_rid = rid_match.group(1)
        rel = orig_by_id.get(old_rid)
        if rel is None:
            continue
        current = next((item for item in new_rels if item["Id"] == old_rid), None)
        rel_key = (rel["Type"], rel["Target"], rel.get("TargetMode"))
        if current is None:
            new_rels.append(dict(rel))
            used_ids.add(old_rid)
            rid = old_rid
        elif (current["Type"], current["Target"], current.get("TargetMode")) != rel_key:
            rid = _next_rid(used_ids)
            snippet = _rewrite_rid(snippet, old_rid, rid)
            used_ids.add(rid)
            new_rels.append({**rel, "Id": rid})
        else:
            rid = old_rid
        if not _is_external_rel(rel):
            child = _resolve_rel_target(orig_part, rel["Target"])
            if child:
                _copy_closure(origin, final, child, set())
        rewritten.append(snippet)
    if not rewritten:
        return
    final[new_part] = _inject_drawing_elements(final[new_part], rewritten)
    final[new_rels_name] = _write_rels(new_rels)


def _graft_workbook_extras(origin: dict[str, bytes], final: dict[str, bytes]) -> None:
    """Keep workbook-level embeds (VBA, custom XML) openpyxl failed to write."""
    rels_name = "xl/_rels/workbook.xml.rels"
    if rels_name not in origin or "xl/workbook.xml" not in final:
        return
    new_rels = _parse_rels(final[rels_name]) if rels_name in final else []
    known_targets = {item["Target"] for item in new_rels}
    used_ids = {item["Id"] for item in new_rels}
    changed = False
    for rel in _parse_rels(origin[rels_name]):
        rel_type = rel.get("Type") or ""
        if not any(token in rel_type for token in _WORKBOOK_PRESERVE_TYPES):
            continue
        if rel["Target"] in known_targets or _is_external_rel(rel):
            continue
        part = _resolve_rel_target("xl/workbook.xml", rel["Target"])
        if part is None or part not in origin or not _restorable(part):
            continue
        new_id = rel["Id"] if rel["Id"] not in used_ids else _next_rid(used_ids)
        used_ids.add(new_id)
        new_rels.append({**rel, "Id": new_id})
        known_targets.add(rel["Target"])
        _copy_closure(origin, final, part, set())
        changed = True
    if changed:
        final[rels_name] = _write_rels(new_rels)


def _prune_content_type_overrides(data: bytes, names: set[str]) -> bytes:
    root = ET.fromstring(data)
    override_tag = f"{{{_CT_NS}}}Override"
    for el in list(root.findall(override_tag)):
        part = (el.get("PartName") or "").lstrip("/")
        if part not in names:
            root.remove(el)
    return _xml_bytes(root)


def _harmonize_content_types(origin: dict[str, bytes], final: dict[str, bytes]) -> None:
    ct_name = "[Content_Types].xml"
    new_ct = final.get(ct_name)
    if new_ct is None:
        return
    orig_ct = origin.get(ct_name)
    merged = _merge_content_types(orig_ct, new_ct) if orig_ct is not None else new_ct
    final[ct_name] = _prune_content_type_overrides(merged, set(final))


def _write_zip(parts: dict[str, bytes], out_path: Path) -> None:
    names = sorted(name for name in parts if name != "[Content_Types].xml")
    if "[Content_Types].xml" in parts:
        names.insert(0, "[Content_Types].xml")
    with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name in names:
            info = zipfile.ZipInfo(filename=name)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.date_time = (1980, 1, 1, 0, 0, 0)
            archive.writestr(info, parts[name])


def _strip_dangling_vba_rel(path: Path) -> None:
    """Drop a vbaProject relationship when the binary is absent.

    openpyxl writes that relationship for every keep_vba load, including
    workbooks that have no VBA. Excel then repairs the file.
    """
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        rels_name = "xl/_rels/workbook.xml.rels"
        if "xl/vbaProject.bin" in names or rels_name not in names:
            return
        rels = _parse_rels(archive.read(rels_name))
        kept = [
            rel
            for rel in rels
            if _resolve_rel_target("xl/workbook.xml", rel["Target"])
            != "xl/vbaProject.bin"
        ]
        if len(kept) == len(rels):
            return
        parts = _read_zip(archive)
    parts[rels_name] = _write_rels(kept)
    if "[Content_Types].xml" in parts:
        parts["[Content_Types].xml"] = _prune_content_type_overrides(
            parts["[Content_Types].xml"], set(parts)
        )
    _write_zip(parts, path)


def _temp_workbook_path(directory: Path, suffix: str) -> Path:
    excel_suffix = suffix if suffix in {".xlsx", ".xlsm", ".xltx", ".xltm"} else ".xlsx"
    fd, name = tempfile.mkstemp(dir=directory, suffix=excel_suffix)
    os.close(fd)
    return Path(name)


def _assert_package_consistent(path: Path) -> None:
    """Reject a package Excel would have to repair. Caller keeps the original."""
    _strip_dangling_vba_rel(path)
    with zipfile.ZipFile(path) as archive:
        bad = archive.testzip()
        if bad is not None:
            raise ValueError(f"Corrupt ZIP entry after save: {bad}")
        names = set(archive.namelist())
        content_types = archive.read("[Content_Types].xml")
        override_tag = f"{{{_CT_NS}}}Override"
        for el in ET.fromstring(content_types).findall(override_tag):
            part = (el.get("PartName") or "").lstrip("/")
            if part not in names:
                raise ValueError(f"Content type points at missing part {part}")
        for name in names:
            if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
                ET.fromstring(archive.read(name))
            if not name.endswith(".rels"):
                continue
            source = _source_part_for_rels(name)
            for rel in _parse_rels(archive.read(name)):
                if _is_external_rel(rel):
                    continue
                target = _resolve_rel_target(source, rel["Target"])
                if target is not None and target not in names:
                    raise ValueError(f"{name} points at missing part {target}")
    workbook = open_workbook(str(path), read_only=True, data_only=False)
    try:
        if not workbook.sheetnames:
            raise ValueError("Workbook has no sheets after save")
    finally:
        workbook.close()


def _merge_workbook_zips(
    original_path: Path, new_buf: io.BytesIO, out_path: Path
) -> None:
    """Use openpyxl's package as structure. Graft back only dropped embeds.

    Sheets are paired by name so a delete or reorder cannot stick another
    sheet's drawings onto the wrong worksheet. Relationship ids that collide
    with openpyxl's own rels are renumbered. Parts openpyxl removed on purpose
    (deleted sheets, calc chain, shared strings) are not copied back.
    """
    new_buf.seek(0)
    with (
        zipfile.ZipFile(original_path, "r") as orig_zip,
        zipfile.ZipFile(new_buf, "r") as new_zip,
    ):
        origin = _read_zip(orig_zip)
        final = _read_zip(new_zip)
    for orig_part, new_part in _pair_sheets(
        _sheet_parts_by_name(origin), _sheet_parts_by_name(final)
    ):
        _graft_sheet_drawings(origin, final, orig_part, new_part)
    _graft_workbook_extras(origin, final)
    _harmonize_content_types(origin, final)
    _write_zip(final, out_path)


def safe_save_workbook(workbook: Workbook, path: str | Path) -> None:
    """Replace the workbook only after the merged package checks out.

    openpyxl writes to memory. The merge result is validated (zip integrity,
    relationship targets, content types, and a reload) before os.replace.
    A failed save leaves the original bytes in place.
    """
    target = Path(path).expanduser().resolve()
    if not target.is_file():
        raise FileNotFoundError(f"File not found: {target}")
    backup = target.with_name(target.name + ".bak")
    shutil.copy2(target, backup)

    buf = io.BytesIO()
    workbook.save(buf)

    tmp_path = _temp_workbook_path(target.parent, target.suffix)
    try:
        _merge_workbook_zips(target, buf, tmp_path)
        _assert_package_consistent(tmp_path)
        os.replace(tmp_path, target)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        backup.unlink(missing_ok=True)
        raise
    try:
        backup.unlink()
    except OSError:
        pass


def save_new_workbook(workbook: Workbook, path: str | Path) -> None:
    """Write a new workbook. An existing file is replaced only after it validates."""
    target = Path(path).expanduser().resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = _temp_workbook_path(target.parent, target.suffix)
    try:
        workbook.save(tmp_path)
        _assert_package_consistent(tmp_path)
        os.replace(tmp_path, target)
    except Exception:
        tmp_path.unlink(missing_ok=True)
        raise
