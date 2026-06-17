from __future__ import annotations

import pytest

from spreadsheet_tools.utils import (
    _extract_drawing_elements,
    _inject_drawing_elements,
    _merge_content_types,
    _merge_rels,
    column_to_index,
    index_to_column,
    keyword_overlap_ratio,
    normalize_scalar,
    parse_cell_address,
    validate_cell_address,
    validate_column,
)

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def test_validate_column_accepts_letters() -> None:
    assert validate_column("ab") == "AB"
    assert column_to_index("C") == 3


def test_validate_column_rejects_invalid() -> None:
    with pytest.raises(ValueError, match="Invalid column"):
        validate_column("A1")


def test_validate_cell_address_rejects_row_zero() -> None:
    with pytest.raises(ValueError, match="Invalid cell address"):
        validate_cell_address("A0")


def test_parse_cell_address() -> None:
    assert parse_cell_address("b12") == ("B", 12)


def test_index_to_column_rejects_non_positive() -> None:
    with pytest.raises(ValueError, match="Column index"):
        index_to_column(0)


def test_normalize_scalar() -> None:
    assert normalize_scalar(None) is None
    assert normalize_scalar("  ") is None
    assert normalize_scalar("  x  ") == "x"
    assert normalize_scalar(42) == 42


def test_keyword_overlap_ratio_no_meaningful_tokens() -> None:
    assert keyword_overlap_ratio("de la", "anything") == 1.0


def test_keyword_overlap_ratio_partial_match() -> None:
    ratio = keyword_overlap_ratio("Ventas regionales", "Plan de ventas regionales 2026")
    assert 0.3 < ratio <= 1.0


def test_merge_rels_appends_new_ids() -> None:
    orig = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_RELS_NS}">
  <Relationship Id="rId1" Type="t1" Target="a"/>
</Relationships>""".encode()
    new = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_RELS_NS}">
  <Relationship Id="rId2" Type="t2" Target="b"/>
</Relationships>""".encode()
    merged = _merge_rels(orig, new).decode()
    assert 'Id="rId1"' in merged
    assert 'Id="rId2"' in merged


def test_merge_content_types_adds_override() -> None:
    orig = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{_CT_NS}">
  <Override PartName="/xl/workbook.xml" ContentType="ct1"/>
</Types>""".encode()
    new = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{_CT_NS}">
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="ct2"/>
</Types>""".encode()
    merged = _merge_content_types(orig, new).decode()
    assert "/xl/workbook.xml" in merged
    assert "/xl/worksheets/sheet2.xml" in merged


def test_extract_and_inject_drawing_elements() -> None:
    orig = b'<worksheet><drawing r:id="rId5"/></worksheet>'
    new = b"<worksheet></worksheet>"
    elements = _extract_drawing_elements(orig)
    patched = _inject_drawing_elements(new, elements).decode()
    assert "drawing" in patched
    assert 'r:id="rId5"' in patched
