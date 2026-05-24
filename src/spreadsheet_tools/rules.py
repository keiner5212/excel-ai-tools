from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

from openpyxl.utils import column_index_from_string, get_column_letter

from spreadsheet_tools.utils import GENERIC_NAME_RE, STOPWORDS, keyword_overlap_ratio

_CELL_ADDR_RE = re.compile(r"^([A-Za-z]+)(\d+)$")

# Type alias used throughout this module
MergeLookup = dict[str, tuple[str, str]]


@dataclass
class RuleResult:
    rule: str
    passed: bool
    message: str
    addresses: list[str] = field(default_factory=list)


def _parse_addr(address: str) -> tuple[str, int]:
    m = _CELL_ADDR_RE.match(address.strip())
    if not m:
        raise ValueError(f"Invalid cell address: {address!r}")
    return m.group(1).upper(), int(m.group(2))


def _cells_in_range(from_addr: str, to_addr: str) -> list[str]:
    """Enumerate all cell addresses from top-left to bottom-right."""
    from_col_str, from_row = _parse_addr(from_addr)
    to_col_str, to_row = _parse_addr(to_addr)
    from_col = column_index_from_string(from_col_str)
    to_col = column_index_from_string(to_col_str)
    return [
        f"{get_column_letter(col)}{row}"
        for row in range(from_row, to_row + 1)
        for col in range(from_col, to_col + 1)
    ]


def _resolve_address(address: str, merge_lookup: MergeLookup) -> str:
    """Redirect a slave cell to its merge master; return unchanged if not a slave."""
    if address in merge_lookup:
        _range_str, master = merge_lookup[address]
        return master
    return address


def _cell_value(
    sheet: Any, address: str, merge_lookup: MergeLookup
) -> object | None:
    """Read cell value, resolving merged-cell slaves to their master.

    Slave cells always have value=None in openpyxl; the actual value lives
    on the master.  Without this redirect every rule applied to a slave
    address would incorrectly see an empty cell.
    """
    effective = _resolve_address(address, merge_lookup)
    try:
        return sheet[effective].value
    except Exception:
        return None


def _is_empty(value: object | None) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def apply_rule(
    rule_text: str,
    sheet: Any,
    merge_lookup: MergeLookup | None = None,
) -> RuleResult:
    """Parse and evaluate one validation rule against a worksheet.

    Rule syntax (colon-separated):

      not-empty:CELL
        Cell must have a non-blank value.

      not-empty-range:FROM_CELL:TO_CELL
        Every cell in the rectangular range must be non-blank.

      price-min-max:MIN_CELL:MAX_CELL
        MIN < MAX (both must be numeric).

      name-matches-desc:NAME_CELL:DESC_CELL
        Keyword overlap between name and description must be >= 30%.

      numeric-range:CELL:MIN:MAX
        Cell value must satisfy MIN <= value <= MAX.

      string-contains:CELL:SUBSTRING
        Cell string value must contain SUBSTRING (case-insensitive).

      no-generic-name:CELL
        Cell value must NOT match the pattern 'Estrategia de X N'.

    All rules correctly handle merged cells: slave addresses are
    transparently redirected to their merge master.
    """
    lookup: MergeLookup = merge_lookup or {}
    stripped = rule_text.strip()

    def _val(addr: str) -> object | None:
        return _cell_value(sheet, addr.upper(), lookup)

    if stripped.startswith("not-empty:"):
        addr = stripped[len("not-empty:"):].strip().upper()
        val = _val(addr)
        passed = not _is_empty(val)
        return RuleResult(
            rule=rule_text,
            passed=passed,
            message=f"{addr}={val!r}" if passed else f"{addr} is empty",
            addresses=[addr],
        )

    if stripped.startswith("not-empty-range:"):
        _, _, range_part = stripped.partition(":")
        from_addr, _, to_addr = range_part.partition(":")
        cells = _cells_in_range(from_addr.strip(), to_addr.strip())
        # Resolve each address through the merge lookup before checking
        empty = [a for a in cells if _is_empty(_val(a))]
        passed = len(empty) == 0
        return RuleResult(
            rule=rule_text,
            passed=passed,
            message=(
                f"All {len(cells)} cells filled"
                if passed
                else f"{len(empty)} empty: {empty}"
            ),
            addresses=empty if not passed else cells,
        )

    if stripped.startswith("price-min-max:"):
        _, _, args = stripped.partition(":")
        min_addr, _, max_addr = args.partition(":")
        min_addr = min_addr.strip().upper()
        max_addr = max_addr.strip().upper()
        min_val = _val(min_addr)
        max_val = _val(max_addr)
        try:
            passed = float(min_val) < float(max_val)  # type: ignore[arg-type]
        except (TypeError, ValueError):
            passed = False
        return RuleResult(
            rule=rule_text,
            passed=passed,
            message=(
                f"{min_addr}={min_val} < {max_addr}={max_val}"
                if passed
                else f"{min_addr}={min_val!r} must be < {max_addr}={max_val!r}"
            ),
            addresses=[min_addr, max_addr],
        )

    if stripped.startswith("name-matches-desc:"):
        _, _, args = stripped.partition(":")
        name_addr, _, desc_addr = args.partition(":")
        name_addr = name_addr.strip().upper()
        desc_addr = desc_addr.strip().upper()
        name_val = _val(name_addr)
        desc_val = _val(desc_addr)
        if not name_val or not desc_val:
            return RuleResult(
                rule=rule_text,
                passed=False,
                message=f"Cannot check: {name_addr}={name_val!r}, {desc_addr}={desc_val!r}",
                addresses=[name_addr, desc_addr],
            )
        ratio = keyword_overlap_ratio(str(name_val), str(desc_val))
        passed = ratio >= 0.3
        return RuleResult(
            rule=rule_text,
            passed=passed,
            message=(
                f"Keyword overlap {ratio:.0%}"
                if passed
                else (
                    f"Mismatch: {name_addr}={name_val!r} not reflected in "
                    f"{desc_addr} (overlap {ratio:.0%})"
                )
            ),
            addresses=[name_addr, desc_addr],
        )

    if stripped.startswith("numeric-range:"):
        parts = rule_text.split(":", maxsplit=3)
        if len(parts) != 4:
            raise ValueError(f"numeric-range requires CELL:MIN:MAX — got {rule_text!r}")
        addr = parts[1].strip().upper()
        min_bound = float(parts[2].strip())
        max_bound = float(parts[3].strip())
        val = _val(addr)
        try:
            passed = min_bound <= float(val) <= max_bound  # type: ignore[arg-type]
        except (TypeError, ValueError):
            passed = False
        return RuleResult(
            rule=rule_text,
            passed=passed,
            message=(
                f"{addr}={val} in [{min_bound}, {max_bound}]"
                if passed
                else f"{addr}={val!r} not in [{min_bound}, {max_bound}]"
            ),
            addresses=[addr],
        )

    if stripped.startswith("string-contains:"):
        parts = rule_text.split(":", maxsplit=2)
        if len(parts) != 3:
            raise ValueError(
                f"string-contains requires CELL:SUBSTRING — got {rule_text!r}"
            )
        addr = parts[1].strip().upper()
        substring = parts[2]
        val = _val(addr)
        val_str = str(val) if val is not None else ""
        passed = substring.lower() in val_str.lower()
        return RuleResult(
            rule=rule_text,
            passed=passed,
            message=(
                f"{addr} contains {substring!r}"
                if passed
                else f"{addr}={val!r} does not contain {substring!r}"
            ),
            addresses=[addr],
        )

    if stripped.startswith("no-generic-name:"):
        addr = stripped[len("no-generic-name:"):].strip().upper()
        val = _val(addr)
        val_str = str(val) if val else ""
        is_generic = bool(GENERIC_NAME_RE.match(val_str))
        passed = not is_generic
        return RuleResult(
            rule=rule_text,
            passed=passed,
            message=(
                f"{addr}={val!r} is not generic"
                if passed
                else f"{addr}={val!r} matches generic 'Estrategia de X N' pattern"
            ),
            addresses=[addr],
        )

    raise ValueError(f"Unknown rule type in: {rule_text!r}")
