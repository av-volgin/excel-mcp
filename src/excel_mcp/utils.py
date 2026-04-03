"""Shared utilities for Excel MCP server."""

from __future__ import annotations

import re
from datetime import date, datetime, time
from typing import Any


def cell_ref_to_coords(ref: str) -> tuple[int, int]:
    """Convert Excel cell reference (e.g. 'B5') to (row, col) 1-based tuple."""
    match = re.match(r"^([A-Z]+)(\d+)$", ref.upper())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")
    col_str, row_str = match.groups()
    col = 0
    for ch in col_str:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return int(row_str), col


def coords_to_cell_ref(row: int, col: int) -> str:
    """Convert (row, col) 1-based tuple to Excel cell reference (e.g. 'B5')."""
    result = ""
    c = col
    while c > 0:
        c, remainder = divmod(c - 1, 26)
        result = chr(65 + remainder) + result
    return f"{result}{row}"


def parse_range(range_str: str) -> tuple[tuple[int, int], tuple[int, int]]:
    """Parse 'A1:D10' into ((min_row, min_col), (max_row, max_col))."""
    if ":" in range_str:
        start, end = range_str.split(":", 1)
        r1, c1 = cell_ref_to_coords(start)
        r2, c2 = cell_ref_to_coords(end)
        return (min(r1, r2), min(c1, c2)), (max(r1, r2), max(c1, c2))
    else:
        r, c = cell_ref_to_coords(range_str)
        return (r, c), (r, c)


def classify_value(value: Any) -> tuple[Any, str]:
    """Return (serializable_value, type_name) for a cell value."""
    if value is None:
        return None, "empty"
    if isinstance(value, bool):
        return value, "boolean"
    if isinstance(value, (int, float)):
        return value, "number"
    if isinstance(value, datetime):
        return value.isoformat(), "datetime"
    if isinstance(value, date):
        return value.isoformat(), "date"
    if isinstance(value, time):
        return value.isoformat(), "time"
    if isinstance(value, str):
        if value.startswith("="):
            return value, "formula"
        return value, "string"
    return str(value), "string"


def is_total_row(row_values: list[Any], bold_flags: list[bool] | None = None) -> bool:
    """Heuristic: detect if a row is a totals/summary row.

    Checks for:
    - Bold formatting on first cell
    - Keywords like "Total", "Итого", "SUBTOTAL", "Sum"
    """
    if bold_flags and any(bold_flags[:3]):
        return True
    for val in row_values[:3]:
        if isinstance(val, str):
            lower = val.strip().lower()
            if lower in ("total", "итого", "всего", "sum", "subtotal"):
                return True
            if "subtotal" in lower or "итого" in lower:
                return True
    return False
