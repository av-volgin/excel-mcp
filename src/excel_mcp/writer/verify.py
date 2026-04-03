"""Post-write verification — re-open generated files and validate."""

from __future__ import annotations

import os
from typing import Any

from ..common.utils import cell_ref_to_coords, classify_value
from ..common.workbook import open_for_read


def verify_workbook(
    file_path: str,
    checks: list[dict[str, Any]],
) -> dict[str, Any]:
    """Re-open a workbook and verify its structure and values.

    Runs a list of checks and returns pass/fail for each.

    Args:
        file_path: Absolute path to the .xlsx file.
        checks: List of check definitions. Each check has a 'type' and type-specific params:
            type: "row_count" — {sheet, expected} — verify row count of a sheet
            type: "cell_value" — {sheet, cell, expected} — verify a cell value (tolerance 0.01)
            type: "sheet_exists" — {sheet} — verify a sheet exists
            type: "file_size_min" — {min_bytes} — verify minimum file size

    Returns:
        {valid: bool, file_size: int, results: [{check_type, passed, detail}]}
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")

    file_size = os.path.getsize(file_path)
    results = []
    all_passed = True

    for check in checks:
        check_type = check["type"]

        if check_type == "row_count":
            result = _check_row_count(file_path, check["sheet"], check["expected"])
        elif check_type == "cell_value":
            result = _check_cell_value(
                file_path, check["sheet"], check["cell"], check["expected"]
            )
        elif check_type == "sheet_exists":
            result = _check_sheet_exists(file_path, check["sheet"])
        elif check_type == "file_size_min":
            result = _check_file_size(file_size, check["min_bytes"])
        else:
            result = {"check_type": check_type, "passed": False, "detail": "Unknown check type"}

        if not result["passed"]:
            all_passed = False
        results.append(result)

    return {
        "valid": all_passed,
        "file_size": file_size,
        "results": results,
    }


def _check_row_count(file_path: str, sheet_name: str, expected: int) -> dict[str, Any]:
    """Check that a sheet has the expected number of rows."""
    wb = open_for_read(file_path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return {
                "check_type": "row_count",
                "passed": False,
                "detail": f"Sheet '{sheet_name}' not found",
            }
        ws = wb[sheet_name]
        actual = ws.max_row or 0
        return {
            "check_type": "row_count",
            "passed": actual == expected,
            "detail": f"sheet={sheet_name}, expected={expected}, actual={actual}",
        }
    finally:
        wb.close()


def _check_cell_value(
    file_path: str, sheet_name: str, cell_ref: str, expected: Any
) -> dict[str, Any]:
    """Check that a cell has the expected value."""
    wb = open_for_read(file_path, data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            return {
                "check_type": "cell_value",
                "passed": False,
                "detail": f"Sheet '{sheet_name}' not found",
            }
        ws = wb[sheet_name]
        row, col = cell_ref_to_coords(cell_ref)
        actual = ws.cell(row=row, column=col).value

        if isinstance(expected, (int, float)) and isinstance(actual, (int, float)):
            match = abs(expected - actual) < 0.01
        else:
            match = expected == actual

        actual_ser, _ = classify_value(actual)
        return {
            "check_type": "cell_value",
            "passed": match,
            "detail": f"cell={cell_ref}, expected={expected}, actual={actual_ser}",
        }
    finally:
        wb.close()


def _check_sheet_exists(file_path: str, sheet_name: str) -> dict[str, Any]:
    """Check that a sheet exists in the workbook."""
    wb = open_for_read(file_path, data_only=True)
    try:
        exists = sheet_name in wb.sheetnames
        return {
            "check_type": "sheet_exists",
            "passed": exists,
            "detail": f"sheet={sheet_name}, exists={exists}",
        }
    finally:
        wb.close()


def _check_file_size(actual_size: int, min_bytes: int) -> dict[str, Any]:
    """Check that file is at least min_bytes."""
    return {
        "check_type": "file_size_min",
        "passed": actual_size >= min_bytes,
        "detail": f"actual={actual_size}, min={min_bytes}",
    }
