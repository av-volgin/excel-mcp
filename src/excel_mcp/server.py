"""Excel MCP Server — read Excel files with formula-aware parsing."""

from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from . import analyzer, reader

mcp = FastMCP(
    "excel-mcp",
    instructions=(
        "MCP server for reading Excel (.xlsx) files. "
        "Provides tools for bulk data extraction (openpyxl) and "
        "formula-aware analysis (Formualizer). "
        "Use list_sheets first to understand file structure, "
        "then read_sheet for tabular data or get_formulas for formula analysis."
    ),
)


@mcp.tool()
def list_sheets(file_path: str) -> list[dict[str, Any]]:
    """List all sheets in an Excel workbook with metadata.

    Returns sheet names, dimensions, and table information.
    Use this first to understand the structure of a file before reading data.

    Args:
        file_path: Absolute path to the .xlsx file.
    """
    return reader.list_sheets(file_path)


@mcp.tool()
def read_sheet(
    file_path: str,
    sheet_name: str | None = None,
    header_row: int | None = None,
    start_row: int | None = None,
    max_rows: int = 500,
    columns: list[str] | None = None,
    skip_totals: bool = True,
) -> dict[str, Any]:
    """Read a sheet as a table with automatic header detection and pagination.

    Best for: flat data tables (transactions, employee lists, timesheets).
    Automatically detects headers and skips total/summary rows.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Sheet to read (default: first sheet).
        header_row: Row number containing headers (default: auto-detect).
        start_row: First data row to read (default: header_row + 1).
        max_rows: Maximum rows to return (default: 500). Use with start_row for pagination.
        columns: Filter to specific column names only.
        skip_totals: Skip rows that look like totals — bold, "Total", "Итого" (default: true).
    """
    return reader.read_sheet(
        file_path,
        sheet_name=sheet_name,
        header_row=header_row,
        start_row=start_row,
        max_rows=max_rows,
        columns=columns,
        skip_totals=skip_totals,
    )


@mcp.tool()
def read_cell_range(
    file_path: str,
    sheet_name: str,
    range: str,
) -> dict[str, Any]:
    """Read raw cell values from a specific range.

    Best for: analytical sheets where data is not in a flat table format.
    Returns individual cells with values and types.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Sheet to read from.
        range: Cell range in Excel notation, e.g. "A1:D10" or "B5".
    """
    return reader.read_cell_range(file_path, sheet_name, range)


@mcp.tool()
def get_formulas(
    file_path: str,
    sheet_name: str | None = None,
    cell_range: str | None = None,
) -> dict[str, Any]:
    """Extract formulas from cells with optional AST parsing.

    With Formualizer installed: returns formula text + AST tree + referenced cells.
    Without Formualizer: returns formula text only.

    The 'engine' field in the response indicates which backend was used.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Sheet to analyze (default: first sheet).
        cell_range: Limit to specific range, e.g. "A1:D10" (default: entire sheet).
    """
    return analyzer.get_formulas(file_path, sheet_name=sheet_name, cell_range=cell_range)


@mcp.tool()
def get_cell_value(
    file_path: str,
    sheet_name: str,
    row: int,
    col: int,
) -> dict[str, Any]:
    """Read a single cell value by sheet name and coordinates.

    Fast targeted read — useful for checking specific cells in analytical reports.
    Row and column are 1-based (row=1, col=1 = cell A1).

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Sheet name.
        row: Row number (1-based).
        col: Column number (1-based).
    """
    return analyzer.get_cell_value(file_path, sheet_name, row, col)


@mcp.tool()
def validate_totals(
    file_path: str,
    sheet_name: str,
    checks: list[dict[str, Any]],
) -> dict[str, Any]:
    """Validate cell values against expected values (e.g. Excel vs database).

    Compares actual cell values with expected values. Useful for verifying
    data integrity after import. Numbers are compared with 0.01 tolerance.

    Args:
        file_path: Absolute path to the .xlsx file.
        sheet_name: Sheet to validate.
        checks: List of checks, each with 'cell' (e.g. "B10") and 'expected_value'.
    """
    return analyzer.validate_totals(file_path, sheet_name, checks)


def main():
    """Entry point for the MCP server."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
