"""MCP tool definitions for the reader module."""

from __future__ import annotations

from typing import Any

from . import reader


def register_reader_tools(mcp):
    """Register reader tools on the MCP server instance."""

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
