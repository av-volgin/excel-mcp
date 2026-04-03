"""MCP tool definitions for the analyzer module."""

from __future__ import annotations

from typing import Any

from . import analyzer


def register_analyzer_tools(mcp):
    """Register analyzer tools on the MCP server instance."""

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
