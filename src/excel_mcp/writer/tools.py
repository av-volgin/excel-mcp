"""MCP tool definitions for the writer module."""

from __future__ import annotations

from typing import Any


def register_writer_tools(mcp):
    """Register writer tools on the MCP server instance."""

    from . import formatter, templates, verify, writer

    @mcp.tool()
    def create_workbook(
        file_path: str,
        sheets: list[str] | None = None,
    ) -> dict[str, Any]:
        """Create a new empty Excel workbook.

        Args:
            file_path: Absolute path for the new .xlsx file.
            sheets: List of sheet names to create (default: ["Sheet1"]).
        """
        return writer.create_workbook(file_path, sheets)

    @mcp.tool()
    def write_data(
        file_path: str,
        sheet_name: str,
        range: str,
        data: list[list[Any]],
    ) -> dict[str, Any]:
        """Write a 2D array of values into a cell range (batch operation).

        Efficient for writing many rows at once. Values can be strings, numbers,
        dates, or None. Formulas should use write_formula instead.

        Args:
            file_path: Absolute path to the .xlsx file.
            sheet_name: Target sheet name.
            range: Starting cell, e.g. "A1" or "B2".
            data: 2D array — list of rows, each row is a list of values.
        """
        return writer.write_data(file_path, sheet_name, range, data)

    @mcp.tool()
    def write_formula(
        file_path: str,
        sheet_name: str,
        cell: str,
        formula: str,
        number_format: str | None = None,
    ) -> dict[str, Any]:
        """Write an Excel formula to a cell.

        Formulas are written as text — Excel evaluates them on open.
        Supports any Excel formula: SUM, SUMIFS, VLOOKUP, IF, cross-sheet refs.

        Args:
            file_path: Absolute path to the .xlsx file.
            sheet_name: Target sheet name.
            cell: Cell reference, e.g. "D5".
            formula: Excel formula, e.g. "=SUM(A1:A10)" or "=SUMIFS(B:B,A:A,\\"Revenue\\")".
            number_format: Optional display format, e.g. "#,##0.00" or "0.0%".
        """
        return writer.write_formula(file_path, sheet_name, cell, formula, number_format)

    @mcp.tool()
    def format_range(
        file_path: str,
        sheet_name: str,
        range: str,
        styles: dict[str, Any],
    ) -> dict[str, Any]:
        """Apply formatting (font, fill, borders, number format) to a range.

        Styles dict supports:
        - font: {bold, italic, underline, size, color, name}
        - fill: {color, type}
        - border: "thin" or {left, right, top, bottom}
        - number_format: "#,##0.00"
        - alignment: {horizontal, vertical, wrap_text}

        Args:
            file_path: Absolute path to the .xlsx file.
            sheet_name: Target sheet name.
            range: Cell range, e.g. "A1:D10".
            styles: Style dictionary (see description above).
        """
        return formatter.format_range(file_path, sheet_name, range, styles)

    @mcp.tool()
    def create_table(
        file_path: str,
        sheet_name: str,
        range: str,
        table_name: str,
        style: str = "TableStyleMedium2",
        totals_row: bool = False,
    ) -> dict[str, Any]:
        """Create an Excel Table (ListObject) from a data range.

        Tables enable structured references, auto-filters, and styled formatting.
        The first row of the range must contain headers.

        Args:
            file_path: Absolute path to the .xlsx file.
            sheet_name: Sheet containing the data.
            range: Data range including headers, e.g. "A1:D20".
            table_name: Display name for the table (must be unique in workbook).
            style: Table style (default: "TableStyleMedium2").
            totals_row: Add a totals row below the table (default: false).
        """
        return writer.create_table(file_path, sheet_name, range, table_name, style, totals_row)

    @mcp.tool()
    def add_conditional_format(
        file_path: str,
        sheet_name: str,
        range: str,
        rule: dict[str, Any],
    ) -> dict[str, Any]:
        """Add conditional formatting to a cell range.

        Rule types:
        - cell_is: {type, operator, value, font?, fill?}
        - color_scale: {type, start_color, end_color}
        - data_bar: {type, color}

        Args:
            file_path: Absolute path to the .xlsx file.
            sheet_name: Target sheet name.
            range: Cell range, e.g. "C2:C100".
            rule: Rule definition (see description above).
        """
        return formatter.add_conditional_format(file_path, sheet_name, range, rule)

    @mcp.tool()
    def open_template(
        template_path: str,
        output_path: str,
    ) -> dict[str, Any]:
        """Open an Excel template and create a working copy.

        Copies the template file preserving all formatting, formulas,
        conditional formatting, and print areas. Use write_data and
        write_formula to fill in data afterwards.

        Args:
            template_path: Absolute path to the template .xlsx file.
            output_path: Absolute path for the output copy.
        """
        return templates.open_template(template_path, output_path)

    @mcp.tool()
    def verify_workbook(
        file_path: str,
        checks: list[dict[str, Any]],
    ) -> dict[str, Any]:
        """Verify a generated workbook by re-opening and checking structure/values.

        Run this after writing to confirm the file is valid. Check types:
        - row_count: {type: "row_count", sheet: "Data", expected: 100}
        - cell_value: {type: "cell_value", sheet: "Summary", cell: "B5", expected: 12345.67}
        - sheet_exists: {type: "sheet_exists", sheet: "Revenue"}
        - file_size_min: {type: "file_size_min", min_bytes: 5000}

        Args:
            file_path: Absolute path to the .xlsx file.
            checks: List of check definitions (see description above).
        """
        return verify.verify_workbook(file_path, checks)
