"""Tests for reader module (openpyxl-based)."""

from pathlib import Path

import pytest

from excel_mcp.reader import list_sheets, read_cell_range, read_sheet  # via reader/__init__.py

FIXTURES = Path(__file__).parent / "fixtures"


class TestListSheets:
    def test_simple_file(self):
        result = list_sheets(str(FIXTURES / "simple_table.xlsx"))
        assert len(result) == 1
        assert result[0]["name"] == "Employees"
        assert result[0]["rows"] >= 6
        assert result[0]["cols"] >= 4

    def test_multi_sheet(self):
        result = list_sheets(str(FIXTURES / "multi_sheet.xlsx"))
        names = [s["name"] for s in result]
        assert "Transactions" in names
        assert "Summary" in names
        assert "Notes" in names
        assert len(result) == 3

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            list_sheets("/nonexistent/file.xlsx")


class TestReadSheet:
    def test_auto_detect_headers(self):
        result = read_sheet(str(FIXTURES / "simple_table.xlsx"))
        assert "Name" in result["headers"]
        assert "Salary" in result["headers"]
        assert result["sheet_name"] == "Employees"

    def test_data_rows(self):
        result = read_sheet(str(FIXTURES / "simple_table.xlsx"))
        names = [r["Name"] for r in result["rows"]]
        assert "Alice" in names
        assert "Bob" in names

    def test_skip_totals(self):
        result = read_sheet(str(FIXTURES / "simple_table.xlsx"), skip_totals=True)
        names = [r.get("Name") for r in result["rows"]]
        assert "Total" not in names

    def test_include_totals(self):
        result = read_sheet(str(FIXTURES / "simple_table.xlsx"), skip_totals=False)
        names = [r.get("Name") for r in result["rows"]]
        assert "Total" in names

    def test_pagination(self):
        result = read_sheet(str(FIXTURES / "simple_table.xlsx"), max_rows=2)
        assert len(result["rows"]) == 2
        assert result["has_more"] is True

    def test_column_filter(self):
        result = read_sheet(
            str(FIXTURES / "simple_table.xlsx"),
            columns=["Name", "Salary"],
        )
        assert result["headers"] == ["Name", "Salary"]
        for row in result["rows"]:
            assert "Department" not in row

    def test_specific_sheet(self):
        result = read_sheet(str(FIXTURES / "multi_sheet.xlsx"), sheet_name="Transactions")
        assert result["sheet_name"] == "Transactions"
        assert "Description" in result["headers"]


class TestReadCellRange:
    def test_single_cell(self):
        result = read_cell_range(
            str(FIXTURES / "simple_table.xlsx"), "Employees", "A1"
        )
        assert len(result["cells"]) == 1
        assert result["cells"][0]["value"] == "Name"
        assert result["cells"][0]["ref"] == "A1"

    def test_range(self):
        result = read_cell_range(
            str(FIXTURES / "simple_table.xlsx"), "Employees", "A1:B3"
        )
        assert len(result["cells"]) == 6  # 3 rows × 2 cols
        values = [c["value"] for c in result["cells"]]
        assert "Name" in values
        assert "Alice" in values

    def test_types(self):
        result = read_cell_range(
            str(FIXTURES / "simple_table.xlsx"), "Employees", "C2"
        )
        cell = result["cells"][0]
        assert cell["type"] == "number"
        assert cell["value"] == 95000
