"""Tests for analyzer module (Formualizer + openpyxl fallback)."""

from pathlib import Path

import pytest

from excel_mcp.analyzer import (  # via analyzer/__init__.py
    get_cell_value,
    get_formulas,
    validate_totals,
)

FIXTURES = Path(__file__).parent / "fixtures"


class TestGetFormulas:
    def test_finds_formulas(self):
        result = get_formulas(str(FIXTURES / "formulas.xlsx"))
        assert len(result["formulas"]) > 0
        assert result["engine"] in ("formualizer", "openpyxl")
        assert result["sheet_name"] == "Revenue"

    def test_formula_text(self):
        result = get_formulas(str(FIXTURES / "formulas.xlsx"))
        formula_texts = [f["formula_text"] for f in result["formulas"]]
        # Should find SUM formulas
        assert any("SUM" in ft for ft in formula_texts)

    def test_cell_range_filter(self):
        result = get_formulas(
            str(FIXTURES / "formulas.xlsx"),
            cell_range="D2:D4",
        )
        # Only D2:D4 should have formulas (=B{i}-C{i})
        cells = [f["cell"] for f in result["formulas"]]
        assert all(c.startswith("D") for c in cells)

    def test_no_formulas_sheet(self):
        result = get_formulas(str(FIXTURES / "simple_table.xlsx"))
        # simple_table has no formulas (values only)
        assert len(result["formulas"]) == 0

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            get_formulas("/nonexistent/file.xlsx")


class TestGetCellValue:
    def test_read_value(self):
        result = get_cell_value(
            str(FIXTURES / "simple_table.xlsx"), "Employees", 2, 1
        )
        assert result["value"] == "Alice"
        assert result["type"] == "string"
        assert result["ref"] == "A2"
        assert result["engine"] in ("formualizer", "openpyxl")

    def test_read_number(self):
        result = get_cell_value(
            str(FIXTURES / "simple_table.xlsx"), "Employees", 2, 3
        )
        assert result["value"] == 95000
        assert result["type"] == "number"

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            get_cell_value("/nonexistent/file.xlsx", "Sheet1", 1, 1)


class TestValidateTotals:
    def test_matching_values(self):
        result = validate_totals(
            str(FIXTURES / "multi_sheet.xlsx"),
            "Summary",
            [
                {"cell": "B2", "expected_value": 242000},
                {"cell": "B3", "expected_value": 250000},
            ],
        )
        assert result["valid"] is True
        assert all(r["match"] for r in result["results"])

    def test_mismatched_value(self):
        result = validate_totals(
            str(FIXTURES / "multi_sheet.xlsx"),
            "Summary",
            [
                {"cell": "B2", "expected_value": 999999},
            ],
        )
        assert result["valid"] is False
        assert result["results"][0]["match"] is False
        assert result["results"][0]["actual"] == 242000

    def test_float_tolerance(self):
        result = validate_totals(
            str(FIXTURES / "multi_sheet.xlsx"),
            "Summary",
            [
                {"cell": "B2", "expected_value": 242000.005},
            ],
        )
        assert result["valid"] is True

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            validate_totals("/nonexistent/file.xlsx", "Sheet1", [])
