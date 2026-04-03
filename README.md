# excel-mcp

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

An MCP server for reading Excel files with formula-aware parsing.

Unlike other Excel MCP servers that only read cell values, **excel-mcp** also understands formulas — it can extract formula AST trees, build dependency graphs, and validate computed values. Powered by [openpyxl](https://openpyxl.readthedocs.io/) for reliable data extraction and [Formualizer](https://github.com/PSU3D0/formualizer) for formula analysis.

## Features

- **Bulk data reading** — read sheets as tables with automatic header detection, pagination, and total-row skipping
- **Cell-level access** — read specific cells or ranges by coordinates
- **Formula extraction** — get formula text, AST trees, and cell dependencies (with Formualizer)
- **Data validation** — compare Excel values against expected values (e.g. database totals)
- **Graceful degradation** — works with openpyxl alone; Formualizer adds formula intelligence

## Installation

```bash
# Basic (openpyxl only)
pip install excel-mcp

# With formula analysis (+ Formualizer)
pip install excel-mcp[formulas]

# Or run directly with uvx
uvx excel-mcp
```

## Configuration

### Claude Desktop / Cursor

Add to your MCP config:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["excel-mcp"]
    }
  }
}
```

### With Formualizer support

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["--from", "excel-mcp[formulas]", "excel-mcp"]
    }
  }
}
```

## Tools

### `list_sheets`

List all sheets in a workbook with metadata (dimensions, tables).

```
list_sheets(file_path="/path/to/report.xlsx")
→ [{name: "Data", rows: 500, cols: 12, has_tables: true, table_names: ["Table1"]}]
```

### `read_sheet`

Read a sheet as a table with automatic header detection and pagination.

```
read_sheet(file_path="/path/to/data.xlsx", sheet_name="Employees", max_rows=100)
→ {headers: ["Name", "Department", "Salary"], rows: [{...}], total_rows: 250, has_more: true}
```

Features:
- Auto-detects header row (first row with ≥2 non-empty cells)
- Skips total/summary rows (bold text, "Total", "Итого")
- Pagination via `max_rows` + `start_row`
- Column filtering via `columns` parameter

### `read_cell_range`

Read raw cell values from a specific range.

```
read_cell_range(file_path="/path/to/report.xlsx", sheet_name="Summary", range="A1:D10")
→ {cells: [{ref: "A1", row: 1, col: 1, value: "Revenue", type: "string"}, ...]}
```

### `get_formulas`

Extract formulas with optional AST parsing.

```
get_formulas(file_path="/path/to/report.xlsx", sheet_name="P&L", cell_range="D2:D20")
→ {formulas: [{cell: "D2", formula_text: "=B2-C2", formula_ast: {...}}], engine: "formualizer"}
```

The `engine` field indicates whether Formualizer ("formualizer") or openpyxl ("openpyxl") was used. Without Formualizer, `formula_ast` and `referenced_cells` will be null.

### `get_cell_value`

Read a single cell by coordinates (1-based).

```
get_cell_value(file_path="/path/to/report.xlsx", sheet_name="Summary", row=5, col=2)
→ {ref: "B5", value: 170000, type: "number", formula: "=SUM(B2:B4)", engine: "formualizer"}
```

### `validate_totals`

Compare cell values against expected values.

```
validate_totals(
  file_path="/path/to/report.xlsx",
  sheet_name="Summary",
  checks=[{cell: "B10", expected_value: 1234567.89}]
)
→ {valid: true, results: [{cell: "B10", expected: 1234567.89, actual: 1234567.89, match: true}]}
```

Numbers are compared with 0.01 tolerance for floating-point precision.

## Use Cases

1. **Import data from Excel to a database** — use `read_sheet` to extract tabular data, then insert into PostgreSQL/SQLite
2. **Understand Excel report logic** — use `get_formulas` to see how values are computed (SUMIFS, VLOOKUP, etc.)
3. **Validate data after import** — use `validate_totals` to verify key figures match between Excel and your database
4. **Read analytical dashboards** — use `read_cell_range` for non-tabular sheets where data is scattered across cells
5. **Audit spreadsheets** — use `get_formulas` to find all formulas and their dependencies

## Architecture

```
┌─────────────────────────────────────────┐
│           excel-mcp (MCP Server)        │
│                                         │
│  ┌─────────────┐  ┌──────────────────┐  │
│  │  reader.py   │  │  analyzer.py     │  │
│  │  (openpyxl)  │  │  (Formualizer)   │  │
│  │              │  │  + openpyxl      │  │
│  │ • list_sheets│  │  fallback        │  │
│  │ • read_sheet │  │                  │  │
│  │ • read_range │  │ • get_formulas   │  │
│  │              │  │ • get_cell_value  │  │
│  │  Bulk data   │  │ • validate_totals│  │
│  │  extraction  │  │                  │  │
│  └─────────────┘  │  Formula-aware    │  │
│                    │  analysis         │  │
│                    └──────────────────┘  │
└─────────────────────────────────────────┘
```

**openpyxl** (required) — reliable workhorse for reading/writing Excel files. Handles bulk data extraction, header detection, and cell value reading.

**Formualizer** (optional) — Rust-based Excel formula engine with Python bindings. Adds AST parsing, dependency graphs, and fast targeted cell reads. Install with `pip install excel-mcp[formulas]`.

## Development

```bash
git clone https://github.com/avolgin/excel-mcp.git
cd excel-mcp
uv venv .venv && source .venv/bin/activate
uv pip install -e ".[dev]"
pytest
ruff check src/
```

## License

MIT — see [LICENSE](LICENSE).
