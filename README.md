# excel-mcp

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.11+](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)

An MCP server for reading, writing, and analyzing Excel files with formula-aware parsing.

Unlike other Excel MCP servers that only read cell values, **excel-mcp** understands formulas — it can extract formula AST trees, build dependency graphs, and validate computed values. It also provides structured writing with batch operations, templates, and verification.

Powered by [openpyxl](https://openpyxl.readthedocs.io/) for reliable data operations and [Formualizer](https://github.com/PSU3D0/formualizer) for formula analysis.

## Features

- **Bulk data reading** — read sheets as tables with automatic header detection, pagination, and total-row skipping
- **Cell-level access** — read specific cells or ranges by coordinates
- **Formula extraction** — get formula text, AST trees, and cell dependencies (with Formualizer)
- **Structured writing** — create workbooks, write data in batches, add formulas, format cells, create tables
- **Template mode** — copy .xlsx templates preserving formatting and formulas, then fill with data
- **Verification** — re-open written files and verify row counts, spot values
- **Data validation** — compare Excel values against expected values (e.g. database totals)
- **Conditional formatting** — cell-is rules, color scales, data bars
- **Multiple transports** — stdio (default, for Claude Code/Cursor) and streamable-http (for Docker sidecars)
- **Graceful degradation** — works with openpyxl alone; Formualizer adds formula intelligence

## Installation

```bash
# Basic (reader + writer, openpyxl only)
pip install excel-mcp

# With formula analysis (+ Formualizer)
pip install "excel-mcp[analyzer]"

# Everything
pip install "excel-mcp[all]"

# Or run directly with uvx
uvx excel-mcp
```

## Configuration

### Claude Code / Claude Desktop / Cursor (stdio)

Add to your `.mcp.json` or MCP settings:

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

With Formualizer support:

```json
{
  "mcpServers": {
    "excel": {
      "command": "uvx",
      "args": ["--from", "excel-mcp[all]", "excel-mcp"]
    }
  }
}
```

### Docker (streamable-http)

For running as a sidecar container alongside AI gateways (OpenClaw, custom setups):

```yaml
# docker-compose.yml
services:
  mcp-excel:
    build:
      context: .
      dockerfile_inline: |
        FROM python:3.11-slim
        WORKDIR /app
        COPY . .
        RUN pip install --no-cache-dir .[all]
        EXPOSE 8000
        CMD ["excel-mcp", "--tools", "all", "--transport", "streamable-http", "--host", "0.0.0.0", "--port", "8000", "--no-dns-rebinding-protection"]
    container_name: mcp-excel
    mem_limit: 256m
    restart: unless-stopped
    volumes:
      - /path/to/excel/files:/data:ro
    healthcheck:
      test: ["CMD", "python", "-c", "import urllib.request; urllib.request.urlopen('http://localhost:8000/healthz')"]
      interval: 30s
      timeout: 15s
      start_period: 60s
      retries: 3
```

The MCP endpoint is available at `http://mcp-excel:8000/mcp` (POST).
Health check endpoint: `http://mcp-excel:8000/healthz` (GET).

### Tool groups

You can enable only the tool groups you need:

```bash
excel-mcp --tools reader      # 3 tools: list_sheets, read_sheet, read_cell_range
excel-mcp --tools writer      # 8 tools: create, write, format, tables, templates, verify
excel-mcp --tools analyzer    # 3 tools: get_formulas, get_cell_value, validate_totals
excel-mcp --tools all         # all 14 tools (default)
```

## Tools

### Reader (3 tools)

| Tool | Description |
|------|-------------|
| `list_sheets` | List all sheets with metadata (dimensions, tables) |
| `read_sheet` | Read a sheet as a table with auto headers, pagination, total-row skipping |
| `read_cell_range` | Read raw cell values from a specific range |

### Writer (8 tools)

| Tool | Description |
|------|-------------|
| `create_workbook` | Create a new .xlsx file with named sheets |
| `write_data` | Write a 2D array of data to a range (batch, not cell-by-cell) |
| `write_formula` | Write a formula to a cell with optional number format |
| `format_range` | Apply styles (font, fill, border, alignment, number format) |
| `create_table` | Create an Excel table with optional totals row |
| `add_conditional_format` | Add conditional formatting (cell-is, color scale, data bar) |
| `open_template` | Copy a .xlsx template preserving formatting and formulas |
| `verify_workbook` | Re-open a file and verify row counts, spot values |

### Analyzer (3 tools)

| Tool | Description |
|------|-------------|
| `get_formulas` | Extract formulas with optional AST parsing (Formualizer) |
| `get_cell_value` | Read a single cell by coordinates with formula info |
| `validate_totals` | Compare cell values against expected values (0.01 tolerance) |

## CLI options

```
excel-mcp [OPTIONS]

Options:
  --tools {reader,writer,analyzer,all}   Tool groups to register (default: all)
  --transport {stdio,streamable-http}    Transport protocol (default: stdio)
  --host HOST                            Host to bind for HTTP (default: 127.0.0.1)
  --port PORT                            Port for HTTP transport (default: 8000)
  --no-dns-rebinding-protection          Disable DNS rebinding protection (Docker/proxy)
```

## Architecture

```
┌──────────────────────────────────────────────────────┐
│                excel-mcp (MCP Server)                │
│                                                      │
│  ┌─────────────┐ ┌──────────────┐ ┌──────────────┐  │
│  │   reader     │ │   writer     │ │  analyzer    │  │
│  │  (openpyxl)  │ │  (openpyxl)  │ │ (Formualizer │  │
│  │              │ │              │ │  + openpyxl   │  │
│  │• list_sheets │ │• create_wb   │ │  fallback)   │  │
│  │• read_sheet  │ │• write_data  │ │              │  │
│  │• read_range  │ │• write_formula│ │• get_formulas│  │
│  │              │ │• format_range│ │• get_cell_val│  │
│  │  Bulk data   │ │• create_table│ │• validate    │  │
│  │  extraction  │ │• cond_format │ │              │  │
│  │              │ │• template    │ │  Formula-    │  │
│  │              │ │• verify      │ │  aware       │  │
│  └─────────────┘ └──────────────┘ └──────────────┘  │
│                                                      │
│  Transport: stdio | streamable-http (/mcp)           │
│  Health:    /healthz (GET, HTTP mode only)            │
└──────────────────────────────────────────────────────┘
```

## Development

```bash
git clone https://github.com/avolgin/excel-mcp.git
cd excel-mcp
uv venv .venv && source .venv/bin/activate
uv pip install -e ".[dev]"
pytest          # 50 tests
ruff check src/
```

## License

MIT — see [LICENSE](LICENSE).
