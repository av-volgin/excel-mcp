"""Excel MCP Server — read, write, and analyze Excel files."""

from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP


def create_server(tools: str = "all") -> FastMCP:
    """Create and configure the MCP server with selected tool groups.

    Args:
        tools: Which tool groups to register. One of: reader, writer, analyzer, all.
    """
    mcp = FastMCP(
        "excel-mcp",
        instructions=(
            "MCP server for Excel (.xlsx) files. "
            "Provides tools for bulk data reading (openpyxl), "
            "formula-aware analysis (Formualizer), and "
            "structured writing with verification. "
            "Use list_sheets first to understand file structure."
        ),
    )

    if tools in ("reader", "all"):
        from .reader.tools import register_reader_tools

        register_reader_tools(mcp)

    if tools in ("analyzer", "all"):
        from .analyzer.tools import register_analyzer_tools

        register_analyzer_tools(mcp)

    if tools in ("writer", "all"):
        try:
            from .writer.tools import register_writer_tools

            register_writer_tools(mcp)
        except ImportError:
            # Writer module not yet implemented or dependencies missing
            pass

    return mcp


def main():
    """Entry point for the MCP server."""
    parser = argparse.ArgumentParser(description="Excel MCP Server")
    parser.add_argument(
        "--tools",
        choices=["reader", "writer", "analyzer", "all"],
        default="all",
        help="Tool groups to register (default: all)",
    )
    parser.add_argument(
        "--transport",
        choices=["stdio", "sse"],
        default="stdio",
        help="Transport protocol (default: stdio)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for SSE transport (default: 8000)",
    )

    args = parser.parse_args()
    mcp = create_server(tools=args.tools)

    if args.transport == "sse":
        mcp.run(transport="sse", host="0.0.0.0", port=args.port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
