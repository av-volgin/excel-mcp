"""Excel MCP Server — read, write, and analyze Excel files."""

from __future__ import annotations

import argparse

from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings


def create_server(
    tools: str = "all",
    *,
    host: str = "127.0.0.1",
    port: int = 8000,
    transport_security: TransportSecuritySettings | None = None,
) -> FastMCP:
    """Create and configure the MCP server with selected tool groups.

    Args:
        tools: Which tool groups to register. One of: reader, writer, analyzer, all.
        host: Host to bind for HTTP transports.
        port: Port to bind for HTTP transports.
        transport_security: Transport security settings (DNS rebinding protection etc.).
    """
    mcp = FastMCP(
        "excel-mcp",
        host=host,
        port=port,
        transport_security=transport_security,
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
            pass

    @mcp.custom_route("/healthz", methods=["GET"])
    async def healthz(request):  # noqa: ARG001
        from starlette.responses import JSONResponse

        return JSONResponse({"status": "ok"})

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
        choices=["stdio", "streamable-http"],
        default="stdio",
        help="Transport protocol (default: stdio)",
    )
    parser.add_argument(
        "--host",
        default="127.0.0.1",
        help="Host to bind for HTTP transport (default: 127.0.0.1)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8000,
        help="Port for HTTP transport (default: 8000)",
    )
    parser.add_argument(
        "--no-dns-rebinding-protection",
        action="store_true",
        help="Disable DNS rebinding protection (for Docker/reverse-proxy setups)",
    )

    args = parser.parse_args()

    transport_security = None
    if args.no_dns_rebinding_protection:
        transport_security = TransportSecuritySettings(
            enable_dns_rebinding_protection=False,
        )

    mcp = create_server(
        tools=args.tools,
        host=args.host,
        port=args.port,
        transport_security=transport_security,
    )

    mcp.run(transport=args.transport)


if __name__ == "__main__":
    main()
