"""Exception hierarchy for Excel MCP server."""

from __future__ import annotations


class ExcelMCPError(Exception):
    """Base exception for all Excel MCP errors."""


class FileValidationError(ExcelMCPError):
    """File does not exist or has unsupported format."""


class RangeError(ExcelMCPError):
    """Invalid cell reference or range."""


class SheetNotFoundError(ExcelMCPError):
    """Requested sheet does not exist in workbook."""
