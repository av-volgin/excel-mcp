"""Template support — open existing .xlsx templates and fill with data."""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

from ..common.workbook import open_for_metadata


def open_template(
    template_path: str,
    output_path: str,
) -> dict[str, Any]:
    """Copy a template .xlsx file and return its structure.

    Creates a working copy of the template at output_path.
    All formatting, formulas, conditional formatting, and print areas
    are preserved from the template. Use write_data/write_formula
    to fill in data afterwards.

    Args:
        template_path: Absolute path to the template .xlsx file.
        output_path: Absolute path for the output copy.

    Returns:
        {template, output, sheets: [{name, rows, cols, has_tables, table_names}]}
    """
    src = Path(template_path)
    if not src.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    # Ensure output directory exists
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    # Copy template to output
    shutil.copy2(template_path, output_path)

    # Read structure of the copy
    wb = open_for_metadata(output_path)
    try:
        sheets = []
        for name in wb.sheetnames:
            ws = wb[name]
            tables = list(ws.tables.keys()) if hasattr(ws, "tables") else []
            sheets.append({
                "name": name,
                "rows": ws.max_row or 0,
                "cols": ws.max_column or 0,
                "has_tables": len(tables) > 0,
                "table_names": tables,
            })
        return {
            "template": template_path,
            "output": output_path,
            "sheets": sheets,
        }
    finally:
        wb.close()
