"""MCP Server for Excel file retrieval and analysis.

Provides tools for LLMs to explore, query, summarize, and analyze Excel
workbooks from 1 KB to 100 MB.  Designed for the Model Context Protocol
(MCP) so it can be used as a tool server by any MCP-compatible client.

Usage:
    python -m excel_mcp_server.server          # stdio transport (default)
"""

from __future__ import annotations

import json
import logging
from typing import Any

from mcp.server.fastmcp import FastMCP

from . import excel_reader

logger = logging.getLogger(__name__)

# ── Create the MCP server instance ───────────────────────────────────────

mcp = FastMCP(
    "Excel Retriever",
    instructions=(
        "MCP server that retrieves, summarizes, and analyses Excel files. "
        "Handles workbooks from 1 KB to 100 MB with paginated data access, "
        "formula explanations, statistics, search, and data-quality checks."
    ),
)


# ── Helper ───────────────────────────────────────────────────────────────

def _json(obj: Any) -> str:
    """Serialize a result dict to pretty JSON for the LLM."""
    return json.dumps(obj, indent=2, default=str)


# ── Tools ────────────────────────────────────────────────────────────────


@mcp.tool()
def list_excel_files(directory: str) -> str:
    """List all Excel files (.xlsx, .xlsm, .xltx) in a directory.

    Returns file names, paths, sizes, and last-modified timestamps.
    Use this first to discover which files are available.

    Args:
        directory: Absolute or relative path to the directory to scan.
    """
    files = excel_reader.list_excel_files(directory)
    if not files:
        return "No Excel files found in the specified directory."
    return _json(files)


@mcp.tool()
def get_workbook_info(file_path: str) -> str:
    """Get high-level metadata about an Excel workbook.

    Returns file size, sheet count, sheet names, and per-sheet dimensions
    (row and column counts).  Call this before drilling into a specific
    sheet to understand the workbook structure.

    Args:
        file_path: Path to the Excel file.
    """
    return _json(excel_reader.get_workbook_info(file_path))


@mcp.tool()
def get_sheet_names(file_path: str) -> str:
    """Return the list of sheet names in a workbook.

    Args:
        file_path: Path to the Excel file.
    """
    return _json(excel_reader.get_sheet_names(file_path))


@mcp.tool()
def get_sheet_summary(
    file_path: str,
    sheet_name: str | None = None,
    detail_level: str = "brief",
) -> str:
    """Get a summary of a specific sheet.

    Use detail_level='brief' for a quick overview (row/column counts,
    headers, dimensions).

    Use detail_level='detailed' for a deep dive that includes:
    - column data types
    - sample data rows
    - statistics for every numeric column
    - every formula in the sheet with a full human-readable explanation
      that uses column header names and row numbers instead of cell
      references (e.g. "adds up 'Revenue' rows 2 to 50")
    - empty-cell counts per column

    Args:
        file_path:    Path to the Excel file.
        sheet_name:   Name of the sheet (defaults to the active sheet).
        detail_level: 'brief' or 'detailed'.
    """
    return _json(
        excel_reader.get_sheet_summary(file_path, sheet_name, detail_level)
    )


@mcp.tool()
def read_sheet_data(
    file_path: str,
    sheet_name: str | None = None,
    start_row: int = 1,
    max_rows: int = 100,
    columns: list[str] | None = None,
) -> str:
    """Read data rows from a sheet with pagination.

    Returns up to *max_rows* rows starting from *start_row*.  The response
    includes a *next_start_row* field for fetching the next page.

    Args:
        file_path:  Path to the Excel file.
        sheet_name: Sheet name (defaults to active sheet).
        start_row:  1-based starting row (1 = header row).
        max_rows:   Maximum data rows to return (default 100).
        columns:    Optional list of column header names to include.
    """
    return _json(
        excel_reader.read_sheet_data(
            file_path, sheet_name, start_row, max_rows, columns
        )
    )


@mcp.tool()
def get_cell_info(
    file_path: str,
    cell_reference: str,
    sheet_name: str | None = None,
) -> str:
    """Get full details about a single cell.

    Returns the cell's value, data type, column header name, and — if the
    cell contains a formula — the raw formula and a plain-English
    explanation that uses column headers and row labels.

    Args:
        file_path:      Path to the Excel file.
        cell_reference: Cell address, e.g. 'B5'.
        sheet_name:     Sheet name (defaults to active sheet).
    """
    return _json(
        excel_reader.get_cell_info(file_path, cell_reference, sheet_name)
    )


@mcp.tool()
def get_range_data(
    file_path: str,
    range_ref: str,
    sheet_name: str | None = None,
) -> str:
    """Read data from a specific cell range (e.g. 'A1:D10').

    Args:
        file_path:  Path to the Excel file.
        range_ref:  Range address, e.g. 'B2:E20'.
        sheet_name: Sheet name (defaults to active sheet).
    """
    return _json(
        excel_reader.get_range_data(file_path, range_ref, sheet_name)
    )


@mcp.tool()
def get_formulas(
    file_path: str,
    sheet_name: str | None = None,
) -> str:
    """Extract every formula from a sheet with human-readable explanations.

    Each formula entry includes:
    - the cell address
    - the column header name
    - the raw Excel formula
    - a plain-English explanation using header names and row numbers

    Args:
        file_path:  Path to the Excel file.
        sheet_name: Sheet name (defaults to active sheet).
    """
    return _json(excel_reader.get_formulas(file_path, sheet_name))


@mcp.tool()
def search_in_sheet(
    file_path: str,
    query: str,
    sheet_name: str | None = None,
    case_sensitive: bool = False,
    max_results: int = 50,
) -> str:
    """Search for a text value across all cells in a sheet.

    Returns matching cells with their addresses, column names, and values.

    Args:
        file_path:      Path to the Excel file.
        query:          Text to search for (substring match).
        sheet_name:     Sheet name (defaults to active sheet).
        case_sensitive: Whether the search is case-sensitive.
        max_results:    Maximum matches to return (default 50).
    """
    return _json(
        excel_reader.search_in_sheet(
            file_path, query, sheet_name, case_sensitive, max_results
        )
    )


@mcp.tool()
def get_sheet_statistics(
    file_path: str,
    sheet_name: str | None = None,
) -> str:
    """Compute descriptive statistics for all numeric columns.

    Returns count, sum, mean, median, min, and max for each column that
    contains numeric data.

    Args:
        file_path:  Path to the Excel file.
        sheet_name: Sheet name (defaults to active sheet).
    """
    return _json(excel_reader.get_sheet_statistics(file_path, sheet_name))


@mcp.tool()
def validate_sheet_data(
    file_path: str,
    sheet_name: str | None = None,
) -> str:
    """Validate data quality in a sheet.

    Checks for:
    - mixed data types in a column
    - columns with many missing / empty values
    - duplicate rows

    Args:
        file_path:  Path to the Excel file.
        sheet_name: Sheet name (defaults to active sheet).
    """
    return _json(excel_reader.validate_sheet_data(file_path, sheet_name))


# ── Entry point ──────────────────────────────────────────────────────────


def main() -> None:
    """Run the MCP server on stdio."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
