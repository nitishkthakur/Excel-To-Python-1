"""MCP Server for Excel file retrieval and analysis.

Exposes each retriever as a separate MCP tool so the LLM can discover and
choose the right one.  Every retriever supports:

* **multi-sheet** — pass ``sheet_name=null`` to operate on all sheets.
* **content_type** — ``"formulas"`` (recommended default), ``"values"``,
  or ``"both"``.
* **output_format** — ``"markdown"``, ``"json"``, or ``"csv"``.

Architecture decision: each retriever is its own tool (rather than one
dispatcher function) because it follows Anthropic's tool-use guidelines —
clear, focused descriptions help the model pick the right tool without
extra routing logic.

Usage::

    python -m excel_mcp_server.server          # stdio transport
"""

from __future__ import annotations

import json
import logging
from typing import Annotated, Any

from mcp.server.fastmcp import FastMCP
from pydantic import Field

from .excel_retrievers import (
    cell_info,
    format_output,
    formula,
    get_sheet_names,
    get_workbook_info,
    list_excel_files,
    range_data,
    search,
    sheet_data,
    statistics,
    summary,
    validation,
)

logger = logging.getLogger(__name__)

# ── MCP server instance ─────────────────────────────────────────────────

mcp = FastMCP(
    "Excel Retriever",
    instructions=(
        "MCP server that retrieves, summarizes, and analyses Excel files "
        "(1 KB – 100 MB).  Each tool is a self-contained retriever.  "
        "Start with list_excel_files → get_workbook_info to discover "
        "files, then use the appropriate retriever.  Default to "
        "content_type='formulas' to understand sheet logic before "
        "requesting computed values."
    ),
)


# ── Helper ───────────────────────────────────────────────────────────────

def _json(obj: Any) -> str:
    return json.dumps(obj, indent=2, default=str)


# ═════════════════════════════════════════════════════════════════════════
#  Discovery tools
# ═════════════════════════════════════════════════════════════════════════


@mcp.tool()
def discover_excel_files(
    directory: Annotated[str, Field(
        description=(
            "Absolute or relative path to the directory to scan for "
            "Excel files."
        ),
    )],
) -> str:
    """Use this as the **first step** to discover which Excel files are
    available before calling any retriever.

    Lists every Excel file (.xlsx / .xlsm / .xltx) in the given directory.
    Returns file names, absolute paths, sizes, and last-modified timestamps.
    Do NOT use this to read file contents — use the appropriate retriever
    instead.
    """
    files = list_excel_files(directory)
    if not files:
        return "No Excel files found in the specified directory."
    return _json(files)


@mcp.tool()
def inspect_workbook(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    output_format: Annotated[str, Field(
        default="json",
        description=(
            "Response format: 'json' (default), 'markdown', or 'csv'."
        ),
    )] = "json",
) -> str:
    """Use this **after** discover_excel_files to understand the structure of
    a specific workbook before drilling into individual sheets.

    Returns file size, sheet count, sheet names, and per-sheet dimensions
    (row and column counts).
    """
    result = get_workbook_info(file_path)
    return format_output(result, output_format)


@mcp.tool()
def list_sheets(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
) -> str:
    """Use this when you just need to know which sheets exist without the
    full metadata returned by inspect_workbook.

    Returns the ordered list of sheet names in a workbook.
    """
    return _json(get_sheet_names(file_path))


# ═════════════════════════════════════════════════════════════════════════
#  Retriever tools
# ═════════════════════════════════════════════════════════════════════════


@mcp.tool()
def retrieve_sheet_data(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description=(
            "Name of the sheet to read.  Pass null to read ALL sheets "
            "and return results for each."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "What cell content to return.  'formulas' (default, "
            "**recommended**) returns raw formulas with human-readable "
            "explanations — use this first to understand the sheet logic.  "
            "'values' returns computed values only.  'both' returns "
            "computed values plus formula details."
        ),
    )] = "formulas",
    output_format: Annotated[str, Field(
        default="markdown",
        description=(
            "Response format.  'markdown' (default) for human-readable "
            "tables.  'json' for structured data.  'csv' for tabular "
            "export."
        ),
    )] = "markdown",
    header_row: Annotated[int, Field(
        default=0,
        description=(
            "Row number containing column headers.  0 (default) means "
            "auto-detect — useful for unstructured sheets where headers "
            "are not in the first row."
        ),
    )] = 0,
    start_row: Annotated[int, Field(
        default=1,
        description="1-based starting row (1 = header row).",
    )] = 1,
    max_rows: Annotated[int, Field(
        default=100,
        description="Maximum data rows to return per sheet (default 100).",
    )] = 100,
    columns: Annotated[list[str] | None, Field(
        default=None,
        description=(
            "Optional list of column header names to include.  When "
            "null, all columns are returned."
        ),
    )] = None,
) -> str:
    """Use this retriever to inspect actual cell contents — data values or
    raw formulas — with pagination for large sheets.

    **Start with content_type='formulas'** to understand the sheet's
    calculation logic before requesting computed values.

    Returns paginated rows with headers, row counts, and a
    ``next_start_row`` field for fetching subsequent pages.  Supports
    column filtering and auto-detection of header rows for unstructured
    sheets.
    """
    result = sheet_data.retrieve(
        file_path,
        sheet_name=sheet_name,
        content_type=content_type,
        header_row=header_row,
        start_row=start_row,
        max_rows=max_rows,
        columns=columns,
    )
    return format_output(result, output_format)


@mcp.tool()
def retrieve_formulas(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description=(
            "Sheet name, or null to extract formulas from ALL sheets."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "'formulas' (default) returns raw formulas + explanations.  "
            "'values' also adds computed values for each formula cell.  "
            "'both' returns everything."
        ),
    )] = "formulas",
    output_format: Annotated[str, Field(
        default="markdown",
        description="'markdown' (default), 'json', or 'csv'.",
    )] = "markdown",
    header_row: Annotated[int, Field(
        default=0,
        description="Header row (0 = auto-detect).",
    )] = 0,
) -> str:
    """Use this retriever for deep reports and lineage — call with
    sheet_name=null to get a complete formula inventory across the entire
    workbook.

    Each formula entry includes the cell address, column header name, raw
    Excel formula, and a human-readable explanation that uses header names
    and row numbers instead of cell references (e.g. "adds up 'Revenue'
    rows 2 to 50").  This is the best starting point for building a
    lineage or understanding how outputs are calculated.
    """
    result = formula.retrieve(
        file_path,
        sheet_name=sheet_name,
        content_type=content_type,
        header_row=header_row,
    )
    return format_output(result, output_format)


@mcp.tool()
def retrieve_summary(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description=(
            "Sheet name, or null to summarize ALL sheets."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "'formulas' (default, recommended) includes formula "
            "inventory in detailed mode.  'values' omits formulas.  "
            "'both' includes formulas with computed values."
        ),
    )] = "formulas",
    output_format: Annotated[str, Field(
        default="markdown",
        description="'markdown' (default), 'json', or 'csv'.",
    )] = "markdown",
    detail_level: Annotated[str, Field(
        default="brief",
        description=(
            "'brief' returns headers, dimensions, row/column counts.  "
            "'detailed' adds column data types, sample data rows, "
            "statistics for every numeric column, empty-cell analysis, "
            "and (when content_type includes formulas) a full formula "
            "inventory with human-readable explanations."
        ),
    )] = "brief",
    header_row: Annotated[int, Field(
        default=0,
        description="Header row (0 = auto-detect).",
    )] = 0,
) -> str:
    """Use this retriever to get a structural overview of one or all sheets —
    from a quick glance (brief) to a comprehensive deep dive (detailed).

    **Brief** (default): quick structural overview — headers, row/column
    counts, dimensions.

    **Detailed**: deep dive including column types, sample rows,
    descriptive statistics, empty-cell counts, and a complete formula
    inventory with explanations.

    Use ``detail_level='detailed'`` with ``sheet_name=null`` to produce a
    comprehensive overview of the entire workbook suitable for executive
    summaries.
    """
    result = summary.retrieve(
        file_path,
        sheet_name=sheet_name,
        content_type=content_type,
        detail_level=detail_level,
        header_row=header_row,
    )
    return format_output(result, output_format)


@mcp.tool()
def retrieve_cell_info(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    cell_reference: Annotated[str, Field(
        description="Cell address, e.g. 'B5' or 'AA12'.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description="Sheet name (defaults to the active sheet).",
    )] = None,
    content_type: Annotated[str, Field(
        default="both",
        description=(
            "'both' (default) returns the computed value AND formula "
            "with explanation.  'formulas' returns formula only.  "
            "'values' returns the computed value only."
        ),
    )] = "both",
    output_format: Annotated[str, Field(
        default="json",
        description="'json' (default), 'markdown', or 'csv'.",
    )] = "json",
    header_row: Annotated[int, Field(
        default=0,
        description="Header row (0 = auto-detect).",
    )] = 0,
) -> str:
    """Use this to drill into a specific cell after identifying it via
    search or sheet data retrieval.

    Returns the cell's value, data type, column header name, and — if the
    cell contains a formula — the raw formula and a plain-English
    explanation using column headers and row labels.
    """
    result = cell_info.retrieve(
        file_path,
        cell_reference=cell_reference,
        sheet_name=sheet_name,
        content_type=content_type,
        header_row=header_row,
    )
    return format_output(result, output_format)


@mcp.tool()
def retrieve_range(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    range_ref: Annotated[str, Field(
        description="Cell range address, e.g. 'B2:E20' or 'A1:D10'.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description="Sheet name (defaults to the active sheet).",
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "'formulas' (default, recommended) shows raw formulas.  "
            "'values' shows computed values.  'both' shows both."
        ),
    )] = "formulas",
    output_format: Annotated[str, Field(
        default="markdown",
        description="'markdown' (default), 'json', or 'csv'.",
    )] = "markdown",
    header_row: Annotated[int, Field(
        default=0,
        description="Header row (0 = auto-detect).",
    )] = 0,
) -> str:
    """Use this when you know the exact cell range you need — e.g. from a
    previous search result or a known layout.

    Returns all cells in the range as a table with column headers resolved
    from the sheet's header row.
    """
    result = range_data.retrieve(
        file_path,
        range_ref=range_ref,
        sheet_name=sheet_name,
        content_type=content_type,
        header_row=header_row,
    )
    return format_output(result, output_format)


@mcp.tool()
def search_workbook(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    query: Annotated[str, Field(
        description="Substring to search for across cell values.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description=(
            "Sheet name, or null to search ALL sheets in the workbook."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "'formulas' (default) searches in raw formula strings.  "
            "'values' searches in computed values.  'both' searches "
            "in both views (deduplicated)."
        ),
    )] = "formulas",
    output_format: Annotated[str, Field(
        default="json",
        description="'json' (default), 'markdown', or 'csv'.",
    )] = "json",
    header_row: Annotated[int, Field(
        default=0,
        description="Header row (0 = auto-detect).",
    )] = 0,
    case_sensitive: Annotated[bool, Field(
        default=False,
        description="Whether the search should be case-sensitive.",
    )] = False,
    max_results: Annotated[int, Field(
        default=50,
        description="Maximum number of matches to return.",
    )] = 50,
) -> str:
    """Use this to locate where a specific variable, label, or value appears
    across the workbook — especially useful for tracing how a variable is
    calculated across multiple sheets.

    Returns matching cells with their sheet name, cell address, column
    header, and value.

    **Tip**: search with ``content_type='formulas'`` to find which cells
    reference a particular variable name in their formulas.
    """
    result = search.retrieve(
        file_path,
        query=query,
        sheet_name=sheet_name,
        content_type=content_type,
        header_row=header_row,
        case_sensitive=case_sensitive,
        max_results=max_results,
    )
    return format_output(result, output_format)


@mcp.tool()
def retrieve_statistics(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description="Sheet name, or null for all sheets.",
    )] = None,
    output_format: Annotated[str, Field(
        default="markdown",
        description="'markdown' (default), 'json', or 'csv'.",
    )] = "markdown",
    header_row: Annotated[int, Field(
        default=0,
        description="Header row (0 = auto-detect).",
    )] = 0,
) -> str:
    """Use this to get a quick quantitative overview of numeric data in a
    sheet before diving deeper.

    Returns count, sum, mean, median, min, and max for each column that
    contains numeric data.  Statistics always operate on computed values.
    """
    result = statistics.retrieve(
        file_path,
        sheet_name=sheet_name,
        header_row=header_row,
    )
    return format_output(result, output_format)


@mcp.tool()
def validate_data(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file.",
    )],
    sheet_name: Annotated[str | None, Field(
        default=None,
        description="Sheet name, or null to validate all sheets.",
    )] = None,
    output_format: Annotated[str, Field(
        default="json",
        description="'json' (default), 'markdown', or 'csv'.",
    )] = "json",
    header_row: Annotated[int, Field(
        default=0,
        description="Header row (0 = auto-detect).",
    )] = 0,
) -> str:
    """Use this before preparing a requirements document to identify which
    columns are inputs vs. computed outputs, and to check data quality.

    Reports:
    - **Mixed types**: columns where cells contain different data types.
    - **Missing values**: columns with > 10 % empty / blank cells.
    - **Duplicate rows**: identical rows (sampled for large sheets).
    """
    result = validation.retrieve(
        file_path,
        sheet_name=sheet_name,
        header_row=header_row,
    )
    return format_output(result, output_format)


# ── Entry point ──────────────────────────────────────────────────────────

def main() -> None:
    """Run the MCP server on stdio."""
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
