"""Excel Retrievers — pluggable retrieval modules for Excel workbooks.

Each sub-module exposes a ``retrieve()`` function with a consistent
interface.  The registry :data:`RETRIEVERS` maps names to functions so
new retrievers can be added without changing calling code.

Quick reference
---------------
``sheet_data``   Read paginated data rows
``formula``      Extract formulas with explanations
``search``       Search values / formulas across sheets
``summary``      Brief or detailed sheet summaries
``cell_info``    Single cell deep-dive
``range_data``   Arbitrary cell range
``statistics``   Numeric column statistics
``validation``   Data-quality checks
"""

from . import (
    cell_info,
    formula,
    range_data,
    search,
    sheet_data,
    statistics,
    summary,
    validation,
)
from ._base import (
    get_sheet_names,
    get_workbook_info,
    list_excel_files,
)
from ._formatters import format_output

RETRIEVERS: dict[str, object] = {
    "sheet_data": sheet_data.retrieve,
    "formula": formula.retrieve,
    "search": search.retrieve,
    "summary": summary.retrieve,
    "cell_info": cell_info.retrieve,
    "range_data": range_data.retrieve,
    "statistics": statistics.retrieve,
    "validation": validation.retrieve,
}

__all__ = [
    "RETRIEVERS",
    "cell_info",
    "format_output",
    "formula",
    "get_sheet_names",
    "get_workbook_info",
    "list_excel_files",
    "range_data",
    "search",
    "sheet_data",
    "statistics",
    "summary",
    "validation",
]
