"""Statistics retriever — compute descriptive stats for numeric columns."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from ._base import (
    compute_statistics,
    get_headers,
    open_workbook,
    package_results,
    resolve_sheets,
    validate_file,
)


def retrieve(
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file (.xlsx, .xlsm, .xltx).",
    )],
    *,
    sheet_name: Annotated[str | None, Field(
        default=None,
        description=(
            "Name of the sheet to compute statistics for. Pass null to "
            "compute statistics for ALL sheets in the workbook."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="values",
        description=(
            "Accepted for interface consistency with other retrievers. "
            "Statistics always operate on computed values regardless of "
            "this setting. Allowed values: 'values', 'formulas', 'both'."
        ),
    )] = "values",
    header_row: Annotated[int, Field(
        default=0,
        description=(
            "1-based row number containing column headers. "
            "0 (default) means auto-detect."
        ),
    )] = 0,
) -> dict[str, Any]:
    """Use this retriever to get a quick quantitative overview of numeric
    data in a sheet before diving deeper.

    Returns count, sum, mean, median, min, and max for each column that
    contains numeric data. Statistics always operate on computed values.
    """
    validate_file(file_path)
    wb = open_workbook(file_path, data_only=True)
    try:
        sheets = resolve_sheets(wb, sheet_name)
        results = []
        for sname, ws in sheets:
            max_col = ws.max_column or 0
            max_row = ws.max_row or 0
            headers, hrow = get_headers(ws, max_col, header_row)
            data_start = hrow + 1
            stats = compute_statistics(
                ws, headers, max_col, max_row, data_start,
            )
            results.append({
                "sheet_name": sname,
                "total_data_rows": max(0, max_row - hrow),
                "statistics": stats,
            })
    finally:
        wb.close()

    return package_results(results, "values")
