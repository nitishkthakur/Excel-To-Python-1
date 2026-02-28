"""Statistics retriever — compute descriptive stats for numeric columns."""

from __future__ import annotations

from typing import Any

from ._base import (
    compute_statistics,
    get_headers,
    open_workbook,
    package_results,
    resolve_sheets,
    validate_file,
)


def retrieve(
    file_path: str,
    *,
    sheet_name: str | None = None,
    content_type: str = "values",
    header_row: int = 0,
) -> dict[str, Any]:
    """Compute descriptive statistics for numeric columns.

    Returns count, sum, mean, median, min, and max for each column that
    contains numeric data.  Statistics always operate on **computed
    values**, so content_type is accepted for interface consistency but
    only ``"values"`` behaviour applies.

    Parameters
    ----------
    file_path:   Path to the Excel file.
    sheet_name:  Sheet name, or ``None`` for all sheets.
    content_type: Accepted for consistency; statistics always use values.
    header_row:  Row containing headers (0 = auto-detect).
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
