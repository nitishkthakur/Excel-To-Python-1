"""Sheet-data retriever — read paginated rows from one or all sheets."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from ._base import (
    DEFAULT_PAGE_SIZE,
    add_formula_explanations,
    get_headers,
    open_workbook,
    package_results,
    resolve_sheets,
    serialize,
    validate_content_type,
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
            "Name of the sheet to read. Pass null to read ALL sheets in "
            "the workbook and return results for each."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "What cell content to return. Allowed values: "
            "'formulas' (default, **recommended**) — returns raw formulas "
            "with human-readable explanations; use this first to understand "
            "the sheet logic. "
            "'values' — returns computed cell values only. "
            "'both' — returns computed values plus formula details."
        ),
    )] = "formulas",
    header_row: Annotated[int, Field(
        default=0,
        description=(
            "1-based row number containing column headers. "
            "0 (default) means auto-detect — the system scans the first "
            "20 rows for the most likely header row; useful for "
            "unstructured sheets where headers are not in row 1."
        ),
    )] = 0,
    start_row: Annotated[int, Field(
        default=1,
        description=(
            "1-based starting row for pagination. 1 (default) means start "
            "from the header row. Data rows begin after the detected or "
            "specified header row."
        ),
    )] = 1,
    max_rows: Annotated[int, Field(
        default=DEFAULT_PAGE_SIZE,
        description=(
            "Maximum number of data rows to return per sheet. "
            f"Default is {DEFAULT_PAGE_SIZE}. Use with start_row for "
            "pagination through large sheets."
        ),
    )] = DEFAULT_PAGE_SIZE,
    columns: Annotated[list[str] | None, Field(
        default=None,
        description=(
            "Optional list of column header names to include. When null "
            "(default), all columns are returned. Column names must match "
            "the headers exactly."
        ),
    )] = None,
) -> dict[str, Any]:
    """Use this retriever to inspect actual cell contents — data values or
    raw formulas — with pagination for large sheets.

    **Start with content_type='formulas'** to understand the sheet's
    calculation logic before requesting computed values.

    Returns paginated rows with headers, row counts, and a
    ``next_start_row`` field for fetching subsequent pages.  Supports
    column filtering and auto-detection of header rows for unstructured
    sheets.
    """
    validate_file(file_path)
    validate_content_type(content_type)

    data_only = content_type in ("values", "both")
    wb = open_workbook(file_path, data_only=data_only)
    try:
        sheets = resolve_sheets(wb, sheet_name)
        results = [
            _read_single(ws, sname, header_row, start_row, max_rows, columns)
            for sname, ws in sheets
        ]
    finally:
        wb.close()

    if content_type in ("formulas", "both"):
        add_formula_explanations(file_path, results, sheet_name)

    return package_results(results, content_type)


def _read_single(
    ws,
    sheet_name: str,
    header_row: int,
    start_row: int,
    max_rows: int,
    columns: list[str] | None,
) -> dict[str, Any]:
    max_col = ws.max_column or 0
    total_rows = ws.max_row or 0
    headers, hrow = get_headers(ws, max_col, header_row)

    # Column filter
    if columns:
        col_indices = [headers.index(c) + 1 for c in columns if c in headers]
        if not col_indices:
            raise ValueError(
                f"None of the requested columns found. Available: {headers}"
            )
        selected_headers = [headers[i - 1] for i in col_indices]
    else:
        col_indices = list(range(1, max_col + 1))
        selected_headers = headers

    # First data row is the row after the header
    data_start = hrow + 1
    if start_row > 1:
        data_start = max(data_start, start_row)
    data_end = min(total_rows, data_start + max_rows - 1)

    rows: list[dict[str, Any]] = []
    for row_idx in range(data_start, data_end + 1):
        row_data: dict[str, Any] = {"_row": row_idx}
        for ci, col_idx in enumerate(col_indices):
            row_data[selected_headers[ci]] = serialize(
                ws.cell(row=row_idx, column=col_idx).value
            )
        rows.append(row_data)

    return {
        "sheet_name": sheet_name,
        "headers": selected_headers,
        "header_row": hrow,
        "start_row": data_start,
        "rows_returned": len(rows),
        "total_data_rows": max(0, total_rows - hrow),
        "has_more": data_end < total_rows,
        "next_start_row": data_end + 1 if data_end < total_rows else None,
        "data": rows,
    }
