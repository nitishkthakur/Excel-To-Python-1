"""Sheet-data retriever — read paginated rows from one or all sheets."""

from __future__ import annotations

from typing import Any

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
    file_path: str,
    *,
    sheet_name: str | None = None,
    content_type: str = "formulas",
    header_row: int = 0,
    start_row: int = 1,
    max_rows: int = DEFAULT_PAGE_SIZE,
    columns: list[str] | None = None,
) -> dict[str, Any]:
    """Read paginated data rows from one or all sheets.

    Parameters
    ----------
    file_path:   Path to the Excel file.
    sheet_name:  Sheet name, or ``None`` for all sheets.
    content_type: ``"values"`` | ``"formulas"`` | ``"both"``.
    header_row:  Row number with headers (0 = auto-detect).
    start_row:   1-based starting row (1 = header row).
    max_rows:    Maximum data rows per sheet.
    columns:     Optional list of header names to include.
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
