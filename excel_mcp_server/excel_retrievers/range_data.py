"""Range-data retriever — read an arbitrary cell range."""

from __future__ import annotations

from typing import Any

from openpyxl.utils import get_column_letter, range_boundaries

from ._base import (
    add_formula_explanations,
    get_headers,
    open_workbook,
    resolve_sheet,
    serialize,
    validate_content_type,
    validate_file,
)


def retrieve(
    file_path: str,
    *,
    range_ref: str,
    sheet_name: str | None = None,
    content_type: str = "formulas",
    header_row: int = 0,
) -> dict[str, Any]:
    """Read data from an explicit cell range such as ``"A1:D10"``.

    Parameters
    ----------
    file_path:   Path to the Excel file.
    range_ref:   Range address (e.g. ``"B2:E20"``).
    sheet_name:  Sheet name (defaults to the active sheet).
    content_type: ``"values"`` | ``"formulas"`` | ``"both"``.
    header_row:  Row containing headers (0 = auto-detect).
    """
    validate_file(file_path)
    validate_content_type(content_type)
    min_col, min_row, max_col_r, max_row_r = range_boundaries(
        range_ref.upper()
    )

    data_only = content_type in ("values", "both")
    wb = open_workbook(file_path, data_only=data_only)
    try:
        ws = resolve_sheet(wb, sheet_name)
        headers, _ = get_headers(ws, header_row=header_row)

        rows: list[dict[str, Any]] = []
        for row_idx in range(min_row, max_row_r + 1):
            row_data: dict[str, Any] = {"_row": row_idx}
            for col_idx in range(min_col, max_col_r + 1):
                header = (
                    headers[col_idx - 1]
                    if col_idx <= len(headers)
                    else get_column_letter(col_idx)
                )
                row_data[header] = serialize(
                    ws.cell(row=row_idx, column=col_idx).value
                )
            rows.append(row_data)

        result: dict[str, Any] = {
            "sheet_name": ws.title,
            "range": range_ref.upper(),
            "rows_returned": len(rows),
            "data": rows,
            "content_type": content_type,
        }
    finally:
        wb.close()

    if content_type in ("formulas", "both"):
        results_list = [result]
        add_formula_explanations(file_path, results_list, sheet_name)
        result = results_list[0]

    return result
