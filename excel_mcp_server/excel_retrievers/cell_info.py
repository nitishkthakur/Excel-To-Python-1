"""Cell-info retriever — deep-dive into a single cell."""

from __future__ import annotations

from typing import Any

from openpyxl.utils import coordinate_to_tuple, get_column_letter

from ._base import (
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
    cell_reference: str,
    sheet_name: str | None = None,
    content_type: str = "formulas",
    header_row: int = 0,
) -> dict[str, Any]:
    """Return detailed information about a single cell.

    Parameters
    ----------
    file_path:      Path to the Excel file.
    cell_reference: Cell address such as ``"B5"``.
    sheet_name:     Sheet name (defaults to the active sheet).
    content_type:   ``"formulas"`` shows the raw formula + explanation.
                    ``"values"`` shows the computed value.
                    ``"both"`` shows both.
    header_row:     Row containing headers (0 = auto-detect).
    """
    validate_file(file_path)
    validate_content_type(content_type)
    row, col = coordinate_to_tuple(cell_reference.upper())

    result: dict[str, Any] = {
        "cell": cell_reference.upper(),
        "row": row,
        "column": col,
        "column_letter": get_column_letter(col),
        "content_type": content_type,
    }

    if content_type in ("values", "both"):
        wb = open_workbook(file_path, data_only=True)
        try:
            ws = resolve_sheet(wb, sheet_name)
            result["sheet_name"] = ws.title
            val = ws.cell(row=row, column=col).value
            headers, _ = get_headers(ws, header_row=header_row)
            result["column_header"] = (
                headers[col - 1] if col <= len(headers) else get_column_letter(col)
            )
            result["value"] = serialize(val)
            result["data_type"] = (
                type(val).__name__ if val is not None else "empty"
            )
        finally:
            wb.close()

    if content_type in ("formulas", "both"):
        wb = open_workbook(file_path, data_only=False)
        try:
            ws = resolve_sheet(wb, sheet_name)
            result["sheet_name"] = ws.title
            cell_val = ws.cell(row=row, column=col).value
            headers, _ = get_headers(ws, header_row=header_row)
            result["column_header"] = (
                headers[col - 1] if col <= len(headers) else get_column_letter(col)
            )
            if isinstance(cell_val, str) and cell_val.startswith("="):
                from ..formula_explainer import explain_formula

                result["formula"] = cell_val
                result["formula_explanation"] = explain_formula(
                    cell_val, headers, row,
                )
            elif content_type == "formulas":
                # No formula — still populate value
                result["value"] = serialize(cell_val)
                result["data_type"] = (
                    type(cell_val).__name__
                    if cell_val is not None
                    else "empty"
                )
        finally:
            wb.close()

    return result
