"""Formula retriever — extract formulas from one or all sheets."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from ._base import (
    add_formula_explanations,
    extract_formulas,
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
            "Name of the sheet to extract formulas from. Pass null to "
            "extract formulas from ALL sheets in the workbook."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "What to return for each formula cell. Allowed values: "
            "'formulas' (default) — returns raw formula strings with "
            "human-readable explanations. "
            "'values' — also includes the computed value for each formula "
            "cell. "
            "'both' — returns formula strings, explanations, and computed "
            "values."
        ),
    )] = "formulas",
    header_row: Annotated[int, Field(
        default=0,
        description=(
            "1-based row number containing column headers. "
            "0 (default) means auto-detect — useful for unstructured "
            "sheets where headers are not in row 1."
        ),
    )] = 0,
) -> dict[str, Any]:
    """Use this retriever to get a complete formula inventory for building
    lineage or understanding how outputs are calculated.

    Each formula entry includes the cell address, column header name, raw
    Excel formula, and a human-readable explanation that uses header names
    and row numbers instead of cell references (e.g. "adds up 'Revenue'
    rows 2 to 50").

    Call with ``sheet_name=null`` for a full workbook formula inventory —
    the best starting point for deep reports.
    """
    validate_file(file_path)
    validate_content_type(content_type)

    wb = open_workbook(file_path, data_only=False)
    try:
        sheets = resolve_sheets(wb, sheet_name)
        results = []
        for sname, ws in sheets:
            max_col = ws.max_column or 0
            max_row = ws.max_row or 0
            headers, _ = get_headers(ws, max_col, header_row)
            formulas = extract_formulas(ws, headers, max_col, max_row)
            results.append({
                "sheet_name": sname,
                "formula_count": len(formulas),
                "formulas": formulas,
            })
    finally:
        wb.close()

    # Optionally add computed values for each formula cell
    if content_type in ("values", "both"):
        wb_val = open_workbook(file_path, data_only=True)
        try:
            sheets_val = resolve_sheets(wb_val, sheet_name)
            for i, (_, ws_val) in enumerate(sheets_val):
                if i >= len(results):
                    break
                for fentry in results[i].get("formulas", []):
                    from openpyxl.utils import coordinate_to_tuple
                    row, col = coordinate_to_tuple(fentry["cell"])
                    fentry["computed_value"] = serialize(
                        ws_val.cell(row=row, column=col).value
                    )
        finally:
            wb_val.close()

    return package_results(results, content_type)
