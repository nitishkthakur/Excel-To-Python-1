"""Summary retriever — brief or detailed sheet summaries."""

from __future__ import annotations

from typing import Any

from ._base import (
    add_formula_explanations,
    analyze_column_types,
    compute_statistics,
    count_empty_cells,
    get_headers,
    open_workbook,
    package_results,
    read_rows,
    resolve_sheets,
    validate_content_type,
    validate_file,
)


def retrieve(
    file_path: str,
    *,
    sheet_name: str | None = None,
    content_type: str = "formulas",
    detail_level: str = "brief",
    header_row: int = 0,
) -> dict[str, Any]:
    """Generate a summary of one or all sheets.

    Parameters
    ----------
    file_path:    Path to the Excel file.
    sheet_name:   Sheet name, or ``None`` for all sheets.
    content_type: ``"formulas"`` (default) includes formula inventory in
                  detailed mode. ``"values"`` omits formulas. ``"both"``
                  includes formulas with computed values.
    detail_level: ``"brief"`` returns headers, dimensions, row/column counts.
                  ``"detailed"`` adds column types, sample rows, statistics,
                  empty-cell analysis, and (when content_type includes
                  formulas) a full formula inventory with explanations.
    header_row:   Row containing headers (0 = auto-detect).
    """
    validate_file(file_path)
    validate_content_type(content_type)
    if detail_level not in ("brief", "detailed"):
        raise ValueError("detail_level must be 'brief' or 'detailed'")

    data_only = content_type in ("values", "both")
    wb = open_workbook(file_path, data_only=data_only)
    try:
        sheets = resolve_sheets(wb, sheet_name)
        results = [
            _summarize(ws, sname, detail_level, header_row)
            for sname, ws in sheets
        ]
    finally:
        wb.close()

    # Formula inventory for detailed summaries
    if detail_level == "detailed" and content_type in ("formulas", "both"):
        add_formula_explanations(file_path, results, sheet_name)

    return package_results(results, content_type)


def _summarize(
    ws,
    sheet_name: str,
    detail_level: str,
    header_row: int,
) -> dict[str, Any]:
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    headers, hrow = get_headers(ws, max_col, header_row)
    data_start = hrow + 1

    summary: dict[str, Any] = {
        "sheet_name": sheet_name,
        "header_row": hrow,
        "total_rows": max(0, max_row - hrow),
        "total_columns": max_col,
        "headers": headers,
        "dimensions": ws.dimensions if hasattr(ws, "dimensions") else None,
        "detail_level": detail_level,
    }

    if detail_level == "detailed":
        summary["column_types"] = analyze_column_types(
            ws, max_col, max_row, data_start,
        )
        summary["sample_rows"] = read_rows(ws, headers, data_start, 5)
        summary["statistics"] = compute_statistics(
            ws, headers, max_col, max_row, data_start,
        )
        summary["empty_cells"] = count_empty_cells(
            ws, max_col, max_row, data_start,
        )

    return summary
