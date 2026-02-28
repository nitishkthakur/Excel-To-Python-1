"""Summary retriever — brief or detailed sheet summaries."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

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
    file_path: Annotated[str, Field(
        description="Absolute path to the Excel file (.xlsx, .xlsm, .xltx).",
    )],
    *,
    sheet_name: Annotated[str | None, Field(
        default=None,
        description=(
            "Name of the sheet to summarize. Pass null to summarize ALL "
            "sheets in the workbook."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="formulas",
        description=(
            "What cell content to include. Allowed values: "
            "'formulas' (default, **recommended**) — includes a formula "
            "inventory with explanations in detailed mode. "
            "'values' — omits formulas. "
            "'both' — includes formulas with their computed values."
        ),
    )] = "formulas",
    detail_level: Annotated[str, Field(
        default="brief",
        description=(
            "Level of detail in the summary. Allowed values: "
            "'brief' (default) — returns headers, dimensions, and "
            "row/column counts. "
            "'detailed' — adds column data types, sample data rows, "
            "descriptive statistics for every numeric column, empty-cell "
            "analysis, and (when content_type includes formulas) a full "
            "formula inventory with human-readable explanations."
        ),
    )] = "brief",
    header_row: Annotated[int, Field(
        default=0,
        description=(
            "1-based row number containing column headers. "
            "0 (default) means auto-detect — useful for unstructured "
            "sheets where headers are not in row 1."
        ),
    )] = 0,
) -> dict[str, Any]:
    """Use this retriever to get a structural overview of one or all sheets —
    from a quick glance (brief) to a comprehensive deep dive (detailed).

    **Brief**: quick structural overview — headers, row/column counts,
    dimensions.

    **Detailed**: deep dive including column types, sample rows,
    descriptive statistics, empty-cell counts, and a complete formula
    inventory with explanations.

    Use ``detail_level='detailed'`` with ``sheet_name=null`` to produce a
    comprehensive overview of the entire workbook suitable for executive
    summaries.
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
