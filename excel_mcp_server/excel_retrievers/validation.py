"""Validation retriever — data-quality checks on one or all sheets."""

from __future__ import annotations

from typing import Annotated, Any

from pydantic import Field

from ._base import (
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
            "Name of the sheet to validate. Pass null to validate ALL "
            "sheets in the workbook."
        ),
    )] = None,
    content_type: Annotated[str, Field(
        default="values",
        description=(
            "Accepted for interface consistency with other retrievers. "
            "Validation always operates on computed values regardless of "
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
    """Use this retriever to identify data-quality issues before preparing
    a requirements document or to separate inputs from computed outputs.

    Reports:
    - **Mixed types**: columns where cells contain different data types.
    - **Missing values**: columns with > 10% empty / blank cells.
    - **Duplicate rows**: identical rows (sampled for large sheets).
    """
    validate_file(file_path)
    wb = open_workbook(file_path, data_only=True)
    try:
        sheets = resolve_sheets(wb, sheet_name)
        results = [
            _validate(ws, sname, header_row)
            for sname, ws in sheets
        ]
    finally:
        wb.close()

    return package_results(results, "values")


def _validate(ws, sheet_name: str, header_row: int) -> dict[str, Any]:
    max_col = ws.max_column or 0
    max_row = ws.max_row or 0
    headers, hrow = get_headers(ws, max_col, header_row=header_row)
    data_start = hrow + 1

    issues: list[dict[str, Any]] = []
    col_types: dict[str, set[str]] = {h: set() for h in headers}
    empty_per_col: dict[str, int] = {h: 0 for h in headers}

    for row_idx in range(data_start, max_row + 1):
        for col_idx in range(1, max_col + 1):
            header = headers[col_idx - 1]
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                empty_per_col[header] += 1
            else:
                col_types[header].add(type(val).__name__)

    # Mixed types
    for header, types in col_types.items():
        if len(types) > 1:
            issues.append({
                "type": "mixed_types",
                "column": header,
                "types_found": sorted(types),
                "severity": "warning",
            })

    # Missing values
    total_data = max(0, max_row - hrow)
    for header, empty_count in empty_per_col.items():
        if total_data > 0 and empty_count > 0:
            pct = round(100 * empty_count / total_data, 1)
            if pct > 10:
                issues.append({
                    "type": "missing_values",
                    "column": header,
                    "empty_count": empty_count,
                    "empty_percent": pct,
                    "severity": "warning" if pct < 50 else "error",
                })

    # Duplicate rows (sample-based for efficiency)
    check_rows = min(max_row, 1000)
    seen: set[tuple] = set()
    dup_count = 0
    for row_idx in range(data_start, check_rows + 1):
        row_tuple = tuple(
            ws.cell(row=row_idx, column=c).value
            for c in range(1, max_col + 1)
        )
        if row_tuple in seen:
            dup_count += 1
        else:
            seen.add(row_tuple)
    if dup_count > 0:
        issues.append({
            "type": "duplicate_rows",
            "count": dup_count,
            "checked_rows": check_rows - data_start + 1,
            "severity": "info",
        })

    return {
        "sheet_name": sheet_name,
        "total_data_rows": total_data,
        "total_columns": max_col,
        "issue_count": len(issues),
        "issues": issues,
        "empty_cells_per_column": empty_per_col,
    }
