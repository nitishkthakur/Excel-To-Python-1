"""Search retriever — find values or formulas across one or all sheets."""

from __future__ import annotations

from typing import Any

from openpyxl.utils import get_column_letter

from ._base import (
    get_headers,
    open_workbook,
    resolve_sheets,
    serialize,
    validate_content_type,
    validate_file,
)


def retrieve(
    file_path: str,
    *,
    query: str,
    sheet_name: str | None = None,
    content_type: str = "formulas",
    header_row: int = 0,
    case_sensitive: bool = False,
    max_results: int = 50,
) -> dict[str, Any]:
    """Search for *query* across cells in one or all sheets.

    Parameters
    ----------
    file_path:      Path to the Excel file.
    query:          Substring to search for.
    sheet_name:     Sheet name, or ``None`` to search every sheet.
    content_type:   ``"values"`` searches computed values.
                    ``"formulas"`` searches raw formula strings.
                    ``"both"`` searches both views (deduplicated).
    header_row:     Row containing headers (0 = auto-detect).
    case_sensitive: Whether to match case.
    max_results:    Maximum matches to return.
    """
    validate_file(file_path)
    validate_content_type(content_type)

    all_matches: list[dict[str, Any]] = []

    views = []
    if content_type in ("values", "both"):
        views.append(True)   # data_only=True
    if content_type in ("formulas", "both"):
        views.append(False)  # data_only=False

    seen: set[tuple[str, str]] = set()  # (sheet_name, cell_ref)

    for data_only in views:
        if len(all_matches) >= max_results:
            break
        wb = open_workbook(file_path, data_only=data_only)
        try:
            sheets = resolve_sheets(wb, sheet_name)
            for sname, ws in sheets:
                if len(all_matches) >= max_results:
                    break
                _search_sheet(
                    ws, sname, query, header_row,
                    case_sensitive, max_results,
                    all_matches, seen,
                )
        finally:
            wb.close()

    return {
        "query": query,
        "content_type": content_type,
        "match_count": len(all_matches),
        "truncated": len(all_matches) >= max_results,
        "matches": all_matches,
    }


def _search_sheet(
    ws,
    sheet_name: str,
    query: str,
    header_row: int,
    case_sensitive: bool,
    max_results: int,
    matches: list[dict[str, Any]],
    seen: set[tuple[str, str]],
) -> None:
    max_col = ws.max_column or 0
    max_row = ws.max_row or 0
    headers, _ = get_headers(ws, max_col, header_row)
    q = query if case_sensitive else query.lower()

    for row_idx in range(1, max_row + 1):
        if len(matches) >= max_results:
            return
        for col_idx in range(1, max_col + 1):
            if len(matches) >= max_results:
                return
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None:
                continue
            val_str = str(val)
            compare = val_str if case_sensitive else val_str.lower()
            if q not in compare:
                continue
            cell_ref = f"{get_column_letter(col_idx)}{row_idx}"
            key = (sheet_name, cell_ref)
            if key in seen:
                continue
            seen.add(key)
            header = (
                headers[col_idx - 1]
                if col_idx <= len(headers)
                else get_column_letter(col_idx)
            )
            matches.append({
                "sheet_name": sheet_name,
                "cell": cell_ref,
                "row": row_idx,
                "column": header,
                "value": serialize(val),
            })
