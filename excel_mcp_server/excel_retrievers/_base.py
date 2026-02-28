"""Shared utilities for all Excel retrievers.

Provides workbook I/O, file validation, header detection (including
auto-detect for unstructured sheets), serialization, discovery functions,
and common analysis helpers used by multiple retrievers.
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import Any

import openpyxl
from openpyxl.utils import get_column_letter


# ── Constants ────────────────────────────────────────────────────────────

LARGE_FILE_THRESHOLD = 5 * 1024 * 1024   # 5 MB — switch to read_only mode
MAX_FILE_SIZE = 100 * 1024 * 1024         # 100 MB hard cap
DEFAULT_PAGE_SIZE = 100

VALID_CONTENT_TYPES = ("values", "formulas", "both")
VALID_OUTPUT_FORMATS = ("json", "markdown", "csv")


# ── Validation helpers ───────────────────────────────────────────────────

def validate_file(file_path: str) -> None:
    """Validate that *file_path* is a supported Excel file within size limits."""
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    size = os.path.getsize(file_path)
    if size > MAX_FILE_SIZE:
        raise ValueError(
            f"File too large ({human_size(size)}). "
            f"Maximum supported size is {human_size(MAX_FILE_SIZE)}."
        )
    if not file_path.lower().endswith((".xlsx", ".xlsm", ".xltx")):
        raise ValueError(
            "Unsupported file format. Supported: .xlsx, .xlsm, .xltx"
        )


def validate_content_type(content_type: str) -> None:
    if content_type not in VALID_CONTENT_TYPES:
        raise ValueError(
            f"Invalid content_type '{content_type}'. "
            f"Must be one of: {VALID_CONTENT_TYPES}"
        )


def validate_output_format(output_format: str) -> None:
    if output_format not in VALID_OUTPUT_FORMATS:
        raise ValueError(
            f"Invalid output_format '{output_format}'. "
            f"Must be one of: {VALID_OUTPUT_FORMATS}"
        )


# ── Workbook I/O ────────────────────────────────────────────────────────

def open_workbook(
    file_path: str,
    *,
    data_only: bool = False,
    read_only: bool | None = None,
) -> openpyxl.Workbook:
    """Open a workbook; auto-selects read_only mode for large files."""
    validate_file(file_path)
    if read_only is None:
        read_only = os.path.getsize(file_path) > LARGE_FILE_THRESHOLD
    return openpyxl.load_workbook(
        file_path, read_only=read_only, data_only=data_only,
    )


# ── Sheet resolution ────────────────────────────────────────────────────

def resolve_sheet(wb: openpyxl.Workbook, sheet_name: str | None):
    """Return a single worksheet; defaults to the active sheet."""
    if sheet_name is None:
        return wb.active
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available: {list(wb.sheetnames)}"
        )
    return wb[sheet_name]


def resolve_sheets(
    wb: openpyxl.Workbook,
    sheet_name: str | None,
) -> list[tuple[str, Any]]:
    """Return ``(name, worksheet)`` pairs for all sheets or one sheet.

    When *sheet_name* is ``None`` every sheet in the workbook is returned.
    """
    if sheet_name is None:
        return [(name, wb[name]) for name in wb.sheetnames]
    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' not found. "
            f"Available: {list(wb.sheetnames)}"
        )
    return [(sheet_name, wb[sheet_name])]


# ── Header detection ────────────────────────────────────────────────────

def detect_header_row(ws, max_col: int | None = None) -> int:
    """Auto-detect which row contains column headers.

    Scans the first 20 rows and picks the first row where ≥ 50 % of cells
    are non-empty *strings*.  Falls back to row 1 if no clear header row
    is found — this handles unstructured sheets gracefully.
    """
    mc = max_col or ws.max_column or 1
    limit = min(20, (ws.max_row or 1))
    for row_idx in range(1, limit + 1):
        values = [
            ws.cell(row=row_idx, column=c).value for c in range(1, mc + 1)
        ]
        non_empty = [v for v in values if v is not None and str(v).strip()]
        if len(non_empty) < max(1, mc * 0.4):
            continue
        str_count = sum(1 for v in non_empty if isinstance(v, str))
        if str_count >= len(non_empty) * 0.5:
            return row_idx
    return 1


def get_headers(
    ws,
    max_col: int | None = None,
    header_row: int = 0,
) -> tuple[list[str], int]:
    """Read column headers from *header_row*.

    Parameters
    ----------
    header_row:
        Explicit 1-based row number.  ``0`` means auto-detect.

    Returns
    -------
    (headers, resolved_header_row)
    """
    mc = max_col or ws.max_column or 0
    if header_row == 0:
        header_row = detect_header_row(ws, mc)
    headers: list[str] = []
    for col_idx in range(1, mc + 1):
        val = ws.cell(row=header_row, column=col_idx).value
        headers.append(
            str(val) if val is not None else get_column_letter(col_idx)
        )
    return headers, header_row


# ── Serialization ────────────────────────────────────────────────────────

def serialize(value: Any) -> Any:
    """Convert a cell value to a JSON-safe Python type."""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.isoformat()
    if isinstance(value, (int, float, bool, str)):
        return value
    return str(value)


def human_size(nbytes: int | float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if nbytes < 1024:
            return f"{nbytes:.1f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.1f} TB"


# ── Discovery functions ─────────────────────────────────────────────────

def list_excel_files(directory: str) -> list[dict[str, Any]]:
    """List all Excel files in *directory* with metadata."""
    if not os.path.isdir(directory):
        raise FileNotFoundError(f"Directory not found: {directory}")
    results: list[dict[str, Any]] = []
    for entry in os.scandir(directory):
        if (
            entry.is_file()
            and entry.name.lower().endswith((".xlsx", ".xlsm", ".xltx"))
        ):
            stat = entry.stat()
            results.append({
                "name": entry.name,
                "path": os.path.abspath(entry.path),
                "size_bytes": stat.st_size,
                "size_human": human_size(stat.st_size),
                "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            })
    results.sort(key=lambda f: f["name"])
    return results


def get_workbook_info(file_path: str) -> dict[str, Any]:
    """Return high-level metadata about a workbook."""
    validate_file(file_path)
    stat = os.stat(file_path)
    wb = open_workbook(file_path, data_only=True)
    try:
        sheets_info = []
        for name in wb.sheetnames:
            ws = wb[name]
            sheets_info.append({
                "name": name,
                "min_row": ws.min_row,
                "max_row": ws.max_row,
                "min_column": ws.min_column,
                "max_column": ws.max_column,
                "dimensions": (
                    ws.dimensions if hasattr(ws, "dimensions") else None
                ),
            })
        return {
            "file_path": os.path.abspath(file_path),
            "file_name": os.path.basename(file_path),
            "size_bytes": stat.st_size,
            "size_human": human_size(stat.st_size),
            "modified": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "sheet_count": len(wb.sheetnames),
            "sheet_names": list(wb.sheetnames),
            "sheets": sheets_info,
        }
    finally:
        wb.close()


def get_sheet_names(file_path: str) -> list[str]:
    """Return sheet names in workbook order."""
    validate_file(file_path)
    wb = open_workbook(file_path, data_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


# ── Common analysis helpers ──────────────────────────────────────────────

def read_rows(
    ws,
    headers: list[str],
    start_row: int,
    max_rows: int,
) -> list[dict[str, Any]]:
    """Read data rows as list of dicts keyed by header name."""
    max_row = ws.max_row or 0
    max_col = ws.max_column or 0
    end_row = min(max_row, start_row + max_rows - 1)
    rows: list[dict[str, Any]] = []
    for row_idx in range(start_row, end_row + 1):
        row_data: dict[str, Any] = {"_row": row_idx}
        for col_idx in range(1, max_col + 1):
            header = (
                headers[col_idx - 1]
                if col_idx <= len(headers)
                else get_column_letter(col_idx)
            )
            row_data[header] = serialize(
                ws.cell(row=row_idx, column=col_idx).value
            )
        rows.append(row_data)
    return rows


def analyze_column_types(
    ws,
    max_col: int,
    max_row: int,
    data_start_row: int = 2,
) -> dict[str, list[str]]:
    """Sample the first 100 data rows to determine column data types."""
    headers, _ = get_headers(ws, max_col)
    col_types: dict[str, set[str]] = {h: set() for h in headers}
    sample_end = min(max_row, data_start_row + 99)
    for row_idx in range(data_start_row, sample_end + 1):
        for col_idx in range(1, max_col + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is not None:
                col_types[headers[col_idx - 1]].add(type(val).__name__)
    return {h: sorted(t) for h, t in col_types.items()}


def compute_statistics(
    ws,
    headers: list[str],
    max_col: int,
    max_row: int,
    data_start_row: int = 2,
) -> dict[str, dict[str, Any]]:
    """Compute count / sum / mean / median / min / max for numeric columns."""
    numeric_cols: dict[str, list[float]] = {}
    for col_idx in range(1, max_col + 1):
        vals: list[float] = []
        for row_idx in range(data_start_row, max_row + 1):
            v = ws.cell(row=row_idx, column=col_idx).value
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                vals.append(float(v))
        if vals:
            numeric_cols[headers[col_idx - 1]] = vals

    stats: dict[str, dict[str, Any]] = {}
    for header, vals in numeric_cols.items():
        s = sorted(vals)
        n = len(s)
        total = sum(s)
        mean = total / n
        mid = n // 2
        median = s[mid] if n % 2 == 1 else (s[mid - 1] + s[mid]) / 2
        stats[header] = {
            "count": n,
            "sum": round(total, 4),
            "mean": round(mean, 4),
            "median": round(median, 4),
            "min": s[0],
            "max": s[-1],
        }
    return stats


def count_empty_cells(
    ws,
    max_col: int,
    max_row: int,
    data_start_row: int = 2,
) -> dict[str, int]:
    """Count empty / blank cells per column."""
    headers, _ = get_headers(ws, max_col)
    empties: dict[str, int] = {h: 0 for h in headers}
    for row_idx in range(data_start_row, max_row + 1):
        for col_idx in range(1, max_col + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if val is None or (isinstance(val, str) and val.strip() == ""):
                empties[headers[col_idx - 1]] += 1
    return empties


def extract_formulas(
    ws,
    headers: list[str],
    max_col: int,
    max_row: int,
) -> list[dict[str, Any]]:
    """Extract all formula cells with human-readable explanations."""
    from ..formula_explainer import explain_formula

    formulas: list[dict[str, Any]] = []
    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            val = ws.cell(row=row_idx, column=col_idx).value
            if isinstance(val, str) and val.startswith("="):
                col_letter = get_column_letter(col_idx)
                col_header = (
                    headers[col_idx - 1]
                    if col_idx <= len(headers)
                    else col_letter
                )
                formulas.append({
                    "cell": f"{col_letter}{row_idx}",
                    "row": row_idx,
                    "column_header": col_header,
                    "formula": val,
                    "explanation": explain_formula(val, headers, row_idx),
                })
    return formulas


def add_formula_explanations(
    file_path: str,
    results: list[dict[str, Any]],
    sheet_name: str | None,
) -> None:
    """Append ``formula_explanations`` to each per-sheet result dict."""
    wb = open_workbook(file_path, data_only=False)
    try:
        sheets = resolve_sheets(wb, sheet_name)
        for i, (_, ws) in enumerate(sheets):
            if i >= len(results):
                break
            headers, _ = get_headers(ws)
            max_col = ws.max_column or 0
            max_row = ws.max_row or 0
            flist = extract_formulas(ws, headers, max_col, max_row)
            results[i]["formula_explanations"] = flist
            results[i]["formula_count"] = len(flist)
    finally:
        wb.close()


def package_results(
    results: list[dict[str, Any]],
    content_type: str,
) -> dict[str, Any]:
    """Wrap per-sheet results into a single return dict.

    Single-sheet results are returned flat; multi-sheet results are
    wrapped under a ``sheets`` key.
    """
    if len(results) == 1:
        out = results[0]
    else:
        out = {"sheet_count": len(results), "sheets": results}
    out["content_type"] = content_type
    return out
