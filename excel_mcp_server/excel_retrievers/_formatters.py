"""Output formatters for retriever results.

Converts structured dicts returned by retrievers into JSON, Markdown,
or CSV strings suitable for returning to an LLM via MCP.
"""

from __future__ import annotations

import csv
import io
import json
from typing import Any


def format_output(result: dict[str, Any], output_format: str) -> str:
    """Format a retriever result dict into the requested string format.

    Parameters
    ----------
    result:
        Structured dict returned by a retriever's ``retrieve()`` function.
    output_format:
        ``"json"``, ``"markdown"``, or ``"csv"``.
    """
    fmt = output_format.lower()
    if fmt == "json":
        return _to_json(result)
    if fmt == "markdown":
        return _to_markdown(result)
    if fmt == "csv":
        return _to_csv(result)
    raise ValueError(
        f"Unsupported output_format '{output_format}'. "
        "Use 'json', 'markdown', or 'csv'."
    )


# ── JSON ─────────────────────────────────────────────────────────────────

def _to_json(result: dict[str, Any]) -> str:
    return json.dumps(result, indent=2, default=str)


# ── Markdown ─────────────────────────────────────────────────────────────

def _to_markdown(result: dict[str, Any]) -> str:
    """Convert a result dict to a readable Markdown string.

    Handles multi-sheet results (``sheets`` key) by creating a section
    per sheet, and renders lists-of-dicts as Markdown tables.
    """
    # Multi-sheet wrapper
    if "sheets" in result and isinstance(result["sheets"], list):
        parts: list[str] = []
        meta = {
            k: v for k, v in result.items()
            if k not in ("sheets",) and not k.startswith("_")
        }
        if meta:
            parts.append(_render_meta(meta))
        for sheet_result in result["sheets"]:
            name = sheet_result.get("sheet_name", "Sheet")
            parts.append(f"\n## {name}\n")
            parts.append(
                _render_section(sheet_result, exclude={"sheet_name"})
            )
        return "\n".join(parts).strip()

    # Single result
    return _render_section(result).strip()


def _render_meta(meta: dict[str, Any]) -> str:
    lines: list[str] = []
    for key, value in meta.items():
        if key.startswith("_"):
            continue
        lines.append(f"**{_title(key)}**: {value}")
    return "\n".join(lines)


def _render_section(
    result: dict[str, Any],
    exclude: set[str] | None = None,
) -> str:
    """Render a single result dict as Markdown."""
    lines: list[str] = []
    exclude = exclude or set()
    for key, value in result.items():
        if key.startswith("_") or key in exclude:
            continue
        # List of dicts → table
        if isinstance(value, list) and value and isinstance(value[0], dict):
            lines.append(f"\n### {_title(key)}\n")
            lines.append(_dicts_to_md_table(value))
        # Dict of dicts → table (e.g. statistics)
        elif isinstance(value, dict) and value and all(
            isinstance(v, dict) for v in value.values()
        ):
            lines.append(f"\n### {_title(key)}\n")
            lines.append(_nested_dict_to_md_table(value))
        # Plain dict → bullet list
        elif isinstance(value, dict):
            lines.append(f"\n### {_title(key)}\n")
            for k, v in value.items():
                lines.append(f"- **{k}**: {v}")
        # Simple list → comma-separated
        elif isinstance(value, list):
            lines.append(
                f"**{_title(key)}**: {', '.join(str(v) for v in value)}"
            )
        # Scalar
        else:
            lines.append(f"**{_title(key)}**: {value}")
    return "\n".join(lines)


# ── CSV ──────────────────────────────────────────────────────────────────

def _to_csv(result: dict[str, Any]) -> str:
    """Convert the first list-of-dicts found in the result to CSV.

    Falls back to JSON if no tabular data is present.
    """
    # Multi-sheet: concatenate all sheets
    if "sheets" in result and isinstance(result["sheets"], list):
        parts: list[str] = []
        for sheet_result in result["sheets"]:
            table = _find_table(sheet_result)
            if table:
                parts.append(_dicts_to_csv(table))
        if parts:
            return "\n".join(parts)

    table = _find_table(result)
    if table:
        return _dicts_to_csv(table)
    return _to_json(result)


def _find_table(d: dict[str, Any]) -> list[dict[str, Any]] | None:
    """Return the first list-of-dicts in *d*, or ``None``."""
    for value in d.values():
        if isinstance(value, list) and value and isinstance(value[0], dict):
            return value
    return None


# ── Helpers ──────────────────────────────────────────────────────────────

def _title(key: str) -> str:
    return key.replace("_", " ").title()


def _dicts_to_md_table(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return "_No data_"
    headers = list(rows[0].keys())
    lines = [
        "| " + " | ".join(str(h) for h in headers) + " |",
        "| " + " | ".join("---" for _ in headers) + " |",
    ]
    for row in rows:
        cells = [str(row.get(h, "")).replace("|", "\\|") for h in headers]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def _nested_dict_to_md_table(data: dict[str, dict]) -> str:
    if not data:
        return "_No data_"
    first = next(iter(data.values()))
    sub_keys = list(first.keys())
    header_row = "| | " + " | ".join(str(k) for k in sub_keys) + " |"
    sep_row = "| --- | " + " | ".join("---" for _ in sub_keys) + " |"
    lines = [header_row, sep_row]
    for key, sub in data.items():
        cells = [str(key)] + [str(sub.get(k, "")) for k in sub_keys]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def _dicts_to_csv(rows: list[dict[str, Any]]) -> str:
    if not rows:
        return ""
    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    for row in rows:
        writer.writerow({k: str(v) if v is not None else "" for k, v in row.items()})
    return output.getvalue()
