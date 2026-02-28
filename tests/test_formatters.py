"""Comprehensive tests for the output formatters."""

from __future__ import annotations

import json

import pytest

from excel_mcp_server.excel_retrievers._formatters import format_output


# ── Sample data ──────────────────────────────────────────────────────────

SINGLE_SHEET_RESULT = {
    "sheet_name": "Sales",
    "headers": ["Product", "Revenue"],
    "total_rows": 3,
    "data": [
        {"_row": 2, "Product": "Widget A", "Revenue": 1000},
        {"_row": 3, "Product": "Widget B", "Revenue": 2500},
    ],
}

MULTI_SHEET_RESULT = {
    "sheet_count": 2,
    "content_type": "values",
    "sheets": [
        {
            "sheet_name": "Sales",
            "data": [
                {"_row": 2, "Product": "Widget A", "Revenue": 1000},
            ],
        },
        {
            "sheet_name": "Info",
            "data": [
                {"_row": 2, "Key": "Year", "Value": "2024"},
            ],
        },
    ],
}

STATS_RESULT = {
    "sheet_name": "Sales",
    "statistics": {
        "Revenue": {"count": 3, "sum": 4300, "mean": 1433.33, "min": 800, "max": 2500},
        "Cost": {"count": 3, "sum": 1850, "mean": 616.67, "min": 350, "max": 1100},
    },
}

FORMULA_RESULT = {
    "sheet_name": "Sales",
    "formula_count": 2,
    "formulas": [
        {
            "cell": "D2",
            "formula": "=B2-C2",
            "explanation": "In row 2: subtracts Cost from Revenue",
        },
        {
            "cell": "B5",
            "formula": "=SUM(B2:B4)",
            "explanation": "Sums Revenue rows 2-4",
        },
    ],
}


# ═════════════════════════════════════════════════════════════════════════
#  JSON format
# ═════════════════════════════════════════════════════════════════════════


class TestJsonFormat:
    def test_single_sheet(self):
        output = format_output(SINGLE_SHEET_RESULT, "json")
        parsed = json.loads(output)
        assert parsed["sheet_name"] == "Sales"
        assert len(parsed["data"]) == 2

    def test_multi_sheet(self):
        output = format_output(MULTI_SHEET_RESULT, "json")
        parsed = json.loads(output)
        assert parsed["sheet_count"] == 2

    def test_stats(self):
        output = format_output(STATS_RESULT, "json")
        parsed = json.loads(output)
        assert "Revenue" in parsed["statistics"]

    def test_formulas(self):
        output = format_output(FORMULA_RESULT, "json")
        parsed = json.loads(output)
        assert parsed["formula_count"] == 2


# ═════════════════════════════════════════════════════════════════════════
#  Markdown format
# ═════════════════════════════════════════════════════════════════════════


class TestMarkdownFormat:
    def test_single_sheet_table(self):
        output = format_output(SINGLE_SHEET_RESULT, "markdown")
        assert "Product" in output
        assert "Revenue" in output
        assert "Widget A" in output
        assert "|" in output  # table syntax

    def test_multi_sheet_sections(self):
        output = format_output(MULTI_SHEET_RESULT, "markdown")
        assert "## Sales" in output
        assert "## Info" in output

    def test_stats_table(self):
        output = format_output(STATS_RESULT, "markdown")
        assert "Revenue" in output
        assert "4300" in output

    def test_formula_table(self):
        output = format_output(FORMULA_RESULT, "markdown")
        assert "=B2-C2" in output
        assert "=SUM(B2:B4)" in output

    def test_scalar_values(self):
        output = format_output({"total_rows": 42}, "markdown")
        assert "42" in output


# ═════════════════════════════════════════════════════════════════════════
#  CSV format
# ═════════════════════════════════════════════════════════════════════════


class TestCsvFormat:
    def test_single_sheet(self):
        output = format_output(SINGLE_SHEET_RESULT, "csv")
        lines = output.strip().split("\n")
        # Header + 2 data rows
        assert len(lines) == 3
        assert "Product" in lines[0]
        assert "Widget A" in lines[1]

    def test_multi_sheet(self):
        output = format_output(MULTI_SHEET_RESULT, "csv")
        assert "Widget A" in output
        assert "Year" in output

    def test_formula_list(self):
        output = format_output(FORMULA_RESULT, "csv")
        assert "=B2-C2" in output

    def test_fallback_to_json(self):
        """When no tabular data, CSV falls back to JSON."""
        output = format_output({"total_rows": 42}, "csv")
        parsed = json.loads(output)
        assert parsed["total_rows"] == 42


# ═════════════════════════════════════════════════════════════════════════
#  Error handling
# ═════════════════════════════════════════════════════════════════════════


class TestFormatErrors:
    def test_invalid_format(self):
        with pytest.raises(ValueError, match="Unsupported"):
            format_output({}, "xml")

    def test_empty_dict(self):
        # Should not raise
        output = format_output({}, "json")
        assert output == "{}"

    def test_empty_dict_markdown(self):
        output = format_output({}, "markdown")
        assert output == ""

    def test_empty_list_table(self):
        output = format_output({"data": []}, "markdown")
        assert "Data" in output
