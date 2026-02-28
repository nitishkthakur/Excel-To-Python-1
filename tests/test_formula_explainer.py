"""Tests for excel_mcp_server.formula_explainer."""

from __future__ import annotations

import pytest

from excel_mcp_server.formula_explainer import explain_formula


HEADERS = ["Product", "Revenue", "Cost", "Profit"]


class TestExplainFormula:
    def test_simple_subtraction(self):
        result = explain_formula("=B2-C2", HEADERS, current_row=2)
        assert "Revenue" in result
        assert "Cost" in result
        assert "row 2" in result

    def test_sum_range(self):
        result = explain_formula("=SUM(B2:B4)", HEADERS, current_row=5)
        assert "Revenue" in result
        assert "rows 2 to 4" in result
        assert "adds up" in result

    def test_average(self):
        result = explain_formula("=AVERAGE(C2:C10)", HEADERS, current_row=11)
        assert "average" in result
        assert "Cost" in result

    def test_if_formula(self):
        result = explain_formula('=IF(B2>1000,"High","Low")', HEADERS, current_row=2)
        assert "checks if" in result or "IF" in result

    def test_vlookup(self):
        result = explain_formula("=VLOOKUP(A2,Sheet2!A1:D10,3,FALSE)", HEADERS, current_row=2)
        assert "looks up" in result

    def test_non_formula_passthrough(self):
        result = explain_formula("hello", HEADERS)
        assert result == "hello"

    def test_empty_string(self):
        result = explain_formula("", HEADERS)
        assert result == ""

    def test_raw_formula_included(self):
        result = explain_formula("=SUM(B2:B4)", HEADERS, current_row=5)
        assert "=SUM(B2:B4)" in result

    def test_nested_formula(self):
        result = explain_formula("=ROUND(AVERAGE(B2:B10),2)", HEADERS, current_row=11)
        assert "ROUND" in result or "rounds" in result

    def test_countif(self):
        result = explain_formula('=COUNTIF(A2:A10,"Widget A")', HEADERS, current_row=11)
        assert "count" in result.lower()
        assert "Product" in result

    def test_concatenate(self):
        result = explain_formula("=CONCATENATE(A2,B2)", HEADERS, current_row=2)
        assert "joins" in result or "CONCATENATE" in result

    def test_no_current_row(self):
        """When current_row is None, row prefix is omitted."""
        result = explain_formula("=SUM(B2:B4)", HEADERS)
        assert "In row" not in result
        assert "adds up" in result
