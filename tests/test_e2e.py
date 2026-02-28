"""End-to-end tests exercising the MCP server tool functions directly.

These tests call the actual MCP tool functions (which compose retrievers
+ formatters) and verify the complete pipeline from file to formatted
output in all three formats.
"""

from __future__ import annotations

import json

import pytest


# Import the server module so tools are registered
from excel_mcp_server.server import (
    discover_excel_files,
    inspect_workbook,
    list_sheets,
    retrieve_cell_info,
    retrieve_formulas,
    retrieve_range,
    retrieve_sheet_data,
    retrieve_statistics,
    retrieve_summary,
    search_workbook,
    validate_data,
)


class TestDiscoveryToolsE2E:
    def test_discover_excel_files(self, sample_workbook, tmp_dir):
        output = discover_excel_files(tmp_dir)
        parsed = json.loads(output)
        assert isinstance(parsed, list)
        assert any(f["name"] == "sample.xlsx" for f in parsed)

    def test_discover_empty_dir(self, tmp_dir):
        output = discover_excel_files(tmp_dir)
        assert "No Excel files" in output

    def test_inspect_workbook_json(self, sample_workbook):
        output = inspect_workbook(sample_workbook, output_format="json")
        parsed = json.loads(output)
        assert parsed["sheet_count"] == 3

    def test_inspect_workbook_markdown(self, sample_workbook):
        output = inspect_workbook(sample_workbook, output_format="markdown")
        assert "Sales" in output

    def test_list_sheets(self, sample_workbook):
        output = list_sheets(sample_workbook)
        parsed = json.loads(output)
        assert "Sales" in parsed
        assert "Expenses" in parsed


class TestSheetDataToolE2E:
    def test_json_format(self, sample_workbook):
        output = retrieve_sheet_data(
            sample_workbook, sheet_name="Sales",
            content_type="values", output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["rows_returned"] == 4

    def test_markdown_format(self, sample_workbook):
        output = retrieve_sheet_data(
            sample_workbook, sheet_name="Sales",
            content_type="values", output_format="markdown",
        )
        assert "Widget A" in output
        assert "|" in output

    def test_csv_format(self, sample_workbook):
        output = retrieve_sheet_data(
            sample_workbook, sheet_name="Sales",
            content_type="values", output_format="csv",
        )
        assert "Product" in output
        assert "Widget A" in output

    def test_all_sheets(self, sample_workbook):
        output = retrieve_sheet_data(
            sample_workbook, output_format="json",
            content_type="values",
        )
        parsed = json.loads(output)
        assert parsed["sheet_count"] == 3


class TestFormulaToolE2E:
    def test_json(self, sample_workbook):
        output = retrieve_formulas(
            sample_workbook, sheet_name="Sales", output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["formula_count"] == 6

    def test_markdown(self, sample_workbook):
        output = retrieve_formulas(
            sample_workbook, sheet_name="Sales", output_format="markdown",
        )
        assert "=SUM" in output
        assert "Revenue" in output

    def test_all_sheets(self, sample_workbook):
        output = retrieve_formulas(
            sample_workbook, output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["sheet_count"] == 3


class TestSummaryToolE2E:
    def test_brief_json(self, sample_workbook):
        output = retrieve_summary(
            sample_workbook, sheet_name="Sales",
            detail_level="brief", output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["total_rows"] == 4
        assert parsed["detail_level"] == "brief"

    def test_detailed_markdown(self, sample_workbook):
        output = retrieve_summary(
            sample_workbook, sheet_name="Sales",
            detail_level="detailed", output_format="markdown",
        )
        assert "Revenue" in output
        assert "Statistics" in output or "statistics" in output.lower()

    def test_all_sheets_detailed(self, sample_workbook):
        output = retrieve_summary(
            sample_workbook, detail_level="detailed",
            output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["sheet_count"] == 3

    def test_unstructured(self, unstructured_workbook):
        output = retrieve_summary(
            unstructured_workbook, sheet_name="Messy",
            detail_level="brief", output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["header_row"] == 3


class TestCellInfoToolE2E:
    def test_both(self, sample_workbook):
        output = retrieve_cell_info(
            sample_workbook, cell_reference="D2",
            sheet_name="Sales", content_type="both",
            output_format="json",
        )
        parsed = json.loads(output)
        assert "formula" in parsed
        assert "value" in parsed

    def test_values_only(self, sample_workbook):
        output = retrieve_cell_info(
            sample_workbook, cell_reference="B2",
            sheet_name="Sales", content_type="values",
            output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["value"] == 1000


class TestRangeToolE2E:
    def test_json(self, sample_workbook):
        output = retrieve_range(
            sample_workbook, range_ref="A1:B3",
            sheet_name="Sales", content_type="values",
            output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["rows_returned"] == 3

    def test_csv(self, sample_workbook):
        output = retrieve_range(
            sample_workbook, range_ref="A1:B3",
            sheet_name="Sales", content_type="values",
            output_format="csv",
        )
        assert "Product" in output
        assert "Widget A" in output


class TestSearchToolE2E:
    def test_search_single_sheet(self, sample_workbook):
        output = search_workbook(
            sample_workbook, query="Widget",
            sheet_name="Sales", content_type="values",
            output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["match_count"] == 3

    def test_search_all_sheets(self, sample_workbook):
        output = search_workbook(
            sample_workbook, query="Total",
            content_type="values", output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["match_count"] >= 2

    def test_search_formulas(self, sample_workbook):
        output = search_workbook(
            sample_workbook, query="SUM",
            sheet_name="Sales", content_type="formulas",
            output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["match_count"] >= 3


class TestStatisticsToolE2E:
    def test_json(self, sample_workbook):
        output = retrieve_statistics(
            sample_workbook, sheet_name="Sales", output_format="json",
        )
        parsed = json.loads(output)
        assert "Revenue" in parsed["statistics"]

    def test_markdown(self, sample_workbook):
        output = retrieve_statistics(
            sample_workbook, sheet_name="Sales", output_format="markdown",
        )
        assert "Revenue" in output

    def test_all_sheets(self, sample_workbook):
        output = retrieve_statistics(
            sample_workbook, output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["sheet_count"] == 3


class TestValidateToolE2E:
    def test_json(self, sample_workbook):
        output = validate_data(
            sample_workbook, sheet_name="Sales", output_format="json",
        )
        parsed = json.loads(output)
        assert "issues" in parsed

    def test_mixed_types(self, mixed_type_workbook):
        output = validate_data(
            mixed_type_workbook, output_format="json",
        )
        parsed = json.loads(output)
        issues = [i for i in parsed["issues"] if i["type"] == "mixed_types"]
        assert len(issues) == 1

    def test_all_sheets(self, sample_workbook):
        output = validate_data(
            sample_workbook, output_format="json",
        )
        parsed = json.loads(output)
        assert parsed["sheet_count"] == 3


class TestOutputFormatsConsistency:
    """Ensure all three formats produce non-empty output for every tool."""

    def _assert_non_empty(self, output: str):
        assert output is not None
        assert len(output.strip()) > 0

    def test_sheet_data_all_formats(self, sample_workbook):
        for fmt in ("json", "markdown", "csv"):
            out = retrieve_sheet_data(
                sample_workbook, sheet_name="Sales",
                content_type="values", output_format=fmt,
            )
            self._assert_non_empty(out)

    def test_formulas_all_formats(self, sample_workbook):
        for fmt in ("json", "markdown", "csv"):
            out = retrieve_formulas(
                sample_workbook, sheet_name="Sales", output_format=fmt,
            )
            self._assert_non_empty(out)

    def test_summary_all_formats(self, sample_workbook):
        for fmt in ("json", "markdown", "csv"):
            out = retrieve_summary(
                sample_workbook, sheet_name="Sales",
                detail_level="detailed", output_format=fmt,
            )
            self._assert_non_empty(out)

    def test_search_all_formats(self, sample_workbook):
        for fmt in ("json", "markdown", "csv"):
            out = search_workbook(
                sample_workbook, query="Widget",
                sheet_name="Sales", content_type="values",
                output_format=fmt,
            )
            self._assert_non_empty(out)

    def test_statistics_all_formats(self, sample_workbook):
        for fmt in ("json", "markdown", "csv"):
            out = retrieve_statistics(
                sample_workbook, sheet_name="Sales", output_format=fmt,
            )
            self._assert_non_empty(out)

    def test_validate_all_formats(self, sample_workbook):
        for fmt in ("json", "markdown", "csv"):
            out = validate_data(
                sample_workbook, sheet_name="Sales", output_format=fmt,
            )
            self._assert_non_empty(out)
