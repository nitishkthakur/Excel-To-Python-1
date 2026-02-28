"""Tests for excel_mcp_server.excel_reader."""

from __future__ import annotations

import os

import pytest

from excel_mcp_server import excel_reader


# ── list_excel_files ─────────────────────────────────────────────────────


class TestListExcelFiles:
    def test_lists_xlsx_files(self, sample_workbook, tmp_dir):
        files = excel_reader.list_excel_files(tmp_dir)
        assert len(files) == 1
        assert files[0]["name"] == "sample.xlsx"
        assert files[0]["size_bytes"] > 0

    def test_missing_directory_raises(self):
        with pytest.raises(FileNotFoundError):
            excel_reader.list_excel_files("/nonexistent/path")

    def test_empty_directory(self, tmp_dir):
        assert excel_reader.list_excel_files(tmp_dir) == []


# ── get_workbook_info ────────────────────────────────────────────────────


class TestGetWorkbookInfo:
    def test_returns_metadata(self, sample_workbook):
        info = excel_reader.get_workbook_info(sample_workbook)
        assert info["sheet_count"] == 2
        assert "Sales" in info["sheet_names"]
        assert "Info" in info["sheet_names"]
        assert info["size_bytes"] > 0

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            excel_reader.get_workbook_info("/no/such/file.xlsx")

    def test_unsupported_format(self, non_excel_file):
        with pytest.raises(ValueError, match="Unsupported"):
            excel_reader.get_workbook_info(non_excel_file)


# ── get_sheet_names ──────────────────────────────────────────────────────


class TestGetSheetNames:
    def test_returns_all_names(self, sample_workbook):
        names = excel_reader.get_sheet_names(sample_workbook)
        assert names == ["Sales", "Info"]


# ── get_sheet_summary ────────────────────────────────────────────────────


class TestGetSheetSummary:
    def test_brief_summary(self, sample_workbook):
        s = excel_reader.get_sheet_summary(sample_workbook, "Sales", "brief")
        assert s["sheet_name"] == "Sales"
        assert s["total_rows"] == 4  # 5 rows minus header
        assert s["total_columns"] == 4
        assert s["headers"] == ["Product", "Revenue", "Cost", "Profit"]
        assert s["detail_level"] == "brief"
        # Brief should NOT include formulas
        assert "formulas" not in s

    def test_detailed_summary_includes_formulas(self, sample_workbook):
        s = excel_reader.get_sheet_summary(sample_workbook, "Sales", "detailed")
        assert s["detail_level"] == "detailed"
        assert "formulas" in s
        assert s["formula_count"] > 0
        # Check formula explanations mention header names
        for f in s["formulas"]:
            assert "explanation" in f
            assert "formula" in f

    def test_detailed_summary_includes_stats(self, sample_workbook):
        s = excel_reader.get_sheet_summary(sample_workbook, "Sales", "detailed")
        assert "statistics" in s
        assert "Revenue" in s["statistics"]
        assert s["statistics"]["Revenue"]["count"] == 3

    def test_detailed_summary_includes_sample_rows(self, sample_workbook):
        s = excel_reader.get_sheet_summary(sample_workbook, "Sales", "detailed")
        assert "sample_rows" in s
        assert len(s["sample_rows"]) > 0
        assert s["sample_rows"][0]["Product"] == "Widget A"

    def test_default_sheet(self, sample_workbook):
        s = excel_reader.get_sheet_summary(sample_workbook)
        assert s["sheet_name"] == "Sales"  # active sheet

    def test_invalid_detail_level(self, sample_workbook):
        with pytest.raises(ValueError, match="detail_level"):
            excel_reader.get_sheet_summary(sample_workbook, "Sales", "ultra")

    def test_nonexistent_sheet(self, sample_workbook):
        with pytest.raises(ValueError, match="not found"):
            excel_reader.get_sheet_summary(sample_workbook, "NoSheet")


# ── read_sheet_data ──────────────────────────────────────────────────────


class TestReadSheetData:
    def test_reads_first_page(self, sample_workbook):
        result = excel_reader.read_sheet_data(sample_workbook, "Sales")
        assert result["rows_returned"] == 4
        assert result["total_data_rows"] == 4
        assert result["headers"] == ["Product", "Revenue", "Cost", "Profit"]

    def test_pagination(self, sample_workbook):
        result = excel_reader.read_sheet_data(
            sample_workbook, "Sales", start_row=1, max_rows=2
        )
        assert result["rows_returned"] == 2
        assert result["has_more"] is True
        assert result["next_start_row"] is not None

    def test_column_filter(self, sample_workbook):
        result = excel_reader.read_sheet_data(
            sample_workbook, "Sales", columns=["Product", "Revenue"]
        )
        assert result["headers"] == ["Product", "Revenue"]
        # Each row should only have the requested columns (plus _row)
        for row in result["data"]:
            assert "Product" in row
            assert "Revenue" in row
            assert "Cost" not in row

    def test_invalid_columns(self, sample_workbook):
        with pytest.raises(ValueError, match="None of the requested columns"):
            excel_reader.read_sheet_data(
                sample_workbook, "Sales", columns=["Nonexistent"]
            )


# ── get_cell_info ────────────────────────────────────────────────────────


class TestGetCellInfo:
    def test_data_cell(self, sample_workbook):
        info = excel_reader.get_cell_info(sample_workbook, "A2", "Sales")
        assert info["value"] == "Widget A"
        assert info["column_header"] == "Product"
        assert info["data_type"] == "str"

    def test_formula_cell(self, sample_workbook):
        info = excel_reader.get_cell_info(sample_workbook, "D2", "Sales")
        assert "formula" in info
        assert info["formula"] == "=B2-C2"
        assert "formula_explanation" in info
        # Explanation should mention column headers
        assert "Revenue" in info["formula_explanation"]
        assert "Cost" in info["formula_explanation"]

    def test_numeric_cell(self, sample_workbook):
        info = excel_reader.get_cell_info(sample_workbook, "B2", "Sales")
        assert info["value"] == 1000
        assert info["data_type"] == "int"


# ── get_formulas ─────────────────────────────────────────────────────────


class TestGetFormulas:
    def test_extracts_all_formulas(self, sample_workbook):
        result = excel_reader.get_formulas(sample_workbook, "Sales")
        assert result["formula_count"] == 6  # D2-D4 + B5, C5, D5
        cells = {f["cell"] for f in result["formulas"]}
        assert "D2" in cells
        assert "B5" in cells

    def test_formula_explanations_use_headers(self, sample_workbook):
        result = excel_reader.get_formulas(sample_workbook, "Sales")
        sum_formula = next(f for f in result["formulas"] if f["cell"] == "B5")
        assert "Revenue" in sum_formula["explanation"]
        assert "SUM" in sum_formula["formula"]

    def test_no_formulas(self, sample_workbook):
        result = excel_reader.get_formulas(sample_workbook, "Info")
        assert result["formula_count"] == 0


# ── search_in_sheet ──────────────────────────────────────────────────────


class TestSearchInSheet:
    def test_finds_text(self, sample_workbook):
        result = excel_reader.search_in_sheet(sample_workbook, "Widget", "Sales")
        assert result["match_count"] == 3  # Widget A, B, C

    def test_case_insensitive(self, sample_workbook):
        result = excel_reader.search_in_sheet(
            sample_workbook, "widget", "Sales", case_sensitive=False
        )
        assert result["match_count"] == 3

    def test_case_sensitive(self, sample_workbook):
        result = excel_reader.search_in_sheet(
            sample_workbook, "widget", "Sales", case_sensitive=True
        )
        assert result["match_count"] == 0

    def test_no_matches(self, sample_workbook):
        result = excel_reader.search_in_sheet(
            sample_workbook, "zzz_nothing", "Sales"
        )
        assert result["match_count"] == 0

    def test_max_results(self, sample_workbook):
        result = excel_reader.search_in_sheet(
            sample_workbook, "Widget", "Sales", max_results=2
        )
        assert result["match_count"] == 2
        assert result["truncated"] is True


# ── get_sheet_statistics ─────────────────────────────────────────────────


class TestGetSheetStatistics:
    def test_numeric_statistics(self, sample_workbook):
        result = excel_reader.get_sheet_statistics(sample_workbook, "Sales")
        stats = result["statistics"]
        assert "Revenue" in stats
        assert stats["Revenue"]["count"] == 3
        assert stats["Revenue"]["sum"] == 4300
        assert stats["Revenue"]["min"] == 800
        assert stats["Revenue"]["max"] == 2500


# ── get_range_data ───────────────────────────────────────────────────────


class TestGetRangeData:
    def test_reads_range(self, sample_workbook):
        result = excel_reader.get_range_data(sample_workbook, "A1:B3", "Sales")
        assert result["rows_returned"] == 3
        assert result["data"][0]["Product"] == "Product"  # header row
        assert result["data"][1]["Product"] == "Widget A"


# ── validate_sheet_data ──────────────────────────────────────────────────


class TestValidateSheetData:
    def test_validation_runs(self, sample_workbook):
        result = excel_reader.validate_sheet_data(sample_workbook, "Sales")
        assert "issues" in result
        assert result["total_data_rows"] == 4
        assert isinstance(result["empty_cells_per_column"], dict)

    def test_detects_mixed_types_if_present(self, tmp_dir):
        """Create a file with mixed types and check detection."""
        import openpyxl as xl

        path = os.path.join(tmp_dir, "mixed.xlsx")
        wb = xl.Workbook()
        ws = wb.active
        ws.append(["ID", "Value"])
        ws.append([1, 100])
        ws.append([2, "text"])  # mixed type in Value column
        ws.append([3, 300])
        wb.save(path)
        wb.close()

        result = excel_reader.validate_sheet_data(path)
        type_issues = [i for i in result["issues"] if i["type"] == "mixed_types"]
        assert len(type_issues) == 1
        assert type_issues[0]["column"] == "Value"
