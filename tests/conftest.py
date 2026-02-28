"""Shared test fixtures for Excel MCP Server tests."""

from __future__ import annotations

import os
import tempfile

import openpyxl
import pytest


@pytest.fixture()
def tmp_dir():
    """Provide a temporary directory that is cleaned up after the test."""
    with tempfile.TemporaryDirectory() as d:
        yield d


@pytest.fixture()
def sample_workbook(tmp_dir):
    """Create a small sample workbook with data and formulas.

    Sheet 'Sales':
        A          B         C        D
    1   Product    Revenue   Cost     Profit
    2   Widget A   1000      400      =B2-C2
    3   Widget B   2500      1100     =B3-C3
    4   Widget C   800       350      =B4-C4
    5   Total      =SUM(B2:B4)  =SUM(C2:C4)  =SUM(D2:D4)

    Sheet 'Info':
        A          B
    1   Key        Value
    2   Year       2024
    3   Region     North

    Sheet 'Expenses':
        A            B
    1   Category     Amount
    2   Travel       500
    3   Office       300
    4   Software     200
    5   Total        =SUM(B2:B4)
    """
    path = os.path.join(tmp_dir, "sample.xlsx")
    wb = openpyxl.Workbook()

    # --- Sales sheet ---
    ws = wb.active
    ws.title = "Sales"
    ws.append(["Product", "Revenue", "Cost", "Profit"])
    ws.append(["Widget A", 1000, 400, None])
    ws["D2"] = "=B2-C2"
    ws.append(["Widget B", 2500, 1100, None])
    ws["D3"] = "=B3-C3"
    ws.append(["Widget C", 800, 350, None])
    ws["D4"] = "=B4-C4"
    ws.append(["Total", None, None, None])
    ws["B5"] = "=SUM(B2:B4)"
    ws["C5"] = "=SUM(C2:C4)"
    ws["D5"] = "=SUM(D2:D4)"

    # --- Info sheet ---
    ws2 = wb.create_sheet("Info")
    ws2.append(["Key", "Value"])
    ws2.append(["Year", 2024])
    ws2.append(["Region", "North"])

    # --- Expenses sheet ---
    ws3 = wb.create_sheet("Expenses")
    ws3.append(["Category", "Amount"])
    ws3.append(["Travel", 500])
    ws3.append(["Office", 300])
    ws3.append(["Software", 200])
    ws3.append(["Total", None])
    ws3["B5"] = "=SUM(B2:B4)"

    wb.save(path)
    wb.close()
    return path


@pytest.fixture()
def unstructured_workbook(tmp_dir):
    """Workbook where headers are NOT in row 1.

    Sheet 'Messy':
        A                       B          C
    1   Financial Report 2024
    2   (blank)
    3   Item                    Amount     Tax
    4   Rent                    5000       =B4*0.1
    5   Utilities               1200       =B5*0.1
    6   Total                   =SUM(B4:B5) =SUM(C4:C5)
    """
    path = os.path.join(tmp_dir, "unstructured.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Messy"
    ws["A1"] = "Financial Report 2024"
    # Row 2 blank
    ws["A3"] = "Item"
    ws["B3"] = "Amount"
    ws["C3"] = "Tax"
    ws["A4"] = "Rent"
    ws["B4"] = 5000
    ws["C4"] = "=B4*0.1"
    ws["A5"] = "Utilities"
    ws["B5"] = 1200
    ws["C5"] = "=B5*0.1"
    ws["A6"] = "Total"
    ws["B6"] = "=SUM(B4:B5)"
    ws["C6"] = "=SUM(C4:C5)"

    wb.save(path)
    wb.close()
    return path


@pytest.fixture()
def empty_workbook(tmp_dir):
    """An Excel file with a single empty sheet."""
    path = os.path.join(tmp_dir, "empty.xlsx")
    wb = openpyxl.Workbook()
    wb.save(path)
    wb.close()
    return path


@pytest.fixture()
def non_excel_file(tmp_dir):
    """A plain text file that is not an Excel file."""
    path = os.path.join(tmp_dir, "data.txt")
    with open(path, "w") as f:
        f.write("not excel")
    return path


@pytest.fixture()
def mixed_type_workbook(tmp_dir):
    """Workbook with mixed data types in a column for validation testing."""
    path = os.path.join(tmp_dir, "mixed.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ID", "Value"])
    ws.append([1, 100])
    ws.append([2, "text"])
    ws.append([3, 300])
    wb.save(path)
    wb.close()
    return path
