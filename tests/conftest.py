"""Shared fixtures: synthetic openpyxl workbooks for testing."""
import io

import openpyxl
import pytest


@pytest.fixture
def simple_workbook():
    """Workbook with basic formula cells used in the plan's example."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = 10
    ws["A2"] = 20
    ws["C5"] = "=SUM(A1:A2)"
    ws["B10"] = "=C5+1.1"
    ws["B11"] = "=C5*2.0"
    return wb


@pytest.fixture
def simple_workbook_path(simple_workbook, tmp_path):
    """Save simple_workbook to a temp file and return the path."""
    p = tmp_path / "test.xlsx"
    simple_workbook.save(p)
    return p


@pytest.fixture
def multi_sheet_workbook():
    """Workbook with cross-sheet formula references."""
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Sheet1"
    ws1["A1"] = 5
    ws1["B1"] = "=Sheet2!A1 + A1"

    ws2 = wb.create_sheet("Sheet2")
    ws2["A1"] = "=Sheet1!A1 * 2"
    return wb


@pytest.fixture
def multi_sheet_workbook_path(multi_sheet_workbook, tmp_path):
    p = tmp_path / "multi.xlsx"
    multi_sheet_workbook.save(p)
    return p


@pytest.fixture
def circular_workbook():
    """Workbook containing a circular reference A1 → B1 → A1."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"
    return wb


@pytest.fixture
def circular_workbook_path(circular_workbook, tmp_path):
    p = tmp_path / "circular.xlsx"
    circular_workbook.save(p)
    return p
