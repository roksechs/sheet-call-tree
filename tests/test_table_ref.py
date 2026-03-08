"""Tests for TableRefNode — parsing, resolution, and serialization."""
import yaml
import pytest

from sheet_call_tree.formula_parser import parse_formula
from sheet_call_tree.models import FunctionNode, TableRefNode
from sheet_call_tree.reader import (
    _build_table_ranges,
    extract_formula_cells,
    extract_formula_cells_from_workbook,
)
from sheet_call_tree.serializer import to_yaml


SHEET = "Sheet1"


def p(formula):
    return parse_formula(formula, SHEET)


class TestTableRefParsing:
    def test_column_ref(self):
        node = p("=SUM(Table1[Amount])")
        assert isinstance(node, FunctionNode)
        assert node.name == "SUM"
        tref = node.args[0]
        assert isinstance(tref, TableRefNode)
        assert tref.table_name == "Table1"
        assert tref.column == "Amount"
        assert tref.this_row is False

    def test_this_row_ref(self):
        node = p("=Table1[@Amount]")
        assert isinstance(node, TableRefNode)
        assert node.table_name == "Table1"
        assert node.column == "Amount"
        assert node.this_row is True

    def test_whole_table_ref(self):
        node = p("=SUM(Table1[])")
        tref = node.args[0]
        assert isinstance(tref, TableRefNode)
        assert tref.table_name == "Table1"
        assert tref.column is None
        assert tref.this_row is False

    def test_unresolved_range_is_none(self):
        node = p("=SUM(Table1[Amount])")
        assert node.args[0].resolved_range is None

    def test_table_ref_in_arithmetic(self):
        node = p("=Table1[@Amount]*2")
        assert isinstance(node, FunctionNode)
        assert node.name == "MUL"
        tref = node.args[0]
        assert isinstance(tref, TableRefNode)
        assert tref.this_row is True

    def test_non_ascii_column_name(self):
        node = p("=SUM(売上テーブル[金額])")
        tref = node.args[0]
        assert isinstance(tref, TableRefNode)
        assert tref.table_name == "売上テーブル"
        assert tref.column == "金額"


class TestTableRangeBuilding:
    def test_build_table_ranges_keys(self, table_workbook):
        tmap = _build_table_ranges(table_workbook)
        assert "Table1" in tmap

    def test_build_table_ranges_sheet(self, table_workbook):
        tmap = _build_table_ranges(table_workbook)
        assert tmap["Table1"]["_sheet"] == "Sheet1"

    def test_build_table_ranges_column(self, table_workbook):
        tmap = _build_table_ranges(table_workbook)
        assert "Amount" in tmap["Table1"]["columns"]
        # Table ref is A1:B3, header in row 1, data in rows 2-3
        assert tmap["Table1"]["columns"]["Amount"] == "Sheet1!B2:B3"

    def test_build_table_ranges_whole(self, table_workbook):
        tmap = _build_table_ranges(table_workbook)
        assert tmap["Table1"]["_range"] == "Sheet1!A1:B3"

    def test_empty_workbook_has_no_tables(self):
        import openpyxl
        wb = openpyxl.Workbook()
        assert _build_table_ranges(wb) == {}


class TestTableRefResolution:
    def test_column_ref_resolved(self, table_workbook_path):
        cells, *_ = extract_formula_cells(table_workbook_path)
        c2 = cells["Sheet1!C2"]
        tref = c2.args[0]
        assert isinstance(tref, TableRefNode)
        assert tref.resolved_range == "Sheet1!B2:B3"

    def test_this_row_ref_resolved(self, table_workbook_path):
        cells, *_ = extract_formula_cells(table_workbook_path)
        d2 = cells["Sheet1!D2"]
        assert isinstance(d2, FunctionNode)
        tref = d2.args[0]
        assert isinstance(tref, TableRefNode)
        # this_row ref resolves to the same column range (not row-specific)
        assert tref.resolved_range == "Sheet1!B2:B3"

    def test_unresolved_when_no_table_map(self, table_workbook):
        cells = extract_formula_cells_from_workbook(table_workbook)
        c2 = cells["Sheet1!C2"]
        tref = c2.args[0]
        assert isinstance(tref, TableRefNode)
        assert tref.resolved_range is None


class TestTableRefSerialization:
    def _tref(self, **kw):
        defaults = dict(table_name="Table1", column="Amount", this_row=False)
        defaults.update(kw)
        return TableRefNode(**defaults)

    def test_depth0_basic(self):
        node = FunctionNode("SUM", [self._tref()])
        out = yaml.safe_load(to_yaml({"Sheet1!A1": node}, depth=0))
        expr = out["book"]["sheets"][0]["cells"][0]["expression"]
        assert expr == {"type": "SUM", "inputs": [{"type": "TABLE_REF", "name": "Table1", "column": "Amount"}]}

    def test_depth0_with_range(self):
        node = FunctionNode("SUM", [self._tref(resolved_range="Sheet1!B2:B3")])
        out = yaml.safe_load(to_yaml({"Sheet1!A1": node}, depth=0))
        tref_dict = out["book"]["sheets"][0]["cells"][0]["expression"]["inputs"][0]
        assert tref_dict["range"] == "Sheet1!B2:B3"

    def test_this_row_flag_present(self):
        node = self._tref(this_row=True)
        out = yaml.safe_load(to_yaml({"Sheet1!A1": node}, depth=0))
        tref_dict = out["book"]["sheets"][0]["cells"][0]["expression"]
        assert tref_dict["this_row"] is True

    def test_this_row_flag_absent_when_false(self):
        node = self._tref(this_row=False)
        out = yaml.safe_load(to_yaml({"Sheet1!A1": node}, depth=0))
        tref_dict = out["book"]["sheets"][0]["cells"][0]["expression"]
        assert "this_row" not in tref_dict

    def test_inline_mode(self):
        node = FunctionNode("SUM", [self._tref()])
        out = yaml.safe_load(to_yaml({"Sheet1!A1": node}, ref_mode="inline"))
        formula = out["book"]["sheets"][0]["cells"][0]["expression"]
        assert formula == "SUM(TABLE_REF(Table1[Amount]))"

    def test_inline_mode_this_row(self):
        node = self._tref(this_row=True)
        out = yaml.safe_load(to_yaml({"Sheet1!A1": node}, ref_mode="inline"))
        formula = out["book"]["sheets"][0]["cells"][0]["expression"]
        assert formula == "TABLE_REF(Table1[@Amount])"
