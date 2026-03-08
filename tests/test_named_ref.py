"""Tests for NamedRefNode — parsing, resolution, and serialization."""
import yaml
import pytest

from sheet_call_tree.formula_parser import parse_formula
from sheet_call_tree.models import FunctionNode, NamedRefNode, RefNode
from sheet_call_tree.reader import (
    _build_named_ranges,
    extract_formula_cells,
    extract_formula_cells_from_workbook,
)
from sheet_call_tree.serializer import to_yaml


SHEET = "Sheet1"


def p(formula):
    return parse_formula(formula, SHEET)


class TestNamedRefParsing:
    def test_simple_named_range(self):
        node = p("=SalesTotal")
        assert isinstance(node, NamedRefNode)
        assert node.name == "SalesTotal"

    def test_named_range_in_arithmetic(self):
        node = p("=SalesTotal+10")
        assert isinstance(node, FunctionNode)
        assert node.name == "ADD"
        named = node.args[0]
        assert isinstance(named, NamedRefNode)
        assert named.name == "SalesTotal"

    def test_named_range_unresolved(self):
        node = p("=SalesTotal")
        assert node.resolved_range is None
        assert node.value is None

    def test_regular_cell_not_named_ref(self):
        node = p("=A1")
        assert isinstance(node, RefNode)
        assert not isinstance(node, NamedRefNode)

    def test_range_not_named_ref(self):
        node = p("=SUM(A1:A10)")
        assert isinstance(node, FunctionNode)

    def test_cross_sheet_not_named_ref(self):
        node = p("=Sheet2!A1")
        assert isinstance(node, RefNode)

    def test_non_ascii_name(self):
        node = p("=売上合計")
        assert isinstance(node, NamedRefNode)
        assert node.name == "売上合計"


class TestNamedRangeBuilding:
    def test_build_named_ranges_keys(self, named_range_workbook):
        nmap = _build_named_ranges(named_range_workbook)
        assert "SalesTotal" in nmap

    def test_build_named_ranges_value(self, named_range_workbook):
        nmap = _build_named_ranges(named_range_workbook)
        assert nmap["SalesTotal"] == "Sheet1!$B$2"

    def test_empty_workbook_no_names(self):
        import openpyxl
        wb = openpyxl.Workbook()
        assert _build_named_ranges(wb) == {}


class TestNamedRefResolution:
    def test_resolved_range_populated(self, named_range_workbook_path):
        cells = extract_formula_cells(named_range_workbook_path)
        c2 = cells["Sheet1!C2"]
        named = c2.args[0]
        assert isinstance(named, NamedRefNode)
        assert named.resolved_range == "Sheet1!$B$2"

    def test_resolved_value_populated(self, named_range_workbook_path):
        cells = extract_formula_cells(named_range_workbook_path)
        c2 = cells["Sheet1!C2"]
        named = c2.args[0]
        assert isinstance(named, NamedRefNode)
        # B2=500 is a constant cell; value and cached_value should be populated
        assert named.value == 500
        assert named.cached_value == 500

    def test_unresolved_when_workbook_only(self, named_range_workbook):
        cells = extract_formula_cells_from_workbook(named_range_workbook)
        c2 = cells["Sheet1!C2"]
        named = c2.args[0]
        assert isinstance(named, NamedRefNode)
        assert named.resolved_range is None


class TestNamedRefSerialization:
    def _nref(self, **kw):
        defaults = dict(name="SalesTotal")
        defaults.update(kw)
        return NamedRefNode(**defaults)

    def test_ref_mode_basic(self):
        node = FunctionNode("ADD", [self._nref(), 10])
        out = yaml.safe_load(to_yaml({"Sheet1!C2": node}, ref_mode="ref"))
        formula = out["book"]["sheets"][0]["cells"][0]["formula"]
        assert formula == {"ADD": [{"NAMED_REF": {"name": "SalesTotal"}}, 10]}

    def test_ref_mode_with_range(self):
        node = self._nref(resolved_range="Sheet1!$B$2")
        out = yaml.safe_load(to_yaml({"Sheet1!C2": node}, ref_mode="ref"))
        named_dict = out["book"]["sheets"][0]["cells"][0]["formula"]["NAMED_REF"]
        assert named_dict["range"] == "Sheet1!$B$2"

    def test_inline_mode(self):
        node = FunctionNode("ADD", [self._nref(), 10])
        out = yaml.safe_load(to_yaml({"Sheet1!C2": node}, ref_mode="inline"))
        formula = out["book"]["sheets"][0]["cells"][0]["formula"]
        assert formula == "ADD(NAMED_REF(SalesTotal), 10)"

    def test_value_mode_returns_cached(self):
        node = self._nref(cached_value=500)
        out = yaml.safe_load(to_yaml({"Sheet1!C2": node}, ref_mode="value"))
        formula = out["book"]["sheets"][0]["cells"][0]["formula"]
        assert formula == 500

    def test_value_mode_none_when_no_cache(self):
        node = self._nref()
        out = yaml.safe_load(to_yaml({"Sheet1!C2": node}, ref_mode="value"))
        formula = out["book"]["sheets"][0]["cells"][0]["formula"]
        assert formula is None

    def test_ast_mode_expands_function_value(self):
        inner = FunctionNode("MUL", [2, 3])
        node = self._nref(value=inner)
        out = yaml.safe_load(to_yaml({"Sheet1!C2": node}, ref_mode="ast"))
        formula = out["book"]["sheets"][0]["cells"][0]["formula"]
        assert "NAMED_REF(SalesTotal)" in formula
        assert formula["NAMED_REF(SalesTotal)"] == {"MUL": [2, 3]}
