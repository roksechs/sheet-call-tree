"""Tests for reader.py — workbook → formula cells dict."""
from sheet_call_tree.models import FunctionNode, RangeNode, RefNode
from sheet_call_tree.reader import extract_formula_cells, extract_formula_cells_from_workbook


def test_extracts_formula_cells_only(simple_workbook):
    result = extract_formula_cells_from_workbook(simple_workbook)
    # Only formula cells should be present; A1 and A2 are constants
    assert "Sheet1!A1" not in result
    assert "Sheet1!A2" not in result
    assert "Sheet1!C5" in result
    assert "Sheet1!B10" in result
    assert "Sheet1!B11" in result


def test_formula_cell_count(simple_workbook):
    result = extract_formula_cells_from_workbook(simple_workbook)
    assert len(result) == 3


def test_cell_ref_format(simple_workbook):
    result = extract_formula_cells_from_workbook(simple_workbook)
    for key in result:
        assert "!" in key, f"Key {key!r} should contain '!'"


def test_returns_function_nodes(simple_workbook):
    result = extract_formula_cells_from_workbook(simple_workbook)
    assert isinstance(result["Sheet1!C5"], FunctionNode)
    assert isinstance(result["Sheet1!B10"], FunctionNode)
    assert isinstance(result["Sheet1!B11"], FunctionNode)


def test_c5_ast_structure(simple_workbook):
    result = extract_formula_cells_from_workbook(simple_workbook)
    c5 = result["Sheet1!C5"]
    assert c5.name == "SUM"
    assert len(c5.args) == 1
    rng = c5.args[0]
    assert isinstance(rng, RangeNode)
    assert rng.start == "Sheet1!A1"
    assert rng.end == "Sheet1!A2"


def test_b10_ast_structure(simple_workbook):
    result = extract_formula_cells_from_workbook(simple_workbook)
    b10 = result["Sheet1!B10"]
    assert b10.name == "ADD"
    refs = [a for a in b10.args if isinstance(a, RefNode)]
    assert any(r.ref == "Sheet1!C5" for r in refs)
    floats = [a for a in b10.args if isinstance(a, float)]
    assert 1.1 in floats


def test_reads_from_file(simple_workbook_path):
    result = extract_formula_cells(simple_workbook_path)
    assert "Sheet1!C5" in result
    assert "Sheet1!B10" in result


def test_file_load_populates_range_values(simple_workbook_path):
    """When loaded from file, RangeNode.values gets populated with cell values."""
    result = extract_formula_cells(simple_workbook_path)
    c5 = result["Sheet1!C5"]
    rng = c5.args[0]
    assert isinstance(rng, RangeNode)
    # A1=10 and A2=20 are constant cells; values should be populated
    assert rng.values == [10, 20]


def test_file_load_populates_formula_ref_values(simple_workbook_path):
    """When loaded from file, formula-cell RefNodes get the referenced FunctionNode."""
    result = extract_formula_cells(simple_workbook_path)
    b10 = result["Sheet1!B10"]
    c5_ref = next(a for a in b10.args if isinstance(a, RefNode) and a.ref == "Sheet1!C5")
    assert isinstance(c5_ref.formula, FunctionNode)
    assert c5_ref.formula is result["Sheet1!C5"]


def test_multi_sheet(multi_sheet_workbook):
    result = extract_formula_cells_from_workbook(multi_sheet_workbook)
    assert "Sheet1!B1" in result
    assert "Sheet2!A1" in result
