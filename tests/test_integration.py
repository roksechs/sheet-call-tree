"""End-to-end tests: .xlsx file → YAML output."""
import yaml

from sheet_call_tree.cli import main
from sheet_call_tree.models import FunctionNode, RangeNode, RefNode
from sheet_call_tree.reader import extract_formula_cells
from sheet_call_tree.serializer import to_yaml


def test_yaml_output_structure(simple_workbook_path):
    cells = extract_formula_cells(simple_workbook_path)
    output = to_yaml(cells)
    parsed = yaml.safe_load(output)

    assert "Sheet1!C5" in parsed
    assert "Sheet1!B10" in parsed
    assert "Sheet1!B11" in parsed


def test_c5_is_sum_node(simple_workbook_path):
    cells = extract_formula_cells(simple_workbook_path)
    c5 = cells["Sheet1!C5"]
    assert isinstance(c5, FunctionNode)
    assert c5.name == "SUM"
    assert len(c5.args) == 1
    rng = c5.args[0]
    assert isinstance(rng, RangeNode)
    assert rng.start.ref == "Sheet1!A1"
    assert rng.end.ref == "Sheet1!A2"


def test_b10_is_add_node(simple_workbook_path):
    cells = extract_formula_cells(simple_workbook_path)
    b10 = cells["Sheet1!B10"]
    assert isinstance(b10, FunctionNode)
    assert b10.name == "ADD"
    refs = [a for a in b10.args if isinstance(a, RefNode)]
    assert any(r.ref == "Sheet1!C5" for r in refs)
    floats = [a for a in b10.args if isinstance(a, float)]
    assert 1.1 in floats


def test_b11_is_mul_node(simple_workbook_path):
    cells = extract_formula_cells(simple_workbook_path)
    b11 = cells["Sheet1!B11"]
    assert isinstance(b11, FunctionNode)
    assert b11.name == "MUL"
    refs = [a for a in b11.args if isinstance(a, RefNode)]
    assert any(r.ref == "Sheet1!C5" for r in refs)
    nums = [a for a in b11.args if isinstance(a, (int, float))]
    assert 2 in nums


def test_cli_stdout(simple_workbook_path, capsys):
    rc = main([str(simple_workbook_path)])
    assert rc == 0
    captured = capsys.readouterr()
    parsed = yaml.safe_load(captured.out)
    assert "Sheet1!C5" in parsed


def test_cli_filter(simple_workbook_path, capsys):
    rc = main([str(simple_workbook_path), "--filter", "Sheet1!C5"])
    assert rc == 0
    captured = capsys.readouterr()
    parsed = yaml.safe_load(captured.out)
    assert list(parsed.keys()) == ["Sheet1!C5"]


def test_cli_output_file(simple_workbook_path, tmp_path):
    out = tmp_path / "result.yaml"
    rc = main([str(simple_workbook_path), "--output", str(out)])
    assert rc == 0
    parsed = yaml.safe_load(out.read_text())
    assert "Sheet1!C5" in parsed


def test_cli_filter_missing_cell(simple_workbook_path, capsys):
    rc = main([str(simple_workbook_path), "--filter", "Sheet1!Z99"])
    assert rc == 1


# ---------------------------------------------------------------------------
# ref-mode tests
# ---------------------------------------------------------------------------

class TestRefModeRef:
    """Default mode: formula-cell refs as '@'-prefixed strings."""

    def test_constant_refs_resolve_to_scalars(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ref")
        parsed = yaml.safe_load(output)
        # C5 = SUM(A1:A2); A1=10, A2=20 are constants → scalars in RANGE
        c5 = parsed["Sheet1!C5"]
        assert c5 == {"SUM": [{"RANGE": [10, 20]}]}

    def test_formula_refs_as_at_strings(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ref")
        parsed = yaml.safe_load(output)
        # B10 = C5+1.1; C5 is a formula cell → '@Sheet1!C5'
        b10 = parsed["Sheet1!B10"]
        assert b10 == {"ADD": ["@Sheet1!C5", 1.1]}

    def test_cli_default_is_ref(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert parsed["Sheet1!B10"] == {"ADD": ["@Sheet1!C5", 1.1]}


class TestRefModeAst:
    """ast mode: formula refs as {`@ref`: <expanded AST>}."""

    def test_formula_ref_expanded(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ast")
        parsed = yaml.safe_load(output)
        b10 = parsed["Sheet1!B10"]
        # Should have ADD with '@Sheet1!C5' mapped to its expanded AST
        assert "ADD" in b10
        args = b10["ADD"]
        ref_entry = next((a for a in args if isinstance(a, dict) and "@Sheet1!C5" in a), None)
        assert ref_entry is not None
        assert ref_entry["@Sheet1!C5"] == {"SUM": [{"RANGE": [10, 20]}]}

    def test_constant_still_scalar(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ast")
        parsed = yaml.safe_load(output)
        c5 = parsed["Sheet1!C5"]
        assert c5 == {"SUM": [{"RANGE": [10, 20]}]}

    def test_cli_ast_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10", "--ref-mode", "ast"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = parsed["Sheet1!B10"]
        args = b10["ADD"]
        assert any(isinstance(a, dict) and "@Sheet1!C5" in a for a in args)


class TestRefModeValue:
    """value mode: formula refs replaced by cached scalar (None if not cached)."""

    def test_constant_refs_resolve(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="value")
        parsed = yaml.safe_load(output)
        # A1 and A2 are constants → scalars in RANGE regardless of mode
        c5 = parsed["Sheet1!C5"]
        assert c5 == {"SUM": [{"RANGE": [10, 20]}]}

    def test_formula_ref_is_none_for_programmatic_xlsx(self, simple_workbook_path):
        """Programmatic xlsx has no cached compute → formula refs render as null."""
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="value")
        parsed = yaml.safe_load(output)
        b10 = parsed["Sheet1!B10"]
        # C5 is a formula cell with no cached value → null
        assert b10 == {"ADD": [None, 1.1]}

    def test_cli_value_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!C5", "--ref-mode", "value"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert "Sheet1!C5" in parsed


class TestRefModeInline:
    """inline mode: each cell as a single FUNC(...) expression string."""

    def test_c5_inline(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline")
        parsed = yaml.safe_load(output)
        assert parsed["Sheet1!C5"] == "SUM(RANGE(10, 20))"

    def test_b10_inline_expands_c5(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline")
        parsed = yaml.safe_load(output)
        assert parsed["Sheet1!B10"] == "ADD(SUM(RANGE(10, 20)), 1.1)"

    def test_b11_inline_expands_c5(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline")
        parsed = yaml.safe_load(output)
        assert parsed["Sheet1!B11"] == "MUL(SUM(RANGE(10, 20)), 2)"

    def test_cli_inline_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--ref-mode", "inline"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert parsed["Sheet1!C5"] == "SUM(RANGE(10, 20))"
        assert parsed["Sheet1!B10"] == "ADD(SUM(RANGE(10, 20)), 1.1)"
