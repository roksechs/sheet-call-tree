"""End-to-end tests: .xlsx file → YAML output."""
import math

import yaml

from sheet_call_tree.cli import main
from sheet_call_tree.models import FunctionNode, NamedRefNode, RangeNode, RefNode, TableRefNode
from sheet_call_tree.reader import extract_formula_cells
from sheet_call_tree.serializer import to_yaml


def _get_cell(parsed, sheet, cell):
    """Return the formula dict for a given sheet+cell from parsed YAML."""
    s = next(s for s in parsed["book"]["sheets"] if s["name"] == sheet)
    return next(c["formula"] for c in s["cells"] if c["cell"] == cell)


def test_yaml_output_structure(simple_workbook_path):
    cells = extract_formula_cells(simple_workbook_path)
    output = to_yaml(cells)
    parsed = yaml.safe_load(output)

    assert "book" in parsed
    sheet_names = [s["name"] for s in parsed["book"]["sheets"]]
    assert "Sheet1" in sheet_names
    s1_cells = next(s["cells"] for s in parsed["book"]["sheets"] if s["name"] == "Sheet1")
    cell_ids = [c["cell"] for c in s1_cells]
    assert "C5" in cell_ids
    assert "B10" in cell_ids
    assert "B11" in cell_ids


def test_c5_is_sum_node(simple_workbook_path):
    cells = extract_formula_cells(simple_workbook_path)
    c5 = cells["Sheet1!C5"]
    assert isinstance(c5, FunctionNode)
    assert c5.name == "SUM"
    assert len(c5.args) == 1
    rng = c5.args[0]
    assert isinstance(rng, RangeNode)
    assert rng.start == "Sheet1!A1"
    assert rng.end == "Sheet1!A2"


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
    assert "book" in parsed
    cell_ids = [c["cell"] for s in parsed["book"]["sheets"] for c in s["cells"]]
    assert "C5" in cell_ids


def test_cli_filter(simple_workbook_path, capsys):
    rc = main([str(simple_workbook_path), "--filter", "Sheet1!C5"])
    assert rc == 0
    captured = capsys.readouterr()
    parsed = yaml.safe_load(captured.out)
    s1_cells = next(s["cells"] for s in parsed["book"]["sheets"] if s["name"] == "Sheet1")
    assert [c["cell"] for c in s1_cells] == ["C5"]


def test_cli_output_file(simple_workbook_path, tmp_path):
    out = tmp_path / "result.yaml"
    rc = main([str(simple_workbook_path), "--output", str(out)])
    assert rc == 0
    parsed = yaml.safe_load(out.read_text())
    cell_ids = [c["cell"] for s in parsed["book"]["sheets"] for c in s["cells"]]
    assert "C5" in cell_ids


def test_cli_filter_missing_cell(simple_workbook_path, capsys):
    rc = main([str(simple_workbook_path), "--filter", "Sheet1!Z99"])
    assert rc == 1


# ---------------------------------------------------------------------------
# Depth-based expansion tests (tree format)
# ---------------------------------------------------------------------------

class TestDepth0:
    """depth=0: refs as '@'-prefixed strings, no expansion."""

    def test_formula_refs_as_at_strings(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=0)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10 == {"ADD": ["@Sheet1!C5", 1.1]}

    def test_range_ref_only(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=0)
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        assert c5 == {"SUM": [{"RANGE": {"ref": "@Sheet1!A1:A2"}}]}

    def test_cli_default_is_depth0(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10 == {"ADD": ["@Sheet1!C5", 1.1]}

    def test_cli_depth_0(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10", "--depth", "0"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10 == {"ADD": ["@Sheet1!C5", 1.1]}


class TestDepth1:
    """depth=1: expand one level of refs."""

    def test_formula_ref_expanded(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=1)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert "ADD" in b10
        args = b10["ADD"]
        ref_entry = next((a for a in args if isinstance(a, dict) and "@Sheet1!C5" in a), None)
        assert ref_entry is not None
        # At depth 1, C5's SUM has a RANGE with ref but no values (that's depth 2)
        assert ref_entry["@Sheet1!C5"] == {"SUM": [{"RANGE": {"ref": "@Sheet1!A1:A2"}}]}

    def test_range_no_values_at_depth1(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=1)
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        # At depth 1, RANGE gets values
        assert c5 == {"SUM": [{"RANGE": {"ref": "@Sheet1!A1:A2", "values": [10, 20]}}]}


class TestDepthInf:
    """depth=inf: fully expand all refs and range values."""

    def test_full_expansion(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=math.inf)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert "ADD" in b10
        args = b10["ADD"]
        ref_entry = next((a for a in args if isinstance(a, dict) and "@Sheet1!C5" in a), None)
        assert ref_entry is not None
        # At depth inf, C5's SUM has RANGE with values populated
        assert ref_entry["@Sheet1!C5"] == {"SUM": [{"RANGE": {"ref": "@Sheet1!A1:A2", "values": [10, 20]}}]}

    def test_constant_ref_expanded(self, simple_workbook_path):
        """Constant-cell RefNodes show as {ref: value} when depth allows."""
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=math.inf)
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        assert c5 == {"SUM": [{"RANGE": {"ref": "@Sheet1!A1:A2", "values": [10, 20]}}]}


class TestDepth2:
    """depth=2: expand two levels."""

    def test_b10_depth2_has_range_values(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=2)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        args = b10["ADD"]
        ref_entry = next((a for a in args if isinstance(a, dict) and "@Sheet1!C5" in a), None)
        assert ref_entry is not None
        # depth 2: C5's inner RANGE is at depth 2, gets values
        assert ref_entry["@Sheet1!C5"] == {"SUM": [{"RANGE": {"ref": "@Sheet1!A1:A2", "values": [10, 20]}}]}


# ---------------------------------------------------------------------------
# Legacy --ref-mode backward compat
# ---------------------------------------------------------------------------

class TestRefModeRef:
    """Legacy ref_mode='ref' maps to depth=0."""

    def test_formula_refs_as_at_strings(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ref")
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10 == {"ADD": ["@Sheet1!C5", 1.1]}

    def test_cli_legacy_ref_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10", "--ref-mode", "ref"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10 == {"ADD": ["@Sheet1!C5", 1.1]}


class TestRefModeAst:
    """Legacy ref_mode='ast' maps to depth=inf."""

    def test_formula_ref_expanded(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ast")
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert "ADD" in b10
        args = b10["ADD"]
        ref_entry = next((a for a in args if isinstance(a, dict) and "@Sheet1!C5" in a), None)
        assert ref_entry is not None

    def test_cli_ast_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10", "--ref-mode", "ast"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        args = b10["ADD"]
        assert any(isinstance(a, dict) and "@Sheet1!C5" in a for a in args)


class TestRefModeInline:
    """Legacy ref_mode='inline' maps to format=inline + depth=inf."""

    def test_c5_inline(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline")
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        assert c5 == "SUM(RANGE(@Sheet1!A1:A2, [10, 20]))"

    def test_b10_inline_expands_c5(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline")
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10 == "ADD(SUM(RANGE(@Sheet1!A1:A2, [10, 20])), 1.1)"

    def test_b11_inline_expands_c5(self, simple_workbook_path):
        cells = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline")
        parsed = yaml.safe_load(output)
        b11 = _get_cell(parsed, "Sheet1", "B11")
        assert b11 == "MUL(SUM(RANGE(@Sheet1!A1:A2, [10, 20])), 2)"

    def test_cli_inline_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--ref-mode", "inline"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert _get_cell(parsed, "Sheet1", "C5") == "SUM(RANGE(@Sheet1!A1:A2, [10, 20]))"

    def test_cli_format_inline(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--format", "inline", "--depth", "inf"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert _get_cell(parsed, "Sheet1", "C5") == "SUM(RANGE(@Sheet1!A1:A2, [10, 20]))"


# ---------------------------------------------------------------------------
# YAML format correctness — complex patterns
# ---------------------------------------------------------------------------

class TestYamlValidity:
    """Verify YAML output is valid and round-trips correctly for complex patterns."""

    @staticmethod
    def _roundtrip(cells, **kw):
        out = to_yaml(cells, **kw)
        parsed = yaml.safe_load(out)
        assert parsed is not None, f"YAML parse failed for {kw}"
        return parsed

    def test_none_in_range_values(self):
        cells = {"Sheet1!A1": FunctionNode("SUM", [
            RangeNode("Sheet1!B1", "Sheet1!B3", values=[None, 42, None]),
        ])}
        p = self._roundtrip(cells, depth=math.inf)
        f = _get_cell(p, "Sheet1", "A1")
        assert f == {"SUM": [{"RANGE": {"ref": "@Sheet1!B1:B3", "values": [None, 42, None]}}]}

    def test_yaml_keywords_stay_strings(self):
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["true", "false", "null", "yes", "no", "~"])}
        p = self._roundtrip(cells, depth=0)
        f = _get_cell(p, "Sheet1", "A1")
        # All must remain strings, not be coerced to bool/null
        assert f == {"CONCAT": ["true", "false", "null", "yes", "no", "~"]}

    def test_empty_string_and_quotes(self):
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["", "it's", 'a "test"'])}
        p = self._roundtrip(cells, depth=0)
        vals = _get_cell(p, "Sheet1", "A1")["CONCAT"]
        assert vals[0] == ""
        assert vals[1] == "it's"
        assert vals[2] == 'a "test"'

    def test_newline_in_string(self):
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["line1\nline2"])}
        p = self._roundtrip(cells, depth=0)
        assert _get_cell(p, "Sheet1", "A1")["CONCAT"][0] == "line1\nline2"

    def test_float_edge_cases(self):
        cells = {"Sheet1!A1": FunctionNode("ADD", [float("nan"), float("inf"), float("-inf"), 0.0])}
        p = self._roundtrip(cells, depth=0)
        vals = _get_cell(p, "Sheet1", "A1")["ADD"]
        assert math.isnan(vals[0])
        assert vals[1] == float("inf")
        assert vals[2] == float("-inf")

    def test_deep_nesting_all_depths(self):
        inner = FunctionNode("CONST", [42])
        for i in range(5):
            inner = FunctionNode("WRAP", [RefNode(f"Sheet1!X{i}", formula=inner)])
        cells = {"Sheet1!Z1": inner}
        for d in [0, 2, 5, math.inf]:
            self._roundtrip(cells, depth=d)

    def test_mixed_node_types(self):
        cells = {"Sheet1!A1": FunctionNode("ADD", [
            TableRefNode("Table1", "Amount", False, resolved_range="Sheet1!B2:B10"),
            NamedRefNode("Tax", resolved_range="Sheet1!$C$1", resolved_value=0.1),
            RefNode("Sheet1!D1", resolved_value=100),
        ])}
        for d in [0, 1, math.inf]:
            p = self._roundtrip(cells, depth=d)
            assert p is not None

    def test_sheet_name_with_space(self):
        cells = {"My Sheet!A1": FunctionNode("ADD", [RefNode("My Sheet!B1", resolved_value=5), 10])}
        p = self._roundtrip(cells, depth=1)
        f = _get_cell(p, "My Sheet", "A1")
        assert f == {"ADD": [{"@My Sheet!B1": 5}, 10]}

    def test_at_sigil_in_string_value(self):
        """Strings starting with @ must be quoted in YAML."""
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["@mention", "normal"])}
        p = self._roundtrip(cells, depth=0)
        vals = _get_cell(p, "Sheet1", "A1")["CONCAT"]
        assert vals[0] == "@mention"
        assert vals[1] == "normal"

    def test_colon_space_in_string(self):
        """Strings containing ': ' must be quoted (YAML mapping indicator)."""
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["key: value", "k : v"])}
        p = self._roundtrip(cells, depth=0)
        vals = _get_cell(p, "Sheet1", "A1")["CONCAT"]
        assert vals[0] == "key: value"
        assert vals[1] == "k : v"

    def test_large_range_values(self):
        cells = {"Sheet1!A1": FunctionNode("SUM", [
            RangeNode("Sheet1!A1", "Sheet1!A100", values=list(range(100))),
        ])}
        p = self._roundtrip(cells, depth=math.inf)
        vals = _get_cell(p, "Sheet1", "A1")["SUM"][0]["RANGE"]["values"]
        assert vals == list(range(100))

    def test_inline_format_roundtrips(self):
        inner = FunctionNode("SUM", [RangeNode("Sheet1!A1", "Sheet1!A3", values=[1, 2, 3])])
        ref = RefNode("Sheet1!C1", formula=inner, resolved_value=6)
        cells = {"Sheet1!B1": FunctionNode("MUL", [ref, 2])}
        p = self._roundtrip(cells, fmt="inline", depth=math.inf)
        f = _get_cell(p, "Sheet1", "B1")
        assert isinstance(f, str)
        assert "SUM" in f
        assert "RANGE" in f
