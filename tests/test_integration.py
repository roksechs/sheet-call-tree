"""End-to-end tests: .xlsx file → YAML output."""
import json
import math

import yaml

from sheet_call_tree.cli import main
from sheet_call_tree.models import FunctionNode, NamedRefNode, RangeNode, RefNode, TableRefNode
from sheet_call_tree.reader import extract_formula_cells
from sheet_call_tree.serializer import to_json, to_yaml


def _get_cell(parsed, sheet, cell):
    """Return the expression dict for a given sheet+cell from parsed YAML."""
    s = next(s for s in parsed["book"]["sheets"] if s["name"] == sheet)
    return next(c for c in s["cells"] if c["cell"] == cell)


def test_yaml_output_structure(simple_workbook_path):
    cells, dv, *_ = extract_formula_cells(simple_workbook_path)
    output = to_yaml(cells, data_values=dv)
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
    cells, *_ = extract_formula_cells(simple_workbook_path)
    c5 = cells["Sheet1!C5"]
    assert isinstance(c5, FunctionNode)
    assert c5.name == "SUM"
    assert len(c5.args) == 1
    rng = c5.args[0]
    assert isinstance(rng, RangeNode)
    assert rng.start == "Sheet1!A1"
    assert rng.end == "Sheet1!A2"


def test_b10_is_add_node(simple_workbook_path):
    cells, *_ = extract_formula_cells(simple_workbook_path)
    b10 = cells["Sheet1!B10"]
    assert isinstance(b10, FunctionNode)
    assert b10.name == "ADD"
    refs = [a for a in b10.args if isinstance(a, RefNode)]
    assert any(r.ref == "Sheet1!C5" for r in refs)
    floats = [a for a in b10.args if isinstance(a, float)]
    assert 1.1 in floats


def test_b11_is_mul_node(simple_workbook_path):
    cells, *_ = extract_formula_cells(simple_workbook_path)
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
    """depth=0: refs as plain strings (no @ sigil), no expansion."""

    def test_formula_refs_as_strings(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=0, data_values=dv)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"] == {"type": "ADD", "inputs": ["Sheet1!C5", 1.1]}

    def test_range_ref_only(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=0, data_values=dv)
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        assert c5["expression"] == {"type": "SUM", "inputs": ["Sheet1!A1:A2"]}

    def test_cli_default_is_depth0(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"] == {"type": "ADD", "inputs": ["Sheet1!C5", 1.1]}

    def test_cli_depth_0(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10", "--depth", "0"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"] == {"type": "ADD", "inputs": ["Sheet1!C5", 1.1]}


class TestDepth1:
    """depth=1: expand one level of refs."""

    def test_formula_ref_expanded(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=1, data_values=dv)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"]["type"] == "ADD"
        inputs = b10["expression"]["inputs"]
        ref_entry = next((a for a in inputs if isinstance(a, dict) and "cell" in a), None)
        assert ref_entry is not None
        assert ref_entry["cell"] == "Sheet1!C5"
        # At depth 1, C5's SUM has a range ref string (not expanded values)
        assert ref_entry["expression"] == {"type": "SUM", "inputs": ["Sheet1!A1:A2"]}

    def test_range_values_at_depth1(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=1, data_values=dv)
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        # At depth 1, range values are inlined
        assert c5["expression"] == {"type": "SUM", "inputs": [10, 20]}


class TestDepthInf:
    """depth=inf: fully expand all refs and range values."""

    def test_full_expansion(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=math.inf, data_values=dv)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"]["type"] == "ADD"
        inputs = b10["expression"]["inputs"]
        ref_entry = next((a for a in inputs if isinstance(a, dict) and "cell" in a), None)
        assert ref_entry is not None
        assert ref_entry["cell"] == "Sheet1!C5"
        # At depth inf, C5's SUM has range values inlined
        assert ref_entry["expression"] == {"type": "SUM", "inputs": [10, 20]}

    def test_constant_ref_expanded(self, simple_workbook_path):
        """Constant-cell RefNodes show as {cell: ref, outputs: value} when depth allows."""
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=math.inf, data_values=dv)
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        assert c5["expression"] == {"type": "SUM", "inputs": [10, 20]}


class TestDepth2:
    """depth=2: expand two levels."""

    def test_b10_depth2_has_range_values(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, depth=2, data_values=dv)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        inputs = b10["expression"]["inputs"]
        ref_entry = next((a for a in inputs if isinstance(a, dict) and "cell" in a), None)
        assert ref_entry is not None
        # depth 2: C5's inner range is at depth 2, gets values inlined
        assert ref_entry["expression"] == {"type": "SUM", "inputs": [10, 20]}


# ---------------------------------------------------------------------------
# Outputs field tests
# ---------------------------------------------------------------------------

class TestOutputs:
    """Top-level cells and expanded refs show outputs when data_values provided."""

    def test_top_level_cell_has_outputs(self):
        """outputs shown when data_values contains the cell's cached value."""
        cells = {"Sheet1!A1": FunctionNode("ADD", [1, 2])}
        dv = {"Sheet1!A1": 3}
        output = to_yaml(cells, depth=0, data_values=dv)
        parsed = yaml.safe_load(output)
        cell = parsed["book"]["sheets"][0]["cells"][0]
        assert cell["outputs"] == 3

    def test_expanded_ref_has_outputs(self):
        """Expanded RefNode shows outputs from resolved_value."""
        inner = FunctionNode("ADD", [1, 2])
        ref = RefNode("Sheet1!B1", formula=inner, resolved_value=3)
        cells = {"Sheet1!A1": FunctionNode("MUL", [ref, 10])}
        output = to_yaml(cells, depth=1)
        parsed = yaml.safe_load(output)
        expr = parsed["book"]["sheets"][0]["cells"][0]["expression"]
        ref_entry = next(a for a in expr["inputs"] if isinstance(a, dict))
        assert ref_entry["outputs"] == 3

    def test_no_outputs_without_data_values(self):
        cells = {"Sheet1!A1": FunctionNode("ADD", [1, 2])}
        output = to_yaml(cells, depth=0)
        parsed = yaml.safe_load(output)
        cell = parsed["book"]["sheets"][0]["cells"][0]
        assert "outputs" not in cell


# ---------------------------------------------------------------------------
# Root-only mode tests
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Sheet filter tests
# ---------------------------------------------------------------------------

class TestSheetFilter:
    """--sheet filters output to a single sheet."""

    def test_cli_sheet_filter(self, multi_sheet_workbook_path, capsys):
        rc = main([str(multi_sheet_workbook_path), "--sheet", "Sheet1"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        sheet_names = [s["name"] for s in parsed["book"]["sheets"]]
        assert sheet_names == ["Sheet1"]

    def test_cli_sheet_filter_not_found(self, multi_sheet_workbook_path, capsys):
        rc = main([str(multi_sheet_workbook_path), "--sheet", "NoSuchSheet"])
        assert rc == 1
        assert "NoSuchSheet" in capsys.readouterr().err

    def test_cli_sheet_and_filter_combined(self, multi_sheet_workbook_path, capsys):
        rc = main([str(multi_sheet_workbook_path), "--sheet", "Sheet1", "--filter", "Sheet1!B1"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        cell_ids = [c["cell"] for s in parsed["book"]["sheets"] for c in s["cells"]]
        assert cell_ids == ["B1"]


class TestRootsOnly:
    """--roots-only outputs only cells not referenced by other formula cells."""

    def test_roots_only_excludes_referenced(self, simple_workbook_path, capsys):
        # B10 and B11 reference C5, so C5 is not a root
        rc = main([str(simple_workbook_path), "--roots-only"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        cell_ids = [c["cell"] for s in parsed["book"]["sheets"] for c in s["cells"]]
        assert "B10" in cell_ids
        assert "B11" in cell_ids
        assert "C5" not in cell_ids

    def test_roots_only_with_filter(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--roots-only", "--filter", "Sheet1!B10"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        cell_ids = [c["cell"] for s in parsed["book"]["sheets"] for c in s["cells"]]
        assert cell_ids == ["B10"]


# ---------------------------------------------------------------------------
# Legacy --ref-mode backward compat
# ---------------------------------------------------------------------------

class TestRefModeRef:
    """Legacy ref_mode='ref' maps to depth=0."""

    def test_formula_refs_as_strings(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ref", data_values=dv)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"] == {"type": "ADD", "inputs": ["Sheet1!C5", 1.1]}

    def test_cli_legacy_ref_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10", "--ref-mode", "ref"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"] == {"type": "ADD", "inputs": ["Sheet1!C5", 1.1]}


class TestRefModeAst:
    """Legacy ref_mode='ast' maps to depth=inf."""

    def test_formula_ref_expanded(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="ast", data_values=dv)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"]["type"] == "ADD"
        inputs = b10["expression"]["inputs"]
        ref_entry = next((a for a in inputs if isinstance(a, dict) and "cell" in a), None)
        assert ref_entry is not None

    def test_cli_ast_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--filter", "Sheet1!B10", "--ref-mode", "ast"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        inputs = b10["expression"]["inputs"]
        assert any(isinstance(a, dict) and "cell" in a for a in inputs)


class TestRefModeInline:
    """Legacy ref_mode='inline' maps to format=inline + depth=inf."""

    def test_c5_inline(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline", data_values=dv)
        parsed = yaml.safe_load(output)
        c5 = _get_cell(parsed, "Sheet1", "C5")
        assert c5["expression"] == "SUM(10, 20)"

    def test_b10_inline_expands_c5(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline", data_values=dv)
        parsed = yaml.safe_load(output)
        b10 = _get_cell(parsed, "Sheet1", "B10")
        assert b10["expression"] == "ADD(SUM(10, 20), 1.1)"

    def test_b11_inline_expands_c5(self, simple_workbook_path):
        cells, dv, *_ = extract_formula_cells(simple_workbook_path)
        output = to_yaml(cells, ref_mode="inline", data_values=dv)
        parsed = yaml.safe_load(output)
        b11 = _get_cell(parsed, "Sheet1", "B11")
        assert b11["expression"] == "MUL(SUM(10, 20), 2)"

    def test_cli_inline_mode(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--ref-mode", "inline"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert _get_cell(parsed, "Sheet1", "C5")["expression"] == "SUM(10, 20)"

    def test_cli_format_inline(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--format", "inline", "--depth", "inf"])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert _get_cell(parsed, "Sheet1", "C5")["expression"] == "SUM(10, 20)"


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
        expr = _get_cell(p, "Sheet1", "A1")["expression"]
        assert expr == {"type": "SUM", "inputs": [None, 42, None]}

    def test_yaml_keywords_stay_strings(self):
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["true", "false", "null", "yes", "no", "~"])}
        p = self._roundtrip(cells, depth=0)
        expr = _get_cell(p, "Sheet1", "A1")["expression"]
        # All must remain strings, not be coerced to bool/null
        assert expr == {"type": "CONCAT", "inputs": ["true", "false", "null", "yes", "no", "~"]}

    def test_empty_string_and_quotes(self):
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["", "it's", 'a "test"'])}
        p = self._roundtrip(cells, depth=0)
        vals = _get_cell(p, "Sheet1", "A1")["expression"]["inputs"]
        assert vals[0] == ""
        assert vals[1] == "it's"
        assert vals[2] == 'a "test"'

    def test_newline_in_string(self):
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["line1\nline2"])}
        p = self._roundtrip(cells, depth=0)
        assert _get_cell(p, "Sheet1", "A1")["expression"]["inputs"][0] == "line1\nline2"

    def test_float_edge_cases(self):
        cells = {"Sheet1!A1": FunctionNode("ADD", [float("nan"), float("inf"), float("-inf"), 0.0])}
        p = self._roundtrip(cells, depth=0)
        vals = _get_cell(p, "Sheet1", "A1")["expression"]["inputs"]
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
        expr = _get_cell(p, "My Sheet", "A1")["expression"]
        assert expr == {"type": "ADD", "inputs": [{"cell": "My Sheet!B1", "outputs": 5}, 10]}

    def test_colon_space_in_string(self):
        """Strings containing ': ' must be quoted (YAML mapping indicator)."""
        cells = {"Sheet1!A1": FunctionNode("CONCAT", ["key: value", "k : v"])}
        p = self._roundtrip(cells, depth=0)
        vals = _get_cell(p, "Sheet1", "A1")["expression"]["inputs"]
        assert vals[0] == "key: value"
        assert vals[1] == "k : v"

    def test_large_range_values(self):
        cells = {"Sheet1!A1": FunctionNode("SUM", [
            RangeNode("Sheet1!A1", "Sheet1!A100", values=list(range(100))),
        ])}
        p = self._roundtrip(cells, depth=math.inf)
        vals = _get_cell(p, "Sheet1", "A1")["expression"]["inputs"]
        assert vals == list(range(100))

    def test_inline_format_roundtrips(self):
        inner = FunctionNode("SUM", [RangeNode("Sheet1!A1", "Sheet1!A3", values=[1, 2, 3])])
        ref = RefNode("Sheet1!C1", formula=inner, resolved_value=6)
        cells = {"Sheet1!B1": FunctionNode("MUL", [ref, 2])}
        p = self._roundtrip(cells, fmt="inline", depth=math.inf)
        expr = _get_cell(p, "Sheet1", "B1")["expression"]
        assert isinstance(expr, str)
        assert "SUM" in expr


# ---------------------------------------------------------------------------
# Issue #11: Cycle detection exit code
# ---------------------------------------------------------------------------

class TestCycleExitCode:
    """CLI returns exit code 1 and error message on circular references."""

    def test_cli_cycle_exit_code(self, circular_workbook_path, capsys):
        rc = main([str(circular_workbook_path)])
        assert rc == 1

    def test_cli_cycle_error_message(self, circular_workbook_path, capsys):
        main([str(circular_workbook_path)])
        err = capsys.readouterr().err
        assert "Circular reference" in err

    def test_cli_no_cycle_check_skips(self, circular_workbook_path, capsys):
        rc = main([str(circular_workbook_path), "--no-cycle-check"])
        assert rc == 0


# ---------------------------------------------------------------------------
# Issue #8: .xlsm support
# ---------------------------------------------------------------------------

class TestXlsmSupport:
    """.xlsm files are read correctly."""

    def test_cli_xlsm_file(self, simple_workbook_xlsm_path, capsys):
        rc = main([str(simple_workbook_xlsm_path)])
        assert rc == 0
        parsed = yaml.safe_load(capsys.readouterr().out)
        assert "book" in parsed
        cell_ids = [c["cell"] for s in parsed["book"]["sheets"] for c in s["cells"]]
        assert "C5" in cell_ids


# ---------------------------------------------------------------------------
# Issue #3: --format json
# ---------------------------------------------------------------------------

class TestJsonOutput:
    """--format json produces valid JSON with the same structure."""

    def test_cli_json_output(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--format", "json"])
        assert rc == 0
        parsed = json.loads(capsys.readouterr().out)
        assert "book" in parsed
        assert "sheets" in parsed["book"]
        cell_ids = [c["cell"] for s in parsed["book"]["sheets"] for c in s["cells"]]
        assert "C5" in cell_ids

    def test_cli_json_filter(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--format", "json", "--filter", "Sheet1!C5"])
        assert rc == 0
        parsed = json.loads(capsys.readouterr().out)
        cells = parsed["book"]["sheets"][0]["cells"]
        assert len(cells) == 1
        assert cells[0]["cell"] == "C5"

    def test_cli_json_depth(self, simple_workbook_path, capsys):
        rc = main([str(simple_workbook_path), "--format", "json", "--depth", "1", "--filter", "Sheet1!B10"])
        assert rc == 0
        parsed = json.loads(capsys.readouterr().out)
        expr = parsed["book"]["sheets"][0]["cells"][0]["expression"]
        assert expr["type"] == "ADD"
        ref_entry = next(a for a in expr["inputs"] if isinstance(a, dict) and "cell" in a)
        assert ref_entry["cell"] == "Sheet1!C5"

    def test_to_json_nan_handling(self):
        cells = {"Sheet1!A1": FunctionNode("ADD", [float("nan"), float("inf")])}
        result = to_json(cells, depth=0)
        parsed = json.loads(result)
        inputs = parsed["book"]["sheets"][0]["cells"][0]["expression"]["inputs"]
        assert inputs[0] is None  # NaN → null
        assert inputs[1] == "Infinity"  # inf → "Infinity"
