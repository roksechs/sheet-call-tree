"""Tests for dependency_graph.py."""
import pytest

from sheet_call_tree.dependency_graph import (
    CircularReferenceError,
    build_dependency_graph,
    detect_cycles,
)
from sheet_call_tree.reader import extract_formula_cells_from_workbook


def test_dependency_graph_simple(simple_workbook):
    cells = extract_formula_cells_from_workbook(simple_workbook)
    graph = build_dependency_graph(cells)
    # B10 depends on C5; B11 depends on C5; C5 has no formula deps
    assert "Sheet1!C5" in graph["Sheet1!B10"]
    assert "Sheet1!C5" in graph["Sheet1!B11"]
    assert graph["Sheet1!C5"] == set()


def test_no_cycle_simple(simple_workbook):
    cells = extract_formula_cells_from_workbook(simple_workbook)
    graph = build_dependency_graph(cells)
    detect_cycles(graph)  # should not raise


def test_cycle_detected(circular_workbook):
    cells = extract_formula_cells_from_workbook(circular_workbook)
    graph = build_dependency_graph(cells)
    with pytest.raises(CircularReferenceError) as exc_info:
        detect_cycles(graph)
    assert len(exc_info.value.cycle) >= 2


def test_circular_error_message(circular_workbook):
    cells = extract_formula_cells_from_workbook(circular_workbook)
    graph = build_dependency_graph(cells)
    with pytest.raises(CircularReferenceError) as exc_info:
        detect_cycles(graph)
    msg = str(exc_info.value)
    assert "Circular reference" in msg
    assert "→" in msg
