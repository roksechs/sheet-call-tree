"""Extract formula cells from an openpyxl workbook."""
from __future__ import annotations

import logging
from pathlib import Path

import openpyxl

from .formula_parser import parse_formula
from .models import FunctionNode, RangeNode, RefNode

log = logging.getLogger(__name__)


def extract_formula_cells(path: str | Path) -> dict[str, FunctionNode]:
    """Load an xlsx file and return a dict mapping cell refs to their ASTs.

    Loads the workbook twice: once for formulas (data_only=False) and once for
    cached values (data_only=True). RefNode.value and RefNode.cached_value are
    populated from both sources.

    Only formula cells (values starting with '=') are included as keys.
    Constant cells appear only as RefNode values inside formula ASTs.

    Args:
        path: Path to the .xlsx file.

    Returns:
        Dict mapping 'SheetName!CellRef' strings to FunctionNode AST roots.
    """
    wb = openpyxl.load_workbook(path, data_only=False)
    wb_data = openpyxl.load_workbook(path, data_only=True)

    cells = extract_formula_cells_from_workbook(wb)
    data_values = _extract_all_values(wb_data)
    _populate_ref_values(cells, data_values)
    return cells


def extract_formula_cells_from_workbook(wb: openpyxl.Workbook) -> dict[str, FunctionNode]:
    """Extract formula cells from an already-loaded workbook.

    RefNode.value and RefNode.cached_value are left as None; call
    _populate_ref_values() after loading a data_only workbook if needed.
    """
    result: dict[str, FunctionNode] = {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if not isinstance(val, str) or not val.startswith("="):
                    continue
                col_letter = cell.column_letter
                row_num = cell.row
                cell_ref = f"{sheet_name}!{col_letter}{row_num}"
                ast = parse_formula(val, default_sheet=sheet_name)
                if ast is not None:
                    result[cell_ref] = ast
                else:
                    log.warning("Could not parse formula at %s: %r", cell_ref, val)

    return result


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _extract_all_values(wb: openpyxl.Workbook) -> dict[str, object]:
    """Return a dict of all non-None cell values from a data_only workbook."""
    result: dict[str, object] = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    ref = f"{sheet_name}!{cell.column_letter}{cell.row}"
                    result[ref] = cell.value
    return result


def _populate_ref_values(
    cells: dict[str, FunctionNode],
    data_values: dict[str, object],
) -> None:
    """Walk every AST and fill RefNode.value / RefNode.cached_value.

    - Formula-cell refs: value = FunctionNode of that cell; cached_value from data_only.
    - Constant-cell refs: value = scalar from data_only; cached_value = same scalar.
    - Unknown refs: both fields remain None.
    """
    known = set(cells)
    for fn in cells.values():
        _fill_node(fn, cells, data_values, known)


def _fill_node(node, cells, data_values, known):
    if isinstance(node, FunctionNode):
        for arg in node.args:
            _fill_node(arg, cells, data_values, known)
    elif isinstance(node, RefNode):
        ref = node.ref
        if ref in known:
            node.value = cells[ref]
            node.cached_value = data_values.get(ref)  # None for programmatic xlsx
        elif ref in data_values:
            node.value = data_values[ref]
            node.cached_value = data_values[ref]
        # else: both fields stay None
    elif isinstance(node, RangeNode):
        _fill_node(node.start, cells, data_values, known)
        _fill_node(node.end, cells, data_values, known)
