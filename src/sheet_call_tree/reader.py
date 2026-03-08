"""Extract formula cells from an openpyxl workbook."""
from __future__ import annotations

import logging
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

from .formula_parser import parse_formula
from .models import FunctionNode, NamedRefNode, RangeNode, RefNode, TableRefNode

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
    table_ranges = _build_table_ranges(wb)
    named_ranges = _build_named_ranges(wb)
    _populate_ref_values(cells, data_values, table_ranges, named_ranges)
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


def _build_table_ranges(wb) -> dict[str, dict]:
    """Build a map of table name → column ranges for all tables in the workbook.

    Returns: { "Table1": {"_sheet": "Sheet1", "_range": "Sheet1!A1:B10",
                          "columns": {"ColumnName": "Sheet1!B2:B10"}} }
    """
    result: dict[str, dict] = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for table in ws.tables.values():
            table_ref = table.ref  # e.g. "A1:D10"
            start, end = table_ref.split(":")
            start_col = "".join(c for c in start if c.isalpha())
            start_row = int("".join(c for c in start if c.isdigit()))
            end_row = int("".join(c for c in end if c.isdigit()))
            data_start_row = start_row + 1  # first row is header
            start_col_idx = column_index_from_string(start_col)
            columns: dict[str, str] = {}
            for i, col in enumerate(table.tableColumns):
                col_letter = get_column_letter(start_col_idx + i)
                columns[col.name] = (
                    f"{sheet_name}!{col_letter}{data_start_row}:{col_letter}{end_row}"
                )
            result[table.displayName] = {
                "_sheet": sheet_name,
                "_range": f"{sheet_name}!{table_ref}",
                "columns": columns,
            }
    return result


def _build_named_ranges(wb) -> dict[str, str]:
    """Build a map of defined name → range text for all named ranges in the workbook.

    Returns: { "SalesTotal": "Sheet1!$B$10" }
    """
    result: dict[str, str] = {}
    for name in wb.defined_names:
        dn = wb.defined_names[name]
        result[name] = dn.attr_text
    return result


def _populate_ref_values(
    cells: dict[str, FunctionNode],
    data_values: dict[str, object],
    table_ranges: dict[str, dict] | None = None,
    named_ranges: dict[str, str] | None = None,
) -> None:
    """Walk every AST and fill ref node values.

    - Formula-cell refs: value = FunctionNode of that cell; cached_value from data_only.
    - Constant-cell refs: value = scalar from data_only; cached_value = same scalar.
    - TableRefNode: resolved_range from table_ranges map.
    - NamedRefNode: resolved_range from named_ranges map; value/cached_value if single cell.
    - Unknown refs: fields remain None.
    """
    known = set(cells)
    for fn in cells.values():
        _fill_node(fn, cells, data_values, known, table_ranges or {}, named_ranges or {})


def _fill_node(node, cells, data_values, known, table_ranges, named_ranges):
    if isinstance(node, FunctionNode):
        for arg in node.args:
            _fill_node(arg, cells, data_values, known, table_ranges, named_ranges)
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
        _fill_node(node.start, cells, data_values, known, table_ranges, named_ranges)
        _fill_node(node.end, cells, data_values, known, table_ranges, named_ranges)
    elif isinstance(node, TableRefNode):
        tbl = table_ranges.get(node.table_name)
        if tbl:
            if node.column and node.column in tbl["columns"]:
                node.resolved_range = tbl["columns"][node.column]
            elif not node.column:
                node.resolved_range = tbl["_range"]
    elif isinstance(node, NamedRefNode):
        attr_text = named_ranges.get(node.name)
        if attr_text:
            node.resolved_range = attr_text
            normalized = attr_text.replace("$", "")
            if normalized in known:
                node.value = cells[normalized]
                node.cached_value = data_values.get(normalized)
            elif normalized in data_values:
                node.value = data_values[normalized]
                node.cached_value = data_values[normalized]
