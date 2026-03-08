"""Serialize typed formula AST nodes to YAML.

Rendering modes (--ref-mode):
  ref    (default) nested YAML dict; formula-cell refs as '@Sheet1!C5' strings
  ast              nested YAML dict; formula-cell refs as {'@Sheet1!C5': <AST>}
  value            nested YAML dict; formula-cell refs as their cached scalar
  inline           each cell is a single YAML string, fully expanded as FUNC(...)
"""
from __future__ import annotations

import yaml

from .models import FunctionNode, RangeNode, RefNode

_REF_MODES = {"ref", "ast", "value", "inline"}


class _NoAliasDumper(yaml.Dumper):
    def ignore_aliases(self, data):
        return True


def to_yaml(
    formula_cells: dict[str, object],
    ref_mode: str = "ref",
    book_name: str = "",
    stream=None,
) -> str | None:
    """Convert a formula_cells dict to YAML.

    Args:
        formula_cells: Mapping of cell ref strings to FunctionNode AST roots.
        ref_mode:      How to render formula-cell references.
                       One of 'ref' (default), 'ast', 'value', 'inline'.
        book_name:     Workbook filename for the top-level 'book.name' field.
        stream:        Optional writable stream. If provided, YAML is written
                       there and None is returned. Otherwise the YAML string
                       is returned.

    Returns:
        YAML string when stream is None, else None.
    """
    if ref_mode not in _REF_MODES:
        raise ValueError(f"ref_mode must be one of {sorted(_REF_MODES)}, got {ref_mode!r}")

    # Group cells by sheet name, preserving insertion order.
    sheets: dict[str, list[dict]] = {}
    for full_ref, node in formula_cells.items():
        sheet, cell = full_ref.split("!", 1)
        if sheet not in sheets:
            sheets[sheet] = []
        if ref_mode == "inline":
            formula = _expr(node)
        else:
            formula = _to_dict(node, ref_mode)
        sheets[sheet].append({"cell": cell, "formula": formula})

    document = {
        "book": {
            "name": book_name,
            "sheets": [
                {"name": sheet_name, "cells": cells}
                for sheet_name, cells in sheets.items()
            ],
        }
    }

    return yaml.dump(
        document,
        Dumper=_NoAliasDumper,
        stream=stream,
        default_flow_style=False,
        allow_unicode=True,
        sort_keys=False,
    )


# ---------------------------------------------------------------------------
# Recursive dict converter (ref / ast / value modes)
# ---------------------------------------------------------------------------

def _to_dict(node, ref_mode: str):
    """Convert a typed AST node to a YAML-serializable structure."""
    if isinstance(node, FunctionNode):
        return {node.name: [_to_dict(arg, ref_mode) for arg in node.args]}
    if isinstance(node, RangeNode):
        return {"RANGE": [_render_ref(node.start, ref_mode), _render_ref(node.end, ref_mode)]}
    if isinstance(node, RefNode):
        return _render_ref(node, ref_mode)
    # Scalar (int, float, bool, str)
    return node


def _render_ref(ref_node: RefNode, ref_mode: str):
    """Render a RefNode according to the current ref_mode.

    Constant-cell refs (value is a scalar) always resolve to that scalar.
    Formula-cell refs (value is a FunctionNode) are rendered per ref_mode.
    Unknown refs (value is None) are treated like formula-cell refs.
    """
    # Constant cell: scalar in all modes
    if ref_node.value is not None and not isinstance(ref_node.value, FunctionNode):
        return ref_node.value

    at_ref = f"@{ref_node.ref}"

    if ref_mode == "ref":
        return at_ref

    if ref_mode == "ast":
        if isinstance(ref_node.value, FunctionNode):
            return {at_ref: _to_dict(ref_node.value, ref_mode)}
        return at_ref  # unknown ref — no AST to expand

    if ref_mode == "value":
        return ref_node.cached_value  # None for programmatic xlsx without cached values

    # Should not reach here for valid ref_mode
    return at_ref


# ---------------------------------------------------------------------------
# Inline expression renderer
# ---------------------------------------------------------------------------

def _expr(node) -> str:
    """Recursively render a typed AST node as a FUNC(arg1, arg2, …) string."""
    if isinstance(node, FunctionNode):
        args_str = ", ".join(_expr(arg) for arg in node.args)
        return f"{node.name}({args_str})"

    if isinstance(node, RangeNode):
        return f"RANGE({_expr(node.start)}, {_expr(node.end)})"

    if isinstance(node, RefNode):
        if isinstance(node.value, FunctionNode):
            return _expr(node.value)  # recursively expand formula cell
        if node.value is not None:
            return _scalar_str(node.value)
        return f"@{node.ref}"  # unknown ref

    # Plain scalar
    return _scalar_str(node)


def _scalar_str(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    return str(v)
