"""Serialize typed formula AST nodes to YAML.

Rendering parameters:
  --depth N       Expansion depth (default: 0). 0 = refs only, inf = full expansion.
  --format F      tree (default) | inline

Legacy --ref-mode mapping:
  ref    → --depth 0
  ast    → --depth inf
  inline → --format inline --depth inf
"""
from __future__ import annotations

import math

from .models import FunctionNode, NamedRefNode, RangeNode, RefNode, TableRefNode

_REF_MODES = {"ref", "ast", "inline"}


# ---------------------------------------------------------------------------
# Custom YAML string builder  (replaces yaml.dump — 10-30× faster)
#
# Rules derived from YAML 1.2 block-scalar grammar:
#   - Plain scalars must not start with indicator chars (see _YAML_UNSAFE_START)
#   - Plain scalars must not contain ': ' or ' #'
#   - Strings matching YAML boolean/null keywords must be quoted
#   - Everything else: single-quote style  '...'  ('' escapes a literal quote)
# ---------------------------------------------------------------------------

_YAML_UNSAFE_START = frozenset(" \t-?:,[]{}#&*!|>'\"@`%")
_YAML_KEYWORDS = frozenset(["true", "false", "yes", "no", "on", "off", "null", "~"])


def _ys(v: str) -> str:
    """YAML scalar for a string: plain when safe, single-quoted otherwise."""
    if not v:
        return "''"
    if "\n" in v or "\r" in v or "\t" in v or "\\" in v:
        # Double-quote style for strings with control characters
        escaped = v.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "\\r").replace("\t", "\\t")
        return '"' + escaped + '"'
    if v[0] in _YAML_UNSAFE_START or ": " in v or " #" in v or v.lower() in _YAML_KEYWORDS:
        return "'" + v.replace("'", "''") + "'"
    return v


def _yn(v) -> str:
    """YAML scalar for a non-string Python value (int, float, bool, None)."""
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        if v != v:
            return ".nan"
        if v == float("inf"):
            return ".inf"
        if v == float("-inf"):
            return "-.inf"
        return str(v)
    return _ys(str(v))


def _yscalar(v) -> str:
    return _ys(v) if isinstance(v, str) else _yn(v)


def _emit_dict(buf: list[str], d: dict, indent: int) -> None:
    pad = " " * indent
    for k, v in d.items():
        ks = _yscalar(k)
        if isinstance(v, list):
            buf.append(f"{pad}{ks}:\n")
            _emit_list(buf, v, indent)
        elif isinstance(v, dict):
            buf.append(f"{pad}{ks}:\n")
            _emit_dict(buf, v, indent + 2)
        else:
            buf.append(f"{pad}{ks}: {_yscalar(v)}\n")


def _emit_list(buf: list[str], lst: list, indent: int) -> None:
    pad = " " * indent
    for item in lst:
        if isinstance(item, dict):
            _emit_dict_item(buf, item, indent)
        elif isinstance(item, list):
            buf.append(f"{pad}-\n")
            _emit_list(buf, item, indent + 2)
        else:
            buf.append(f"{pad}- {_yscalar(item)}\n")


def _emit_dict_item(buf: list[str], d: dict, indent: int) -> None:
    """Emit a dict as a list item: first key on the same line as '-'."""
    prefix = " " * indent + "- "
    sub = " " * (indent + 2)
    first = True
    for k, v in d.items():
        ks = _yscalar(k)
        if first:
            first = False
            if isinstance(v, list):
                buf.append(f"{prefix}{ks}:\n")
                _emit_list(buf, v, indent + 2)
            elif isinstance(v, dict):
                buf.append(f"{prefix}{ks}:\n")
                _emit_dict(buf, v, indent + 4)
            else:
                buf.append(f"{prefix}{ks}: {_yscalar(v)}\n")
        else:
            if isinstance(v, list):
                buf.append(f"{sub}{ks}:\n")
                _emit_list(buf, v, indent + 2)
            elif isinstance(v, dict):
                buf.append(f"{sub}{ks}:\n")
                _emit_dict(buf, v, indent + 4)
            else:
                buf.append(f"{sub}{ks}: {_yscalar(v)}\n")


# ---------------------------------------------------------------------------
# Backward-compat ref_mode → depth mapping
# ---------------------------------------------------------------------------

def _resolve_depth(depth: int | float | None, ref_mode: str | None) -> float:
    """Resolve depth from explicit depth or legacy ref_mode."""
    if depth is not None:
        return float(depth)
    if ref_mode is not None:
        if ref_mode == "ref":
            return 0.0
        if ref_mode in ("ast", "inline"):
            return math.inf
        raise ValueError(f"ref_mode must be one of {sorted(_REF_MODES)}, got {ref_mode!r}")
    return 0.0  # default: depth 0 (refs only)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def to_yaml(
    formula_cells: dict[str, object],
    *,
    depth: int | float | None = None,
    fmt: str = "tree",
    ref_mode: str | None = None,
    book_name: str = "",
    data_values: dict[str, object] | None = None,
    stream=None,
) -> str | None:
    """Convert a formula_cells dict to YAML.

    Args:
        formula_cells: Mapping of cell ref strings to FunctionNode AST roots.
        depth:         Expansion depth. 0 = refs only, inf = full expansion.
        fmt:           Output format: 'tree' (default) or 'inline'.
        ref_mode:      Legacy parameter. Maps: 'ref'→depth 0, 'ast'→depth inf,
                       'inline'→inline format + depth inf.
        book_name:     Workbook filename for the top-level 'book.name' field.
        data_values:   Mapping of cell refs to cached scalar values (from
                       data_only workbook). Used to populate 'outputs' fields.
        stream:        Optional writable stream. If provided, YAML is written
                       there and None is returned. Otherwise the YAML string
                       is returned.

    Returns:
        YAML string when stream is None, else None.
    """
    # Handle legacy ref_mode
    if ref_mode == "inline":
        fmt = "inline"
    resolved_depth = _resolve_depth(depth, ref_mode)
    dv = data_values or {}

    is_inline = fmt == "inline"
    _inline_cache: dict[str, str] | None = {} if is_inline else None
    sheets: dict[str, list] = {}
    for full_ref, node in formula_cells.items():
        sheet, cell = full_ref.split("!", 1)
        if sheet not in sheets:
            sheets[sheet] = []
        if is_inline:
            formula = _expr(node, resolved_depth, 0, _inline_cache)
        else:
            formula = _to_dict(node, resolved_depth, 0)
        cell_output = dv.get(full_ref)
        sheets[sheet].append((cell, formula, cell_output))

    buf: list[str] = ["book:\n", f"  name: {_ys(book_name)}\n", "  sheets:\n"]
    for sheet_name, cells in sheets.items():
        buf.append(f"  - name: {_ys(sheet_name)}\n")
        buf.append("    cells:\n")
        for cell_ref, formula, cell_output in cells:
            buf.append(f"    - cell: {_ys(cell_ref)}\n")
            if cell_output is not None:
                buf.append(f"      outputs: {_yscalar(cell_output)}\n")
            if isinstance(formula, dict):
                buf.append("      expression:\n")
                _emit_dict(buf, formula, 8)
            elif isinstance(formula, list):
                buf.append("      expression:\n")
                _emit_list(buf, formula, 8)
            else:
                buf.append(f"      expression: {_yscalar(formula)}\n")

    result = "".join(buf)
    if stream is not None:
        stream.write(result)
        return None
    return result


# ---------------------------------------------------------------------------
# Recursive dict converter (tree mode)
# ---------------------------------------------------------------------------

def _range_ref_str(node: RangeNode) -> str:
    """Build the 'Sheet1!A1:A9' string for a RangeNode (no @ sigil)."""
    if "!" in node.end:
        _, end_cell = node.end.rsplit("!", 1)
    else:
        end_cell = node.end
    return f"{node.start}:{end_cell}"


# Sentinel to signal that a RangeNode was expanded into a list of values
# that should be flattened into the parent's inputs.
_RANGE_EXPAND = type("_RANGE_EXPAND", (), {"__slots__": ("values",)})


def _to_dict(node, depth: float, current_depth: int):
    """Convert a typed AST node to a YAML-serializable structure.

    Format:
      FunctionNode  → {"type": name, "inputs": [...]}
      RefNode       → "Sheet1!C5" (at depth limit)
                    → {"cell": ref, "outputs": value, "expression": ...} (expanded)
      RangeNode     → "Sheet1!A1:A2" (at depth limit)
                    → flattened list of values (expanded, merged into parent inputs)
      TableRefNode  → {"type": "TABLE_REF", "name": ..., ...}
      NamedRefNode  → {"type": "NAMED_REF", "name": ..., ...} (at depth limit)
                    → {"named_ref": name, "outputs": ..., "expression": ...} (expanded)
      Scalar        → raw value
    """
    if isinstance(node, TableRefNode):
        d: dict = {"type": "TABLE_REF", "name": node.table_name}
        if node.column:
            d["column"] = node.column
        if node.this_row:
            d["this_row"] = True
        if node.resolved_range:
            d["range"] = node.resolved_range
        return d

    if isinstance(node, NamedRefNode):
        if current_depth < depth and isinstance(node.formula, FunctionNode):
            d = {"named_ref": node.name}
            if node.resolved_value is not None:
                d["outputs"] = node.resolved_value
            d["expression"] = _to_dict(node.formula, depth, current_depth + 1)
            return d
        if current_depth < depth and node.resolved_value is not None:
            return {"named_ref": node.name, "outputs": node.resolved_value}
        d = {"type": "NAMED_REF", "name": node.name}
        if node.resolved_range:
            d["range"] = node.resolved_range
        return d

    if isinstance(node, FunctionNode):
        inputs = []
        for arg in node.args:
            result = _to_dict(arg, depth, current_depth)
            if isinstance(result, _RANGE_EXPAND):
                inputs.extend(result.values)
            else:
                inputs.append(result)
        return {"type": node.name, "inputs": inputs}

    if isinstance(node, RangeNode):
        if current_depth < depth and node.values is not None:
            marker = _RANGE_EXPAND()
            marker.values = node.values
            return marker
        return _range_ref_str(node)

    if isinstance(node, RefNode):
        ref = node.ref
        if current_depth >= depth:
            return ref
        if node.formula is not None:
            d = {"cell": ref}
            if node.resolved_value is not None:
                d["outputs"] = node.resolved_value
            d["expression"] = _to_dict(node.formula, depth, current_depth + 1)
            return d
        if node.resolved_value is not None:
            return {"cell": ref, "outputs": node.resolved_value}
        return ref  # unknown ref

    # Scalar (int, float, bool, str)
    return node


# ---------------------------------------------------------------------------
# Inline expression renderer
# ---------------------------------------------------------------------------

def _expr(node, depth: float, current_depth: int, _cache: dict[str, str] | None = None) -> str:
    """Recursively render a typed AST node as a FUNC(arg1, arg2, …) string."""
    if isinstance(node, TableRefNode):
        at = "@" if node.this_row else ""
        col = f"[{at}{node.column}]" if node.column is not None else "[]"
        return f"TABLE_REF({node.table_name}{col})"

    if isinstance(node, NamedRefNode):
        return f"NAMED_REF({node.name})"

    if isinstance(node, FunctionNode):
        parts = []
        for arg in node.args:
            if isinstance(arg, RangeNode) and current_depth < depth and arg.values is not None:
                parts.extend(_scalar_str(v) for v in arg.values)
            else:
                parts.append(_expr(arg, depth, current_depth, _cache))
        args_str = ", ".join(parts)
        return f"{node.name}({args_str})"

    if isinstance(node, RangeNode):
        ref_str = _range_ref_str(node)
        if current_depth < depth and node.values is not None:
            vals = ", ".join(_scalar_str(v) for v in node.values)
            return f"[{vals}]"
        return ref_str

    if isinstance(node, RefNode):
        ref = node.ref
        if current_depth >= depth:
            return ref
        if node.formula is not None:
            if _cache is not None:
                cached = _cache.get(node.ref)
                if cached is not None:
                    return cached
                result = _expr(node.formula, depth, current_depth + 1, _cache)
                _cache[node.ref] = result
                return result
            return _expr(node.formula, depth, current_depth + 1, _cache)
        if node.resolved_value is not None:
            return _scalar_str(node.resolved_value)
        return ref  # unknown ref

    # Plain scalar
    return _scalar_str(node)


def _scalar_str(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if v is None:
        return "null"
    return str(v)
