"""Serialize typed formula AST nodes to YAML.

Rendering modes (--ref-mode):
  ref    (default) nested YAML dict; formula-cell refs as '@Sheet1!C5' strings
  ast              nested YAML dict; formula-cell refs as {'@Sheet1!C5': <AST>}
  value            nested YAML dict; formula-cell refs as their cached scalar
  inline           each cell is a single YAML string, fully expanded as FUNC(...)
"""
from __future__ import annotations

from .models import FunctionNode, NamedRefNode, RangeNode, RefNode, TableRefNode

_REF_MODES = {"ref", "ast", "value", "inline"}


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
# Public API
# ---------------------------------------------------------------------------

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

    _inline_cache: dict[str, str] | None = {} if ref_mode == "inline" else None
    sheets: dict[str, list] = {}
    for full_ref, node in formula_cells.items():
        sheet, cell = full_ref.split("!", 1)
        if sheet not in sheets:
            sheets[sheet] = []
        if ref_mode == "inline":
            formula = _expr(node, _inline_cache)
        else:
            formula = _to_dict(node, ref_mode)
        sheets[sheet].append((cell, formula))

    buf: list[str] = ["book:\n", f"  name: {_ys(book_name)}\n", "  sheets:\n"]
    for sheet_name, cells in sheets.items():
        buf.append(f"  - name: {_ys(sheet_name)}\n")
        buf.append("    cells:\n")
        for cell_ref, formula in cells:
            buf.append(f"    - cell: {_ys(cell_ref)}\n")
            if isinstance(formula, dict):
                buf.append("      formula:\n")
                _emit_dict(buf, formula, 8)
            elif isinstance(formula, list):
                buf.append("      formula:\n")
                _emit_list(buf, formula, 8)
            else:
                buf.append(f"      formula: {_yscalar(formula)}\n")

    result = "".join(buf)
    if stream is not None:
        stream.write(result)
        return None
    return result


# ---------------------------------------------------------------------------
# Recursive dict converter (ref / ast / value modes)
# ---------------------------------------------------------------------------

def _to_dict(node, ref_mode: str):
    """Convert a typed AST node to a YAML-serializable structure."""
    if isinstance(node, TableRefNode):
        if ref_mode == "value":
            return node.cached_value
        d: dict = {"name": node.table_name}
        if node.column:
            d["column"] = node.column
        if node.this_row:
            d["this_row"] = True
        if node.resolved_range:
            d["range"] = node.resolved_range
        return {"TABLE_REF": d}
    if isinstance(node, NamedRefNode):
        if ref_mode == "value":
            return node.cached_value
        if ref_mode == "ast" and isinstance(node.value, FunctionNode):
            return {f"NAMED_REF({node.name})": _to_dict(node.value, ref_mode)}
        d = {"name": node.name}
        if node.resolved_range:
            d["range"] = node.resolved_range
        return {"NAMED_REF": d}
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

def _expr(node, _cache: dict[str, str] | None = None) -> str:
    """Recursively render a typed AST node as a FUNC(arg1, arg2, …) string."""
    if isinstance(node, TableRefNode):
        at = "@" if node.this_row else ""
        col = f"[{at}{node.column}]" if node.column is not None else "[]"
        return f"TABLE_REF({node.table_name}{col})"

    if isinstance(node, NamedRefNode):
        return f"NAMED_REF({node.name})"

    if isinstance(node, FunctionNode):
        args_str = ", ".join(_expr(arg, _cache) for arg in node.args)
        return f"{node.name}({args_str})"

    if isinstance(node, RangeNode):
        return f"RANGE({_expr(node.start, _cache)}, {_expr(node.end, _cache)})"

    if isinstance(node, RefNode):
        if isinstance(node.value, FunctionNode):
            if _cache is not None:
                cached = _cache.get(node.ref)
                if cached is not None:
                    return cached
                result = _expr(node.value, _cache)
                _cache[node.ref] = result
                return result
            return _expr(node.value, _cache)
        if node.value is not None:
            return _scalar_str(node.value)
        return f"@{node.ref}"  # unknown ref

    # Plain scalar
    return _scalar_str(node)


def _scalar_str(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    return str(v)
