"""Parse Excel formula strings into typed AST nodes.

Uses openpyxl's Tokenizer for lexing, then a modified Shunting-Yard
algorithm to build the AST directly.

Token types from openpyxl Tokenizer:
  OPERAND / RANGE     - cell ref ('A1', 'A1:A9', 'Sheet2!C5')
  OPERAND / NUMBER    - numeric literal
  OPERAND / TEXT      - string literal (with surrounding quotes)
  OPERAND / LOGICAL   - TRUE or FALSE
  OPERAND / ERROR     - #DIV/0! etc.
  FUNC / OPEN         - function name + '(' e.g. 'SUM('
  FUNC / CLOSE        - ')'
  SEP / ARG           - ',' argument separator
  OPERATOR-INFIX      - +, -, *, /, ^, &, =, <>, <, >, <=, >=
  OPERATOR-PREFIX     - unary -, +
  PAREN / OPEN        - '(' grouping
  PAREN / CLOSE       - ')' grouping
  WHITE-SPACE         - spaces (skipped)
"""
from __future__ import annotations

import logging
import re

from openpyxl.formula import Tokenizer

from .models import FunctionNode, NamedRefNode, RangeNode, RefNode, TableRefNode

log = logging.getLogger(__name__)

# Infix operators: (precedence, right_associative, AST node name)
_INFIX: dict[str, tuple[int, bool, str]] = {
    "^":  (5, True,  "POW"),
    "*":  (4, False, "MUL"),
    "/":  (4, False, "DIV"),
    "+":  (3, False, "ADD"),
    "-":  (3, False, "SUB"),
    "&":  (2, False, "CONCAT"),
    "=":  (1, False, "EQ"),
    "<>": (1, False, "NEQ"),
    "<":  (1, False, "LT"),
    ">":  (1, False, "GT"),
    "<=": (1, False, "LTE"),
    ">=": (1, False, "GTE"),
}

_PREFIX_NODES: dict[str, str] = {"-": "NEG", "+": "POS"}


# ---------------------------------------------------------------------------
# Token → operand conversion
# ---------------------------------------------------------------------------

_CELL_RE = re.compile(r"^(\$?[A-Za-z]+\$?\d+)(:\$?[A-Za-z]+\$?\d+)?$")


def _is_cell_ref(val: str) -> bool:
    """Return True if val matches a cell/range coordinate after stripping sheet qualifier and $."""
    bare = val.split("!")[-1].replace("$", "")
    return bool(_CELL_RE.match(bare))


def _parse_table_ref(val: str) -> TableRefNode:
    """'Table1[Column]' or 'Table1[@Column]' → TableRefNode (unresolved)."""
    table_name, rest = val.split("[", 1)
    column = rest.rstrip("]")
    this_row = column.startswith("@")
    if this_row:
        column = column[1:]
    return TableRefNode(
        table_name=table_name,
        column=column if column else None,
        this_row=this_row,
    )


def _qualify(ref: str, sheet: str) -> str:
    """Qualify a cell reference as 'Sheet!Ref' (no @ sigil)."""
    return ref if "!" in ref else f"{sheet}!{ref}"


def _parse_range_token(raw: str, sheet: str):
    """Convert a RANGE token value to a typed AST node.

    - 'A1'          → RefNode('Sheet1!A1')
    - 'A1:A9'       → RangeNode('Sheet1!A1', 'Sheet1!A9')
    - 'Sheet2!C5'   → RefNode('Sheet2!C5')
    - 'Sheet2!A1:B9'→ RangeNode('Sheet2!A1', 'Sheet2!B9')
    """
    clean = raw.replace("$", "")
    if ":" in clean:
        left, right = clean.split(":", 1)
        # Sheet qualifier appears on the left side only for cross-sheet ranges
        if "!" in left and "!" not in right:
            sh, lcell = left.rsplit("!", 1)
            return RangeNode(f"{sh}!{lcell}", f"{sh}!{right}")
        return RangeNode(_qualify(left, sheet), _qualify(right, sheet))
    return RefNode(_qualify(clean, sheet))


def _make_operand(tok, sheet: str):
    sub = tok.subtype
    val = tok.value
    if sub == "RANGE":
        if "[" in val:
            return _parse_table_ref(val)
        if ":" in val or "!" in val or _is_cell_ref(val):
            return _parse_range_token(val, sheet)
        return NamedRefNode(name=val.replace("$", ""))
    if sub == "NUMBER":
        v = float(val)
        return int(v) if v == int(v) else v
    if sub == "TEXT":
        return val[1:-1]  # strip surrounding quotes
    if sub == "LOGICAL":
        return val.upper() == "TRUE"
    return val  # ERROR or unknown subtype — keep as string


# ---------------------------------------------------------------------------
# Shunting-Yard helpers
# ---------------------------------------------------------------------------

def _apply_op(op_type: str, op_val: str, output: list) -> None:
    """Pop operand(s) from output and push the result FunctionNode."""
    if op_type == "INFIX":
        node_name = _INFIX[op_val][2]
        right = output.pop()
        left = output.pop()
        output.append(FunctionNode(node_name, [left, right]))
    elif op_type == "PREFIX":
        node_name = _PREFIX_NODES.get(op_val, f"PREFIX_{op_val}")
        output.append(FunctionNode(node_name, [output.pop()]))


def _should_pop_for_infix(top: tuple, prec: int, right_assoc: bool) -> bool:
    """Return True if the top-of-ops entry should be popped before pushing an infix op."""
    top_type = top[0]
    if top_type == "PREFIX":
        return True  # prefix operators always bind tighter than any infix
    if top_type == "INFIX":
        top_prec = _INFIX.get(top[1], (0,))[0]
        return top_prec > prec or (top_prec == prec and not right_assoc)
    return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def parse_formula(formula: str, default_sheet: str):
    """Parse an Excel formula string into a typed AST.

    Args:
        formula:       The formula string, with or without a leading '='.
        default_sheet: Sheet name used to qualify unqualified cell references.

    Returns:
        A FunctionNode, RefNode, RangeNode, or scalar (int/float/bool/str).
        Returns None on empty input, or the raw formula string on parse failure.
    """
    if not formula.startswith("="):
        formula = "=" + formula
    try:
        tokenizer = Tokenizer(formula)
    except Exception as exc:
        log.warning("Failed to tokenize %r: %s", formula, exc)
        return formula

    tokens = [t for t in tokenizer.items if t.type != "WHITE-SPACE"]
    return _shunting_yard(tokens, default_sheet)


def _shunting_yard(tokens, sheet: str):
    """Build a typed AST from a filtered token list via modified Shunting-Yard.

    The ops stack stores tuples of (type, value, depth_at_open):
      - ("INFIX",  "+",    None)  – binary infix operator
      - ("PREFIX", "-",    None)  – unary prefix operator
      - ("FUNC",   "SUM",  depth) – function call; depth = len(output) when pushed
      - ("PAREN",  "(",    depth) – grouping paren
    """
    output: list = []  # operand/result stack
    ops: list[tuple] = []  # operator stack

    for tok in tokens:
        tt, tv = tok.type, tok.value

        if tt == "OPERAND":
            output.append(_make_operand(tok, sheet))

        elif tt == "FUNC" and tok.subtype == "OPEN":
            ops.append(("FUNC", tv[:-1], len(output)))  # strip trailing '('

        elif tt == "SEP" and tok.subtype == "ARG":
            # Flush pending operators up to the enclosing function/paren boundary
            while ops and ops[-1][0] not in ("FUNC", "PAREN"):
                _apply_op(*ops.pop()[:2], output)

        elif tt == "OPERATOR-INFIX":
            if tv not in _INFIX:
                log.warning("Unsupported infix operator %r — skipping", tv)
                continue
            prec, right_assoc, _ = _INFIX[tv]
            while ops and _should_pop_for_infix(ops[-1], prec, right_assoc):
                _apply_op(*ops.pop()[:2], output)
            ops.append(("INFIX", tv, None))

        elif tt == "OPERATOR-PREFIX":
            ops.append(("PREFIX", tv, None))

        elif tt == "FUNC" and tok.subtype == "CLOSE":
            while ops and ops[-1][0] not in ("FUNC", "PAREN"):
                _apply_op(*ops.pop()[:2], output)
            if ops and ops[-1][0] == "FUNC":
                _, func_name, depth = ops.pop()
                arity = len(output) - depth
                args = [output.pop() for _ in range(arity)][::-1]
                output.append(FunctionNode(func_name, args))

        elif tt == "PAREN" and tok.subtype == "OPEN":
            ops.append(("PAREN", "(", len(output)))

        elif tt == "PAREN" and tok.subtype == "CLOSE":
            while ops and ops[-1][0] != "PAREN":
                _apply_op(*ops.pop()[:2], output)
            if ops:
                ops.pop()  # discard the matching PAREN entry

    # Flush any remaining operators
    while ops:
        _apply_op(*ops.pop()[:2], output)

    if not output:
        return None
    if len(output) > 1:
        log.warning("Unexpected extra items on output stack: %r", output)
    return output[0]
