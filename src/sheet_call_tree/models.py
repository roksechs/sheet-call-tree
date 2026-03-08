"""Typed dataclasses for the formula AST."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Union

# Forward-declared via strings; resolved at runtime by Union.
Node = Union["FunctionNode", "RefNode", "RangeNode", "TableRefNode", "NamedRefNode", int, float, bool, str]


@dataclass
class FunctionNode:
    name: str       # "SUM", "ADD", "IF", "NEG", …
    args: list      # list[Node]


@dataclass
class RefNode:
    ref: str                          # "Sheet1!A1"  (no @ sigil)
    formula: FunctionNode | None = None   # formula cell AST (None for constant/unknown)
    resolved_value: object = None     # scalar value (constant cell value or cached compute)


@dataclass
class RangeNode:
    start: str                            # "Sheet1!A1"
    end: str                              # "Sheet1!A9"
    values: list[object] | None = None    # all cell values in range (populated by reader)


@dataclass
class TableRefNode:
    table_name: str            # "Table1"
    column: str | None         # "Amount"; None for whole-table reference
    this_row: bool             # True when @-prefixed (Table1[@Amount])
    resolved_range: str | None = None   # "Sheet1!D2:D100"
    cached_value: object = None


@dataclass
class NamedRefNode:
    name: str                  # "SalesTotal"
    resolved_range: str | None = None   # "Sheet1!$B$10"
    formula: FunctionNode | None = None   # formula cell AST after resolution
    resolved_value: object = None         # scalar value


@dataclass
class CellEntry:
    ref: str        # "Sheet1!C5"
    ast: FunctionNode
