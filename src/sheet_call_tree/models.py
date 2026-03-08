"""Typed dataclasses for the formula AST."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Union

# Forward-declared via strings; resolved at runtime by Union.
Node = Union["FunctionNode", "RefNode", "RangeNode", "TableRefNode", "NamedRefNode", int, float, bool, str]


@dataclass
class FunctionNode:
    name: str       # "SUM", "ADD", "IF", "NEG", …
    args: list      # list[Node]


@dataclass
class RefNode:
    ref: str                    # "Sheet1!A1"  (no @ sigil)
    value: object = None        # scalar for constant cells; FunctionNode for formula cells; None if unknown
    cached_value: object = None # data_only computed value (for --ref-mode value); None if unavailable


@dataclass
class RangeNode:
    start: RefNode
    end: RefNode


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
    value: object = None       # scalar or FunctionNode after resolution
    cached_value: object = None


@dataclass
class CellEntry:
    ref: str        # "Sheet1!C5"
    ast: FunctionNode
