"""Typed dataclasses for the formula AST."""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Union

# Forward-declared via strings; resolved at runtime by Union.
Node = Union["FunctionNode", "CellNode", "RangeNode", "TableRefNode", "NamedRefNode", int, float, bool, str]


@dataclass
class FunctionNode:
    type: str       # "SUM", "ADD", "IF", "NEG", …
    inputs: list    # list[Node]


@dataclass
class CellNode:
    cell: str                                  # "Sheet1!A1"  (no @ sigil)
    outputs: object = None                     # scalar value (constant cell value or cached compute)
    labels: dict | None = None                 # semantic labels
    expression: FunctionNode | None = None     # formula cell AST (None for constant/unknown)


@dataclass
class RangeNode:
    start: str                                 # "Sheet1!A1"
    end: str                                   # "Sheet1!A9"
    cells: list[CellNode] | None = None        # all cells in range (populated by reader)


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
    cell: CellNode | None = None        # resolved cell (formula + value)
