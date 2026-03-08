"""Typed dataclasses for the formula AST."""
from __future__ import annotations

from dataclasses import dataclass
from typing import Union

# Forward-declared via strings; resolved at runtime by Union.
Node = Union["FunctionNode", "RefNode", "RangeNode", int, float, bool, str]


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
class CellEntry:
    ref: str        # "Sheet1!C5"
    ast: FunctionNode
