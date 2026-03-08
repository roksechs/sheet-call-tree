"""Build a cell dependency graph and detect circular references."""
from __future__ import annotations

from .models import FunctionNode, RangeNode, RefNode


class CircularReferenceError(Exception):
    """Raised when a circular formula dependency is detected."""

    def __init__(self, cycle: list[str]) -> None:
        self.cycle = cycle
        path = " → ".join(cycle)
        super().__init__(f"Circular reference detected: {path}")


def find_root_cells(graph: dict[str, set[str]]) -> set[str]:
    """Return cells that are not referenced by any other formula cell.

    These are the "entry point" cells — they depend on other cells but
    nothing else depends on them.
    """
    referenced: set[str] = set()
    for deps in graph.values():
        referenced.update(deps)
    return set(graph) - referenced


def build_dependency_graph(formula_cells: dict[str, object]) -> dict[str, set[str]]:
    """Return a mapping from each formula cell to the set of cell refs it depends on.

    Only dependencies that appear in formula_cells are tracked (i.e. formula
    cells that reference other formula cells). Constant cells are not included.

    Args:
        formula_cells: Output of extract_formula_cells — maps cell ref → AST.

    Returns:
        Dict[cell_ref, Set[dependency_cell_ref]]
    """
    known = set(formula_cells)
    graph: dict[str, set[str]] = {ref: set() for ref in known}
    for ref, ast in formula_cells.items():
        _collect_deps(ast, graph[ref], known)
    return graph


def _collect_deps(node, deps: set[str], known: set[str]) -> None:
    """Recursively walk a typed AST node and collect cell reference dependencies."""
    if isinstance(node, RefNode):
        if node.ref in known:
            deps.add(node.ref)
    elif isinstance(node, FunctionNode):
        for child in node.args:
            _collect_deps(child, deps, known)
    elif isinstance(node, RangeNode):
        if node.start in known:
            deps.add(node.start)
        if node.end in known:
            deps.add(node.end)


def detect_cycles(graph: dict[str, set[str]]) -> None:
    """Raise CircularReferenceError if any cycle is found in the dependency graph.

    Uses iterative DFS with a path stack to reconstruct the cycle for reporting.
    """
    visited: set[str] = set()
    in_stack: set[str] = set()

    def dfs(node: str, path: list[str]) -> None:
        visited.add(node)
        in_stack.add(node)
        path.append(node)
        for neighbour in graph.get(node, set()):
            if neighbour not in visited:
                dfs(neighbour, path)
            elif neighbour in in_stack:
                # Found a back-edge — reconstruct cycle
                cycle_start = path.index(neighbour)
                raise CircularReferenceError(path[cycle_start:] + [neighbour])
        path.pop()
        in_stack.discard(node)

    for node in list(graph):
        if node not in visited:
            dfs(node, [])
