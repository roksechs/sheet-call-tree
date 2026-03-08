"""CLI entrypoint for sheet-call-tree."""
from __future__ import annotations

import argparse
import math
import sys
import warnings
from pathlib import Path

from ._i18n import get_strings
from .dependency_graph import build_dependency_graph, detect_cycles, find_root_cells
from .reader import extract_formula_cells
from .serializer import to_yaml


def _parse_depth(value: str) -> float:
    """Parse --depth value: integer or 'inf'."""
    if value.lower() == "inf":
        return math.inf
    return int(value)


def main(argv=None) -> int:
    s = get_strings()
    parser = argparse.ArgumentParser(
        prog="sheet-call-tree",
        description=s["description"],
    )
    parser.add_argument("input", help=s["input"])
    parser.add_argument(
        "--filter",
        metavar="CELL",
        dest="filter_cell",
        help=s["filter"],
    )
    parser.add_argument(
        "--output",
        metavar="FILE",
        help=s["output"],
    )
    parser.add_argument(
        "--no-cycle-check",
        action="store_true",
        help=s["no_cycle"],
    )
    parser.add_argument(
        "--depth",
        metavar="N",
        type=_parse_depth,
        default=None,
        help="Expansion depth: 0 = refs only (default), inf = full expansion.",
    )
    parser.add_argument(
        "--format",
        choices=["tree", "inline"],
        default="tree",
        dest="fmt",
        help="Output format: tree (default) or inline.",
    )
    parser.add_argument(
        "--roots-only",
        action="store_true",
        help="Output only root cells (cells not referenced by any other formula cell).",
    )
    # Legacy --ref-mode (deprecated)
    parser.add_argument(
        "--ref-mode",
        choices=["ref", "ast", "inline"],
        default=None,
        dest="ref_mode",
        help=argparse.SUPPRESS,
    )

    args = parser.parse_args(argv)

    # Handle legacy --ref-mode
    ref_mode = None
    depth = args.depth
    fmt = args.fmt
    if args.ref_mode is not None:
        warnings.warn("--ref-mode is deprecated; use --depth and --format instead", DeprecationWarning, stacklevel=2)
        ref_mode = args.ref_mode
        if args.ref_mode == "inline":
            fmt = "inline"

    formula_cells, data_values, label_map = extract_formula_cells(args.input)

    if not args.no_cycle_check:
        graph = build_dependency_graph(formula_cells)
        detect_cycles(graph)

    if args.roots_only:
        if not args.no_cycle_check:
            roots = find_root_cells(graph)
        else:
            graph = build_dependency_graph(formula_cells)
            roots = find_root_cells(graph)
        formula_cells = {ref: ast for ref, ast in formula_cells.items() if ref in roots}

    if args.filter_cell:
        if args.filter_cell not in formula_cells:
            print(
                s["err_not_found"].format(ref=args.filter_cell),
                file=sys.stderr,
            )
            return 1
        formula_cells = {args.filter_cell: formula_cells[args.filter_cell]}

    yaml_kw = dict(depth=depth, fmt=fmt, ref_mode=ref_mode, book_name=Path(args.input).name, data_values=data_values, label_map=label_map)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as fh:
            to_yaml(formula_cells, stream=fh, **yaml_kw)
    else:
        print(to_yaml(formula_cells, **yaml_kw), end="")

    return 0


if __name__ == "__main__":
    sys.exit(main())
