"""CLI entrypoint for sheet-call-tree."""
from __future__ import annotations

import argparse
import sys

from ._i18n import get_strings
from .dependency_graph import build_dependency_graph, detect_cycles
from .reader import extract_formula_cells
from .serializer import to_yaml


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
        "--ref-mode",
        choices=["ref", "ast", "value", "inline"],
        default="ref",
        dest="ref_mode",
        help=s["ref_mode"],
    )

    args = parser.parse_args(argv)

    formula_cells = extract_formula_cells(args.input)

    if not args.no_cycle_check:
        graph = build_dependency_graph(formula_cells)
        detect_cycles(graph)

    if args.filter_cell:
        if args.filter_cell not in formula_cells:
            print(
                s["err_not_found"].format(ref=args.filter_cell),
                file=sys.stderr,
            )
            return 1
        formula_cells = {args.filter_cell: formula_cells[args.filter_cell]}

    if args.output:
        with open(args.output, "w", encoding="utf-8") as fh:
            to_yaml(formula_cells, ref_mode=args.ref_mode, stream=fh)
    else:
        print(to_yaml(formula_cells, ref_mode=args.ref_mode), end="")

    return 0


if __name__ == "__main__":
    sys.exit(main())
