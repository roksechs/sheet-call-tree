# sheet-call-tree

Visualize Excel formula dependencies as a YAML AST call tree.

[日本語版 README](README.ja.md)

## What it does

Excel formulas can reference other cells that themselves contain formulas, creating
dependency chains that are hard to follow in a spreadsheet UI. `sheet-call-tree`
loads an `.xlsx` file, parses every formula cell into an abstract syntax tree (AST),
resolves cross-cell references, and emits the complete dependency tree as YAML — one
top-level key per formula cell.

## Installation

```bash
uv pip install git+https://github.com/roksechs/sheet-call-tree.git
```

See [user_manuals/installation.md](user_manuals/installation.md) for full details and
editable dev install instructions.

## 60-second quickstart

Given a workbook with these cells:

| Cell | Value |
|------|-------|
| A1 | `10` |
| A2 | `20` |
| C5 | `=SUM(A1:A2)` |
| B10 | `=C5+1.1` |
| B11 | `=C5*2` |

```bash
sheet-call-tree myfile.xlsx
```

Output:

```yaml
Sheet1!B10:
  ADD:
  - '@Sheet1!C5'
  - 1.1
Sheet1!B11:
  MUL:
  - '@Sheet1!C5'
  - 2
Sheet1!C5:
  SUM:
  - RANGE:
    - 10
    - 20
```

Each formula cell becomes a top-level YAML key. Its value is the parsed AST of the
formula, with constant cell references resolved to scalars and formula cell references
shown as `@Sheet!Cell` strings (in the default `ref` mode).

## CLI flag overview

| Flag | Default | Description |
|------|---------|-------------|
| `INPUT` | *(required)* | Path to the `.xlsx` file |
| `--filter CELL` | — | Output only the named cell, e.g. `Sheet1!B10` |
| `--output FILE` | stdout | Write YAML to FILE instead of stdout |
| `--no-cycle-check` | off | Skip circular reference detection |
| `--ref-mode MODE` | `ref` | How to render formula-cell refs (see below) |

Full reference: [user_manuals/cli-reference.md](user_manuals/cli-reference.md)

## Output modes (`--ref-mode`)

| Mode | Formula-cell refs render as |
|------|-----------------------------|
| `ref` (default) | `@Sheet1!C5` — cross-reference string |
| `ast` | `'@Sheet1!C5': {SUM: ...}` — key with expanded sub-tree |
| `value` | cached scalar from Excel (`null` if not cached) |
| `inline` | fully expanded expression string, e.g. `ADD(SUM(RANGE(10, 20)), 1.1)` |

Full details with examples: [user_manuals/output-formats.md](user_manuals/output-formats.md)

## Python API

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells = extract_formula_cells("myfile.xlsx")
print(to_yaml(cells))                        # ref mode (default)
print(to_yaml(cells, ref_mode="inline"))     # inline mode
```

Full API reference: [user_manuals/python-api.md](user_manuals/python-api.md)

## Documentation

- [Installation](user_manuals/installation.md)
- [Quickstart walkthrough](user_manuals/quickstart.md)
- [CLI reference](user_manuals/cli-reference.md)
- [Output formats](user_manuals/output-formats.md)
- [Python API](user_manuals/python-api.md)
