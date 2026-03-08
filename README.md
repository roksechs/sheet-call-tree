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
uv tool install git+https://github.com/roksechs/sheet-call-tree.git
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
book:
  name: myfile.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: B10
      formula:
        ADD:
        - '@Sheet1!C5'
        - 1.1
    - cell: B11
      formula:
        MUL:
        - '@Sheet1!C5'
        - 2
    - cell: C5
      formula:
        SUM:
        - RANGE:
            ref: '@Sheet1!A1:A2'
```

The output is grouped as `book → sheets → cells`. Each cell entry has a `cell` coordinate
and a `formula` with the parsed AST. Constant cell references resolve to scalars; formula
cell references appear as `@Sheet!Cell` strings at the default depth (0). Range references show as `RANGE: {ref: '@Sheet1!A1:A2'}`.

## CLI flag overview

| Flag | Default | Description |
|------|---------|-------------|
| `INPUT` | *(required)* | Path to the `.xlsx` file |
| `--filter CELL` | — | Output only the named cell, e.g. `Sheet1!B10` |
| `--output FILE` | stdout | Write YAML to FILE instead of stdout |
| `--no-cycle-check` | off | Skip circular reference detection |
| `--depth N` | `0` | Expansion depth: 0 = refs only, inf = full expansion |
| `--format FORMAT` | `tree` | Output format: `tree` or `inline` |

Full reference: [user_manuals/cli-reference.md](user_manuals/cli-reference.md)

## Output modes (`--depth` / `--format`)

| Setting | Formula-cell refs render as |
|---------|-----------------------------|
| `--depth 0` (default) | `@Sheet1!C5` — cross-reference string |
| `--depth inf` | `'@Sheet1!C5': {SUM: ...}` — key with expanded sub-tree |
| `--format inline` | `SUM(RANGE(@Sheet1!A1:A2))` — fully expanded expression string |

The deprecated `--ref-mode` flag is still accepted (`ref`→depth 0, `ast`→depth inf, `inline`→format inline).

Full details with examples: [user_manuals/output-formats.md](user_manuals/output-formats.md)

## Python API

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells = extract_formula_cells("myfile.xlsx")
print(to_yaml(cells))                        # ref mode (default)
print(to_yaml(cells, fmt="inline"))          # inline format
```

Full API reference: [user_manuals/python-api.md](user_manuals/python-api.md)

## Documentation

- [Installation](user_manuals/installation.md)
- [Quickstart walkthrough](user_manuals/quickstart.md)
- [CLI reference](user_manuals/cli-reference.md)
- [Output formats](user_manuals/output-formats.md)
- [Python API](user_manuals/python-api.md)
