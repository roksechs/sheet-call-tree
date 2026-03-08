# sheet-call-tree

Visualize Excel formula dependencies as a YAML AST call tree.

[日本語版 README](README.ja.md)

## What it does

Excel formulas can reference other cells that themselves contain formulas, creating
dependency chains that are hard to follow in a spreadsheet UI. `sheet-call-tree`
loads an `.xlsx` file, parses every formula cell into an abstract syntax tree (AST),
resolves cross-cell references, and emits the complete dependency tree as YAML — one
top-level key per formula cell.

It also **detects semantic labels** for each formula cell by classifying surrounding
cells as headers or data using a trained RandomForest classifier, then scanning for
the nearest header cells in each direction.

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
| A1 | `Name` |
| B1 | `Q1` |
| C1 | `Q2` |
| D1 | `Total` |
| A2 | `Alice` |
| B2 | `100` |
| C2 | `200` |
| D2 | `=SUM(B2:C2)` |

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
    - cell: D2
      labels:
        row:
        - Alice
        column:
        - Total
      expression:
        type: SUM
        inputs:
        - Sheet1!B2:C2
```

The output is grouped as `book → sheets → cells`. Each cell entry has:
- `cell` — the cell coordinate
- `labels` — detected semantic labels (`row` and `column` headers, nearest first)
- `expression` — the parsed AST with `type` (function name) and `inputs` (arguments)

Range references appear as `Sheet1!A1:A2` strings. Formula cell references appear as
`Sheet1!C5` strings at the default depth (0).

## CLI flag overview

| Flag | Default | Description |
|------|---------|-------------|
| `INPUT` | *(required)* | Path to the `.xlsx` file |
| `--filter CELL` | — | Output only the named cell, e.g. `Sheet1!B10` |
| `--output FILE` | stdout | Write YAML to FILE instead of stdout |
| `--no-cycle-check` | off | Skip circular reference detection |
| `--depth N` | `0` | Expansion depth: 0 = refs only, inf = full expansion |
| `--format FORMAT` | `tree` | Output format: `tree` or `inline` |
| `--roots-only` | off | Output only root cells (not referenced by other formulas) |

Full reference: [user_manuals/cli-reference.md](user_manuals/cli-reference.md)

## Output modes (`--depth` / `--format`)

| Setting | Formula-cell refs render as |
|---------|-----------------------------|
| `--depth 0` (default) | `Sheet1!C5` — cross-reference string |
| `--depth inf` | `{cell: Sheet1!C5, expression: {type: SUM, inputs: [...]}}` — expanded sub-tree |
| `--format inline` | `SUM(Sheet1!B2:C2)` — fully expanded expression string |

The deprecated `--ref-mode` flag is still accepted (`ref`→depth 0, `ast`→depth inf, `inline`→format inline).

Full details with examples: [user_manuals/output-formats.md](user_manuals/output-formats.md)

## Python API

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells, data_values, label_map = extract_formula_cells("myfile.xlsx")
print(to_yaml(cells, data_values=data_values, label_map=label_map))
```

Full API reference: [user_manuals/python-api.md](user_manuals/python-api.md)

## Documentation

- [Installation](user_manuals/installation.md)
- [Quickstart walkthrough](user_manuals/quickstart.md)
- [CLI reference](user_manuals/cli-reference.md)
- [Output formats](user_manuals/output-formats.md)
- [Python API](user_manuals/python-api.md)
