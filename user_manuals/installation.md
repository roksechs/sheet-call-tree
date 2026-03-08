# Installation

## Requirements

- Python ≥ 3.11
- [`uv`](https://github.com/astral-sh/uv) (recommended) or any PEP 517-compatible installer

## Install from PyPI

```bash
uv pip install sheet-call-tree
```

## Install from source (editable dev install)

Clone the repository and install in editable mode with dev dependencies:

```bash
git clone https://github.com/yourorg/sheet-call-tree.git
cd sheet-call-tree
uv venv --clear
uv pip install -e ".[dev]"
```

The `[dev]` extra adds `pytest` and `pytest-cov` for running the test suite.

## Verify the installation

```bash
sheet-call-tree --help
```

Expected output:

```
usage: sheet-call-tree [-h] [--filter CELL] [--output FILE] [--no-cycle-check]
                       [--ref-mode {ref,ast,value,inline}]
                       input

Visualize Excel formula dependencies as YAML AST.

positional arguments:
  input                 Path to the .xlsx file

options:
  -h, --help            show this help message and exit
  --filter CELL         Output only the specified cell (e.g. 'Sheet1!B10')
  --output FILE         Write YAML to FILE instead of stdout
  --no-cycle-check      Skip circular reference detection
  --ref-mode {ref,ast,value,inline}
                        How to render formula-cell references in YAML output.
```

## Running tests

```bash
.venv/bin/pytest
```
