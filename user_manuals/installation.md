# Installation

## Requirements

- Python ≥ 3.11
- [`uv`](https://github.com/astral-sh/uv) (recommended) or any PEP 517-compatible installer

## Install from GitHub

```bash
uv tool install git+https://github.com/roksechs/sheet-call-tree.git
```

## Install from source (editable dev install)

Clone the repository and install in editable mode with dev dependencies:

```bash
git clone https://github.com/roksechs/sheet-call-tree.git
cd sheet-call-tree
uv sync --dev
```

The `--dev` flag includes `pytest` and `pytest-cov` for running the test suite.

## Verify the installation

```bash
sheet-call-tree --help
```

Expected output:

```
usage: sheet-call-tree [-h] [--filter CELL] [--output FILE] [--no-cycle-check]
                       [--depth N] [--format {tree,inline}]
                       input

Excel の数式依存関係を YAML AST として可視化します。

positional arguments:
  input                 .xlsx ファイルのパス

options:
  -h, --help            show this help message and exit
  --filter CELL         指定したセルのみ出力する（例: 'Sheet1!B10'）
  --output FILE         YAML を stdout ではなくファイルに書き出す
  --no-cycle-check      循環参照の検出をスキップする
  --depth N             Expansion depth: 0 = refs only (default), inf = full
                        expansion.
  --format {tree,inline}
                        Output format: tree (default) or inline.
```

## Running tests

```bash
.venv/bin/pytest
```
