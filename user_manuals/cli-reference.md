# CLI Reference

## Synopsis

```
sheet-call-tree INPUT [OPTIONS]
```

## Positional arguments

| Argument | Description |
|----------|-------------|
| `INPUT` | Path to the `.xlsx`/`.xlsm` file to analyse. Required. |

## Options

### `--sheet SHEETNAME`

Output only cells in the specified sheet.

If the sheet does not exist or has no formula cells, `sheet-call-tree` prints an error
to stderr and exits with code `1`.

```bash
sheet-call-tree myfile.xlsx --sheet Sheet1
```

### `--filter CELL`

Output only the named cell. `CELL` must be in `SheetName!ColRow` format,
e.g. `Sheet1!B10`.

If the cell does not exist or has no formula, `sheet-call-tree` prints an error
to stderr and exits with code `1`.

```bash
sheet-call-tree myfile.xlsx --filter Sheet1!B10
```

### `--output FILE`

Write YAML to `FILE` instead of stdout. The file is created (or overwritten) in
UTF-8 encoding. When this flag is used, nothing is printed to stdout.

```bash
sheet-call-tree myfile.xlsx --output deps.yaml
```

### `--no-cycle-check`

By default, `sheet-call-tree` builds a dependency graph and raises an error if it
detects a circular reference (e.g. A1 depends on B1 which depends on A1). Use
`--no-cycle-check` to skip this check and produce output even if cycles exist.

```bash
sheet-call-tree myfile.xlsx --no-cycle-check
```

### `--depth N`

Controls the expansion depth for formula-cell references. Default: `0`.

| Value | Description |
|-------|-------------|
| `0` | Refs only — formula-cell references render as `Sheet1!C5` cross-reference strings; range references render as `Sheet1!A1:A2` without values |
| `inf` | Full expansion — formula-cell references render as `{cell: Sheet1!C5, expression: ...}` with the AST expanded in place; range values are flattened into parent `inputs` |

Intermediate integer values (1, 2, …) expand to that many levels of nesting.

See [output-formats.md](output-formats.md) for concrete YAML examples.

### `--format {tree,inline,json}`

Controls the output format. Default: `tree`.

| Value | Description |
|-------|-------------|
| `tree` | Nested YAML structure with `type` and `inputs` keys (default) |
| `inline` | Each cell emitted as a single `FUNC(arg1, arg2, …)` expression string |
| `json` | Same structure as `tree`, output as JSON with 2-space indentation |

Constant-cell references (cells containing plain values, not formulas) always resolve
to their scalar values in all modes.

### `--roots-only`

Output only root cells — cells that are not referenced by any other formula cell.
This is useful for finding the "entry points" of a workbook's calculation graph.

```bash
sheet-call-tree myfile.xlsx --roots-only
```

### `--ref-mode` *(deprecated)*

The legacy `--ref-mode` flag is still accepted for backwards compatibility but hidden from `--help`.

| Legacy value | Equivalent |
|-------------|------------|
| `ref` | `--depth 0` |
| `ast` | `--depth inf` |
| `inline` | `--format inline --depth inf` |

The `value` mode has been removed.

## Exit codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | Cell not found (`--filter`), sheet not found (`--sheet`), circular reference detected, or other error |
