# CLI Reference

## Synopsis

```
sheet-call-tree INPUT [OPTIONS]
```

## Positional arguments

| Argument | Description |
|----------|-------------|
| `INPUT` | Path to the `.xlsx` file to analyse. Required. |

## Options

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

### `--ref-mode {ref,ast,value,inline}`

Controls how formula-cell references are rendered in the YAML output.
Default: `ref`.

| Value | Description |
|-------|-------------|
| `ref` | Cross-reference string prefixed with `@`, e.g. `'@Sheet1!C5'` |
| `ast` | The `@`-prefixed name used as a key whose value is the expanded sub-tree |
| `value` | Cached computed scalar from Excel's `data_only` workbook; `null` if not cached |
| `inline` | Each cell emitted as a single `FUNC(arg1, arg2, …)` expression string |

Constant-cell references (cells containing plain values, not formulas) always resolve
to their scalar values in all modes.

See [output-formats.md](output-formats.md) for concrete YAML examples of each mode.

## Exit codes

| Code | Meaning |
|------|---------|
| `0` | Success |
| `1` | Cell not found (when `--filter` is used) or other error |
