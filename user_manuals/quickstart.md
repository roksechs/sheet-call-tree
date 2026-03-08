# Quickstart

This walkthrough uses a small example workbook to show the most common `sheet-call-tree`
workflows.

## 1. Create the example workbook

Save the following as `example.xlsx` (or create it with Python):

| Cell | Value |
|------|-------|
| A1 | `10` |
| A2 | `20` |
| C5 | `=SUM(A1:A2)` |
| B10 | `=C5+1.1` |
| B11 | `=C5*2` |

```python
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sheet1"
ws["A1"] = 10
ws["A2"] = 20
ws["C5"] = "=SUM(A1:A2)"
ws["B10"] = "=C5+1.1"
ws["B11"] = "=C5*2"
wb.save("example.xlsx")
```

## 2. Run `sheet-call-tree`

```bash
sheet-call-tree example.xlsx
```

Output (default `ref` mode):

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

**Reading the output:**

- Each top-level key (`Sheet1!B10`, `Sheet1!B11`, `Sheet1!C5`) is a formula cell.
- The value is the parsed AST of that cell's formula.
- `ADD`, `MUL`, `SUM` are the parsed operators/functions; their arguments are YAML lists.
- `RANGE: [10, 20]` represents `A1:A2` with both endpoints resolved to their scalar values.
- `'@Sheet1!C5'` is a cross-reference to another formula cell. The `@` prefix distinguishes
  formula-cell refs from plain strings.
- Output keys are sorted alphabetically (`Sheet1!B10` < `Sheet1!B11` < `Sheet1!C5`).

## 3. Drill into a single cell

Use `--filter` to focus on one cell:

```bash
sheet-call-tree example.xlsx --filter Sheet1!B10
```

```yaml
Sheet1!B10:
  ADD:
  - '@Sheet1!C5'
  - 1.1
```

## 4. Expand the full AST inline (`--ref-mode ast`)

`ref` mode keeps cross-references compact. Use `--ref-mode ast` to expand each
referenced formula cell's sub-tree in place:

```bash
sheet-call-tree example.xlsx --ref-mode ast
```

```yaml
Sheet1!B10:
  ADD:
  - '@Sheet1!C5':
      SUM:
      - RANGE:
        - 10
        - 20
  - 1.1
Sheet1!B11:
  MUL:
  - '@Sheet1!C5':
      SUM:
      - RANGE:
        - 10
        - 20
  - 2
Sheet1!C5:
  SUM:
  - RANGE:
    - 10
    - 20
```

The `@Sheet1!C5` key now carries its full sub-tree as its value, so you can read the
complete dependency inline without jumping between entries.

## 5. Compact expression strings (`--ref-mode inline`)

`inline` mode renders each cell as a single fully-expanded expression string:

```bash
sheet-call-tree example.xlsx --ref-mode inline
```

```yaml
Sheet1!B10: ADD(SUM(RANGE(10, 20)), 1.1)
Sheet1!B11: MUL(SUM(RANGE(10, 20)), 2)
Sheet1!C5: SUM(RANGE(10, 20))
```

This is useful for quick human reading or for pasting into comments/tickets.

## 6. Write output to a file

Use `--output` to save the YAML instead of printing to stdout:

```bash
sheet-call-tree example.xlsx --output result.yaml
```

No output is printed; `result.yaml` contains the full YAML.

## Next steps

- [CLI reference](cli-reference.md) — all flags documented
- [Output formats](output-formats.md) — all four `--ref-mode` values with examples
- [Python API](python-api.md) — use `sheet-call-tree` as a library
