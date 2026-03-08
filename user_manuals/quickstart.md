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

Output (default: depth 0, tree format):

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: C5
      expression:
        type: SUM
        inputs:
        - Sheet1!A1:A2
    - cell: B10
      expression:
        type: ADD
        inputs:
        - Sheet1!C5
        - 1.1
    - cell: B11
      expression:
        type: MUL
        inputs:
        - Sheet1!C5
        - 2
```

**Reading the output:**

- The top-level `book` key groups cells by workbook and sheet.
- Each entry in `cells` has a `cell` coordinate (`B10`, `C5`, …) and an `expression` AST.
- `expression` contains `type` (the function/operator name: `ADD`, `MUL`, `SUM`, …) and `inputs` (the arguments as a YAML list).
- `Sheet1!A1:A2` is a range reference. At depth 0, the cell values are not included.
- `Sheet1!C5` is a cross-reference to another formula cell.

## 3. Drill into a single cell

Use `--filter` to focus on one cell:

```bash
sheet-call-tree example.xlsx --filter Sheet1!B10
```

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: B10
      expression:
        type: ADD
        inputs:
        - Sheet1!C5
        - 1.1
```

## 4. Expand the full AST inline (`--depth inf`)

The default depth (0) keeps cross-references compact. Use `--depth inf` to expand each
referenced formula cell's sub-tree in place:

```bash
sheet-call-tree example.xlsx --depth inf
```

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: C5
      expression:
        type: SUM
        inputs:
        - 10
        - 20
    - cell: B10
      expression:
        type: ADD
        inputs:
        - cell: Sheet1!C5
          expression:
            type: SUM
            inputs:
            - 10
            - 20
        - 1.1
    - cell: B11
      expression:
        type: MUL
        inputs:
        - cell: Sheet1!C5
          expression:
            type: SUM
            inputs:
            - 10
            - 20
        - 2
```

The `Sheet1!C5` entry now carries its full sub-tree as `expression`, so you can read the
complete dependency inline without jumping between entries. Range values are flattened
directly into the `inputs` list.

## 5. Compact expression strings (`--format inline`)

`inline` format renders each cell as a single fully-expanded expression string:

```bash
sheet-call-tree example.xlsx --format inline --depth inf
```

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: C5
      expression: SUM(10, 20)
    - cell: B10
      expression: ADD(SUM(10, 20), 1.1)
    - cell: B11
      expression: MUL(SUM(10, 20), 2)
```

This is useful for quick human reading or for pasting into comments/tickets.

## 6. Semantic labels

When your workbook has header rows/columns, labels are automatically detected:

```python
import openpyxl
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sheet1"
ws["A1"] = "Name"; ws["B1"] = "Q1"; ws["C1"] = "Total"
ws["A2"] = "Alice"; ws["B2"] = 100; ws["C2"] = "=SUM(B2:B2)"
wb.save("labeled.xlsx")
```

```bash
sheet-call-tree labeled.xlsx
```

```yaml
book:
  name: labeled.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: C2
      labels:
        row:
        - Alice
        column:
        - Total
      expression:
        type: SUM
        inputs:
        - Sheet1!B2:B2
```

The `labels` block shows the nearest header cells: `row` labels scan left, `column` labels
scan up. Up to 5 candidates per direction, deduplicated.

## 7. Write output to a file

Use `--output` to save the YAML instead of printing to stdout:

```bash
sheet-call-tree example.xlsx --output result.yaml
```

No output is printed; `result.yaml` contains the full YAML.

## Next steps

- [CLI reference](cli-reference.md) — all flags documented
- [Output formats](output-formats.md) — depth-based expansion and format options with examples
- [Python API](python-api.md) — use `sheet-call-tree` as a library
