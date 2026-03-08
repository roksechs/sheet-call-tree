# Output Formats

`sheet-call-tree` uses two parameters to control output:

- **`--depth N`** — controls how deeply formula-cell references are expanded. `0` (default) shows references only; `inf` shows the full expanded AST.
- **`--format {tree,inline}`** — controls the output structure. `tree` (default) produces nested YAML; `inline` renders each cell as a single expression string.

Constant-cell references (cells containing plain values, not formulas) always resolve to their scalar values regardless of depth or format.

All examples use the same workbook:

| Cell | Value |
|------|-------|
| A1 | `10` |
| A2 | `20` |
| C5 | `=SUM(A1:A2)` |
| B10 | `=C5+1.1` |
| B11 | `=C5*2` |

---

## Depth 0 — refs only (default)

**Use when:** you want a compact dependency map where each formula-cell reference is a
navigable cross-reference string. Good for human review and diff-friendly output.

Formula-cell refs render as `Sheet1!C5` strings. Range references render as
`Sheet1!A1:A2` strings.

```bash
sheet-call-tree example.xlsx
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

`B10` calls `ADD` with two arguments: a reference to the formula cell `Sheet1!C5`
and the constant `1.1`. `C5` appears as its own `cells` entry with its full AST.

---

## Depth inf — full expansion

**Use when:** you want the full dependency tree expanded in place, without needing to
follow cross-references. Useful for detailed analysis or feeding into downstream tools.

Formula-cell refs render as `{cell: Sheet1!C5, expression: ...}` mappings,
embedding the referenced cell's AST inline. Range values are flattened into the
parent function's inputs.

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

The `Sheet1!C5` entry carries its full `SUM(...)` sub-tree as its `expression`,
visible directly inside `B10` and `B11` without jumping to another entry.
Range values (`10`, `20`) are flattened into the `inputs` list.

---

## Inline format

**Use when:** you want the most compact human-readable representation, or a single
expression string per cell for pasting into comments, tickets, or logs.

Each formula cell's entire dependency tree is rendered as a single
`FUNC(arg1, arg2, …)` expression string. The `--depth` parameter controls whether
references are expanded or kept as reference strings.

### Inline at depth 0 (default)

```bash
sheet-call-tree example.xlsx --format inline
```

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: C5
      expression: SUM(Sheet1!A1:A2)
    - cell: B10
      expression: ADD(Sheet1!C5, 1.1)
    - cell: B11
      expression: MUL(Sheet1!C5, 2)
```

### Inline at depth inf (fully expanded)

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

Note that `C5` is fully inlined into `B10` and `B11` —
`ADD(SUM(10, 20), 1.1)` — rather than referenced by name.

---

## Semantic labels

When a workbook has header rows or columns, `sheet-call-tree` automatically detects
them using a trained classifier and emits `labels` for each formula cell:

```yaml
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

Each label list contains up to 5 candidates, nearest first, deduplicated. `row` labels
come from scanning left in the same row; `column` labels from scanning up in the same
column.

---

## Constant-cell refs in all modes

A cell is a *constant cell* if its value is a plain scalar (number, string, boolean),
not a formula. In all modes, constant-cell refs resolve directly to their scalar
value. In the examples above, `A1=10` and `A2=20` are flattened into the `inputs`
list at depth > 0, or omitted at depth 0.

---

## Deprecated `--ref-mode`

The legacy `--ref-mode` flag is still accepted for backwards compatibility:

| Legacy value | Equivalent |
|-------------|------------|
| `ref` | `--depth 0` |
| `ast` | `--depth inf` |
| `inline` | `--format inline --depth inf` |

The `value` mode has been removed.
