# Output Formats

`sheet-call-tree` supports four rendering modes, selected with `--ref-mode`. The mode
controls how **formula-cell references** appear in the output. Constant-cell references
(cells containing plain values, not formulas) always resolve to their scalar values
regardless of mode.

All examples use the same workbook:

| Cell | Value |
|------|-------|
| A1 | `10` |
| A2 | `20` |
| C5 | `=SUM(A1:A2)` |
| B10 | `=C5+1.1` |
| B11 | `=C5*2` |

---

## `ref` (default)

**Use when:** you want a compact dependency map where each formula-cell reference is a
navigable cross-reference string. Good for human review and diff-friendly diffs.

Formula-cell refs render as `'@SheetName!CellRef'` strings. You can follow the
reference to another top-level key in the same YAML document. Constant-cell refs
(A1=10, A2=20) are resolved inline to their scalar values.

```bash
sheet-call-tree example.xlsx
```

```yaml
book:
  name: example.xlsx
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
          - 10
          - 20
```

`B10` calls `ADD` with two arguments: a reference to the formula cell `Sheet1!C5`
and the constant `1.1`. `C5` appears as its own `cells` entry with its full AST.

---

## `ast`

**Use when:** you want the full dependency tree expanded in place, without needing to
follow cross-references. Useful for detailed analysis or feeding into downstream tools.

Formula-cell refs render as a single-key mapping `{'@SheetName!CellRef': <sub-tree>}`,
embedding the referenced cell's AST inline. This means the same sub-tree may appear
multiple times if several cells depend on the same formula cell.

```bash
sheet-call-tree example.xlsx --ref-mode ast
```

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: B10
      formula:
        ADD:
        - '@Sheet1!C5':
            SUM:
            - RANGE:
              - 10
              - 20
        - 1.1
    - cell: B11
      formula:
        MUL:
        - '@Sheet1!C5':
            SUM:
            - RANGE:
              - 10
              - 20
        - 2
    - cell: C5
      formula:
        SUM:
        - RANGE:
          - 10
          - 20
```

The `@Sheet1!C5` key carries its full `SUM(RANGE(10, 20))` sub-tree as its value,
visible directly inside `B10` and `B11` without jumping to another entry.

---

## `value`

**Use when:** you want to see the computed result of each formula-cell reference as
cached by Excel, rather than its AST structure. Useful for spot-checking results.

Formula-cell refs render as the cached scalar value stored in the workbook's
`data_only` representation. If the workbook was created programmatically (e.g. with
openpyxl) and never opened in Excel, computed values are not cached and render as
`null`.

```bash
sheet-call-tree example.xlsx --ref-mode value
```

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: B10
      formula:
        ADD:
        - null
        - 1.1
    - cell: B11
      formula:
        MUL:
        - null
        - 2
    - cell: C5
      formula:
        SUM:
        - RANGE:
          - 10
          - 20
```

The `null` values for `Sheet1!C5` refs appear because the example workbook was created
programmatically and has no cached formula results. Workbooks saved by Excel will have
real scalar values here.

---

## `inline`

**Use when:** you want the most compact human-readable representation, or a single
expression string per cell for pasting into comments, tickets, or logs.

Each formula cell's entire dependency tree is rendered as a single
`FUNC(arg1, arg2, …)` expression string. Referenced formula cells are recursively
expanded. The result is a flat YAML mapping of cell ref → expression string.

```bash
sheet-call-tree example.xlsx --ref-mode inline
```

```yaml
book:
  name: example.xlsx
  sheets:
  - name: Sheet1
    cells:
    - cell: B10
      formula: ADD(SUM(RANGE(10, 20)), 1.1)
    - cell: B11
      formula: MUL(SUM(RANGE(10, 20)), 2)
    - cell: C5
      formula: SUM(RANGE(10, 20))
```

Note that `C5` is fully inlined into `B10` and `B11` —
`ADD(SUM(RANGE(10, 20)), 1.1)` — rather than referenced by name.

---

## Constant-cell refs in all modes

A cell is a *constant cell* if its value is a plain scalar (number, string, boolean),
not a formula. In all four modes, constant-cell refs resolve directly to their scalar
value. In the examples above, `A1=10` and `A2=20` always appear as `10` and `20` in
the `RANGE` node, regardless of `--ref-mode`.
