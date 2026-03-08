# Python API

`sheet-call-tree` can be used as a library. The public API consists of three functions
importable from the top-level package.

## Public API

```python
from sheet_call_tree import extract_formula_cells, to_json, to_yaml
```

---

### `extract_formula_cells(path) -> tuple[dict, dict, dict]`

Load an `.xlsx`/`.xlsm` file and return a 3-tuple of formula cells, data values, and label map.

```python
cells, data_values, label_map = extract_formula_cells("myfile.xlsx")
# cells == {
#   'Sheet1!B10': FunctionNode(name='ADD', args=[RefNode(...), 1.1]),
#   'Sheet1!C5':  FunctionNode(name='SUM', args=[RangeNode(...)]),
# }
# data_values == {'Sheet1!A1': 10, 'Sheet1!A2': 20, ...}
# label_map == {'Sheet1!D2': {'row': ['Alice'], 'column': ['Total']}, ...}
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `path` | `str \| Path` | Path to the `.xlsx`/`.xlsm` file |

**Returns:** `tuple[dict[str, FunctionNode], dict[str, object], dict[str, dict]]`

| Element | Type | Description |
|---------|------|-------------|
| `formula_cells` | `dict[str, FunctionNode]` | Keys are `'SheetName!CellRef'` strings; values are parsed AST roots |
| `data_values` | `dict[str, object]` | Keys are cell refs; values are cached scalar values from `data_only` mode |
| `label_map` | `dict[str, dict[str, object]]` | Keys are cell refs; values have `'row'` and `'column'` keys with lists of label strings (up to 5 per direction, nearest first) |

**Behaviour:**
- The workbook is loaded twice: once for formula text (`data_only=False`) and once for
  cached computed values (`data_only=True`).
- Only formula cells (those whose value starts with `=`) appear in `formula_cells`.
- Constant cells appear only as `RefNode` values inside formula ASTs.
- `label_map` is built by a trained RandomForest classifier that identifies header vs data
  cells, then scans for the nearest headers above and to the left of each formula cell.

---

### `to_yaml(cells, *, depth=None, fmt="tree", book_name="", data_values=None, label_map=None, stream=None) -> str | None`

Serialise a `formula_cells` dict to YAML.

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells, data_values, label_map = extract_formula_cells("myfile.xlsx")

# Return YAML string (depth 0, tree format)
yaml_str = to_yaml(cells, data_values=data_values, label_map=label_map)

# Full expansion
yaml_str = to_yaml(cells, depth=float('inf'), data_values=data_values, label_map=label_map)

# Inline format
with open("deps.yaml", "w") as fh:
    to_yaml(cells, fmt="inline", data_values=data_values, label_map=label_map, stream=fh)
```

**Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `cells` | `dict[str, FunctionNode]` | *(required)* | Output of `extract_formula_cells` |
| `depth` | `int \| float \| None` | `None` (→ 0) | Expansion depth: 0 = refs only, `inf` = full expansion |
| `fmt` | `str` | `"tree"` | Output format: `"tree"` or `"inline"` |
| `book_name` | `str` | `""` | Workbook filename for the top-level `book.name` field |
| `data_values` | `dict[str, object] \| None` | `None` | Cached scalar values for `outputs` fields |
| `label_map` | `dict[str, dict] \| None` | `None` | Label map for `labels` blocks |
| `stream` | writable file object | `None` | If provided, YAML is written to the stream |
| `ref_mode` | `str \| None` | `None` | *(deprecated)* Legacy rendering mode; use `depth`/`fmt` instead |

**Returns:** YAML string when `stream` is `None`; `None` when `stream` is provided.

---

### `to_json(cells, *, stream=None, **kw) -> str | None`

Serialise a `formula_cells` dict to JSON. Accepts the same keyword arguments as `to_yaml`
(`depth`, `fmt`, `ref_mode`, `book_name`, `data_values`, `label_map`).

```python
from sheet_call_tree import extract_formula_cells, to_json

cells, data_values, label_map = extract_formula_cells("myfile.xlsx")

# Return JSON string
json_str = to_json(cells, data_values=data_values, label_map=label_map)

# Write to file
with open("deps.json", "w") as fh:
    to_json(cells, data_values=data_values, label_map=label_map, stream=fh)
```

**Returns:** JSON string when `stream` is `None`; `None` when `stream` is provided.

Special float handling: `NaN` → `null`, `Infinity` → `"Infinity"`, `-Infinity` → `"-Infinity"`.

---

**Output structure (tree format, shared by `to_yaml` and `to_json`):**

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
      outputs: 300
      expression:
        type: SUM
        inputs:
        - Sheet1!B2:C2
```

---

## Data model (`sheet_call_tree.models`)

The AST is built from five dataclasses and a type alias:

```python
from sheet_call_tree.models import FunctionNode, RefNode, RangeNode, TableRefNode, NamedRefNode, Node
```

---

### `FunctionNode`

Represents a parsed function call or operator.

```python
@dataclass
class FunctionNode:
    name: str    # "SUM", "ADD", "MUL", "IF", "NEG", …
    args: list   # list[Node]
```

**Examples:**
- `=SUM(A1:A2)` → `FunctionNode(name='SUM', args=[RangeNode(...)])`
- `=C5+1.1`   → `FunctionNode(name='ADD', args=[RefNode('Sheet1!C5', ...), 1.1])`
- `=C5*2`     → `FunctionNode(name='MUL', args=[RefNode('Sheet1!C5', ...), 2])`

Operators are normalised to their function names: `+` → `ADD`, `*` → `MUL`, `-` →
`SUB` or unary `NEG`, `/` → `DIV`.

In YAML tree output, renders as:
```yaml
type: SUM
inputs:
- Sheet1!A1:A2
```

---

### `RefNode`

Represents a single-cell reference.

```python
@dataclass
class RefNode:
    ref: str                              # "Sheet1!A1"
    formula: FunctionNode | None = None   # formula cell AST (None for constant/unknown)
    resolved_value: object = None         # scalar value (constant cell value or cached compute)
```

**Fields:**

| Field | Description |
|-------|-------------|
| `ref` | Fully-qualified cell reference string, e.g. `"Sheet1!C5"` |
| `formula` | The cell's parsed AST: a `FunctionNode` for formula cells; `None` for constant cells or if the cell was not found in the workbook |
| `resolved_value` | The scalar value: for constant cells this is the cell value (`int`, `float`, `str`, `bool`); for formula cells this is the `data_only` cached result (may be `None` for programmatic workbooks) |

In YAML at depth 0: `Sheet1!C5`
At depth > 0: `{cell: Sheet1!C5, expression: ...}`

---

### `RangeNode`

Represents a cell range reference (e.g. `A1:A2`).

```python
@dataclass
class RangeNode:
    start: str                            # "Sheet1!A1"
    end: str                              # "Sheet1!A9"
    values: list[object] | None = None    # all cell values in range (populated by reader)
```

**Fields:**

| Field | Description |
|-------|-------------|
| `start` | Fully-qualified start cell reference string, e.g. `"Sheet1!A1"` |
| `end` | Fully-qualified end cell reference string, e.g. `"Sheet1!A9"` |
| `values` | List of all cell values in the range, or `None` if not populated. At depth > 0 these values are flattened into the parent's `inputs` list. |

In YAML at depth 0: `Sheet1!A1:A2`
At depth > 0: values are flattened directly into parent `inputs`

---

### `TableRefNode`

Represents a structured table reference (e.g. `Table1[Amount]` or `Table1[@Amount]`).

```python
@dataclass
class TableRefNode:
    table_name: str            # "Table1"
    column: str | None         # "Amount"; None for whole-table reference
    this_row: bool             # True when @-prefixed (Table1[@Amount])
    resolved_range: str | None = None   # "Sheet1!D2:D100"
    cached_value: object = None
```

**Fields:**

| Field | Description |
|-------|-------------|
| `table_name` | The table name, e.g. `"Table1"` |
| `column` | The column specifier, or `None` for a whole-table reference |
| `this_row` | `True` when the reference uses `@` for the current row (e.g. `Table1[@Amount]`) |
| `resolved_range` | The resolved cell range, e.g. `"Sheet1!D2:D100"` |
| `cached_value` | Cached computed value, if available |

In YAML tree output, renders as:
```yaml
type: TABLE_REF
name: Table1
column: Amount
this_row: true
range: Sheet1!D2:D100
```

In inline output, renders as: `TABLE_REF(Table1[@Amount])`

---

### `NamedRefNode`

Represents a named range or defined name reference (e.g. `SalesTotal`).

```python
@dataclass
class NamedRefNode:
    name: str                  # "SalesTotal"
    resolved_range: str | None = None   # "Sheet1!$B$10"
    formula: FunctionNode | None = None   # formula cell AST after resolution
    resolved_value: object = None         # scalar value
```

**Fields:**

| Field | Description |
|-------|-------------|
| `name` | The defined name, e.g. `"SalesTotal"` |
| `resolved_range` | The resolved cell range or reference |
| `formula` | The formula cell AST, if the named reference resolves to a formula cell |
| `resolved_value` | The scalar value, if available |

In YAML tree output at depth 0:
```yaml
type: NAMED_REF
name: SalesTotal
range: Sheet1!$B$10
```

At depth > 0, if the named reference resolves to a formula, the AST is expanded:
```yaml
named_ref: SalesTotal
expression:
  type: SUM
  inputs:
  - ...
```

In inline output, renders as: `NAMED_REF(SalesTotal)`

---

### `Node` type alias

```python
Node = Union[FunctionNode, RefNode, RangeNode, TableRefNode, NamedRefNode, int, float, bool, str]
```

Every element in `FunctionNode.args` is a `Node`.
Plain scalars (`int`, `float`, `bool`, `str`) appear as literal argument values.

---

## Advanced: lower-level modules

The following internal functions are not part of the stable public API but may be
useful for advanced use cases.

| Symbol | Module | Description |
|--------|--------|-------------|
| `parse_formula(formula, default_sheet)` | `sheet_call_tree.formula_parser` | Parse a single formula string into a `FunctionNode` AST |
| `extract_formula_cells_from_workbook(wb)` | `sheet_call_tree.reader` | Extract formula cells from an already-loaded `openpyxl.Workbook` |
| `build_dependency_graph(formula_cells)` | `sheet_call_tree.dependency_graph` | Return a `dict[str, set[str]]` of cell dependencies |
| `detect_cycles(graph)` | `sheet_call_tree.dependency_graph` | Raise `CircularReferenceError` if a cycle is found |
| `CircularReferenceError` | `sheet_call_tree.dependency_graph` | Exception class; `.cycle` attribute holds the cycle as a list of cell refs |
| `build_label_map(...)` | `sheet_call_tree.labeler` | Build label map using trained classifier |

---

## Example: walk the AST

```python
from sheet_call_tree import extract_formula_cells
from sheet_call_tree.models import FunctionNode, RefNode, RangeNode, TableRefNode, NamedRefNode


def walk(node, depth=0):
    indent = "  " * depth
    if isinstance(node, FunctionNode):
        print(f"{indent}{node.name}(")
        for arg in node.args:
            walk(arg, depth + 1)
        print(f"{indent})")
    elif isinstance(node, RangeNode):
        print(f"{indent}RANGE({node.start} .. {node.end})")
    elif isinstance(node, RefNode):
        print(f"{indent}@{node.ref} = {node.formula!r}")
    elif isinstance(node, TableRefNode):
        print(f"{indent}TABLE_REF({node.table_name}[{node.column}])")
    elif isinstance(node, NamedRefNode):
        print(f"{indent}NAMED_REF({node.name})")
    else:
        print(f"{indent}{node!r}")


cells, data_values, label_map = extract_formula_cells("myfile.xlsx")
for cell_ref, ast in cells.items():
    print(f"=== {cell_ref} ===")
    walk(ast)
```

Example output for `Sheet1!C5 = SUM(A1:A2)`:

```
=== Sheet1!C5 ===
SUM(
  RANGE(Sheet1!A1 .. Sheet1!A2)
)
```
