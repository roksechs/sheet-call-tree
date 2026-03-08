# Python API

`sheet-call-tree` can be used as a library. The public API consists of two functions
importable from the top-level package.

## Public API

```python
from sheet_call_tree import extract_formula_cells, to_yaml
```

---

### `extract_formula_cells(path) -> dict[str, FunctionNode]`

Load an `.xlsx` file and return a mapping of cell references to their AST roots.

```python
cells = extract_formula_cells("myfile.xlsx")
# cells == {
#   'Sheet1!B10': FunctionNode(name='ADD', args=[RefNode(...), 1.1]),
#   'Sheet1!B11': FunctionNode(name='MUL', args=[RefNode(...), 2]),
#   'Sheet1!C5':  FunctionNode(name='SUM', args=[RangeNode(...)]),
# }
```

**Parameters:**

| Parameter | Type | Description |
|-----------|------|-------------|
| `path` | `str \| Path` | Path to the `.xlsx` file |

**Returns:** `dict[str, FunctionNode]` — keys are `'SheetName!CellRef'` strings;
values are the parsed AST root nodes for each formula cell.

**Behaviour:**
- The workbook is loaded twice: once for formula text (`data_only=False`) and once for
  cached computed values (`data_only=True`).
- Only formula cells (those whose value starts with `=`) appear as top-level keys.
- Constant cells appear only as `RefNode` values inside formula ASTs.
- `RefNode.value` is populated with the referenced cell's scalar (for constant cells)
  or `FunctionNode` (for formula cells). `RefNode.cached_value` holds the `data_only`
  scalar (may be `None` for programmatic workbooks).

---

### `to_yaml(cells, ref_mode="ref", stream=None) -> str | None`

Serialise a `formula_cells` dict to YAML.

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells = extract_formula_cells("myfile.xlsx")

# Return YAML string
yaml_str = to_yaml(cells)

# Write to a file
with open("deps.yaml", "w") as fh:
    to_yaml(cells, ref_mode="inline", stream=fh)
```

**Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `cells` | `dict[str, FunctionNode]` | *(required)* | Output of `extract_formula_cells` |
| `ref_mode` | `str` | `"ref"` | Rendering mode: `"ref"`, `"ast"`, `"value"`, or `"inline"` |
| `stream` | writable file object | `None` | If provided, YAML is written to the stream |

**Returns:** YAML string when `stream` is `None`; `None` when `stream` is provided.

**Raises:** `ValueError` if `ref_mode` is not one of the four valid values.

---

## Data model (`sheet_call_tree.models`)

The AST is built from three dataclasses and a type alias:

```python
from sheet_call_tree.models import FunctionNode, RefNode, RangeNode, Node
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

---

### `RefNode`

Represents a single-cell reference.

```python
@dataclass
class RefNode:
    ref: str                    # "Sheet1!A1"  (no @ sigil)
    value: object = None        # scalar for constant cells; FunctionNode for formula cells; None if unknown
    cached_value: object = None # data_only computed value (for --ref-mode value); None if unavailable
```

**Fields:**

| Field | Description |
|-------|-------------|
| `ref` | Fully-qualified cell reference string without the `@` prefix, e.g. `"Sheet1!C5"` |
| `value` | The cell's content: a scalar (`int`, `float`, `str`, `bool`) for constant cells; a `FunctionNode` for formula cells; `None` if the cell was not found in the workbook |
| `cached_value` | The `data_only` computed value. For workbooks saved by Excel this is the last computed result. For programmatic workbooks (openpyxl-created, never opened in Excel) this is `None`. |

---

### `RangeNode`

Represents a cell range reference (e.g. `A1:A2`).

```python
@dataclass
class RangeNode:
    start: RefNode   # first cell of the range
    end: RefNode     # last cell of the range
```

`start` and `end` are `RefNode` instances with their `value` / `cached_value` fields
populated from the workbook.

---

### `Node` type alias

```python
Node = Union[FunctionNode, RefNode, RangeNode, int, float, bool, str]
```

Every element in `FunctionNode.args` and every `RangeNode.start` / `end` is a `Node`.
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

---

## Example: walk the AST

```python
from sheet_call_tree import extract_formula_cells
from sheet_call_tree.models import FunctionNode, RefNode, RangeNode


def walk(node, depth=0):
    indent = "  " * depth
    if isinstance(node, FunctionNode):
        print(f"{indent}{node.name}(")
        for arg in node.args:
            walk(arg, depth + 1)
        print(f"{indent})")
    elif isinstance(node, RangeNode):
        print(f"{indent}RANGE({node.start.ref} .. {node.end.ref})")
    elif isinstance(node, RefNode):
        print(f"{indent}@{node.ref} = {node.value!r}")
    else:
        print(f"{indent}{node!r}")


cells = extract_formula_cells("myfile.xlsx")
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
