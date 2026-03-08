# Python API

`sheet-call-tree` はライブラリとして使用できます。公開 API はトップレベルパッケージからインポートできる 2 つの関数で構成されています。

## 公開 API

```python
from sheet_call_tree import extract_formula_cells, to_yaml
```

---

### `extract_formula_cells(path) -> dict[str, FunctionNode]`

`.xlsx` ファイルを読み込み、セル参照から AST ルートへのマッピングを返します。

```python
cells = extract_formula_cells("myfile.xlsx")
# cells == {
#   'Sheet1!B10': FunctionNode(name='ADD', args=[RefNode(...), 1.1]),
#   'Sheet1!B11': FunctionNode(name='MUL', args=[RefNode(...), 2]),
#   'Sheet1!C5':  FunctionNode(name='SUM', args=[RangeNode(...)]),
# }
```

**パラメーター：**

| パラメーター | 型 | 説明 |
|------------|-----|------|
| `path` | `str \| Path` | `.xlsx` ファイルへのパス |

**戻り値：** `dict[str, FunctionNode]` — キーは `'シート名!セル参照'` 文字列；値は各数式セルの解析済み AST ルートノード。

**動作：**
- ワークブックは 2 回読み込まれます：数式テキスト用（`data_only=False`）とキャッシュ済み計算値用（`data_only=True`）。
- 数式セル（値が `=` で始まるセル）のみがトップレベルキーとして現れます。
- 定数セルは数式 AST 内の `RefNode` 値としてのみ現れます。
- `RefNode.formula` には参照先セルの `FunctionNode`（数式セルの場合）または `None`（定数/不明セルの場合）が格納されます。`RefNode.resolved_value` にはスカラー値（定数セルの値または `data_only` のキャッシュ結果）が格納されます（プログラムで作成されたワークブックでは `None` の場合があります）。

---

### `to_yaml(cells, *, depth=None, fmt="tree", book_name="", stream=None) -> str | None`

`formula_cells` 辞書を YAML にシリアライズします。

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells = extract_formula_cells("myfile.xlsx")

# YAML 文字列を返す（深さ 0、ツリーフォーマット）
yaml_str = to_yaml(cells)

# 完全展開
yaml_str = to_yaml(cells, depth=float('inf'))

# インラインフォーマット
with open("deps.yaml", "w") as fh:
    to_yaml(cells, fmt="inline", stream=fh)
```

**パラメーター：**

| パラメーター | 型 | デフォルト | 説明 |
|------------|-----|----------|------|
| `cells` | `dict[str, FunctionNode]` | *(必須)* | `extract_formula_cells` の出力 |
| `depth` | `int \| float \| None` | `None`（→ 0） | 展開の深さ: 0 = 参照のみ、`inf` = 完全展開 |
| `fmt` | `str` | `"tree"` | 出力フォーマット: `"tree"` または `"inline"` |
| `book_name` | `str` | `""` | トップレベル `book.name` フィールドのワークブックファイル名 |
| `stream` | 書き込み可能なファイルオブジェクト | `None` | 指定した場合、YAML はストリームに書き出される |
| `ref_mode` | `str \| None` | `None` | *（非推奨）* レガシー描画モード；代わりに `depth`/`fmt` を使用 |

**戻り値：** `stream` が `None` の場合は YAML 文字列；`stream` が指定された場合は `None`。

---

## データモデル（`sheet_call_tree.models`）

AST は 5 つのデータクラスと 1 つの型エイリアスで構成されています：

```python
from sheet_call_tree.models import FunctionNode, RefNode, RangeNode, TableRefNode, NamedRefNode, Node
```

---

### `FunctionNode`

解析された関数呼び出しまたは演算子を表します。

```python
@dataclass
class FunctionNode:
    name: str    # "SUM", "ADD", "MUL", "IF", "NEG", …
    args: list   # list[Node]
```

**例：**
- `=SUM(A1:A2)` → `FunctionNode(name='SUM', args=[RangeNode(...)])`
- `=C5+1.1`   → `FunctionNode(name='ADD', args=[RefNode('Sheet1!C5', ...), 1.1])`
- `=C5*2`     → `FunctionNode(name='MUL', args=[RefNode('Sheet1!C5', ...), 2])`

演算子は関数名に正規化されます：`+` → `ADD`、`*` → `MUL`、`-` → `SUB` または単項 `NEG`、`/` → `DIV`。

---

### `RefNode`

単一セル参照を表します。

```python
@dataclass
class RefNode:
    ref: str                              # "Sheet1!A1"（@ シジルなし）
    formula: FunctionNode | None = None   # 数式セルの AST（定数/不明セルでは None）
    resolved_value: object = None         # スカラー値（定数セルの値またはキャッシュ済み計算結果）
```

**フィールド：**

| フィールド | 説明 |
|----------|------|
| `ref` | `@` プレフィックスなしの完全修飾セル参照文字列（例: `"Sheet1!C5"`） |
| `formula` | セルの解析済み AST：数式セルでは `FunctionNode`；定数セルまたはセルがワークブックに見つからない場合は `None` |
| `resolved_value` | スカラー値：定数セルではセルの値（`int`、`float`、`str`、`bool`）；数式セルでは `data_only` のキャッシュ結果（プログラムで作成されたワークブックでは `None` の場合あり） |

---

### `RangeNode`

セル範囲参照（例: `A1:A2`）を表します。

```python
@dataclass
class RangeNode:
    start: str                            # "Sheet1!A1"
    end: str                              # "Sheet1!A9"
    values: list[object] | None = None    # 範囲内の全セル値（リーダーにより設定）
```

**フィールド：**

| フィールド | 説明 |
|----------|------|
| `start` | 完全修飾の開始セル参照文字列（例: `"Sheet1!A1"`） |
| `end` | 完全修飾の終了セル参照文字列（例: `"Sheet1!A9"`） |
| `values` | 範囲内の全セル値のリスト。未設定の場合は `None`。深さ > 0 ではこれらの値が YAML 出力に含まれる。 |

---

### `TableRefNode`

構造化テーブル参照（例: `Table1[Amount]` や `Table1[@Amount]`）を表します。

```python
@dataclass
class TableRefNode:
    table_name: str            # "Table1"
    column: str | None         # "Amount"；テーブル全体の参照では None
    this_row: bool             # @プレフィックス時に True（Table1[@Amount]）
    resolved_range: str | None = None   # "Sheet1!D2:D100"
    cached_value: object = None
```

**フィールド：**

| フィールド | 説明 |
|----------|------|
| `table_name` | テーブル名（例: `"Table1"`） |
| `column` | カラム指定子。テーブル全体の参照では `None` |
| `this_row` | 参照が現在行の `@` を使用している場合 `True`（例: `Table1[@Amount]`） |
| `resolved_range` | 解決済みのセル範囲（例: `"Sheet1!D2:D100"`） |
| `cached_value` | キャッシュ済み計算値（利用可能な場合） |

YAML ツリー出力では以下のように表示されます：
```yaml
TABLE_REF:
  name: Table1
  column: Amount
  this_row: true
  range: Sheet1!D2:D100
```

インライン出力では: `TABLE_REF(Table1[@Amount])`

---

### `NamedRefNode`

名前付き範囲または定義済み名前の参照（例: `SalesTotal`）を表します。

```python
@dataclass
class NamedRefNode:
    name: str                  # "SalesTotal"
    resolved_range: str | None = None   # "Sheet1!$B$10"
    formula: FunctionNode | None = None   # 解決後の数式セル AST
    resolved_value: object = None         # スカラー値
```

**フィールド：**

| フィールド | 説明 |
|----------|------|
| `name` | 定義済み名前（例: `"SalesTotal"`） |
| `resolved_range` | 解決済みのセル範囲または参照 |
| `formula` | 名前付き参照が数式セルに解決される場合の数式セル AST |
| `resolved_value` | スカラー値（利用可能な場合） |

深さ 0 の YAML ツリー出力では以下のように表示されます：
```yaml
NAMED_REF:
  name: SalesTotal
  range: Sheet1!$B$10
```

深さ > 0 で名前付き参照が数式に解決される場合、AST が展開されます：
```yaml
NAMED_REF(SalesTotal):
  SUM:
  - ...
```

インライン出力では: `NAMED_REF(SalesTotal)`

---

### `Node` 型エイリアス

```python
Node = Union[FunctionNode, RefNode, RangeNode, TableRefNode, NamedRefNode, int, float, bool, str]
```

`FunctionNode.args` の全要素は `Node` です。プレーンなスカラー（`int`、`float`、`bool`、`str`）はリテラル引数値として現れます。

---

## 応用：低レベルモジュール

以下の内部関数は安定した公開 API の一部ではありませんが、高度なユースケースに役立つ場合があります。

| シンボル | モジュール | 説明 |
|---------|----------|------|
| `parse_formula(formula, default_sheet)` | `sheet_call_tree.formula_parser` | 単一の数式文字列を `FunctionNode` AST に解析する |
| `extract_formula_cells_from_workbook(wb)` | `sheet_call_tree.reader` | すでに読み込まれた `openpyxl.Workbook` から数式セルを抽出する |
| `build_dependency_graph(formula_cells)` | `sheet_call_tree.dependency_graph` | セルの依存関係の `dict[str, set[str]]` を返す |
| `detect_cycles(graph)` | `sheet_call_tree.dependency_graph` | 循環が見つかった場合に `CircularReferenceError` を発生させる |
| `CircularReferenceError` | `sheet_call_tree.dependency_graph` | 例外クラス；`.cycle` 属性にセル参照のリストとして循環が格納される |

---

## 例：AST を走査する

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


cells = extract_formula_cells("myfile.xlsx")
for cell_ref, ast in cells.items():
    print(f"=== {cell_ref} ===")
    walk(ast)
```

`Sheet1!C5 = SUM(A1:A2)` の出力例：

```
=== Sheet1!C5 ===
SUM(
  RANGE(Sheet1!A1 .. Sheet1!A2)
)
```
