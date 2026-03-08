# Python API

`sheet-call-tree` はライブラリとして使用できます。公開 API はトップレベルパッケージからインポートできる 3 つの関数で構成されています。

## 公開 API

```python
from sheet_call_tree import extract_formula_cells, to_json, to_yaml
```

---

### `extract_formula_cells(path) -> tuple[dict, dict, dict]`

`.xlsx`/`.xlsm` ファイルを読み込み、数式セル・データ値・ラベルマップの 3-タプルを返します。

```python
cells, data_values, label_map = extract_formula_cells("myfile.xlsx")
# cells == {
#   'Sheet1!B10': FunctionNode(name='ADD', args=[RefNode(...), 1.1]),
#   'Sheet1!C5':  FunctionNode(name='SUM', args=[RangeNode(...)]),
# }
# data_values == {'Sheet1!A1': 10, 'Sheet1!A2': 20, ...}
# label_map == {'Sheet1!D2': {'row': ['Alice'], 'column': ['Total']}, ...}
```

**パラメーター：**

| パラメーター | 型 | 説明 |
|------------|-----|------|
| `path` | `str \| Path` | `.xlsx`/`.xlsm` ファイルへのパス |

**戻り値：** `tuple[dict[str, FunctionNode], dict[str, object], dict[str, dict]]`

| 要素 | 型 | 説明 |
|------|-----|------|
| `formula_cells` | `dict[str, FunctionNode]` | キーは `'シート名!セル参照'` 文字列；値は解析済み AST ルート |
| `data_values` | `dict[str, object]` | キーはセル参照；値は `data_only` モードからのキャッシュ済みスカラー値 |
| `label_map` | `dict[str, dict[str, object]]` | キーはセル参照；値は `'row'` と `'column'` キーを持つ辞書（各方向最大 5 件のラベル文字列リスト、近い順） |

**動作：**
- ワークブックは 2 回読み込まれます：数式テキスト用（`data_only=False`）とキャッシュ済み計算値用（`data_only=True`）。
- 数式セル（値が `=` で始まるセル）のみが `formula_cells` に含まれます。
- 定数セルは数式 AST 内の `RefNode` 値としてのみ現れます。
- `label_map` は学習済み RandomForest 分類器によりヘッダーとデータセルを識別し、各数式セルの上方と左方の最も近いヘッダーを走査して構築されます。

---

### `to_yaml(cells, *, depth=None, fmt="tree", book_name="", data_values=None, label_map=None, stream=None) -> str | None`

`formula_cells` 辞書を YAML にシリアライズします。

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells, data_values, label_map = extract_formula_cells("myfile.xlsx")

# YAML 文字列を返す（深さ 0、ツリーフォーマット）
yaml_str = to_yaml(cells, data_values=data_values, label_map=label_map)

# 完全展開
yaml_str = to_yaml(cells, depth=float('inf'), data_values=data_values, label_map=label_map)

# インラインフォーマット
with open("deps.yaml", "w") as fh:
    to_yaml(cells, fmt="inline", data_values=data_values, label_map=label_map, stream=fh)
```

**パラメーター：**

| パラメーター | 型 | デフォルト | 説明 |
|------------|-----|----------|------|
| `cells` | `dict[str, FunctionNode]` | *(必須)* | `extract_formula_cells` の出力 |
| `depth` | `int \| float \| None` | `None`（→ 0） | 展開の深さ: 0 = 参照のみ、`inf` = 完全展開 |
| `fmt` | `str` | `"tree"` | 出力フォーマット: `"tree"` または `"inline"` |
| `book_name` | `str` | `""` | トップレベル `book.name` フィールドのワークブックファイル名 |
| `data_values` | `dict[str, object] \| None` | `None` | `outputs` フィールド用のキャッシュ済みスカラー値 |
| `label_map` | `dict[str, dict] \| None` | `None` | `labels` ブロック用のラベルマップ |
| `stream` | 書き込み可能なファイルオブジェクト | `None` | 指定した場合、YAML はストリームに書き出される |
| `ref_mode` | `str \| None` | `None` | *（非推奨）* レガシー描画モード；代わりに `depth`/`fmt` を使用 |

**戻り値：** `stream` が `None` の場合は YAML 文字列；`stream` が指定された場合は `None`。

---

### `to_json(cells, *, stream=None, **kw) -> str | None`

`formula_cells` 辞書を JSON にシリアライズします。`to_yaml` と同じキーワード引数（`depth`、`fmt`、`ref_mode`、`book_name`、`data_values`、`label_map`）を受け付けます。

```python
from sheet_call_tree import extract_formula_cells, to_json

cells, data_values, label_map = extract_formula_cells("myfile.xlsx")

# JSON 文字列を返す
json_str = to_json(cells, data_values=data_values, label_map=label_map)

# ファイルに書き出す
with open("deps.json", "w") as fh:
    to_json(cells, data_values=data_values, label_map=label_map, stream=fh)
```

**戻り値：** `stream` が `None` の場合は JSON 文字列；`stream` が指定された場合は `None`。

特殊な浮動小数点値の処理：`NaN` → `null`、`Infinity` → `"Infinity"`、`-Infinity` → `"-Infinity"`。

---

**出力構造（ツリーフォーマット、`to_yaml` と `to_json` で共通）：**

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

YAML ツリー出力では以下のように表示されます：
```yaml
type: SUM
inputs:
- Sheet1!A1:A2
```

---

### `RefNode`

単一セル参照を表します。

```python
@dataclass
class RefNode:
    ref: str                              # "Sheet1!A1"
    formula: FunctionNode | None = None   # 数式セルの AST（定数/不明セルでは None）
    resolved_value: object = None         # スカラー値（定数セルの値またはキャッシュ済み計算結果）
```

**フィールド：**

| フィールド | 説明 |
|----------|------|
| `ref` | 完全修飾セル参照文字列（例: `"Sheet1!C5"`） |
| `formula` | セルの解析済み AST：数式セルでは `FunctionNode`；定数セルまたはセルがワークブックに見つからない場合は `None` |
| `resolved_value` | スカラー値：定数セルではセルの値（`int`、`float`、`str`、`bool`）；数式セルでは `data_only` のキャッシュ結果（プログラムで作成されたワークブックでは `None` の場合あり） |

YAML の深さ 0: `Sheet1!C5`
深さ > 0: `{cell: Sheet1!C5, expression: ...}`

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
| `values` | 範囲内の全セル値のリスト。未設定の場合は `None`。深さ > 0 では親の `inputs` リストにフラット化される。 |

YAML の深さ 0: `Sheet1!A1:A2`
深さ > 0: 値が親の `inputs` に直接フラット化

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
type: TABLE_REF
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
type: NAMED_REF
name: SalesTotal
range: Sheet1!$B$10
```

深さ > 0 で名前付き参照が数式に解決される場合、AST が展開されます：
```yaml
named_ref: SalesTotal
expression:
  type: SUM
  inputs:
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
| `build_label_map(...)` | `sheet_call_tree.labeler` | 学習済み分類器を使ってラベルマップを構築する |

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


cells, data_values, label_map = extract_formula_cells("myfile.xlsx")
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
