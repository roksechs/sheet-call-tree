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
- `RefNode.value` には参照先セルのスカラー値（定数セルの場合）または `FunctionNode`（数式セルの場合）が格納されます。`RefNode.cached_value` には `data_only` のスカラー値が格納されます（プログラムで作成されたワークブックでは `None` の場合があります）。

---

### `to_yaml(cells, ref_mode="ref", stream=None) -> str | None`

`formula_cells` 辞書を YAML にシリアライズします。

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells = extract_formula_cells("myfile.xlsx")

# YAML 文字列を返す
yaml_str = to_yaml(cells)

# ファイルに書き込む
with open("deps.yaml", "w") as fh:
    to_yaml(cells, ref_mode="inline", stream=fh)
```

**パラメーター：**

| パラメーター | 型 | デフォルト | 説明 |
|------------|-----|----------|------|
| `cells` | `dict[str, FunctionNode]` | *(必須)* | `extract_formula_cells` の出力 |
| `ref_mode` | `str` | `"ref"` | 描画モード: `"ref"`、`"ast"`、`"value"`、または `"inline"` |
| `stream` | 書き込み可能なファイルオブジェクト | `None` | 指定した場合、YAML はストリームに書き出される |

**戻り値：** `stream` が `None` の場合は YAML 文字列；`stream` が指定された場合は `None`。

**例外：** `ref_mode` が 4 つの有効な値のいずれでもない場合、`ValueError` を発生させます。

---

## データモデル（`sheet_call_tree.models`）

AST は 3 つのデータクラスと 1 つの型エイリアスで構成されています：

```python
from sheet_call_tree.models import FunctionNode, RefNode, RangeNode, Node
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
    ref: str                    # "Sheet1!A1"（@ シジルなし）
    value: object = None        # 定数セルはスカラー；数式セルは FunctionNode；不明の場合は None
    cached_value: object = None # data_only の計算値（--ref-mode value 用）；利用不可の場合は None
```

**フィールド：**

| フィールド | 説明 |
|----------|------|
| `ref` | `@` プレフィックスなしの完全修飾セル参照文字列（例: `"Sheet1!C5"`） |
| `value` | セルの内容：定数セルではスカラー（`int`、`float`、`str`、`bool`）；数式セルでは `FunctionNode`；セルがワークブックに見つからない場合は `None` |
| `cached_value` | `data_only` の計算値。Excel で保存されたワークブックではこれが最後に計算された結果です。プログラムで作成されたワークブック（openpyxl で作成され、Excel で開かれたことがないもの）では `None` です。 |

---

### `RangeNode`

セル範囲参照（例: `A1:A2`）を表します。

```python
@dataclass
class RangeNode:
    start: RefNode   # 範囲の最初のセル
    end: RefNode     # 範囲の最後のセル
```

`start` と `end` はワークブックから `value` / `cached_value` フィールドが設定された `RefNode` インスタンスです。

---

### `Node` 型エイリアス

```python
Node = Union[FunctionNode, RefNode, RangeNode, int, float, bool, str]
```

`FunctionNode.args` の全要素と `RangeNode.start` / `end` は `Node` です。プレーンなスカラー（`int`、`float`、`bool`、`str`）はリテラル引数値として現れます。

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

`Sheet1!C5 = SUM(A1:A2)` の出力例：

```
=== Sheet1!C5 ===
SUM(
  RANGE(Sheet1!A1 .. Sheet1!A2)
)
```
