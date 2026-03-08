# 出力フォーマット

`sheet-call-tree` は 2 つのパラメーターで出力を制御します：

- **`--depth N`** — 数式セル参照の展開の深さを制御します。`0`（デフォルト）は参照のみを表示、`inf` は完全に展開された AST を表示します。
- **`--format {tree,inline}`** — 出力構造を制御します。`tree`（デフォルト）はネストされた YAML を生成、`inline` は各セルを単一の式文字列として表示します。

定数セル参照（数式ではなく通常の値を含むセル）は深さやフォーマットに関わらず常にスカラー値に解決されます。

全ての例で同じワークブックを使用します：

| セル | 値 |
|------|-----|
| A1 | `10` |
| A2 | `20` |
| C5 | `=SUM(A1:A2)` |
| B10 | `=C5+1.1` |
| B11 | `=C5*2` |

---

## 深さ 0 — 参照のみ（デフォルト）

**使用場面：** 各数式セル参照がナビゲート可能なクロスリファレンス文字列として表示される、コンパクトな依存関係マップが欲しいとき。人間によるレビューや差分に適しています。

数式セル参照は `Sheet1!C5` 文字列として表示されます。範囲参照は `Sheet1!A1:A2` 文字列として表示されます。

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

`B10` は `ADD` を 2 つの引数で呼び出しています：数式セル `Sheet1!C5` への参照と定数 `1.1` です。`C5` はフル AST を持つ独立した `cells` エントリとして現れます。

---

## 深さ inf — 完全展開

**使用場面：** クロスリファレンスを追わずに完全な依存関係ツリーをその場で展開したいとき。詳細分析や後続ツールへの入力に便利です。

数式セル参照は `{cell: Sheet1!C5, expression: ...}` マッピングとして表示され、参照先セルの AST をインラインに埋め込みます。範囲の値は親関数の `inputs` にフラット化されます。

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

`Sheet1!C5` エントリにはフルの `SUM(...)` サブツリーが `expression` として付加されており、別のエントリにジャンプすることなく `B10` と `B11` の中で直接確認できます。範囲の値（`10`、`20`）は `inputs` リストにフラット化されます。

---

## インラインフォーマット

**使用場面：** 最もコンパクトな人間が読める表現が欲しいとき、またはコメントやチケット、ログへの貼り付け用に 1 セルあたり単一の式文字列が必要なとき。

各数式セルの依存関係ツリー全体が単一の `FUNC(arg1, arg2, …)` 式文字列として表示されます。`--depth` パラメーターにより、参照を展開するか参照文字列のまま保持するかを制御します。

### インライン（深さ 0、デフォルト）

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

### インライン（深さ inf、完全展開）

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

`C5` は `B10` と `B11` に完全にインライン展開されていることに注意してください — `ADD(SUM(10, 20), 1.1)` — 名前で参照されるのではなく直接埋め込まれています。

---

## セマンティックラベル

ワークブックにヘッダー行やヘッダー列がある場合、`sheet-call-tree` は学習済み分類器を使用して自動的にそれらを検出し、各数式セルに `labels` を出力します：

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

各ラベルリストは最大 5 件の候補を含み、近い順に並び、重複は除外されます。`row` ラベルは同じ行を左方向に走査して取得、`column` ラベルは同じ列を上方向に走査して取得します。

---

## 全モードにおける定数セル参照

セルの値が通常のスカラー（数値、文字列、真偽値）であり数式でない場合、そのセルは*定数セル*です。全てのモードにおいて、定数セル参照は直接スカラー値に解決されます。上の例では、`A1=10` と `A2=20` は深さ > 0 では `inputs` リストにフラット化され、深さ 0 では省略されます。

---

## 非推奨の `--ref-mode`

レガシーの `--ref-mode` フラグは後方互換性のために引き続き使用できます：

| レガシー値 | 同等の設定 |
|-----------|-----------|
| `ref` | `--depth 0` |
| `ast` | `--depth inf` |
| `inline` | `--format inline --depth inf` |

`value` モードは削除されました。
