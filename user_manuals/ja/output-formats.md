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

数式セル参照は `'@シート名!セル参照'` 文字列として表示されます。範囲参照は `RANGE: {ref: '@Sheet1!A1:A2'}` として解決済みの値なしで表示されます。

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
      formula:
        SUM:
        - RANGE:
            ref: '@Sheet1!A1:A2'
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
```

`B10` は `ADD` を 2 つの引数で呼び出しています：数式セル `Sheet1!C5` への参照と定数 `1.1` です。`C5` はフル AST を持つ独立した `cells` エントリとして現れます。`RANGE` ノードは値を展開せずに範囲参照を表示しています。

---

## 深さ inf — 完全展開

**使用場面：** クロスリファレンスを追わずに完全な依存関係ツリーをその場で展開したいとき。詳細分析や後続ツールへの入力に便利です。

数式セル参照は単一キーのマッピング `{'@シート名!セル参照': <サブツリー>}` として表示され、参照先セルの AST をインラインに埋め込みます。範囲参照は解決済みの `values` を含みます。

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
      formula:
        SUM:
        - RANGE:
            ref: '@Sheet1!A1:A2'
            values:
            - 10
            - 20
    - cell: B10
      formula:
        ADD:
        - '@Sheet1!C5':
            SUM:
            - RANGE:
                ref: '@Sheet1!A1:A2'
                values:
                - 10
                - 20
        - 1.1
    - cell: B11
      formula:
        MUL:
        - '@Sheet1!C5':
            SUM:
            - RANGE:
                ref: '@Sheet1!A1:A2'
                values:
                - 10
                - 20
        - 2
```

`@Sheet1!C5` キーはフルの `SUM(RANGE(...))` サブツリーを値として持ち、別のエントリにジャンプすることなく `B10` と `B11` の中で直接確認できます。`RANGE` ノードは解決済みのセル値を含む `values` リストを含んでいます。

---

## インラインフォーマット

**使用場面：** 最もコンパクトな人間が読める表現が欲しいとき、またはコメントやチケット、ログへの貼り付け用に 1 セルあたり単一の式文字列が必要なとき。

各数式セルの依存関係ツリー全体が単一の `FUNC(arg1, arg2, …)` 式文字列として表示されます。`--depth` パラメーターにより、参照を展開するか `@` プレフィックス付き文字列のまま保持するかを制御します。

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
      formula: SUM(RANGE(@Sheet1!A1:A2))
    - cell: B10
      formula: ADD(@Sheet1!C5, 1.1)
    - cell: B11
      formula: MUL(@Sheet1!C5, 2)
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
      formula: SUM(RANGE(@Sheet1!A1:A2, [10, 20]))
    - cell: B10
      formula: ADD(SUM(RANGE(@Sheet1!A1:A2, [10, 20])), 1.1)
    - cell: B11
      formula: MUL(SUM(RANGE(@Sheet1!A1:A2, [10, 20])), 2)
```

`C5` は `B10` と `B11` に完全にインライン展開されていることに注意してください — `ADD(SUM(RANGE(@Sheet1!A1:A2, [10, 20])), 1.1)` — 名前で参照されるのではなく直接埋め込まれています。

---

## 全モードにおける定数セル参照

セルの値が通常のスカラー（数値、文字列、真偽値）であり数式でない場合、そのセルは*定数セル*です。全てのモードにおいて、定数セル参照は直接スカラー値に解決されます。上の例では、`A1=10` と `A2=20` は `RANGE` ノードの `values` リスト内に現れます（深さ > 0 の場合）。深さ 0 では省略されます。

---

## 非推奨の `--ref-mode`

レガシーの `--ref-mode` フラグは後方互換性のために引き続き使用できます：

| レガシー値 | 同等の設定 |
|-----------|-----------|
| `ref` | `--depth 0` |
| `ast` | `--depth inf` |
| `inline` | `--format inline --depth inf` |

`value` モードは削除されました。
