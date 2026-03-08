# クイックスタート

このウォークスルーでは、小さなサンプルワークブックを使って `sheet-call-tree` の主なワークフローを紹介します。

## 1. サンプルワークブックの作成

以下の内容を `example.xlsx` として保存するか、Python で作成します：

| セル | 値 |
|------|-----|
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

## 2. `sheet-call-tree` を実行する

```bash
sheet-call-tree example.xlsx
```

出力（デフォルト：深さ 0、ツリーフォーマット）：

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

**出力の読み方：**

- トップレベルの `book` キーがワークブックとシートでセルをグループ化します。
- `cells` の各エントリには `cell` 座標（`B10`、`C5` など）と `formula` AST があります。
- `ADD`、`MUL`、`SUM` は解析されたオペレーター/関数で、その引数は YAML のリストです。
- `RANGE: {ref: '@Sheet1!A1:A2'}` は `A1:A2` を範囲参照として表します。深さ 0 では
  セルの値は含まれません。
- `'@Sheet1!C5'` は別の数式セルへのクロスリファレンスです。`@` プレフィックスにより通常の文字列と区別されます。

## 3. 単一セルを詳しく調べる

`--filter` を使って特定のセルに絞り込みます：

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
      formula:
        ADD:
        - '@Sheet1!C5'
        - 1.1
```

## 4. フル AST をインラインで展開する（`--depth inf`）

デフォルトの深さ（0）ではクロスリファレンスがコンパクトに保たれます。`--depth inf` を使うと、参照された数式セルのサブツリーをその場で展開します：

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

`@Sheet1!C5` キーにはフルのサブツリーが値として付加されており、別のエントリにジャンプすることなく完全な依存関係をインラインで読めます。`RANGE` ノードは解決済みのセル値を含む `values` を含んでいます。

## 5. コンパクトな式文字列（`--format inline`）

`inline` フォーマットは各セルを単一の完全展開済み式文字列として表示します：

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

コメントやチケットへの貼り付け、または素早い人間による読み取りに便利です。

## 6. 出力をファイルに書き出す

`--output` を使って YAML を標準出力ではなくファイルに保存します：

```bash
sheet-call-tree example.xlsx --output result.yaml
```

何も出力されず、`result.yaml` にフルの YAML が書き込まれます。

## 次のステップ

- [CLI リファレンス](cli-reference.md) — 全フラグのドキュメント
- [出力フォーマット](output-formats.md) — 深さベースの展開とフォーマットオプションの例
- [Python API](python-api.md) — `sheet-call-tree` をライブラリとして使用する
