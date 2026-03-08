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

出力（デフォルトの `ref` モード）：

```yaml
Sheet1!B10:
  ADD:
  - '@Sheet1!C5'
  - 1.1
Sheet1!B11:
  MUL:
  - '@Sheet1!C5'
  - 2
Sheet1!C5:
  SUM:
  - RANGE:
    - 10
    - 20
```

**出力の読み方：**

- トップレベルのキー（`Sheet1!B10`、`Sheet1!B11`、`Sheet1!C5`）は数式セルです。
- 値はそのセルの数式の解析済み AST です。
- `ADD`、`MUL`、`SUM` は解析されたオペレーター/関数で、その引数は YAML のリストです。
- `RANGE: [10, 20]` は `A1:A2` を表し、両端点がスカラー値に解決されています。
- `'@Sheet1!C5'` は別の数式セルへのクロスリファレンスです。`@` プレフィックスにより通常の文字列と区別されます。
- 出力キーはアルファベット順にソートされています（`Sheet1!B10` < `Sheet1!B11` < `Sheet1!C5`）。

## 3. 単一セルを詳しく調べる

`--filter` を使って特定のセルに絞り込みます：

```bash
sheet-call-tree example.xlsx --filter Sheet1!B10
```

```yaml
Sheet1!B10:
  ADD:
  - '@Sheet1!C5'
  - 1.1
```

## 4. フル AST をインラインで展開する（`--ref-mode ast`）

`ref` モードはクロスリファレンスをコンパクトに保ちます。`--ref-mode ast` を使うと、参照された数式セルのサブツリーをその場で展開します：

```bash
sheet-call-tree example.xlsx --ref-mode ast
```

```yaml
Sheet1!B10:
  ADD:
  - '@Sheet1!C5':
      SUM:
      - RANGE:
        - 10
        - 20
  - 1.1
Sheet1!B11:
  MUL:
  - '@Sheet1!C5':
      SUM:
      - RANGE:
        - 10
        - 20
  - 2
Sheet1!C5:
  SUM:
  - RANGE:
    - 10
    - 20
```

`@Sheet1!C5` キーにはフルのサブツリーが値として付加されており、別のエントリにジャンプすることなく完全な依存関係をインラインで読めます。

## 5. コンパクトな式文字列（`--ref-mode inline`）

`inline` モードは各セルを単一の完全展開済み式文字列として表示します：

```bash
sheet-call-tree example.xlsx --ref-mode inline
```

```yaml
Sheet1!B10: ADD(SUM(RANGE(10, 20)), 1.1)
Sheet1!B11: MUL(SUM(RANGE(10, 20)), 2)
Sheet1!C5: SUM(RANGE(10, 20))
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
- [出力フォーマット](output-formats.md) — 4 つの `--ref-mode` 値と例
- [Python API](python-api.md) — `sheet-call-tree` をライブラリとして使用する
