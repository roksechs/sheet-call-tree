# CLI リファレンス

## 書式

```
sheet-call-tree INPUT [OPTIONS]
```

## 位置引数

| 引数 | 説明 |
|------|------|
| `INPUT` | 解析する `.xlsx`/`.xlsm` ファイルへのパス。必須。 |

## オプション

### `--sheet SHEETNAME`

指定したシートのセルのみ出力します。

シートが存在しないか数式セルがない場合、`sheet-call-tree` はエラーを stderr に出力してコード `1` で終了します。

```bash
sheet-call-tree myfile.xlsx --sheet Sheet1
```

### `--filter CELL`

指定したセルのみ出力します。`CELL` は `シート名!列行` 形式（例: `Sheet1!B10`）で指定します。

セルが存在しないか数式を持たない場合、`sheet-call-tree` はエラーを stderr に出力してコード `1` で終了します。

```bash
sheet-call-tree myfile.xlsx --filter Sheet1!B10
```

### `--output FILE`

YAML を標準出力ではなく `FILE` に書き出します。ファイルは UTF-8 エンコードで作成（または上書き）されます。このフラグを使用した場合、標準出力には何も出力されません。

```bash
sheet-call-tree myfile.xlsx --output deps.yaml
```

### `--no-cycle-check`

デフォルトでは、`sheet-call-tree` は依存関係グラフを構築し、循環参照（例: A1 が B1 に依存し、B1 が A1 に依存する）が検出された場合にエラーを発生させます。`--no-cycle-check` を使用すると、このチェックをスキップして循環が存在しても出力を生成します。

```bash
sheet-call-tree myfile.xlsx --no-cycle-check
```

### `--depth N`

数式セル参照の展開の深さを制御します。デフォルト: `0`。

| 値 | 説明 |
|----|------|
| `0` | 参照のみ — 数式セル参照は `Sheet1!C5` クロスリファレンス文字列として表示；範囲参照は `Sheet1!A1:A2` として値なしで表示 |
| `inf` | 完全展開 — 数式セル参照は `{cell: Sheet1!C5, expression: ...}` として AST をその場で展開；範囲の値は親の `inputs` にフラット化 |

中間の整数値（1、2、…）はそのレベル数まで展開します。

各モードの具体的な YAML 例については [output-formats.md](output-formats.md) を参照してください。

### `--format {tree,inline,json}`

出力フォーマットを制御します。デフォルト: `tree`。

| 値 | 説明 |
|----|------|
| `tree` | `type` と `inputs` キーを持つネストされた YAML 構造（デフォルト） |
| `inline` | 各セルを単一の `FUNC(arg1, arg2, …)` 式文字列として出力 |
| `json` | `tree` と同じ構造を 2 スペースインデントの JSON として出力 |

定数セル参照（数式ではなく通常の値を含むセル）はどのモードでも常にスカラー値に解決されます。

### `--roots-only`

ルートセルのみ出力 — 他の数式セルから参照されていないセルです。ワークブックの計算グラフの「エントリポイント」を見つけるのに便利です。

```bash
sheet-call-tree myfile.xlsx --roots-only
```

### `--ref-mode`（非推奨）

レガシーの `--ref-mode` フラグは後方互換性のために引き続き使用できますが、`--help` には表示されません。

| レガシー値 | 同等の設定 |
|-----------|-----------|
| `ref` | `--depth 0` |
| `ast` | `--depth inf` |
| `inline` | `--format inline --depth inf` |

`value` モードは削除されました。

## 終了コード

| コード | 意味 |
|--------|------|
| `0` | 成功 |
| `1` | セルが見つからない（`--filter`）、シートが見つからない（`--sheet`）、循環参照の検出、またはその他のエラー |
