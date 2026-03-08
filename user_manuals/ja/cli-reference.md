# CLI リファレンス

## 書式

```
sheet-call-tree INPUT [OPTIONS]
```

## 位置引数

| 引数 | 説明 |
|------|------|
| `INPUT` | 解析する `.xlsx` ファイルへのパス。必須。 |

## オプション

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

### `--ref-mode {ref,ast,value,inline}`

YAML 出力における数式セル参照の描画方法を制御します。デフォルト: `ref`。

| 値 | 説明 |
|----|------|
| `ref` | `@` プレフィックス付きのクロスリファレンス文字列（例: `'@Sheet1!C5'`） |
| `ast` | `@` プレフィックス付きの名前をキーとし、展開したサブツリーを値とするマッピング |
| `value` | Excel の `data_only` ワークブックからのキャッシュ済みスカラー値。キャッシュがない場合は `null` |
| `inline` | 各セルを単一の `FUNC(arg1, arg2, …)` 式文字列として出力 |

定数セル参照（数式ではなく通常の値を含むセル）はどのモードでも常にスカラー値に解決されます。

各モードの具体的な YAML 例については [output-formats.md](output-formats.md) を参照してください。

## 終了コード

| コード | 意味 |
|--------|------|
| `0` | 成功 |
| `1` | セルが見つからない（`--filter` 使用時）またはその他のエラー |
