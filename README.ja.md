# sheet-call-tree

Excel の数式依存関係を YAML AST コールツリーとして可視化します。

## 概要

Excel の数式は他のセルを参照でき、それ自体がさらに数式を含む場合もあるため、スプレッドシート上で依存関係チェーンを追うのは困難です。`sheet-call-tree` は `.xlsx`/`.xlsm` ファイルを読み込み、すべての数式セルを抽象構文木（AST）に解析し、セル間の参照を解決して、完全な依存関係ツリーを YAML として出力します。数式セルごとに 1 つのトップレベルキーが生成されます。

また、学習済み RandomForest 分類器を使用して周囲のセルをヘッダーまたはデータとして分類し、各方向で最も近いヘッダーセルを走査することで、各数式セルの**セマンティックラベルを検出**します。

## インストール

```bash
uv tool install git+https://github.com/roksechs/sheet-call-tree.git
```

詳細および開発用のインストール手順については [user_manuals/ja/installation.md](user_manuals/ja/installation.md) を参照してください。

## 60 秒クイックスタート

以下のセルを含むワークブックに対して：

| セル | 値 |
|------|-----|
| A1 | `Name` |
| B1 | `Q1` |
| C1 | `Q2` |
| D1 | `Total` |
| A2 | `Alice` |
| B2 | `100` |
| C2 | `200` |
| D2 | `=SUM(B2:C2)` |

```bash
sheet-call-tree myfile.xlsx
```

出力：

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
      expression:
        type: SUM
        inputs:
        - Sheet1!B2:C2
```

出力は `book → sheets → cells` の階層で整理されます。各セルエントリは以下を持ちます：
- `cell` — セル座標
- `labels` — 検出されたセマンティックラベル（`row` と `column` のヘッダー、近い順）
- `expression` — 解析済み AST（`type`：関数名、`inputs`：引数）

範囲参照は `Sheet1!A1:A2` 文字列として表示されます。数式セル参照はデフォルトの深さ（0）で `Sheet1!C5` 文字列として表示されます。

## CLI フラグ一覧

| フラグ | デフォルト | 説明 |
|--------|----------|------|
| `INPUT` | *(必須)* | `.xlsx`/`.xlsm` ファイルへのパス |
| `--sheet SHEETNAME` | — | 指定したシートのセルのみ出力する |
| `--filter CELL` | — | 指定したセルのみ出力（例: `Sheet1!B10`） |
| `--output FILE` | stdout | YAML をファイルに書き出す |
| `--no-cycle-check` | off | 循環参照の検出をスキップする |
| `--depth N` | `0` | 展開の深さ: 0 = 参照のみ、inf = 完全展開 |
| `--format FORMAT` | `tree` | 出力フォーマット: `tree`、`inline`、または `json` |
| `--roots-only` | off | ルートセルのみ出力（他の数式から参照されないセル） |

完全なリファレンス: [user_manuals/ja/cli-reference.md](user_manuals/ja/cli-reference.md)

## 出力モード（`--depth` / `--format`）

| 設定 | 数式セル参照の描画 |
|------|-----------------|
| `--depth 0`（デフォルト） | `Sheet1!C5` — クロスリファレンス文字列 |
| `--depth inf` | `{cell: Sheet1!C5, expression: {type: SUM, inputs: [...]}}` — 展開したサブツリー |
| `--format inline` | `SUM(Sheet1!B2:C2)` — 完全に展開された式文字列 |
| `--format json` | tree と同じ構造を JSON として出力 |

非推奨の `--ref-mode` フラグは引き続き使用可能です（`ref`→depth 0、`ast`→depth inf、`inline`→format inline）。

例付きの詳細: [user_manuals/ja/output-formats.md](user_manuals/ja/output-formats.md)

## Python API

```python
from sheet_call_tree import extract_formula_cells, to_json, to_yaml

cells, data_values, label_map = extract_formula_cells("myfile.xlsx")
print(to_yaml(cells, data_values=data_values, label_map=label_map))
# JSON として出力:
print(to_json(cells, data_values=data_values, label_map=label_map))
```

完全な API リファレンス: [user_manuals/ja/python-api.md](user_manuals/ja/python-api.md)

## ドキュメント

- [インストール](user_manuals/ja/installation.md)
- [クイックスタート](user_manuals/ja/quickstart.md)
- [CLI リファレンス](user_manuals/ja/cli-reference.md)
- [出力フォーマット](user_manuals/ja/output-formats.md)
- [Python API](user_manuals/ja/python-api.md)
