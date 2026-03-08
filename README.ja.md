# sheet-call-tree

Excel の数式依存関係を YAML AST コールツリーとして可視化します。

## 概要

Excel の数式は他のセルを参照でき、それ自体がさらに数式を含む場合もあるため、スプレッドシート上で依存関係チェーンを追うのは困難です。`sheet-call-tree` は `.xlsx` ファイルを読み込み、すべての数式セルを抽象構文木（AST）に解析し、セル間の参照を解決して、完全な依存関係ツリーを YAML として出力します。数式セルごとに 1 つのトップレベルキーが生成されます。

## インストール

```bash
uv tool install git+https://github.com/roksechs/sheet-call-tree.git
```

詳細および開発用のインストール手順については [user_manuals/ja/installation.md](user_manuals/ja/installation.md) を参照してください。

## 60 秒クイックスタート

以下のセルを含むワークブックに対して：

| セル | 値 |
|------|-----|
| A1 | `10` |
| A2 | `20` |
| C5 | `=SUM(A1:A2)` |
| B10 | `=C5+1.1` |
| B11 | `=C5*2` |

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
    - cell: C5
      formula:
        SUM:
        - RANGE:
          - 10
          - 20
```

出力は `book → sheets → cells` の階層で整理されます。各セルエントリは `cell` 座標と `formula`（解析済み AST）を持ちます。定数セル参照はスカラーに解決され、数式セル参照はデフォルトの `ref` モードで `@シート名!セル参照` 文字列として表示されます。

## CLI フラグ一覧

| フラグ | デフォルト | 説明 |
|--------|----------|------|
| `INPUT` | *(必須)* | `.xlsx` ファイルへのパス |
| `--filter CELL` | — | 指定したセルのみ出力（例: `Sheet1!B10`） |
| `--output FILE` | stdout | YAML をファイルに書き出す |
| `--no-cycle-check` | off | 循環参照の検出をスキップする |
| `--ref-mode MODE` | `ref` | 数式セル参照の描画方法（下記参照） |

完全なリファレンス: [user_manuals/ja/cli-reference.md](user_manuals/ja/cli-reference.md)

## 出力モード（`--ref-mode`）

| モード | 数式セル参照の描画 |
|--------|-----------------|
| `ref`（デフォルト） | `@Sheet1!C5` — クロスリファレンス文字列 |
| `ast` | `'@Sheet1!C5': {SUM: ...}` — 展開したサブツリーを値として持つキー |
| `value` | Excel のキャッシュ済みスカラー値（キャッシュがない場合は `null`） |
| `inline` | 完全に展開された式文字列（例: `ADD(SUM(RANGE(10, 20)), 1.1)`） |

例付きの詳細: [user_manuals/ja/output-formats.md](user_manuals/ja/output-formats.md)

## Python API

```python
from sheet_call_tree import extract_formula_cells, to_yaml

cells = extract_formula_cells("myfile.xlsx")
print(to_yaml(cells))                        # ref モード（デフォルト）
print(to_yaml(cells, ref_mode="inline"))     # inline モード
```

完全な API リファレンス: [user_manuals/ja/python-api.md](user_manuals/ja/python-api.md)

## ドキュメント

- [インストール](user_manuals/ja/installation.md)
- [クイックスタート](user_manuals/ja/quickstart.md)
- [CLI リファレンス](user_manuals/ja/cli-reference.md)
- [出力フォーマット](user_manuals/ja/output-formats.md)
- [Python API](user_manuals/ja/python-api.md)
