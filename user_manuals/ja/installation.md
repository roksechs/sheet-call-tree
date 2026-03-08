# インストール

## 必要条件

- Python ≥ 3.11
- [`uv`](https://github.com/astral-sh/uv)（推奨）または PEP 517 互換インストーラー

## GitHub からインストール

```bash
uv tool install git+https://github.com/roksechs/sheet-call-tree.git
```

## ソースからインストール（開発用の編集可能インストール）

リポジトリをクローンし、開発用依存関係付きで編集可能モードでインストールします：

```bash
git clone https://github.com/roksechs/sheet-call-tree.git
cd sheet-call-tree
uv sync --dev
```

`--dev` フラグにより、テストスイートを実行するための `pytest` と `pytest-cov` が含まれます。

## インストールの確認

```bash
sheet-call-tree --help
```

期待される出力（日本語ロケールの場合）：

```
usage: sheet-call-tree [-h] [--filter CELL] [--output FILE] [--no-cycle-check]
                       [--ref-mode {ref,ast,value,inline}]
                       input

Excel の数式依存関係を YAML AST として可視化します。

positional arguments:
  input                 .xlsx ファイルのパス

options:
  -h, --help            show this help message and exit
  --filter CELL         指定したセルのみ出力する（例: 'Sheet1!B10'）
  --output FILE         YAML を stdout ではなくファイルに書き出す
  --no-cycle-check      循環参照の検出をスキップする
  --ref-mode {ref,ast,value,inline}
                        YAML 出力における数式セル参照の描画方法。
```

## テストの実行

```bash
.venv/bin/pytest
```
