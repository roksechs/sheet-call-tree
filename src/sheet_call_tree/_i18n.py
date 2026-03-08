"""Locale detection and localised CLI strings."""
from __future__ import annotations
import locale
import os


def detect_language() -> str:
    """Return 'ja' if system locale is Japanese, else 'en'."""
    for env_var in ("LC_MESSAGES", "LC_ALL", "LANG"):
        val = os.environ.get(env_var, "")
        if val.startswith("ja"):
            return "ja"
    try:
        lang = locale.getlocale()[0] or ""
        if lang.startswith("ja"):
            return "ja"
    except Exception:
        pass
    return "en"


STRINGS: dict[str, dict[str, str]] = {
    "en": {
        "description": "Visualize Excel formula dependencies as YAML AST.",
        "input":       "Path to the .xlsx/.xlsm file",
        "filter":      "Output only the specified cell (e.g. 'Sheet1!B10')",
        "output":      "Write YAML to FILE instead of stdout",
        "no_cycle":    "Skip circular reference detection",
        "ref_mode":    (
            "How to render formula-cell references in YAML output. "
            "'ref' (default): @-prefixed cross-ref strings. "
            "'ast': ref name as key with expanded AST as value. "
            "'value': cached computed scalar from data_only workbook. "
            "'inline': each cell as a single FUNC(...) expression string."
        ),
        "sheet":          "Output only cells in the specified sheet",
        "err_not_found": "Error: cell {ref!r} not found or has no formula.",
        "err_sheet_not_found": "Error: no formula cells found in sheet {sheet!r}.",
        "err_cycle": "Error: {msg}",
    },
    "ja": {
        "description": "Excel の数式依存関係を YAML AST として可視化します。",
        "input":       ".xlsx/.xlsm ファイルのパス",
        "filter":      "指定したセルのみ出力する（例: 'Sheet1!B10'）",
        "output":      "YAML を stdout ではなくファイルに書き出す",
        "no_cycle":    "循環参照の検出をスキップする",
        "ref_mode":    (
            "YAML 出力における数式セル参照の描画方法。"
            "'ref'（デフォルト）: @ プレフィックス付き相互参照文字列。"
            "'ast': 参照名をキーとし、展開した AST を値とする。"
            "'value': data_only ワークブックのキャッシュ済みスカラー値。"
            "'inline': 各セルを単一の FUNC(...) 式文字列として出力。"
        ),
        "sheet":          "指定したシートのセルのみ出力する",
        "err_not_found": "エラー: セル {ref!r} が見つからないか、数式がありません。",
        "err_sheet_not_found": "エラー: シート {sheet!r} に数式セルが見つかりません。",
        "err_cycle": "エラー: {msg}",
    },
}


def get_strings(lang: str | None = None) -> dict[str, str]:
    """Return the string dict for *lang* (detected automatically if None)."""
    if lang is None:
        lang = detect_language()
    return STRINGS.get(lang, STRINGS["en"])
