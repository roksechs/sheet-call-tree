"""Detect semantic labels for cells using a trained cell-role classifier.

Uses a RandomForest classifier trained on CTC (CIUS + SAUS) data to predict
whether each cell is a header or data cell. Then assigns labels to formula
cells by scanning up (column labels) and left (row labels) for the nearest
header cells, returning up to top_k candidates for each direction.

No hardcoded weights or thresholds — all parameters are learned from data.
"""
from __future__ import annotations

import bisect
import re
from collections import defaultdict
from pathlib import Path

import joblib
import numpy as np
from openpyxl.utils import column_index_from_string, get_column_letter

from .models import FunctionNode

_CELL_REF_RE = re.compile(r"^(.+)!([A-Za-z]+)(\d+)$")

_HEADER = 0
_DATA = 1
_N_FEATURES = 23

_DEFAULT_TOP_K = 5

_clf = None


def _load_classifier():
    global _clf
    if _clf is None:
        model_path = Path(__file__).parent / "cell_classifier.joblib"
        _clf = joblib.load(model_path)
    return _clf


def _is_numeric_str(s: str) -> bool:
    """Heuristic: does this string look like a number?"""
    if not s:
        return False
    s = s.strip().replace(",", "").replace("$", "").replace("%", "")
    if s.startswith("(") and s.endswith(")"):
        s = s[1:-1]
    if s.startswith("-"):
        s = s[1:]
    return s.replace(".", "", 1).isdigit()


def _parse_ref(full_ref: str) -> tuple[str, int, int] | None:
    m = _CELL_REF_RE.match(full_ref)
    if not m:
        return None
    return m.group(1), column_index_from_string(m.group(2)), int(m.group(3))


def _make_ref(sheet: str, col_idx: int, row: int) -> str:
    return f"{sheet}!{get_column_letter(col_idx)}{row}"


def _classify_cells(
    all_refs: set[str],
    data_values: dict[str, object],
    formula_cells: dict[str, FunctionNode],
    bold_cells: set[str],
    sheets_bounds: dict[str, tuple[int, int, int, int]],
) -> dict[str, int]:
    """Classify all cells as header (0) or data (1) using vectorised features."""
    clf = _load_classifier()

    # Phase 1: collect raw cell info into parallel lists
    ref_list: list[str] = []
    rows: list[int] = []
    cols: list[int] = []
    n_rows_list: list[int] = []
    n_cols_list: list[int] = []
    is_numeric_list: list[bool] = []
    is_bold_list: list[bool] = []
    val_strs: list[str] = []

    formula_set = set(formula_cells)
    bold_set = bold_cells

    for ref in all_refs:
        parsed = _parse_ref(ref)
        if not parsed:
            continue
        sheet, col_idx, row_num = parsed
        bounds = sheets_bounds.get(sheet)
        if bounds is None:
            continue

        min_r, max_r, min_c, max_c = bounds

        if ref in formula_set:
            val_str = ""
            is_num = True
        else:
            val = data_values.get(ref)
            if val is None:
                continue
            is_num = isinstance(val, (int, float))
            val_str = "" if is_num else (val if isinstance(val, str) else str(val))

        ref_list.append(ref)
        rows.append(row_num - min_r)
        cols.append(col_idx - min_c)
        n_rows_list.append(max_r - min_r + 1)
        n_cols_list.append(max_c - min_c + 1)
        is_numeric_list.append(is_num)
        is_bold_list.append(ref in bold_set)
        val_strs.append(val_str)

    n = len(ref_list)
    if n == 0:
        return {}

    # Phase 2: build feature matrix with numpy (vectorised)
    r = np.array(rows, dtype=np.float32)
    c = np.array(cols, dtype=np.float32)
    nr = np.array(n_rows_list, dtype=np.float32)
    nc = np.array(n_cols_list, dtype=np.float32)
    bold = np.array(is_bold_list, dtype=np.float32)

    is_empty = np.empty(n, dtype=np.float32)
    is_numeric_feat = np.empty(n, dtype=np.float32)
    is_long = np.empty(n, dtype=np.float32)
    for i, vs in enumerate(val_strs):
        if is_numeric_list[i]:
            is_empty[i] = 0.0
            is_numeric_feat[i] = 1.0
            is_long[i] = 0.0
        else:
            is_empty[i] = float(len(vs) == 0)
            is_numeric_feat[i] = float(_is_numeric_str(vs))
            is_long[i] = float(len(vs) > 50)

    nr_denom = np.maximum(nr - 1, 1)
    nc_denom = np.maximum(nc - 1, 1)

    X = np.zeros((n, _N_FEATURES), dtype=np.float32)
    X[:, 0] = r
    X[:, 1] = c
    X[:, 2] = r / nr_denom
    X[:, 3] = c / nc_denom
    X[:, 4] = (r == 0).astype(np.float32)
    X[:, 5] = (c == 0).astype(np.float32)
    X[:, 6] = (r < 2).astype(np.float32)
    X[:, 7] = is_empty
    X[:, 8] = is_numeric_feat
    X[:, 9] = is_long
    X[:, 10] = nr
    X[:, 11] = nc
    X[:, 12] = bold

    # Phase 3: batch predict
    predictions = clf.predict(X)
    return dict(zip(ref_list, predictions))


def _build_header_indices(
    cell_roles: dict[str, int],
    data_values: dict[str, object],
    merged_map: dict[str, str],
) -> tuple[
    dict[tuple[str, int], list[tuple[int, str]]],   # col_headers: (sheet, col) → sorted [(row, text)]
    dict[tuple[str, int], list[tuple[int, str]]],    # row_headers: (sheet, row) → sorted [(col, text)]
]:
    """Pre-build sorted header indices for fast lookup.

    col_headers[(sheet, col)] = [(row1, text1), (row2, text2), ...]  sorted by row
    row_headers[(sheet, row)] = [(col1, text1), (col2, text2), ...]  sorted by col
    """
    col_headers: dict[tuple[str, int], list[tuple[int, str]]] = defaultdict(list)
    row_headers: dict[tuple[str, int], list[tuple[int, str]]] = defaultdict(list)

    for ref, role in cell_roles.items():
        if role != _HEADER:
            continue
        resolved_ref = merged_map.get(ref, ref)
        val = data_values.get(resolved_ref)
        if val is None or not isinstance(val, str) or not val.strip():
            continue
        parsed = _parse_ref(ref)
        if not parsed:
            continue
        sheet, col_idx, row_num = parsed
        text = val.strip()
        col_headers[(sheet, col_idx)].append((row_num, text))
        row_headers[(sheet, row_num)].append((col_idx, text))

    # Sort each list for binary search
    for v in col_headers.values():
        v.sort()
    for v in row_headers.values():
        v.sort()

    return col_headers, row_headers


def build_label_map(
    formula_cells: dict[str, FunctionNode],
    data_values: dict[str, object],
    merged_map: dict[str, str],
    bold_cells: set[str],
    *,
    top_k: int = _DEFAULT_TOP_K,
) -> dict[str, dict[str, object]]:
    """Detect labels for all formula cells using a trained classifier.

    Returns: {cell_ref: {"row": [label1, ...], "column": [label1, ...]}}
    Each list contains up to top_k label candidates, nearest first.
    """
    all_refs = set(formula_cells) | set(data_values)
    sheets: dict[str, tuple[int, int, int, int]] = {}

    for ref in all_refs:
        parsed = _parse_ref(ref)
        if not parsed:
            continue
        sheet, col_idx, row_num = parsed
        if sheet not in sheets:
            sheets[sheet] = (row_num, row_num, col_idx, col_idx)
        else:
            mn_r, mx_r, mn_c, mx_c = sheets[sheet]
            sheets[sheet] = (
                min(mn_r, row_num), max(mx_r, row_num),
                min(mn_c, col_idx), max(mx_c, col_idx),
            )

    cell_roles = _classify_cells(
        all_refs, data_values, formula_cells, bold_cells, sheets,
    )

    # Pre-build sorted header indices for fast scanning
    col_headers, row_headers = _build_header_indices(
        cell_roles, data_values, merged_map,
    )

    label_map: dict[str, dict[str, object]] = {}

    for cell_ref in formula_cells:
        parsed = _parse_ref(cell_ref)
        if not parsed:
            continue
        sheet, col_idx, row_num = parsed

        # Column labels: headers above in same column, nearest first
        col_labels = _find_headers_before(
            col_headers.get((sheet, col_idx), []), row_num, top_k,
        )
        # Row labels: headers to the left in same row, nearest first
        row_labels = _find_headers_before(
            row_headers.get((sheet, row_num), []), col_idx, top_k,
        )

        labels: dict[str, object] = {}
        if row_labels:
            labels["row"] = row_labels
        if col_labels:
            labels["column"] = col_labels
        if labels:
            label_map[cell_ref] = labels

    return label_map


def _find_headers_before(
    sorted_headers: list[tuple[int, str]],
    position: int,
    top_k: int,
) -> list[str]:
    """Find up to top_k unique header texts before `position`, nearest first.

    sorted_headers is a list of (pos, text) sorted by pos ascending.
    We want entries where pos < position, in reverse order (nearest first).
    """
    if not sorted_headers:
        return []

    # Binary search for insertion point of `position`
    idx = bisect.bisect_left(sorted_headers, (position,))

    results: list[str] = []
    seen: set[str] = set()
    for i in range(idx - 1, -1, -1):
        if len(results) >= top_k:
            break
        text = sorted_headers[i][1]
        if text not in seen:
            seen.add(text)
            results.append(text)

    return results
