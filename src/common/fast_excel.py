"""Faster ``pandas.read_excel`` for .xlsx files.

pandas' openpyxl reader builds one Python ``Cell`` object per cell, which makes
large transactional files slow: a 60,000-row, 46-column file takes ~45s.  Reading
plain values with ``iter_rows(values_only=True)`` avoids the Cell objects and is
roughly 3x faster.  The replacement below keeps pandas' exact conversion rules
(ints stay ints, numeric errors become NaN, empty cells become "", trailing empty
cells / rows are trimmed, rows are padded to equal width), so everything built on
top of it - header detection, dtype inference, column names - behaves the same.

Call ``enable_fast_excel()`` once; it is idempotent and silently does nothing if a
future pandas changes the internals it relies on.
"""

from __future__ import annotations

import hashlib
import os
import tempfile
import time
from pathlib import Path

import numpy as np
import pandas as pd

# Cached formula errors surface as these strings when cell types are not available.
_EXCEL_ERRORS = {"#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!"}

_enabled = False


def _convert(value):
    """Same result as pandas' ``OpenpyxlReader._convert_cell`` for a plain value."""
    if value is None:
        return ""
    t = type(value)
    if t is str:
        return np.nan if value in _EXCEL_ERRORS else value
    if t is int:
        return value
    if t is float:
        iv = int(value)
        return iv if iv == value else value
    return value  # bool, datetime, date, time, timedelta ... unchanged


def _fast_get_sheet_data(self, sheet, file_rows_needed=None):
    if self.book.read_only:
        sheet.reset_dimensions()

    data = []
    last_row_with_data = -1
    for row_number, row in enumerate(sheet.iter_rows(values_only=True)):
        converted_row = [_convert(v) for v in row]
        while converted_row and converted_row[-1] == "":
            converted_row.pop()  # trim trailing empty elements
        if converted_row:
            last_row_with_data = row_number
        data.append(converted_row)
        if file_rows_needed is not None and len(data) >= file_rows_needed:
            break

    data = data[: last_row_with_data + 1]  # trim trailing empty rows

    if len(data) > 0:
        max_width = max(len(r) for r in data)
        if min(len(r) for r in data) < max_width:
            empty_cell = [""]
            data = [r + (max_width - len(r)) * empty_cell for r in data]
    return data


_CACHE_DIR = Path(tempfile.gettempdir()) / "financial_automation_cache"
_CACHE_MAX_AGE_DAYS = 7


def _file_digest(path: str) -> str:
    h = hashlib.sha1()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def read_excel_cached(path, sheet_name, header):
    """``pd.read_excel`` for one sheet, remembered on disk between runs.

    The cache key is the file's *content* (SHA-1) plus sheet, header row and the
    pandas version, so an edited or different file is never served stale data and
    re-uploading the same file reuses the parse.  Anything that goes wrong with the
    cache (read-only disk, corrupt entry, in-memory file) falls back to a normal read.
    """
    if not isinstance(path, (str, os.PathLike)) or not os.path.isfile(path):
        return pd.read_excel(path, sheet_name=sheet_name, header=header)
    try:
        key = hashlib.sha1(
            f"{_file_digest(str(path))}|{sheet_name}|{header}|{pd.__version__}|{_enabled}".encode()
        ).hexdigest()
        cached = _CACHE_DIR / f"{key}.pkl"
        if cached.is_file():
            return pd.read_pickle(cached)
    except Exception:
        cached = None
    df = pd.read_excel(path, sheet_name=sheet_name, header=header)
    if cached is not None:
        try:
            _CACHE_DIR.mkdir(parents=True, exist_ok=True)
            tmp = cached.with_suffix(f".{os.getpid()}.tmp")
            df.to_pickle(tmp)
            os.replace(tmp, cached)
            cutoff = time.time() - _CACHE_MAX_AGE_DAYS * 86400
            for old in _CACHE_DIR.glob("*.pkl"):
                if old.stat().st_mtime < cutoff:
                    old.unlink(missing_ok=True)
        except Exception:
            pass
    return df


def enable_fast_excel() -> bool:
    """Install the faster sheet reader.  Returns True when it is active."""
    global _enabled
    if _enabled:
        return True
    try:
        from pandas.io.excel import _openpyxl

        reader = _openpyxl.OpenpyxlReader
        if not hasattr(reader, "get_sheet_data") or not hasattr(reader, "_convert_cell"):
            return False
        reader.get_sheet_data = _fast_get_sheet_data
        _enabled = True
    except Exception:
        return False
    return True
