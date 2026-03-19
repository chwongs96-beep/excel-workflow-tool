"""
Shared utility functions for Excel/CSV file handling.

Centralises helpers that were previously duplicated across
excel_nodes.py and merge_nodes.py.
"""

import re
import pandas as pd
from pathlib import Path
from typing import Any, List, Tuple

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE


# ---------------------------------------------------------------------------
# Value normalisation
# ---------------------------------------------------------------------------

def sanitize_value(value: Any) -> Any:
    """Remove illegal XML characters from string values (required by openpyxl)."""
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub('', value)
    return value


def normalize_excel_value(value: Any) -> Any:
    """Normalize a value for Excel writing.

    * Numeric-looking text → int / float
    * Leading-zero identifiers (e.g. "00123") → kept as text
    * Formulas (starts with "=") → kept as text
    * Everything else → unchanged
    """
    value = sanitize_value(value)

    if not isinstance(value, str):
        return value

    text = value.strip()
    if text == "":
        return ""

    # Keep formulas
    if text.startswith("="):
        return text

    # Keep identifiers with leading zeros (e.g. "00123")
    if re.fullmatch(r"[+-]?0\d+", text):
        return text

    # Remove thousand separators for numeric parsing
    compact = text.replace(",", "")

    # Integer
    if re.fullmatch(r"[+-]?\d+", compact):
        try:
            return int(compact)
        except (ValueError, OverflowError):
            return text

    # Decimal / scientific notation
    if re.fullmatch(r"[+-]?(\d+\.\d*|\d*\.\d+|\d+)([eE][+-]?\d+)?", compact):
        try:
            return float(compact)
        except (ValueError, OverflowError):
            return text

    return text


def normalize_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Apply :func:`normalize_excel_value` to every cell in *df*."""
    normalized = df.copy()
    for col in normalized.columns:
        normalized[col] = normalized[col].map(normalize_excel_value)
    return normalized


# ---------------------------------------------------------------------------
# Sheet-name helpers
# ---------------------------------------------------------------------------

def sanitize_sheet_name(name: str) -> str:
    """Sanitize a sheet name so Excel accepts it (max 31 chars, no special chars)."""
    name = str(name)
    for ch in ('\\', '/', '?', '*', ':', '[', ']'):
        name = name.replace(ch, '_')
    return name[:31]


def make_unique_sheet_name(name: str, used_names=None) -> str:
    """Return an Excel-safe sheet name that is not in *used_names*."""
    if used_names is None:
        used_names = set()

    base_name = sanitize_sheet_name(name) or "Sheet"
    candidate = base_name
    counter = 1
    while candidate in used_names:
        suffix = f"_{counter}"
        candidate = f"{base_name[:31 - len(suffix)]}{suffix}"
        counter += 1
    return candidate


# ---------------------------------------------------------------------------
# File reading helpers
# ---------------------------------------------------------------------------

def read_excel_with_engine(file_path, sheet_name: Any = 0, header: int = 0, dtype=None):
    """Read an Excel file using the correct engine based on extension.

    Parameters
    ----------
    dtype : optional
        Passed to :func:`pd.read_excel`.  Use ``object`` to preserve all
        values as strings (prevents leading-zero loss, etc.).
    """
    ext = str(file_path).lower()
    engine = 'xlrd' if ext.endswith('.xls') else 'openpyxl'
    kwargs: dict[str, Any] = dict(
        sheet_name=sheet_name, header=header, engine=engine,
    )
    if dtype is not None:
        kwargs['dtype'] = dtype
    return pd.read_excel(file_path, **kwargs)


def get_xls_merged_cells(file_path, sheet_name=None) -> List[Tuple[int, int, int, int]]:
    """Return merged-cell ranges from an ``.xls`` file via *xlrd*.

    Each range is ``(row_start, row_end, col_start, col_end)`` using
    **0-indexed, exclusive-end** convention (same as xlrd).

    Returns an empty list when merged-cell info cannot be read.
    """
    try:
        import xlrd
    except ImportError:
        return []
    try:
        wb = xlrd.open_workbook(str(file_path), formatting_info=True)
    except Exception:
        try:
            wb = xlrd.open_workbook(str(file_path))
        except Exception:
            return []

    try:
        if sheet_name:
            ws = wb.sheet_by_name(sheet_name)
        else:
            ws = wb.sheet_by_index(0)
        return list(ws.merged_cells)
    except Exception:
        return []


def get_excel_file(file_path) -> pd.ExcelFile:
    """Create a :class:`pd.ExcelFile` with the correct engine for the extension."""
    ext = str(file_path).lower()
    engine = 'xlrd' if ext.endswith('.xls') else 'openpyxl'
    return pd.ExcelFile(file_path, engine=engine)


def convert_text_to_numeric(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    """Heuristically convert text columns that are mostly numeric.

    Returns the (possibly modified) DataFrame and a list of converted column names.
    """
    converted_cols: List[str] = []
    for col in df.columns:
        if df[col].dtype != 'object':
            continue
        try:
            numeric_series = pd.to_numeric(df[col], errors='coerce')
            non_null_count = df[col].notna().sum()
            if non_null_count > 0:
                converted_count = numeric_series.notna().sum()
                if converted_count / non_null_count > 0.5:
                    df[col] = numeric_series
                    converted_cols.append(col)
        except Exception:
            pass
    return df, converted_cols


def read_csv_with_options(
    file_path,
    encoding_opt: str = 'auto',
    delimiter_opt: str = 'auto',
    header_row: int = 0,
    *,
    dtype=None,
    auto_convert_numeric: bool = True,
):
    """Read a CSV with flexible encoding/delimiter options.

    Parameters
    ----------
    dtype : optional
        Passed to :func:`pd.read_csv`.  Use ``object`` to preserve leading
        zeros; leave *None* for standard dtype inference.
    auto_convert_numeric : bool
        When *True* (and *dtype* is not ``object``), run
        :func:`convert_text_to_numeric` after reading.
    """
    # Determine delimiter / engine
    sep = None
    engine = 'python'

    if delimiter_opt == 'comma':
        sep, engine = ',', 'c'
    elif delimiter_opt == 'tab':
        sep, engine = '\t', 'c'
    elif delimiter_opt == 'semicolon':
        sep, engine = ';', 'c'
    elif delimiter_opt == 'pipe':
        sep, engine = '|', 'python'
    elif delimiter_opt == 'space':
        sep, engine = r'\s+', 'python'

    # Determine encoding list
    if encoding_opt and encoding_opt != 'auto':
        encodings = [encoding_opt]
    else:
        encodings = ['utf-8', 'gbk', 'utf-8-sig', 'gb18030', 'latin1']

    last_error: Exception | None = None
    for enc in encodings:
        try:
            # First pass: count total lines so we can detect dropped rows
            bad_line_count = 0

            def _on_bad_line(bad_line):
                nonlocal bad_line_count
                bad_line_count += 1
                return None

            read_kwargs: dict[str, Any] = {
                'sep': sep,
                'encoding': enc,
                'header': header_row,
                'engine': 'python',  # callable on_bad_lines requires python engine
                'keep_default_na': True,
                'skipinitialspace': True,
                'on_bad_lines': _on_bad_line,
            }
            if dtype is not None:
                read_kwargs['dtype'] = dtype

            df = pd.read_csv(file_path, **read_kwargs)

            if bad_line_count > 0:
                print(f"⚠️ CSV读取: 跳过了 {bad_line_count} 行格式异常的数据 ({Path(file_path).name})")

            if auto_convert_numeric and dtype is None:
                df, converted_cols = convert_text_to_numeric(df)
                if converted_cols:
                    print(f"已转换文本数字为数值类型: {', '.join(converted_cols)}")

            if enc != 'utf-8':
                print(f"成功使用 {enc} 编码读取CSV: {Path(file_path).name}")

            return df
        except Exception as e:
            last_error = e
            continue

    raise ValueError(f"无法读取CSV文件 (尝试了编码: {encodings}): {last_error}")


def read_tabular_file(
    file_path: str,
    sheet_name: Any = 0,
    header_row: int = 0,
    csv_encoding: str = "auto",
    csv_delimiter: str = "auto",
):
    """Read CSV / XLS / XLSX using the correct engine and options.

    Always reads with ``dtype=object`` so leading-zero strings are preserved.
    """
    lower_path = str(file_path).lower()

    if lower_path.endswith('.csv'):
        return read_csv_with_options(
            file_path,
            encoding_opt=csv_encoding,
            delimiter_opt=csv_delimiter,
            header_row=header_row,
            dtype=object,
            auto_convert_numeric=False,
        )

    engine = 'xlrd' if lower_path.endswith('.xls') else 'openpyxl'
    return pd.read_excel(
        file_path, sheet_name=sheet_name, header=header_row,
        engine=engine, dtype=object,
    )
