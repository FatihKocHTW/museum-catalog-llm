"""File and data helpers shared by linking, generation and the ablation study."""

from __future__ import annotations

import ast
import base64
import datetime
import glob
import logging
import os
import re
from pathlib import Path
from typing import Any, Iterable, List, Optional, Tuple

import pandas as pd

from .config import T_COLUMNS

log = logging.getLogger(__name__)

_PARENS_RE = re.compile(r"\s*\([^)]*\)")
_YEAR_RE = re.compile(r"\b(\d{4})\b")


def strip_parens(value: Any) -> str:
    """Remove "(...)" annotations and turn None/NaN into an empty string."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    return _PARENS_RE.sub("", text).strip()


def normalize_year(value: Any) -> str:
    """Extract a four-digit year from any value, or return ""."""
    match = _YEAR_RE.search("" if value is None else str(value))
    return match.group(1) if match else ""


def read_excel_any(path: str | Path) -> pd.DataFrame:
    """Read .xls or .xlsx; return an empty frame (and log) if the file is unreadable."""
    path = Path(path)
    ext = path.suffix.lower()
    try:
        if ext == ".xlsx":
            return pd.read_excel(path, engine="openpyxl")
        if ext == ".xls":
            try:
                return pd.read_excel(path)
            except Exception:
                return pd.read_excel(path, engine="xlrd")
        raise ValueError(f"Unsupported extension: {ext}")
    except Exception as exc:  # noqa: BLE001 - we want to keep going on bad files
        log.warning("Could not read %s (%s)", path.name, exc)
        return pd.DataFrame()


def resolve_excel_path(excel_dir: str | Path, filename: str) -> Path:
    """Return the .xls path, falling back to an .xlsx file with the same stem."""
    path = Path(excel_dir) / filename
    if not path.exists():
        alt = path.with_suffix(".xlsx")
        if alt.exists():
            return alt
    return path


def load_meta_rows(excel_dir: str | Path, specs: Iterable[Tuple[str, Tuple[int, int]]]) -> pd.DataFrame:
    """Load the configured row ranges, keep only t1..t14 (lower-cased) and clean the values."""
    wanted = set(T_COLUMNS)
    parts: List[pd.DataFrame] = []

    for filename, (start, stop) in specs:
        df = read_excel_any(resolve_excel_path(excel_dir, filename))
        if df.empty:
            continue

        lower_map = {str(c).lower(): c for c in df.columns}
        keys = [k for k in T_COLUMNS if k in wanted and k in lower_map]
        if not keys:
            continue

        sub = df.iloc[start:stop][[lower_map[k] for k in keys]].copy()
        sub = sub.rename(columns={lower_map[k]: k for k in keys})
        for col in sub.columns:
            sub[col] = sub[col].map(strip_parens)
        parts.append(sub)

    if not parts:
        return pd.DataFrame(columns=["t1"])

    all_cols = sorted(set().union(*(p.columns for p in parts)), key=lambda c: int(c[1:]))
    return pd.concat([p.reindex(columns=all_cols) for p in parts], ignore_index=True)


def meta_for_object(meta_df: pd.DataFrame, obj_id: str) -> dict:
    """Look up the metadata row for an inventory number (column t1)."""
    if meta_df.empty or "t1" not in meta_df.columns:
        return {"t1": obj_id}
    hit = meta_df[meta_df["t1"].astype(str).str.strip() == str(obj_id).strip()]
    return hit.iloc[0].to_dict() if not hit.empty else {"t1": obj_id}


def parse_images_field(field: Any) -> List[str]:
    """Accept a real list, a stringified list "['a', 'b']" or a comma-separated string."""
    if field is None or (isinstance(field, float) and pd.isna(field)):
        return []
    if isinstance(field, list):
        return [str(x).strip() for x in field if str(x).strip()]
    text = str(field).strip()
    if not text:
        return []
    if text.startswith("[") and text.endswith("]"):
        try:
            value = ast.literal_eval(text)
            if isinstance(value, list):
                return [str(x).strip() for x in value if str(x).strip()]
        except (ValueError, SyntaxError):
            pass
    parts = [t.strip().strip("'\"[]") for t in text.split(",")]
    return [p for p in parts if p]


def find_file_anywhere(root: str | Path, year: Any, filename: str) -> Optional[str]:
    """Look in <root>/<year>/, then <root>/, then recursively."""
    root = str(root)
    year = normalize_year(year)
    if year:
        candidate = os.path.join(root, year, filename)
        if os.path.exists(candidate):
            return candidate
    candidate = os.path.join(root, filename)
    if os.path.exists(candidate):
        return candidate
    hits = glob.glob(os.path.join(root, "**", filename), recursive=True)
    return hits[0] if hits else None


def resolve_images(image_dir: str | Path, year: Any, images_field: Any) -> List[str]:
    """Turn the image names stored in linked.csv into existing file paths."""
    paths: List[str] = []
    for name in parse_images_field(images_field):
        if os.path.isabs(name) and os.path.exists(name):
            paths.append(name)
            continue
        found = find_file_anywhere(image_dir, year, name)
        if found:
            paths.append(found)
    return paths


def encode_image(path: str | Path) -> str:
    with open(path, "rb") as fh:
        return base64.b64encode(fh.read()).decode("utf-8")


def safe_write_excel(df: pd.DataFrame, path: str | Path) -> Path:
    """Write to Excel; if the file is locked (open in Excel/OneDrive), use a timestamped name."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_excel(path, index=False, engine="openpyxl")
        return path
    except PermissionError:
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = path.with_name(f"{path.stem}_{stamp}{path.suffix}")
        df.to_excel(alt, index=False, engine="openpyxl")
        log.warning("%s was locked - saved as %s", path.name, alt.name)
        return alt
