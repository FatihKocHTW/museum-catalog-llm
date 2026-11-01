"""Step 1: link inventory numbers from the Excel lists to object photos.

Inventory numbers look like "1/1997/1063 0" (group/year/number series).
Photos are stored per year as "1-1997-1063-000-001.JPG". An object is linked to every
file in <image_dir>/<year>/ whose name starts with "1-1997-1063-"; the series number
is ignored so all series of an object end up in one group.
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

import pandas as pd

from .config import IMAGE_EXTENSIONS
from .io_utils import read_excel_any, resolve_excel_path

log = logging.getLogger(__name__)

OBJECT_NUMBER_RE = re.compile(r"(\d+)\s*/\s*(\d{4})\s*/\s*(\d{3,4})\s+(\d+)")
LINKED_COLUMNS = ["obj_id_raw", "year", "base_prefix", "images", "image_count", "_source_files"]


def parse_object_number(raw: object) -> Optional[dict]:
    """Parse "g1/yyyy/g3 g4" into its year and the file-name prefix "g1-yyyy-g3-"."""
    if not isinstance(raw, str):
        return None
    match = OBJECT_NUMBER_RE.search(raw)
    if not match:
        return None
    group, year, number, _series = match.groups()
    return {"year": year, "base_prefix": f"{group}-{year}-{number}-"}


def find_matching_files_in_year(image_root: str | Path, year: str, base_prefix: str) -> List[str]:
    """Non-recursive, case-insensitive prefix match inside <image_root>/<year>/."""
    year_dir = os.path.join(str(image_root), year)
    if not os.path.isdir(year_dir):
        return []
    prefix = base_prefix.lower()
    matches = []
    for name in os.listdir(year_dir):
        path = os.path.join(year_dir, name)
        if (
            os.path.isfile(path)
            and os.path.splitext(name)[1].lower() in IMAGE_EXTENSIONS
            and name.lower().startswith(prefix)
        ):
            matches.append(path)
    return sorted(matches)


def load_inventory_numbers(excel_dir: str | Path, specs: Iterable[Tuple[str, Tuple[int, int]]]) -> pd.DataFrame:
    """Load only column t1 from the configured row ranges."""
    parts = []
    for filename, (start, stop) in specs:
        path = resolve_excel_path(excel_dir, filename)
        if not path.exists():
            log.warning("Missing Excel file: %s", path)
            continue
        df = read_excel_any(path)
        if df.empty:
            continue
        t1 = {str(c).lower(): c for c in df.columns}.get("t1")
        if not t1:
            log.warning("%s has no column t1 - skipped", path.name)
            continue
        slim = df.iloc[start:stop][[t1]].rename(columns={t1: "obj_id_raw"})
        slim["_source_file"] = path.name
        parts.append(slim)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["obj_id_raw", "_source_file"])


def build_links(meta: pd.DataFrame, image_root: str | Path) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Return (linked, misses). One row per unique (year, prefix) object group."""
    if meta.empty:
        return pd.DataFrame(columns=LINKED_COLUMNS), pd.DataFrame()

    sources: dict = {}
    for raw, source in zip(meta["obj_id_raw"], meta["_source_file"], strict=True):
        parsed = parse_object_number(raw)
        if parsed:
            key = (parsed["year"], parsed["base_prefix"].lower())
            sources.setdefault(key, set()).add(source)

    ids = meta["obj_id_raw"].dropna().map(lambda x: x.strip() if isinstance(x, str) else x).drop_duplicates()

    rows, misses, seen = [], [], set()
    for raw in ids:
        parsed = parse_object_number(raw)
        if not parsed:
            misses.append({"obj_id_raw": raw, "reason": "no object number in format 'g1/yyyy/g3 g4'"})
            continue
        year, prefix = parsed["year"], parsed["base_prefix"]
        key = (year, prefix.lower())
        if key in seen:
            continue
        seen.add(key)

        files = find_matching_files_in_year(image_root, year, prefix)
        if not files:
            misses.append({"obj_id_raw": raw, "year": year, "base_prefix": prefix,
                           "reason": f"no files in {year}/ starting with '{prefix}'"})
            continue

        rows.append({
            "obj_id_raw": raw,
            "year": year,
            "base_prefix": prefix,
            "images": [os.path.basename(p) for p in files],
            "image_count": len(files),
            "_source_files": ", ".join(sorted(sources.get(key, set()))),
        })

    return pd.DataFrame(rows, columns=LINKED_COLUMNS), pd.DataFrame(misses)


def run(excel_dir: Path, image_dir: Path, out_csv: Path, misses_csv: Path,
        specs: Iterable[Tuple[str, Tuple[int, int]]]) -> pd.DataFrame:
    meta = load_inventory_numbers(excel_dir, specs)
    linked, misses = build_links(meta, image_dir)

    out_csv.parent.mkdir(parents=True, exist_ok=True)
    linked.to_csv(out_csv, index=False, encoding="utf-8")
    if not misses.empty:
        misses.to_csv(misses_csv, index=False, encoding="utf-8")

    log.info("Excel rows: %d | object groups with images: %d | misses: %d",
             len(meta), len(linked), len(misses))
    log.info("Wrote %s", out_csv)
    return linked
