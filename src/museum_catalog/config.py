"""Central configuration: paths, data subsets, column labels and model settings."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Tuple

# Project root = two levels above this file (src/museum_catalog/config.py)
PROJECT_ROOT = Path(__file__).resolve().parents[2]
DATA_DIR = Path(os.getenv("MUSEUM_DATA_DIR", PROJECT_ROOT / "data"))
OUTPUT_DIR = Path(os.getenv("MUSEUM_OUTPUT_DIR", PROJECT_ROOT / "output"))

# Excel inventory lists and the row ranges (0-based, end exclusive) used in the project
EXCEL_SPECS: List[Tuple[str, Tuple[int, int]]] = [
    ("Liste_AK Kommunikation.xls", (300, 600)),       # rows 301–600
    ("Liste_AEG Produktsammlung.xls", (2500, 3000)),  # rows 2501–3000
]

# Raw export columns are named t1..t14; these labels make them readable inside the prompt
LABELS: Dict[str, str] = {
    "t1": "Inventar-Nr",
    "t2": "Hersteller",
    "t3": "Materialien",
    "t4": "Maße",
    "t5": "Beschreibung (Excel)",
    "t6": "Kategorie",
    "t7": "Bezeichnung",
    "t8": "Standort / Depot",
    "t9": "Provenienz",
    "t10": "Zustand",
    "t11": "Signatur",
    "t12": "Inventarnr. (alt)",
    "t13": "Bildangabe (Excel)",
    "t14": "Datierung / Jahr",
}
T_COLUMNS: List[str] = [f"t{i}" for i in range(1, 15)]

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}


@dataclass
class ModelSettings:
    """Settings for the OpenAI chat completions call."""

    model: str = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
    url: str = "https://api.openai.com/v1/chat/completions"
    max_tokens: int = 700
    temperature: float | None = None  # None = API default
    max_images_per_call: int = 3
    max_retries: int = 3
    base_backoff_s: float = 2.0
    timeout_s: int = 120
    pause_between_objects_s: float = 2.0


@dataclass
class Paths:
    excel_dir: Path = field(default_factory=lambda: DATA_DIR / "Objektdaten")
    image_dir: Path = field(default_factory=lambda: DATA_DIR / "Objektbilder")
    linked_csv: Path = field(default_factory=lambda: OUTPUT_DIR / "linked.csv")
    misses_csv: Path = field(default_factory=lambda: OUTPUT_DIR / "linked_misses.csv")
