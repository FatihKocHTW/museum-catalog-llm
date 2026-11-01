"""Prompt templates and the metadata block that is injected into them."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Any, Mapping

from ..config import LABELS, T_COLUMNS
from ..io_utils import strip_parens

_PROMPT_DIR = Path(__file__).parent

# Name -> template file. Templates contain a single "{meta_block}" placeholder.
TEMPLATES = {
    "curator_en": "curator_en.txt",            # full catalogue entry (generation)
    "curator_de": "curator_de_ablation.txt",   # short German description (ablation study)
}


@lru_cache(maxsize=None)
def load_template(name: str) -> str:
    try:
        filename = TEMPLATES[name]
    except KeyError:
        raise ValueError(f"Unknown prompt template '{name}'. Choose one of {sorted(TEMPLATES)}") from None
    return (_PROMPT_DIR / filename).read_text(encoding="utf-8")


def format_meta_block(meta_row: Mapping[str, Any]) -> str:
    """Render "- Label [Tn]: value" lines for every non-empty t-column, in t1..t14 order."""
    lines = []
    for col in T_COLUMNS:
        if col in meta_row:
            value = strip_parens(meta_row.get(col))
            if value:
                lines.append(f"- {LABELS.get(col, col.upper())} [{col.upper()}]: {value}")
    return "\n".join(lines) if lines else "- (keine Metadaten gefunden)"


def build_prompt(meta_row: Mapping[str, Any], template: str = "curator_en") -> str:
    # str.replace instead of str.format, so literal braces in a template never break it
    return load_template(template).replace("{meta_block}", format_meta_block(meta_row))
