"""Step 2 and 3: generate catalogue entries, and compare text-only vs. text+image input."""

from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Iterable, Tuple

import pandas as pd

from .io_utils import load_meta_rows, meta_for_object, normalize_year, resolve_images, safe_write_excel
from .llm import LLMClient
from .prompts import build_prompt

log = logging.getLogger(__name__)


def _load_inputs(linked_csv: Path, excel_dir: Path, specs: Iterable[Tuple[str, Tuple[int, int]]]):
    if not linked_csv.exists():
        raise FileNotFoundError(f"{linked_csv} not found - run the 'link' step first.")
    return pd.read_csv(linked_csv), load_meta_rows(excel_dir, specs)


def generate(client: LLMClient, linked_csv: Path, excel_dir: Path, image_dir: Path, out_xlsx: Path,
             specs, limit: int = 3, start_at: int = 0, template: str = "curator_en",
             show_prompts: bool = False) -> Path:
    """Create one catalogue entry per object (metadata + photos)."""
    linked, meta_df = _load_inputs(linked_csv, excel_dir, specs)
    results = []

    for _, row in linked.iloc[start_at:start_at + limit].iterrows():
        obj_id = str(row.get("obj_id_raw", "")).strip()
        prompt = build_prompt(meta_for_object(meta_df, obj_id), template)
        images = resolve_images(image_dir, row.get("year", ""), row.get("images", ""))
        if show_prompts:
            log.info("--- prompt for %s ---\n%s", obj_id, prompt)

        log.info("%s: %d images -> requesting description", obj_id, len(images))
        try:
            description = client.describe(prompt, images)
        except Exception as exc:  # noqa: BLE001 - keep the batch running, record the error
            log.error("%s failed: %s", obj_id, exc)
            description = f"ERROR: {exc}"

        results.append({"objekt_id": obj_id, "description": description})
        time.sleep(client.settings.pause_between_objects_s)

    saved = safe_write_excel(pd.DataFrame(results, columns=["objekt_id", "description"]), out_xlsx)
    log.info("Saved %d entries to %s", len(results), saved)
    return saved


def ablation(client: LLMClient, linked_csv: Path, excel_dir: Path, image_dir: Path, out_xlsx: Path,
             specs, limit: int = 10, start_at: int = 0, template: str = "curator_de",
             show_prompts: bool = False) -> Path:
    """Same prompt twice per object: once without and once with photos, for side-by-side review."""
    linked, meta_df = _load_inputs(linked_csv, excel_dir, specs)
    results = []

    for _, row in linked.iloc[start_at:start_at + limit].iterrows():
        obj_id = str(row.get("obj_id_raw", "")).strip()
        year = row.get("year", "")
        prompt = build_prompt(meta_for_object(meta_df, obj_id), template)
        images = resolve_images(image_dir, year, row.get("images", ""))
        if show_prompts:
            log.info("--- prompt for %s ---\n%s", obj_id, prompt)

        errors = []
        outputs = {}
        for variant, imgs in (("text_only", []), ("text_plus_images", images)):
            log.info("%s: %s (%d images available)", obj_id, variant, len(images))
            try:
                outputs[variant] = client.describe(prompt, imgs)
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{variant}: {exc}")
                outputs[variant] = f"ERROR: {exc}"
            time.sleep(client.settings.pause_between_objects_s)

        results.append({
            "objekt_id": obj_id,
            "year": normalize_year(year),
            "image_count_total": len(images),
            "image_count_sent": min(len(images), client.settings.max_images_per_call),
            "prompt_text": prompt,
            "description_text_only": outputs["text_only"],
            "description_text_plus_images": outputs["text_plus_images"],
            "errors": "; ".join(errors),
        })

    saved = safe_write_excel(pd.DataFrame(results), out_xlsx)
    log.info("Saved %d comparisons to %s", len(results), saved)
    return saved
