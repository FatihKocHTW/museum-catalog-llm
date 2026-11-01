"""Link museum inventory data to photos and generate catalogue entries with an LLM."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path

from dotenv import load_dotenv

from . import linking, pipeline
from .config import EXCEL_SPECS, OUTPUT_DIR, ModelSettings, Paths


def _common(p: argparse.ArgumentParser, default_limit: int) -> None:
    p.add_argument("--limit", type=int, default=default_limit, help="number of objects to process")
    p.add_argument("--start-at", type=int, default=0, help="row offset in linked.csv")
    p.add_argument("--model", default=None, help="override OPENAI_MODEL")
    p.add_argument("--max-images", type=int, default=None, help="images sent per request")
    p.add_argument("--show-prompts", action="store_true", help="log the full prompt per object")


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(prog="museum_catalog", description=__doc__)
    parser.add_argument("--excel-dir", type=Path, default=Paths().excel_dir)
    parser.add_argument("--image-dir", type=Path, default=Paths().image_dir)
    parser.add_argument("-v", "--verbose", action="store_true")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("link", help="link inventory numbers to photo files -> output/linked.csv")

    gen = sub.add_parser("generate", help="create catalogue entries -> output/catalog_results.xlsx")
    _common(gen, default_limit=3)
    gen.add_argument("--template", default="curator_en")

    abl = sub.add_parser("ablation", help="text-only vs. text+images -> output/ablation_results.xlsx")
    _common(abl, default_limit=10)
    abl.add_argument("--template", default="curator_de")
    abl.add_argument("--temperature", type=float, default=0.0, help="0.0 for comparable runs")

    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    load_dotenv()
    paths = Paths(excel_dir=args.excel_dir, image_dir=args.image_dir)

    if args.command == "link":
        linking.run(paths.excel_dir, paths.image_dir, paths.linked_csv, paths.misses_csv, EXCEL_SPECS)
        return

    settings = ModelSettings()
    if args.model:
        settings.model = args.model
    if args.max_images is not None:
        settings.max_images_per_call = args.max_images
    if args.command == "ablation":
        settings.temperature = args.temperature
        if args.max_images is None:
            settings.max_images_per_call = 5

    from .llm import LLMClient  # imported here so `link` works without an API key
    client = LLMClient(settings)
    run = pipeline.generate if args.command == "generate" else pipeline.ablation
    out = OUTPUT_DIR / ("catalog_results.xlsx" if args.command == "generate" else "ablation_results.xlsx")
    run(client, paths.linked_csv, paths.excel_dir, paths.image_dir, out, EXCEL_SPECS,
        limit=args.limit, start_at=args.start_at, template=args.template, show_prompts=args.show_prompts)


if __name__ == "__main__":
    main()
