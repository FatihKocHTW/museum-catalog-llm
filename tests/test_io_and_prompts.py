import math

import pytest

from museum_catalog.io_utils import (
    load_meta_rows,
    meta_for_object,
    normalize_year,
    parse_images_field,
    resolve_images,
    strip_parens,
)
from museum_catalog.prompts import build_prompt, format_meta_block, load_template


@pytest.mark.parametrize("value, expected", [
    ("AEG (Berlin)", "AEG"),
    ("Stahl (lackiert), Glas (?)", "Stahl, Glas"),
    (None, ""),
    (math.nan, ""),
    ("nan", ""),
    ("  Olympia  ", "Olympia"),
])
def test_strip_parens(value, expected):
    assert strip_parens(value) == expected


@pytest.mark.parametrize("value, expected", [("1997", "1997"), (1997.0, "1997"), ("ca. 1925-30", "1925"), ("", "")])
def test_normalize_year(value, expected):
    assert normalize_year(value) == expected


@pytest.mark.parametrize("field, expected", [
    ("['a.JPG', 'b.JPG']", ["a.JPG", "b.JPG"]),
    ("a.JPG, b.JPG", ["a.JPG", "b.JPG"]),
    (["a.JPG", " "], ["a.JPG"]),
    ("", []),
    (None, []),
    (math.nan, []),
])
def test_parse_images_field(field, expected):
    assert parse_images_field(field) == expected


def test_load_meta_rows_normalizes_columns_and_values(sample_data):
    excel_dir, _, specs = sample_data
    meta = load_meta_rows(excel_dir, specs)
    assert list(meta.columns) == ["t1", "t2", "t3"]
    row = meta_for_object(meta, "1/1997/1063 0")
    assert row["t2"] == "AEG"  # parentheses removed


def test_meta_for_unknown_object_falls_back_to_id(sample_data):
    excel_dir, _, specs = sample_data
    meta = load_meta_rows(excel_dir, specs)
    assert meta_for_object(meta, "9/9999/9999 0") == {"t1": "9/9999/9999 0"}


def test_resolve_images_uses_year_folder(sample_data):
    _, image_dir, _ = sample_data
    paths = resolve_images(image_dir, "1997", "['1-1997-1063-000-000.JPG', 'missing.JPG']")
    assert len(paths) == 1 and paths[0].endswith("1997/1-1997-1063-000-000.JPG")


def test_format_meta_block_orders_and_labels():
    block = format_meta_block({"t3": "Bakelit (schwarz)", "t1": "1/2024/0349 0", "t2": ""})
    assert block.splitlines() == [
        "- Inventar-Nr [T1]: 1/2024/0349 0",
        "- Materialien [T3]: Bakelit",
    ]


def test_format_meta_block_empty():
    assert format_meta_block({}) == "- (keine Metadaten gefunden)"


@pytest.mark.parametrize("template", ["curator_en", "curator_de"])
def test_templates_contain_placeholder_once(template):
    assert load_template(template).count("{meta_block}") == 1


def test_build_prompt_injects_metadata():
    prompt = build_prompt({"t1": "1/1997/1063 0"}, "curator_en")
    assert "{meta_block}" not in prompt
    assert prompt.rstrip().endswith("- Inventar-Nr [T1]: 1/1997/1063 0")


def test_unknown_template():
    with pytest.raises(ValueError):
        load_template("nope")
