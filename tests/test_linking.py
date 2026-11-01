import pytest

from museum_catalog.linking import build_links, find_matching_files_in_year, load_inventory_numbers, parse_object_number


@pytest.mark.parametrize("raw, expected", [
    ("1/1997/1063 0", {"year": "1997", "base_prefix": "1-1997-1063-"}),
    ("1 / 2024 / 0349   2", {"year": "2024", "base_prefix": "1-2024-0349-"}),
    ("Inv. 3/1985/123 0", {"year": "1985", "base_prefix": "3-1985-123-"}),
])
def test_parse_object_number(raw, expected):
    assert parse_object_number(raw) == expected


@pytest.mark.parametrize("raw", ["kaputt", "1/97/1063 0", "1/1997/1063", None, 42])
def test_parse_object_number_rejects_invalid(raw):
    assert parse_object_number(raw) is None


def test_find_matching_files_filters_prefix_and_extension(sample_data):
    _, image_dir, _ = sample_data
    files = find_matching_files_in_year(image_dir, "1997", "1-1997-1063-")
    names = [p.rsplit("/", 1)[-1] for p in files]
    assert names == ["1-1997-1063-000-000.JPG", "1-1997-1063-001-000.jpg"]


def test_find_matching_files_missing_year(sample_data):
    _, image_dir, _ = sample_data
    assert find_matching_files_in_year(image_dir, "1850", "1-1850-0001-") == []


def test_build_links_groups_series_and_reports_misses(sample_data):
    excel_dir, image_dir, specs = sample_data
    meta = load_inventory_numbers(excel_dir, specs)
    linked, misses = build_links(meta, image_dir)

    # "1/1997/1063 0" and "... 1" are the same object -> one group with both photos
    assert linked["base_prefix"].tolist() == ["1-1997-1063-", "1-2024-0349-"]
    assert linked["image_count"].tolist() == [2, 1]
    assert linked.iloc[0]["_source_files"] == "Liste_Test.xlsx"

    reasons = dict(zip(misses["obj_id_raw"], misses["reason"], strict=True))
    assert "format" in reasons["kaputt"]
    assert "no files" in reasons["1/2001/0007 0"]


def test_build_links_empty_input():
    import pandas as pd
    linked, misses = build_links(pd.DataFrame(), "/nowhere")
    assert linked.empty and misses.empty
