import pandas as pd

from museum_catalog import linking, pipeline
from museum_catalog.config import ModelSettings


class StubClient:
    """Records calls instead of talking to the API."""

    def __init__(self, fail_on_images=False):
        self.settings = ModelSettings(pause_between_objects_s=0, max_images_per_call=3)
        self.calls = []
        self.fail_on_images = fail_on_images

    def describe(self, prompt, image_paths=()):
        self.calls.append((prompt, list(image_paths)))
        if self.fail_on_images and image_paths:
            raise RuntimeError("boom")
        return f"text with {len(image_paths)} images"


def _link(tmp_path, sample_data):
    excel_dir, image_dir, specs = sample_data
    linked_csv = tmp_path / "out" / "linked.csv"
    linking.run(excel_dir, image_dir, linked_csv, tmp_path / "out" / "misses.csv", specs)
    return excel_dir, image_dir, specs, linked_csv


def test_generate_end_to_end(tmp_path, sample_data):
    excel_dir, image_dir, specs, linked_csv = _link(tmp_path, sample_data)
    client = StubClient()
    out = pipeline.generate(client, linked_csv, excel_dir, image_dir, tmp_path / "out" / "cat.xlsx", specs, limit=5)

    df = pd.read_excel(out)
    assert df["objekt_id"].tolist() == ["1/1997/1063 0", "1/2024/0349 0"]
    assert df["description"].tolist() == ["text with 2 images", "text with 1 images"]
    assert "Hersteller [T2]: AEG" in client.calls[0][0]


def test_ablation_sends_same_prompt_with_and_without_images(tmp_path, sample_data):
    excel_dir, image_dir, specs, linked_csv = _link(tmp_path, sample_data)
    client = StubClient(fail_on_images=True)
    out = pipeline.ablation(client, linked_csv, excel_dir, image_dir, tmp_path / "out" / "abl.xlsx", specs, limit=1)

    (p1, imgs1), (p2, imgs2) = client.calls
    assert p1 == p2 and imgs1 == [] and len(imgs2) == 2

    row = pd.read_excel(out).iloc[0]
    assert row["description_text_only"] == "text with 0 images"
    assert row["description_text_plus_images"] == "ERROR: boom"  # error recorded, batch continues
    assert row["image_count_sent"] == 2
