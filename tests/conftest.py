from pathlib import Path

import pandas as pd
import pytest


@pytest.fixture
def sample_data(tmp_path: Path):
    """A tiny fake collection: one Excel list and photos in year folders."""
    excel_dir = tmp_path / "Objektdaten"
    image_dir = tmp_path / "Objektbilder"
    excel_dir.mkdir()

    pd.DataFrame({
        "T1": ["1/1997/1063 0", "1/1997/1063 1", "1/2024/0349 0", "kaputt", "1/2001/0007 0"],
        "T2": ["AEG (Berlin)", "AEG", "Siemens", "", "Olympia"],
        "T3": ["Stahl, lackiert", "", "Bakelit", "", ""],
    }).to_excel(excel_dir / "Liste_Test.xlsx", index=False)

    for rel in [
        "1997/1-1997-1063-000-000.JPG",
        "1997/1-1997-1063-001-000.jpg",
        "1997/2-1997-1063-000-000.JPG",   # different group -> must not match
        "1997/1-1997-1063-000-000.txt",   # wrong extension
        "2024/1-2024-0349-000-000.JPG",
    ]:
        path = image_dir / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"\xff\xd8fake")

    specs = [("Liste_Test.xls", (0, 10))]  # .xls name on purpose: tests the .xlsx fallback
    return excel_dir, image_dir, specs
