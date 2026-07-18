from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_PATH = ROOT / "data" / "FoodForward_Project2_Analysis.xlsx"


def test_project2_workbook_exists_and_contains_expected_sheets():
    assert DATA_PATH.exists(), f"Expected workbook at {DATA_PATH}"

    workbook = pd.ExcelFile(DATA_PATH)
    expected_sheets = {
        "Category Analysis",
        "Monthly Profiles",
        "Site Profiles",
        "Basket Quality",
        "Gap Analysis",
        "Candidate Products",
        "Scenario Testing",
        "Methodology",
    }

    missing = expected_sheets.difference(set(workbook.sheet_names))
    assert not missing, f"Missing expected sheets: {sorted(missing)}"


def test_project2_summary_metrics_are_readable():
    category = pd.read_excel(DATA_PATH, sheet_name="Category Analysis", header=2)
    monthly = pd.read_excel(DATA_PATH, sheet_name="Monthly Profiles", header=2)
    site = pd.read_excel(DATA_PATH, sheet_name="Site Profiles", header=2)

    assert not category.empty
    assert not monthly.empty
    assert not site.empty

    assert monthly["Total kg"].sum() > 0
    assert site["Site"].nunique() > 0
    assert category["Share of basket"].sum() > 0.9
