import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv
from supabase import create_client


# ------------------------------------------------------------
# Upload FoodForward cleaned anonymised data to Supabase
# ------------------------------------------------------------
# Before running this script:
# 1. Create a Supabase project.
# 2. Run the SQL in supabase_schema.sql in the Supabase SQL editor.
# 3. Create a .env file using .env.example as a guide.
# 4. Run:
#       python load_to_supabase.py
# ------------------------------------------------------------

load_dotenv()

EXCEL_PATH = Path("data/foodforward_project1_cleaned_anonymised_dataset.xlsx")

TABLE_MAP = {
    "Site_Lookup": "ff_sites",
    "BO_Profile_Anonymised": "ff_bo_profile",
    "Food_Distributed_Clean": "ff_food_distributed",
    "Participant_Metadata": "ff_participants",
    "Measurement_Long_Anonymised": "ff_measurements",
    "Data_Quality_Log": "ff_data_quality_log",
}


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = (
        out.columns.astype(str)
        .str.strip()
        .str.replace(r"[^0-9a-zA-Z]+", "_", regex=True)
        .str.strip("_")
        .str.lower()
    )
    return out


def clean_for_json(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare a DataFrame for Supabase JSON insertion."""
    out = df.copy()

    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d")

    # Convert pandas NaN/NaT values to None for JSON compatibility.
    out = out.where(pd.notnull(out), None)

    # Supabase JSON serialisation does not like Timestamp objects.
    for col in out.columns:
        out[col] = out[col].map(
            lambda x: x.isoformat() if hasattr(x, "isoformat") else x
        )

    return out


def chunk_records(records: list[dict], size: int = 500):
    for i in range(0, len(records), size):
        yield records[i : i + size]


def upload_table(client, table_name: str, df: pd.DataFrame, truncate_first: bool = True):
    if df.empty:
        print(f"Skipping {table_name}: no records.")
        return

    if truncate_first:
        # This requires suitable permissions. If it fails, manually clear the table in Supabase.
        try:
            client.table(table_name).delete().neq("id", -1).execute()
            print(f"Cleared existing records in {table_name}.")
        except Exception as exc:
            print(f"Could not clear {table_name}. Continuing with insert. Reason: {exc}")

    records = clean_for_json(df).to_dict(orient="records")

    for batch in chunk_records(records):
        client.table(table_name).insert(batch).execute()

    print(f"Uploaded {len(records)} records to {table_name}.")


def main():
    url = os.getenv("SUPABASE_URL")
    service_key = os.getenv("SUPABASE_SERVICE_ROLE_KEY")

    if not url or not service_key:
        raise RuntimeError(
            "Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY. "
            "Create a .env file first."
        )

    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Could not find {EXCEL_PATH}")

    client = create_client(url, service_key)

    for sheet_name, table_name in TABLE_MAP.items():
        print(f"\nReading sheet: {sheet_name}")
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
        df = normalise_columns(df)
        upload_table(client, table_name, df)

    print("\nDone. The cleaned anonymised dataset has been uploaded to Supabase.")


if __name__ == "__main__":
    main()
