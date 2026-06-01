import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from dotenv import load_dotenv

try:
    from supabase import create_client
except Exception:
    create_client = None


# ------------------------------------------------------------
# FoodForward Mother and Child Programme Dashboard
# ------------------------------------------------------------
# This dashboard can read from either:
# 1. The cleaned anonymised Excel workbook, or
# 2. Supabase tables populated using load_to_supabase.py
# ------------------------------------------------------------

load_dotenv()

DEFAULT_EXCEL_PATH = Path("data/foodforward_project1_cleaned_anonymised_dataset.xlsx")

TABLES = {
    "sites": "ff_sites",
    "bo_profile": "ff_bo_profile",
    "food_distributed": "ff_food_distributed",
    "participants": "ff_participants",
    "measurements": "ff_measurements",
    "data_quality": "ff_data_quality_log",
}


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy of the DataFrame with snake_case column names."""
    out = df.copy()
    out.columns = (
        out.columns.astype(str)
        .str.strip()
        .str.replace(r"[^0-9a-zA-Z]+", "_", regex=True)
        .str.strip("_")
        .str.lower()
    )
    return out


@st.cache_data(show_spinner=False)
def load_from_excel(path: str) -> dict[str, pd.DataFrame]:
    """Load the cleaned anonymised Excel workbook."""
    xls = pd.ExcelFile(path)
    data = {}

    sheet_map = {
        "sites": "Site_Lookup",
        "bo_profile": "BO_Profile_Anonymised",
        "food_distributed": "Food_Distributed_Clean",
        "participants": "Participant_Metadata",
        "measurements": "Measurement_Long_Anonymised",
        "data_quality": "Data_Quality_Log",
    }

    for key, sheet in sheet_map.items():
        if sheet in xls.sheet_names:
            data[key] = normalise_columns(pd.read_excel(path, sheet_name=sheet))
        else:
            data[key] = pd.DataFrame()

    return data


def get_supabase_client():
    if create_client is None:
        st.error("The 'supabase' package is not installed. Install it with: pip install supabase")
        st.stop()

    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_ANON_KEY") or os.getenv("SUPABASE_SERVICE_ROLE_KEY")

    if not url or not key:
        st.error("Missing SUPABASE_URL and SUPABASE_ANON_KEY or SUPABASE_SERVICE_ROLE_KEY in your .env file.")
        st.stop()

    return create_client(url, key)


@st.cache_data(show_spinner=False)
def load_table_from_supabase(table_name: str) -> pd.DataFrame:
    """Read one full table from Supabase in batches."""
    client = get_supabase_client()
    page_size = 1000
    start = 0
    records = []

    while True:
        response = client.table(table_name).select("*").range(start, start + page_size - 1).execute()
        batch = response.data or []
        records.extend(batch)

        if len(batch) < page_size:
            break

        start += page_size

    return pd.DataFrame(records)


def load_from_supabase() -> dict[str, pd.DataFrame]:
    return {key: load_table_from_supabase(table) for key, table in TABLES.items()}


def prepare_dates(data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Convert date/month fields to pandas datetime where possible."""
    out = {k: v.copy() for k, v in data.items()}

    if not out["food_distributed"].empty:
        fd = out["food_distributed"]
        if "posting_date" in fd.columns:
            fd["posting_date"] = pd.to_datetime(fd["posting_date"], errors="coerce")
        if "posting_month" in fd.columns:
            fd["posting_month"] = pd.to_datetime(fd["posting_month"], errors="coerce")
        out["food_distributed"] = fd

    if not out["measurements"].empty:
        ms = out["measurements"]
        if "measurement_month" in ms.columns:
            ms["measurement_month"] = pd.to_datetime(ms["measurement_month"], errors="coerce")
        if "birth_month" in ms.columns:
            ms["birth_month"] = pd.to_datetime(ms["birth_month"], errors="coerce")
        out["measurements"] = ms

    if not out["participants"].empty:
        pt = out["participants"]
        if "birth_month" in pt.columns:
            pt["birth_month"] = pd.to_datetime(pt["birth_month"], errors="coerce")
        out["participants"] = pt

    return out


def metric_card(label: str, value, help_text: str | None = None):
    st.metric(label, value if value is not None else "—", help=help_text)


def safe_sum(df: pd.DataFrame, column: str) -> float:
    if df.empty or column not in df.columns:
        return 0.0
    return pd.to_numeric(df[column], errors="coerce").fillna(0).sum()


def safe_nunique(df: pd.DataFrame, column: str) -> int:
    if df.empty or column not in df.columns:
        return 0
    return df[column].nunique(dropna=True)


def apply_filters(data: dict[str, pd.DataFrame], selected_sites: list[str]) -> dict[str, pd.DataFrame]:
    if not selected_sites:
        return data

    out = {}
    for key, df in data.items():
        if not df.empty and "site_id" in df.columns:
            out[key] = df[df["site_id"].isin(selected_sites)].copy()
        else:
            out[key] = df.copy()
    return out


def render_overview(data: dict[str, pd.DataFrame]):
    sites = data["sites"]
    food = data["food_distributed"]
    participants = data["participants"]
    measurements = data["measurements"]
    dq = data["data_quality"]

    st.subheader("Programme overview")

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric_card("Sites", safe_nunique(sites, "site_id"))
    with c2:
        metric_card("Participants", safe_nunique(participants, "participant_id"))
    with c3:
        metric_card("Distribution lines", f"{len(food):,}")
    with c4:
        metric_card("Total line weight", f"{safe_sum(food, 'line_weight'):,.1f} kg")
    with c5:
        metric_card("Measurement records", f"{len(measurements):,}")

    st.divider()

    left, right = st.columns(2)

    with left:
        st.markdown("#### Participants by site")
        if not participants.empty and {"site_id", "participant_id"}.issubset(participants.columns):
            chart_df = (
                participants.groupby("site_id", as_index=False)["participant_id"]
                .nunique()
                .rename(columns={"participant_id": "participants"})
                .sort_values("participants", ascending=False)
            )
            fig = px.bar(chart_df, x="site_id", y="participants", text="participants")
            fig.update_layout(xaxis_title="Site", yaxis_title="Participants")
            st.plotly_chart(fig, use_container_width=True, key="overview_participants_by_site")
        else:
            st.info("Participant/site data is not available.")

    with right:
        st.markdown("#### Data quality issues")
        if not dq.empty and "issue" in dq.columns:
            issue_df = (
                dq["issue"].fillna("Unspecified issue")
                .value_counts()
                .head(10)
                .reset_index()
            )
            issue_df.columns = ["issue", "count"]
            fig = px.bar(issue_df, x="count", y="issue", orientation="h", text="count")
            fig.update_layout(xaxis_title="Count", yaxis_title="Issue")
            st.plotly_chart(fig, use_container_width=True, key="overview_data_quality_issues")
        else:
            st.info("No data quality log available.")


def render_distribution(data: dict[str, pd.DataFrame]):
    food = data["food_distributed"]

    st.subheader("Food distribution analysis")

    if food.empty:
        st.info("No food distribution records are available.")
        return

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Total quantity", f"{safe_sum(food, 'quantity'):,.0f}")
    with c2:
        metric_card("Total weight", f"{safe_sum(food, 'line_weight'):,.1f} kg")
    with c3:
        metric_card("Unique items", safe_nunique(food, "item_code"))
    with c4:
        metric_card("Categories", safe_nunique(food, "category"))

    st.divider()

    left, right = st.columns(2)

    with left:
        st.markdown("#### Monthly distribution weight")
        if {"posting_month", "line_weight"}.issubset(food.columns):
            monthly = (
                food.dropna(subset=["posting_month"])
                .groupby("posting_month", as_index=False)["line_weight"]
                .sum()
                .sort_values("posting_month")
            )
            fig = px.line(monthly, x="posting_month", y="line_weight", markers=True)
            fig.update_layout(xaxis_title="Month", yaxis_title="Line weight (kg)")
            st.plotly_chart(fig, use_container_width=True, key="distribution_monthly_weight")
        else:
            st.info("Monthly distribution fields are not available.")

    with right:
        st.markdown("#### Distribution by category")
        if {"category", "line_weight"}.issubset(food.columns):
            cat = (
                food.groupby("category", as_index=False)["line_weight"]
                .sum()
                .sort_values("line_weight", ascending=False)
                .head(12)
            )
            fig = px.bar(cat, x="category", y="line_weight", text_auto=".1f")
            fig.update_layout(xaxis_title="Category", yaxis_title="Line weight (kg)")
            st.plotly_chart(fig, use_container_width=True, key="distribution_by_category")
        else:
            st.info("Category and line-weight fields are not available.")

    st.markdown("#### Top distributed items by weight")
    if {"description", "line_weight", "quantity"}.issubset(food.columns):
        top_items = (
            food.groupby("description", as_index=False)
            .agg(total_weight_kg=("line_weight", "sum"), total_quantity=("quantity", "sum"))
            .sort_values("total_weight_kg", ascending=False)
            .head(15)
        )
        st.dataframe(top_items, use_container_width=True, hide_index=True)

    st.markdown("#### Distribution records")
    st.dataframe(food, use_container_width=True, hide_index=True)


def render_measurements(data: dict[str, pd.DataFrame]):
    measurements = data["measurements"]

    st.subheader("Child measurement analysis")

    if measurements.empty:
        st.info("No measurement records are available.")
        return

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Participants measured", safe_nunique(measurements, "participant_id"))
    with c2:
        avg_weight = pd.to_numeric(measurements.get("weight_kg"), errors="coerce").mean() if "weight_kg" in measurements.columns else None
        metric_card("Average weight", f"{avg_weight:.1f} kg" if pd.notna(avg_weight) else "—")
    with c3:
        avg_height = pd.to_numeric(measurements.get("height_cm"), errors="coerce").mean() if "height_cm" in measurements.columns else None
        metric_card("Average height", f"{avg_height:.1f} cm" if pd.notna(avg_height) else "—")
    with c4:
        metric_card("Flagged records", measurements["data_quality_flag"].notna().sum() if "data_quality_flag" in measurements.columns else 0)

    st.divider()

    left, right = st.columns(2)

    with left:
        st.markdown("#### Average weight over time")
        if {"measurement_month", "weight_kg"}.issubset(measurements.columns):
            weight = (
                measurements.dropna(subset=["measurement_month"])
                .assign(weight_kg=pd.to_numeric(measurements["weight_kg"], errors="coerce"))
                .groupby("measurement_month", as_index=False)["weight_kg"]
                .mean()
                .sort_values("measurement_month")
            )
            fig = px.line(weight, x="measurement_month", y="weight_kg", markers=True)
            fig.update_layout(xaxis_title="Measurement month", yaxis_title="Average weight (kg)")
            st.plotly_chart(fig, use_container_width=True, key="measurements_average_weight_over_time")

    with right:
        st.markdown("#### Average height over time")
        if {"measurement_month", "height_cm"}.issubset(measurements.columns):
            height = (
                measurements.dropna(subset=["measurement_month"])
                .assign(height_cm=pd.to_numeric(measurements["height_cm"], errors="coerce"))
                .groupby("measurement_month", as_index=False)["height_cm"]
                .mean()
                .sort_values("measurement_month")
            )
            fig = px.line(height, x="measurement_month", y="height_cm", markers=True)
            fig.update_layout(xaxis_title="Measurement month", yaxis_title="Average height (cm)")
            st.plotly_chart(fig, use_container_width=True, key="measurements_average_height_over_time")

    st.markdown("#### Weight and height relationship")
    if {"weight_kg", "height_cm", "age_months", "site_id"}.issubset(measurements.columns):
        scatter_df = measurements.copy()
        scatter_df["weight_kg"] = pd.to_numeric(scatter_df["weight_kg"], errors="coerce")
        scatter_df["height_cm"] = pd.to_numeric(scatter_df["height_cm"], errors="coerce")
        scatter_df = scatter_df.dropna(subset=["weight_kg", "height_cm"])
        fig = px.scatter(
            scatter_df,
            x="height_cm",
            y="weight_kg",
            color="site_id",
            hover_data=["participant_id", "measurement_month", "age_months"],
        )
        fig.update_layout(xaxis_title="Height (cm)", yaxis_title="Weight (kg)")
        st.plotly_chart(fig, use_container_width=True, key="measurements_weight_height_relationship")

    st.markdown("#### Measurement records")
    st.dataframe(measurements, use_container_width=True, hide_index=True)


def render_site_profile(data: dict[str, pd.DataFrame]):
    profile = data["bo_profile"]
    sites = data["sites"]

    st.subheader("Site profile")

    if profile.empty:
        st.info("No site profile data is available.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("Profile records", len(profile))
    with c2:
        metric_card("Provinces", safe_nunique(profile, "province"))
    with c3:
        metric_card("Suburbs", safe_nunique(profile, "suburb"))

    st.divider()

    left, right = st.columns(2)

    with left:
        st.markdown("#### Total beneficiaries by site")
        if {"site_id", "total_beneficiaries"}.issubset(profile.columns):
            ben = profile[["site_id", "total_beneficiaries"]].copy()
            ben["total_beneficiaries"] = pd.to_numeric(ben["total_beneficiaries"], errors="coerce")
            ben = ben.sort_values("total_beneficiaries", ascending=False)
            fig = px.bar(ben, x="site_id", y="total_beneficiaries", text_auto=".0f")
            fig.update_layout(xaxis_title="Site", yaxis_title="Total beneficiaries")
            st.plotly_chart(fig, use_container_width=True, key="site_total_beneficiaries_by_site")

    with right:
        st.markdown("#### BO size")
        if "bo_size" in profile.columns:
            size_df = profile["bo_size"].fillna("Unknown").value_counts().reset_index()
            size_df.columns = ["bo_size", "count"]
            fig = px.pie(size_df, names="bo_size", values="count")
            st.plotly_chart(fig, use_container_width=True, key="site_bo_size_pie")

    st.markdown("#### Site lookup")
    st.dataframe(sites, use_container_width=True, hide_index=True)

    st.markdown("#### BO profile")
    st.dataframe(profile, use_container_width=True, hide_index=True)


def render_data_quality(data: dict[str, pd.DataFrame]):
    dq = data["data_quality"]

    st.subheader("Data quality log")

    if dq.empty:
        st.info("No data quality issues were recorded.")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        metric_card("Logged issues", len(dq))
    with c2:
        metric_card("Participants affected", safe_nunique(dq, "participant_id"))
    with c3:
        metric_card("Months affected", safe_nunique(dq, "measurement_month"))

    st.divider()

    if "issue" in dq.columns:
        issue_df = dq["issue"].fillna("Unspecified issue").value_counts().reset_index()
        issue_df.columns = ["issue", "count"]
        fig = px.bar(issue_df.head(15), x="count", y="issue", orientation="h", text="count")
        fig.update_layout(xaxis_title="Count", yaxis_title="Issue")
        st.plotly_chart(fig, use_container_width=True, key="data_quality_issue_counts")

    st.dataframe(dq, use_container_width=True, hide_index=True)


def main():
    st.set_page_config(
        page_title="FoodForward Mother and Child Dashboard",
        page_icon="🍲",
        layout="wide",
    )

    st.title("FoodForward Mother and Child Programme Dashboard")
    st.caption("Dashboard built from the cleaned and anonymised Project 1 dataset.")

    with st.sidebar:
        st.header("Data source")
        source = st.radio("Choose data source", ["Excel workbook", "Supabase"], index=0)

        if source == "Excel workbook":
            excel_path = st.text_input("Excel file path", value=str(DEFAULT_EXCEL_PATH))
            data = load_from_excel(excel_path)
        else:
            st.info("Make sure your .env file contains the Supabase credentials.")
            data = load_from_supabase()

        data = prepare_dates(data)

        site_options = []
        for key in ["sites", "participants", "food_distributed", "measurements"]:
            if key in data and not data[key].empty and "site_id" in data[key].columns:
                site_options = sorted(data[key]["site_id"].dropna().unique().tolist())
                if site_options:
                    break

        selected_sites = st.multiselect("Filter by site", site_options, default=site_options)
        data = apply_filters(data, selected_sites)

        st.divider()
        st.markdown("**Notes**")
        st.write("The dashboard uses anonymised participant IDs. Direct personal identifiers are not displayed.")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["Overview", "Distribution", "Measurements", "Site profile", "Data quality"]
    )

    with tab1:
        render_overview(data)
    with tab2:
        render_distribution(data)
    with tab3:
        render_measurements(data)
    with tab4:
        render_site_profile(data)
    with tab5:
        render_data_quality(data)


if __name__ == "__main__":
    main()
