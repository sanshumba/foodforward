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
# Preferred deployed data source: Supabase
# Local fallback data source: cleaned anonymised Excel workbook
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

EXCEL_SHEETS = {
    "sites": "Site_Lookup",
    "bo_profile": "BO_Profile_Anonymised",
    "food_distributed": "Food_Distributed_Clean",
    "participants": "Participant_Metadata",
    "measurements": "Measurement_Long_Anonymised",
    "data_quality": "Data_Quality_Log",
}


# -----------------------------
# General helper functions
# -----------------------------

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


def safe_sum(df: pd.DataFrame, column: str) -> float:
    if df.empty or column not in df.columns:
        return 0.0
    return pd.to_numeric(df[column], errors="coerce").fillna(0).sum()


def safe_mean(df: pd.DataFrame, column: str) -> float | None:
    if df.empty or column not in df.columns:
        return None
    value = pd.to_numeric(df[column], errors="coerce").mean()
    return None if pd.isna(value) else float(value)


def safe_nunique(df: pd.DataFrame, column: str) -> int:
    if df.empty or column not in df.columns:
        return 0
    return int(df[column].nunique(dropna=True))


def metric_card(label: str, value, help_text: str | None = None):
    st.metric(label, value if value is not None else "—", help=help_text)


def format_kg(value: float) -> str:
    return f"{value:,.1f} kg"


def dataframe_download_button(df: pd.DataFrame, label: str, file_name: str, key: str):
    if df.empty:
        return
    st.download_button(
        label=label,
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=file_name,
        mime="text/csv",
        key=key,
    )


def missing_excel_message(path: str):
    st.error(
        "The Excel data file could not be found. Please check that the data folder was uploaded correctly."
    )
    st.code(path)
    st.stop()


# -----------------------------
# Data loading
# -----------------------------

@st.cache_data(show_spinner=False)
def load_from_excel(path: str) -> dict[str, pd.DataFrame]:
    """Load the cleaned anonymised Excel workbook."""
    if not Path(path).exists():
        missing_excel_message(path)

    xls = pd.ExcelFile(path)
    data = {}

    for key, sheet in EXCEL_SHEETS.items():
        if sheet in xls.sheet_names:
            data[key] = normalise_columns(pd.read_excel(path, sheet_name=sheet))
        else:
            data[key] = pd.DataFrame()

    return data


def get_supabase_client():
    if create_client is None:
        raise RuntimeError("The 'supabase' package is not installed. Install it with: pip install supabase")

    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_ANON_KEY") or os.getenv("SUPABASE_SERVICE_ROLE_KEY")

    if not url or not key:
        raise RuntimeError(
            "Missing SUPABASE_URL and SUPABASE_ANON_KEY. Add them as Render environment variables."
        )

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

    return normalise_columns(pd.DataFrame(records)) if records else pd.DataFrame()


def load_from_supabase() -> dict[str, pd.DataFrame]:
    return {key: load_table_from_supabase(table) for key, table in TABLES.items()}


def prepare_dates(data: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Convert date/month fields to pandas datetime where possible."""
    out = {k: v.copy() for k, v in data.items()}

    if not out.get("food_distributed", pd.DataFrame()).empty:
        fd = out["food_distributed"]
        for col in ["posting_date", "posting_month"]:
            if col in fd.columns:
                fd[col] = pd.to_datetime(fd[col], errors="coerce")
        out["food_distributed"] = fd

    if not out.get("measurements", pd.DataFrame()).empty:
        ms = out["measurements"]
        for col in ["measurement_month", "birth_month"]:
            if col in ms.columns:
                ms[col] = pd.to_datetime(ms[col], errors="coerce")
        out["measurements"] = ms

    if not out.get("participants", pd.DataFrame()).empty:
        pt = out["participants"]
        if "birth_month" in pt.columns:
            pt["birth_month"] = pd.to_datetime(pt["birth_month"], errors="coerce")
        out["participants"] = pt

    if not out.get("data_quality", pd.DataFrame()).empty:
        dq = out["data_quality"]
        if "measurement_month" in dq.columns:
            dq["measurement_month"] = pd.to_datetime(dq["measurement_month"], errors="coerce")
        out["data_quality"] = dq

    return out


def get_data_source_from_env() -> str:
    """Use Supabase by default for deployment, with Excel as a local option."""
    source = os.getenv("DATA_SOURCE", "supabase").strip().lower()
    return "excel" if source in {"excel", "workbook", "xlsx"} else "supabase"


def load_dashboard_data(source: str, excel_path: str) -> tuple[dict[str, pd.DataFrame], str]:
    """Load the selected data source and return data plus the source actually used."""
    if source == "supabase":
        try:
            data = load_from_supabase()
            return prepare_dates(data), "Supabase"
        except Exception as exc:
            st.warning(
                "Supabase could not be loaded. The dashboard has fallen back to the Excel workbook. "
                "Check SUPABASE_URL and SUPABASE_ANON_KEY in Render if you want to use Supabase."
            )
            with st.expander("Show Supabase error"):
                st.exception(exc)
            data = load_from_excel(excel_path)
            return prepare_dates(data), "Excel workbook fallback"

    data = load_from_excel(excel_path)
    return prepare_dates(data), "Excel workbook"


# -----------------------------
# Filtering
# -----------------------------

def build_site_label_map(data: dict[str, pd.DataFrame]) -> dict[str, str]:
    sites = data.get("sites", pd.DataFrame())
    if sites.empty or "site_id" not in sites.columns:
        return {}

    labels = {}
    for _, row in sites.iterrows():
        site_id = row.get("site_id")
        if pd.isna(site_id):
            continue
        name = row.get("bo_name") if "bo_name" in sites.columns else None
        labels[str(site_id)] = f"{site_id} — {name}" if pd.notna(name) else str(site_id)
    return labels


def get_site_options(data: dict[str, pd.DataFrame]) -> list[str]:
    site_ids = set()
    for key in ["sites", "bo_profile", "participants", "food_distributed", "measurements"]:
        df = data.get(key, pd.DataFrame())
        if not df.empty and "site_id" in df.columns:
            site_ids.update(df["site_id"].dropna().astype(str).unique().tolist())
    return sorted(site_ids)


def apply_site_filters(data: dict[str, pd.DataFrame], allowed_sites: list[str]) -> dict[str, pd.DataFrame]:
    if not allowed_sites:
        return {key: df.copy() for key, df in data.items()}

    allowed = set(allowed_sites)
    out = {}
    for key, df in data.items():
        if not df.empty and "site_id" in df.columns:
            out[key] = df[df["site_id"].astype(str).isin(allowed)].copy()
        else:
            out[key] = df.copy()
    return out


def apply_category_filter(data: dict[str, pd.DataFrame], selected_categories: list[str]) -> dict[str, pd.DataFrame]:
    out = {key: df.copy() for key, df in data.items()}
    food = out.get("food_distributed", pd.DataFrame())
    if selected_categories and not food.empty and "category" in food.columns:
        out["food_distributed"] = food[food["category"].fillna("Unknown").isin(selected_categories)].copy()
    return out


def apply_date_filters(
    data: dict[str, pd.DataFrame],
    food_range: tuple[pd.Timestamp, pd.Timestamp] | None,
    measurement_range: tuple[pd.Timestamp, pd.Timestamp] | None,
) -> dict[str, pd.DataFrame]:
    out = {key: df.copy() for key, df in data.items()}

    if food_range is not None:
        food = out.get("food_distributed", pd.DataFrame())
        if not food.empty and "posting_month" in food.columns:
            start, end = food_range
            out["food_distributed"] = food[
                food["posting_month"].between(start, end, inclusive="both")
            ].copy()

    if measurement_range is not None:
        start, end = measurement_range
        measurements = out.get("measurements", pd.DataFrame())
        if not measurements.empty and "measurement_month" in measurements.columns:
            out["measurements"] = measurements[
                measurements["measurement_month"].between(start, end, inclusive="both")
            ].copy()

        dq = out.get("data_quality", pd.DataFrame())
        if not dq.empty and "measurement_month" in dq.columns:
            out["data_quality"] = dq[dq["measurement_month"].between(start, end, inclusive="both")].copy()

    return out


def date_range_widget(df: pd.DataFrame, column: str, label: str, key: str):
    if df.empty or column not in df.columns:
        return None
    valid_dates = df[column].dropna()
    if valid_dates.empty:
        return None

    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()
    selected = st.date_input(label, value=(min_date, max_date), min_value=min_date, max_value=max_date, key=key)

    if isinstance(selected, tuple) and len(selected) == 2:
        start = pd.to_datetime(selected[0])
        end = pd.to_datetime(selected[1])
        return start, end

    return None


# -----------------------------
# Dashboard sections
# -----------------------------

def render_overview(data: dict[str, pd.DataFrame]):
    sites = data["sites"]
    profile = data["bo_profile"]
    food = data["food_distributed"]
    participants = data["participants"]
    measurements = data["measurements"]
    dq = data["data_quality"]

    st.subheader("Programme overview")

    flagged_measurements = 0
    if not measurements.empty and "data_quality_flag" in measurements.columns:
        flagged_measurements = int(measurements["data_quality_flag"].notna().sum())
    flagged_pct = (flagged_measurements / len(measurements) * 100) if len(measurements) else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Sites", safe_nunique(sites, "site_id"))
    with c2:
        metric_card("Unique child participants", safe_nunique(participants, "participant_id"))
    with c3:
        metric_card("Total food distributed", format_kg(safe_sum(food, "line_weight")))
    with c4:
        metric_card("Flagged measurement rate", f"{flagged_pct:.1f}%")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        metric_card("Distribution records", f"{len(food):,}")
    with c6:
        metric_card("Measurement records", f"{len(measurements):,}")
    with c7:
        per_child = safe_sum(food, "line_weight") / safe_nunique(participants, "participant_id") if safe_nunique(participants, "participant_id") else 0
        metric_card("Average kg per child", f"{per_child:,.1f} kg")
    with c8:
        metric_card("Logged data-quality issues", f"{len(dq):,}")

    st.info(
        "This overview combines the food distribution, participant, measurement and data-quality sheets. "
        "Use the filters in the sidebar to narrow the dashboard by site, province, BO size, category and date range."
    )

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
            if not chart_df.empty:
                top = chart_df.iloc[0]
                st.caption(f"The site with the most unique child participants is {top['site_id']} with {int(top['participants'])} participants.")
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
            if not issue_df.empty:
                st.caption(f"The most common logged issue is: {issue_df.iloc[0]['issue']}.")
        else:
            st.info("No data quality log available.")

    st.markdown("#### Quick downloads")
    d1, d2, d3 = st.columns(3)
    with d1:
        dataframe_download_button(food, "Download filtered distribution data", "filtered_food_distribution.csv", "download_overview_food")
    with d2:
        dataframe_download_button(measurements, "Download filtered measurement data", "filtered_measurements.csv", "download_overview_measurements")
    with d3:
        dataframe_download_button(dq, "Download data-quality log", "filtered_data_quality_log.csv", "download_overview_dq")


def render_distribution(data: dict[str, pd.DataFrame]):
    food = data["food_distributed"]

    st.subheader("Food distribution analysis")

    if food.empty:
        st.info("No food distribution records are available for the selected filters.")
        return

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Total quantity", f"{safe_sum(food, 'quantity'):,.0f}")
    with c2:
        metric_card("Total weight", format_kg(safe_sum(food, "line_weight")))
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
            if not monthly.empty:
                top_month = monthly.loc[monthly["line_weight"].idxmax()]
                st.caption(
                    f"The highest distribution month in the selected data is {top_month['posting_month'].strftime('%B %Y')} "
                    f"with {top_month['line_weight']:,.1f} kg."
                )
        else:
            st.info("Monthly distribution fields are not available.")

    with right:
        st.markdown("#### Distribution by category")
        if {"category", "line_weight"}.issubset(food.columns):
            cat = (
                food.assign(category=food["category"].fillna("Unknown"))
                .groupby("category", as_index=False)["line_weight"]
                .sum()
                .sort_values("line_weight", ascending=False)
                .head(12)
            )
            fig = px.bar(cat, x="category", y="line_weight", text_auto=".1f")
            fig.update_layout(xaxis_title="Category", yaxis_title="Line weight (kg)")
            st.plotly_chart(fig, use_container_width=True, key="distribution_by_category")
            if not cat.empty:
                st.caption(f"The largest category by distributed weight is {cat.iloc[0]['category']}.")
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

    dataframe_download_button(food, "Download filtered food distribution data", "filtered_food_distribution.csv", "download_distribution_food")

    with st.expander("View distribution records"):
        st.dataframe(food, use_container_width=True, hide_index=True)


def render_measurements(data: dict[str, pd.DataFrame]):
    measurements = data["measurements"]

    st.subheader("Child measurement analysis")

    if measurements.empty:
        st.info("No measurement records are available for the selected filters.")
        return

    flagged_records = measurements["data_quality_flag"].notna().sum() if "data_quality_flag" in measurements.columns else 0
    flagged_percentage = flagged_records / len(measurements) * 100 if len(measurements) else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        metric_card("Participants measured", safe_nunique(measurements, "participant_id"))
    with c2:
        avg_age = safe_mean(measurements, "age_months")
        metric_card("Average age", f"{avg_age:.1f} months" if avg_age is not None else "—")
    with c3:
        avg_weight = safe_mean(measurements, "weight_kg")
        metric_card("Average weight", f"{avg_weight:.1f} kg" if avg_weight is not None else "—")
    with c4:
        avg_height = safe_mean(measurements, "height_cm")
        metric_card("Average height", f"{avg_height:.1f} cm" if avg_height is not None else "—")
    with c5:
        metric_card("Flagged records", f"{flagged_records:,} ({flagged_percentage:.1f}%)")

    st.divider()

    left, right = st.columns(2)

    with left:
        st.markdown("#### Average weight over time")
        if {"measurement_month", "weight_kg"}.issubset(measurements.columns):
            weight = measurements.dropna(subset=["measurement_month"]).copy()
            weight["weight_kg"] = pd.to_numeric(weight["weight_kg"], errors="coerce")
            weight = (
                weight.groupby("measurement_month", as_index=False)["weight_kg"]
                .mean()
                .sort_values("measurement_month")
            )
            fig = px.line(weight, x="measurement_month", y="weight_kg", markers=True)
            fig.update_layout(xaxis_title="Measurement month", yaxis_title="Average weight (kg)")
            st.plotly_chart(fig, use_container_width=True, key="measurements_average_weight_over_time")
            st.caption("This shows the average recorded weight per measurement month for the selected children/sites.")

    with right:
        st.markdown("#### Average height over time")
        if {"measurement_month", "height_cm"}.issubset(measurements.columns):
            height = measurements.dropna(subset=["measurement_month"]).copy()
            height["height_cm"] = pd.to_numeric(height["height_cm"], errors="coerce")
            height = (
                height.groupby("measurement_month", as_index=False)["height_cm"]
                .mean()
                .sort_values("measurement_month")
            )
            fig = px.line(height, x="measurement_month", y="height_cm", markers=True)
            fig.update_layout(xaxis_title="Measurement month", yaxis_title="Average height (cm)")
            st.plotly_chart(fig, use_container_width=True, key="measurements_average_height_over_time")
            st.caption("This shows the average recorded height per measurement month for the selected children/sites.")

    left2, right2 = st.columns(2)

    with left2:
        st.markdown("#### Weight distribution")
        if "weight_kg" in measurements.columns:
            tmp = measurements.copy()
            tmp["weight_kg"] = pd.to_numeric(tmp["weight_kg"], errors="coerce")
            tmp = tmp.dropna(subset=["weight_kg"])
            fig = px.histogram(tmp, x="weight_kg", nbins=20)
            fig.update_layout(xaxis_title="Weight (kg)", yaxis_title="Number of records")
            st.plotly_chart(fig, use_container_width=True, key="measurements_weight_distribution")

    with right2:
        st.markdown("#### Height distribution")
        if "height_cm" in measurements.columns:
            tmp = measurements.copy()
            tmp["height_cm"] = pd.to_numeric(tmp["height_cm"], errors="coerce")
            tmp = tmp.dropna(subset=["height_cm"])
            fig = px.histogram(tmp, x="height_cm", nbins=20)
            fig.update_layout(xaxis_title="Height (cm)", yaxis_title="Number of records")
            st.plotly_chart(fig, use_container_width=True, key="measurements_height_distribution")

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
        st.caption("This chart helps identify unusual combinations of weight and height that may need further checking.")

    dataframe_download_button(measurements, "Download filtered measurement data", "filtered_measurements.csv", "download_measurements")

    with st.expander("View measurement records"):
        st.dataframe(measurements, use_container_width=True, hide_index=True)


def render_site_profile(data: dict[str, pd.DataFrame]):
    profile = data["bo_profile"]
    sites = data["sites"]

    st.subheader("Site profile")

    if profile.empty:
        st.info("No site profile data is available for the selected filters.")
        return

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Profile records", len(profile))
    with c2:
        metric_card("Provinces", safe_nunique(profile, "province"))
    with c3:
        metric_card("Suburbs", safe_nunique(profile, "suburb"))
    with c4:
        metric_card("Total beneficiaries", f"{safe_sum(profile, 'total_beneficiaries'):,.0f}")

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

    d1, d2 = st.columns(2)
    with d1:
        dataframe_download_button(sites, "Download site lookup", "filtered_site_lookup.csv", "download_site_lookup")
    with d2:
        dataframe_download_button(profile, "Download BO profile", "filtered_bo_profile.csv", "download_bo_profile")

    with st.expander("View site lookup"):
        st.dataframe(sites, use_container_width=True, hide_index=True)

    with st.expander("View BO profile"):
        st.dataframe(profile, use_container_width=True, hide_index=True)


def render_data_quality(data: dict[str, pd.DataFrame]):
    dq = data["data_quality"]
    measurements = data["measurements"]

    st.subheader("Data quality log")

    if dq.empty:
        st.info("No data quality issues were recorded for the selected filters.")
        return

    measurement_records = len(measurements)
    logged_issue_rate = len(dq) / measurement_records * 100 if measurement_records else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Logged issues", f"{len(dq):,}")
    with c2:
        metric_card("Participants affected", safe_nunique(dq, "participant_id"))
    with c3:
        metric_card("Months affected", safe_nunique(dq, "measurement_month"))
    with c4:
        metric_card("Issues vs measurements", f"{logged_issue_rate:.1f}%")

    st.info(
        "This section shows records that were flagged during data cleaning. These should be treated as data-quality notes, "
        "not as direct clinical conclusions."
    )

    st.divider()

    if "issue" in dq.columns:
        issue_df = dq["issue"].fillna("Unspecified issue").value_counts().reset_index()
        issue_df.columns = ["issue", "count"]
        fig = px.bar(issue_df.head(15), x="count", y="issue", orientation="h", text="count")
        fig.update_layout(xaxis_title="Count", yaxis_title="Issue")
        st.plotly_chart(fig, use_container_width=True, key="data_quality_issue_counts")

    dataframe_download_button(dq, "Download filtered data-quality log", "filtered_data_quality_log.csv", "download_dq")

    with st.expander("View data-quality records"):
        st.dataframe(dq, use_container_width=True, hide_index=True)


# -----------------------------
# Main app
# -----------------------------

def main():
    st.set_page_config(
        page_title="FoodForward Mother and Child Dashboard",
        page_icon="🍲",
        layout="wide",
    )

    st.title("FoodForward Mother and Child Programme Dashboard")
    st.caption("Dashboard built from the cleaned and anonymised Project 1 dataset.")

    with st.sidebar:
        st.header("Dashboard controls")

        excel_path = str(DEFAULT_EXCEL_PATH)
        selected_source = get_data_source_from_env()

        with st.expander("Developer options", expanded=False):
            st.caption("For Render, set DATA_SOURCE=supabase in the environment variables.")
            source_label = st.selectbox(
                "Data source",
                options=["Supabase", "Excel workbook"],
                index=0 if selected_source == "supabase" else 1,
                help="This is hidden in an expander so normal users do not need to choose a source.",
            )
            selected_source = "supabase" if source_label == "Supabase" else "excel"
            excel_path = st.text_input("Excel file path", value=excel_path)
            if st.button("Clear cached data"):
                st.cache_data.clear()
                st.rerun()

        with st.spinner("Loading dashboard data..."):
            data, actual_source = load_dashboard_data(selected_source, excel_path)

        st.success(f"Data source: {actual_source}")

        st.divider()
        st.subheader("Filters")

        site_label_map = build_site_label_map(data)
        all_sites = get_site_options(data)

        profile_for_filters = data.get("bo_profile", pd.DataFrame())
        selected_provinces = []
        selected_bo_sizes = []
        selected_categories = []

        if not profile_for_filters.empty and "province" in profile_for_filters.columns:
            province_options = sorted(profile_for_filters["province"].dropna().astype(str).unique().tolist())
            selected_provinces = st.multiselect("Province", province_options, default=province_options)

        if not profile_for_filters.empty and "bo_size" in profile_for_filters.columns:
            bo_size_options = sorted(profile_for_filters["bo_size"].fillna("Unknown").astype(str).unique().tolist())
            selected_bo_sizes = st.multiselect("BO size", bo_size_options, default=bo_size_options)

        allowed_sites_from_profile = all_sites
        if not profile_for_filters.empty and "site_id" in profile_for_filters.columns:
            tmp_profile = profile_for_filters.copy()
            if selected_provinces and "province" in tmp_profile.columns:
                tmp_profile = tmp_profile[tmp_profile["province"].astype(str).isin(selected_provinces)]
            if selected_bo_sizes and "bo_size" in tmp_profile.columns:
                tmp_profile = tmp_profile[tmp_profile["bo_size"].fillna("Unknown").astype(str).isin(selected_bo_sizes)]
            allowed_sites_from_profile = sorted(tmp_profile["site_id"].dropna().astype(str).unique().tolist())

        site_options = [site for site in all_sites if site in allowed_sites_from_profile]
        selected_sites = st.multiselect(
            "Site",
            site_options,
            default=site_options,
            format_func=lambda x: site_label_map.get(str(x), str(x)),
        )

        food_for_filters = data.get("food_distributed", pd.DataFrame())
        if not food_for_filters.empty and "category" in food_for_filters.columns:
            category_options = sorted(food_for_filters["category"].fillna("Unknown").astype(str).unique().tolist())
            selected_categories = st.multiselect("Food category", category_options, default=category_options)

        food_date_range = date_range_widget(food_for_filters, "posting_month", "Food distribution date range", "food_date_filter")
        measurement_date_range = date_range_widget(
            data.get("measurements", pd.DataFrame()),
            "measurement_month",
            "Measurement date range",
            "measurement_date_filter",
        )

        filtered_data = apply_site_filters(data, selected_sites)
        filtered_data = apply_category_filter(filtered_data, selected_categories)
        filtered_data = apply_date_filters(filtered_data, food_date_range, measurement_date_range)

        st.divider()
        st.markdown("**Notes**")
        st.write("The dashboard uses anonymised participant IDs. Direct personal identifiers are not displayed.")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["Overview", "Distribution", "Measurements", "Site profile", "Data quality"]
    )

    with tab1:
        render_overview(filtered_data)
    with tab2:
        render_distribution(filtered_data)
    with tab3:
        render_measurements(filtered_data)
    with tab4:
        render_site_profile(filtered_data)
    with tab5:
        render_data_quality(filtered_data)


if __name__ == "__main__":
    main()
