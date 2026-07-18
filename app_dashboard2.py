import json
import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


ROOT = Path(__file__).resolve().parent


def resolve_data_source() -> Path | None:
    candidates = []

    env_path = os.getenv("PROJECT2_WORKBOOK_PATH")
    if env_path:
        candidates.append(Path(env_path))

    candidates.extend(
        [
            ROOT / "data" / "FoodForward_Project2_Analysis.xlsx",
            ROOT / "data" / "project2_analysis.json",
            Path.cwd() / "data" / "FoodForward_Project2_Analysis.xlsx",
            Path.cwd() / "data" / "project2_analysis.json",
            Path("/opt/render/project/src/data/FoodForward_Project2_Analysis.xlsx"),
            Path("/opt/render/project/src/data/project2_analysis.json"),
            Path("/app/data/FoodForward_Project2_Analysis.xlsx"),
            Path("/app/data/project2_analysis.json"),
        ]
    )

    for candidate in candidates:
        if candidate.exists():
            return candidate

    return None


def load_project2_data_from_json(path: Path) -> dict[str, pd.DataFrame]:
    with path.open("r", encoding="utf-8") as handle:
        payload = json.load(handle)

    return {name: pd.DataFrame(records) for name, records in payload.items()}


@st.cache_data(show_spinner=False)
def load_project2_data():
    data_source = resolve_data_source()
    if data_source is None:
        st.error("The Project 2 workbook and fallback data file could not be found.")
        st.stop()

    if data_source.suffix.lower() == ".json":
        return load_project2_data_from_json(data_source)

    return {
        "category_analysis": pd.read_excel(data_source, sheet_name="Category Analysis", header=2),
        "monthly_profiles": pd.read_excel(data_source, sheet_name="Monthly Profiles", header=2),
        "site_profiles": pd.read_excel(data_source, sheet_name="Site Profiles", header=2),
        "basket_quality": pd.read_excel(data_source, sheet_name="Basket Quality", header=2),
        "gap_analysis": pd.read_excel(data_source, sheet_name="Gap Analysis", header=2),
        "candidate_products": pd.read_excel(data_source, sheet_name="Candidate Products", header=2),
        "scenario_testing": pd.read_excel(data_source, sheet_name="Scenario Testing", header=2),
        "methodology": pd.read_excel(data_source, sheet_name="Methodology", header=2),
    }


def format_kg(value: float) -> str:
    return f"{value:,.2f} kg"


def format_pct(value) -> str:
    if pd.isna(value):
        return "—"
    return f"{float(value):,.1%}"


def format_currency(value) -> str:
    return f"R {float(value):,.2f}"


def render_page_title(title: str, subtitle: str) -> None:
    st.markdown(f"## {title}")
    st.caption(subtitle)


def page_executive_summary(data: dict) -> None:
    render_page_title(
        "Executive Summary",
        "Observed findings from the Project 2 basket-composition analysis only.",
    )

    category_analysis = data["category_analysis"].copy()
    monthly_profiles = data["monthly_profiles"].copy()
    site_profiles = data["site_profiles"].copy()
    basket_quality = data["basket_quality"].copy()

    total_weight = float(monthly_profiles["Total kg"].sum())
    products_analyzed = int(category_analysis["Distinct products"].sum())
    sites = int(site_profiles["Site"].nunique())
    months = int(monthly_profiles["Month"].nunique())

    category_max = category_analysis.loc[category_analysis["Share of basket"].idxmax()]
    quality_mean = float(basket_quality["Quality score"].mean())
    fruit_veg_mean = float(basket_quality["Fruit & veg %"].mean())
    plant_protein_mean = float(basket_quality["Plant protein %"].mean())

    col1, col2, col3 = st.columns(3)
    col1.metric("Total recorded food weight", format_kg(total_weight))
    col2.metric("Number of products analysed", f"{products_analyzed}")
    col3.metric("Number of sites", f"{sites}")

    col4, col5, col6 = st.columns(3)
    col4.metric("Number of months", f"{months}")
    col5.metric("Largest food category", str(category_max["Category"]))
    col6.metric("Relative basket quality", f"{quality_mean:.1f}/100")

    st.divider()
    st.markdown(
        "**Analytical finding:** the observed basket was dominated by grain-products, while fruit-and-vegetable and plant-protein shares remained below the levels typically associated with a more balanced basket."
    )
    st.markdown(
        f"Average observed quality score: {quality_mean:.1f}/100. Average fruit-and-vegetable share: {fruit_veg_mean:.1%}. Average plant-protein share: {plant_protein_mean:.1%}."
    )


def page_basket_composition(data: dict) -> None:
    render_page_title(
        "Basket Composition",
        "Observed composition of the food basket by category and weight.",
    )

    category_analysis = data["category_analysis"].copy()
    category_analysis = category_analysis.sort_values("Share of basket", ascending=False)

    left_col, right_col = st.columns(2)

    with left_col:
        fig_share = px.bar(
            category_analysis,
            x="Category",
            y="Share of basket",
            text=[f"{value:.1%}" for value in category_analysis["Share of basket"]],
            color="Share of basket",
            color_continuous_scale="Viridis",
        )
        fig_share.update_layout(
            xaxis_title="Food category",
            yaxis_title="Share of basket",
            showlegend=False,
            template="seaborn",
        )
        st.plotly_chart(fig_share, use_container_width=True)

    with right_col:
        fig_weight = px.bar(
            category_analysis,
            x="Category",
            y="Total weight (kg)",
            text=[f"{value:,.0f}" for value in category_analysis["Total weight (kg)"]],
            color="Total weight (kg)",
            color_continuous_scale="Viridis",
        )
        fig_weight.update_layout(
            xaxis_title="Food category",
            yaxis_title="Weight (kg)",
            showlegend=False,
            template="seaborn",
        )
        st.plotly_chart(fig_weight, use_container_width=True)

    st.info(
        "Observed pattern: grain-products made up the largest share of the basket, while several nutrient-sensitive categories remained comparatively small."
    )


def page_site_comparison(data: dict) -> None:
    render_page_title(
        "Site Comparison",
        "Relative basket quality and nutrient-sensitive shares across sites.",
    )

    site_profiles = data["site_profiles"].copy()
    site_profiles = site_profiles.sort_values("Quality score", ascending=False)

    quality_fig = px.bar(
        site_profiles,
        x="Site",
        y="Quality score",
        text=[f"{value:.1f}" for value in site_profiles["Quality score"]],
        color="Quality score",
        color_continuous_scale="Viridis",
    )
    quality_fig.update_layout(xaxis_title="Site", yaxis_title="Quality score", template="seaborn")
    st.plotly_chart(quality_fig, use_container_width=True)

    share_long = site_profiles[["Site", "Fruit & veg %", "Plant protein %", "Discretionary %"]].melt(
        id_vars="Site",
        var_name="Metric",
        value_name="Share",
    )
    share_fig = px.bar(
        share_long,
        x="Site",
        y="Share",
        color="Metric",
        barmode="group",
        color_discrete_map={
            "Fruit & veg %": "#2ca25f",
            "Plant protein %": "#3182bd",
            "Discretionary %": "#de2d26",
        },
    )
    share_fig.update_layout(xaxis_title="Site", yaxis_title="Share of basket", template="seaborn")
    st.plotly_chart(share_fig, use_container_width=True)

    st.subheader("Analytical interpretation")
    site_summary = site_profiles[["Site", "Quality score", "Fruit & veg %", "Plant protein %", "Discretionary %"]].copy()
    site_summary.columns = ["Site", "Quality score", "Fruit & veg share", "Plant protein share", "Discretionary share"]
    st.dataframe(site_summary, use_container_width=True, hide_index=True)


def page_monthly_patterns(data: dict) -> None:
    render_page_title(
        "Monthly Patterns",
        "Observed monthly shifts in basket composition and basket quality.",
    )

    monthly_profiles = data["monthly_profiles"].copy()
    monthly_profiles = monthly_profiles.sort_values("Month")

    major_categories = [
        "GRAIN-PRODUCTS",
        "FRUITS-VEGETABLE",
        "PULSES-LEGUMES",
        "PROTEIN",
        "ROOTS",
        "VEGETABLE OILS",
        "ULTRA-PROCESSED",
    ]
    composition_frame = monthly_profiles[["Month"] + major_categories].copy()
    composition_fig = px.bar(
        composition_frame,
        x="Month",
        y=major_categories,
        barmode="stack",
        color_discrete_sequence=px.colors.qualitative.Set2,
    )
    composition_fig.update_layout(xaxis_title="Month", yaxis_title="Weight (kg)", template="seaborn")
    st.plotly_chart(composition_fig, use_container_width=True)

    quality_fig = px.line(
        monthly_profiles,
        x="Month",
        y="Quality score",
        markers=True,
        line_shape="spline",
    )
    quality_fig.update_layout(xaxis_title="Month", yaxis_title="Quality score", template="seaborn")
    st.plotly_chart(quality_fig, use_container_width=True)

    lowest_fruit_month = monthly_profiles.loc[monthly_profiles["Fruit & veg %"].idxmin(), "Month"]
    lowest_protein_month = monthly_profiles.loc[monthly_profiles["Plant protein %"].idxmin(), "Month"]
    st.info(
        f"Observed pattern: fruit-and-vegetable share was weakest in {lowest_fruit_month}, while plant-protein share was weakest in {lowest_protein_month}."
    )


def page_gap_analysis(data: dict) -> None:
    render_page_title(
        "Basket Gap Analysis",
        "Recurring analytical gaps visible across sites and months.",
    )

    gap_analysis = data["gap_analysis"].copy()

    gap_counts = gap_analysis["Gap"].value_counts().reset_index()
    gap_counts.columns = ["Gap", "Count"]
    gap_counts = gap_counts.sort_values("Count", ascending=False)

    fig = px.bar(gap_counts, x="Gap", y="Count", text="Count", color="Count", color_continuous_scale="Blues")
    fig.update_layout(xaxis_title="Observed gap", yaxis_title="Occurrences", template="seaborn")
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Recurring analytical findings")
    st.dataframe(
        gap_analysis[["Scope", "Period", "Gap", "Evidence", "Priority", "Suggested response"]],
        use_container_width=True,
        hide_index=True,
    )

    st.info(
        "Observed pattern: plant-protein and fruit-and-vegetable shortfalls recur most frequently, with discretionary-food shares also appearing above the analytical threshold in several comparisons."
    )


def page_external_context(data: dict) -> None:
    render_page_title(
        "External Product Context",
        "External nutrition guidance and retail benchmark prices for products relevant to the observed gaps.",
    )

    candidate_products = data["candidate_products"].copy()
    candidate_products = candidate_products.sort_values("Weighted score", ascending=False)

    scatter_fig = px.scatter(
        candidate_products,
        x="Price/kg (R)",
        y="Protein grams per rand",
        size="Gap alignment (1-5)",
        hover_name="Candidate product",
        color="Gap alignment (1-5)",
        color_continuous_scale="Viridis",
        size_max=18,
    )
    scatter_fig.update_layout(xaxis_title="Retail benchmark price per kg (R)", yaxis_title="Protein grams per rand", template="seaborn")
    st.plotly_chart(scatter_fig, use_container_width=True)

    st.subheader("Relevant products")
    context_table = candidate_products[[
        "Candidate product",
        "Price/kg (R)",
        "Protein grams per rand",
        "Gap alignment (1-5)",
        "Nutrition (1-5)",
        "Affordability (1-5)",
        "Storage (1-5)",
    ]].copy()
    context_table.columns = [
        "Product",
        "Retail benchmark price/kg (R)",
        "Protein grams per rand",
        "Gap alignment",
        "Nutrition score",
        "Affordability score",
        "Storage score",
    ]
    st.dataframe(context_table, use_container_width=True, hide_index=True)

    st.info("External market context only: these products are illustrative context for the observed analytical gaps and are not procurement instructions.")


def page_cost_scenarios(data: dict) -> None:
    render_page_title(
        "Illustrative Cost Scenarios",
        "Retail-cost examples for basket additions that could address the observed gaps.",
    )

    scenario_testing = data["scenario_testing"].copy()
    scenario_mapping = {
        "Low-cost plant protein": "Pulse-led illustration",
        "Balanced addition": "Balanced illustration",
        "Protein priority": "Protein-diversity illustration",
        "Micronutrient and diversity": "Vegetable-diversity illustration",
    }
    scenario_testing["Scenario label"] = scenario_testing["Scenario"].map(scenario_mapping)

    for _, row in scenario_testing.iterrows():
        with st.container():
            st.markdown(f"### {row['Scenario label']}")
            col1, col2, col3 = st.columns(3)
            col1.metric("Estimated cost", format_currency(row["Estimated cost (R)"]))
            col2.metric("Estimated added weight", f"{row['Estimated added kg']:.2f} kg")
            col3.metric("Estimated protein", f"{row['Estimated protein kg']:.2f} kg")
            st.caption(row["Operational interpretation"])
            st.divider()

    cost_fig = px.bar(
        scenario_testing,
        x="Scenario label",
        y="Estimated cost (R)",
        text=[f"R {value:,.0f}" for value in scenario_testing["Estimated cost (R)"]],
        color="Estimated cost (R)",
        color_continuous_scale="Viridis",
    )
    cost_fig.update_layout(xaxis_title="Illustrative scenario", yaxis_title="Estimated cost (R)", template="seaborn")
    st.plotly_chart(cost_fig, use_container_width=True)


def page_methodology(data: dict) -> None:
    render_page_title(
        "Methodology and Limitations",
        "The analytical logic behind the basket-quality view and the assumptions behind the interpretation.",
    )

    st.subheader("Basket-quality methodology")
    st.markdown(
        "- The analysis uses the Project 2 workbook outputs rather than raw transaction-level exploration."
    )
    st.markdown(
        "- Basket quality is interpreted as a relative indicator derived from observed category shares and the proportion of the basket made up of nutrient-sensitive foods."
    )
    st.markdown(
        "- Higher scores indicate a basket that appears more balanced in the context of the analytical framework, but do not imply a complete nutrition adequacy assessment."
    )

    st.subheader("Assumptions and limitations")
    st.markdown(
        "- The analysis is based on aggregated basket outputs supplied in the Project 2 workbook and should be treated as an analytical summary rather than a full dietary assessment."
    )
    st.markdown(
        "- Some sites had very low recorded distribution volume, which can make comparisons less stable."
    )
    st.markdown(
        "- The retail benchmark prices are illustrative market context and may differ from actual procurement conditions, supplier contracts, or seasonal price changes."
    )
    st.markdown(
        "- The external product context is provided to support interpretation of the observed gaps, not to define a procurement plan."
    )

    st.subheader("Retail price assumptions")
    st.markdown(
        "- Retail price values are taken from the workbook’s external product context sheet and should be treated as indicative public-market benchmarks."
    )
    st.markdown(
        "- The cost scenarios are illustrative examples only and reflect simplified combinations of products rather than full operational procurement realities."
    )


def main() -> None:
    st.set_page_config(page_title="FoodForward SA Project 2", page_icon="🥦", layout="wide")

    st.title("FoodForward SA Project 2")
    st.caption("A findings-driven dashboard of Project 2 basket-composition outputs only.")
    st.markdown(
        "This view focuses on analytical findings, observed basket patterns, and external market context. It does not repeat the general data-exploration pages from Dashboard 1."
    )

    pages = {
        "Executive Summary": page_executive_summary,
        "Basket Composition": page_basket_composition,
        "Site Comparison": page_site_comparison,
        "Monthly Patterns": page_monthly_patterns,
        "Basket Gap Analysis": page_gap_analysis,
        "External Product Context": page_external_context,
        "Illustrative Cost Scenarios": page_cost_scenarios,
        "Methodology and Limitations": page_methodology,
    }

    page_name = st.sidebar.radio("Project 2 analysis pages", list(pages.keys()))
    data = load_project2_data()
    pages[page_name](data)


if __name__ == "__main__":
    main()
