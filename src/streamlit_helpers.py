"""Small Streamlit helper functions shared by dashboard pages."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from src.data_loader import load_or_prepare_data
from src.utils import filter_dataframe


def configure_page(title: str) -> None:
    """Apply a consistent Streamlit page setup."""

    st.set_page_config(
        page_title=title,
        layout="wide",
        initial_sidebar_state="expanded",
    )


@st.cache_data(show_spinner="Loading and cleaning retail dataset...")
def get_dashboard_data():
    """Load and cache dashboard data for Streamlit."""

    return load_or_prepare_data(prefer_processed=True)


def sidebar_filters(clean_sales: pd.DataFrame) -> pd.DataFrame:
    """Render common date and country filters."""

    st.sidebar.header("Filters")
    min_date = clean_sales["InvoiceDate"].min().date()
    max_date = clean_sales["InvoiceDate"].max().date()
    selected_dates = st.sidebar.date_input(
        "Invoice date range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
    )

    countries = sorted(clean_sales["Country"].dropna().unique().tolist())
    selected_countries = st.sidebar.multiselect(
        "Countries / markets",
        options=countries,
        default=[],
        help="Leave empty to include all countries.",
    )

    date_range = None
    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
        date_range = (pd.to_datetime(selected_dates[0]), pd.to_datetime(selected_dates[1]))

    return filter_dataframe(clean_sales, date_range=date_range, countries=selected_countries)


def show_data_quality(quality_summary: dict) -> None:
    """Display key cleaning facts in the sidebar."""

    with st.sidebar.expander("Data quality notes"):
        st.write(f"Total rows: {quality_summary['total_rows']:,}")
        st.write(f"Valid sales rows: {quality_summary['valid_sales_rows']:,}")
        st.write(f"Missing customer rows: {quality_summary['missing_customer_rows']:,}")
        st.write(f"Cancelled/return rows: {quality_summary['cancelled_or_return_rows']:,}")
        st.write(f"Zero or negative price rows: {quality_summary['zero_or_negative_price_rows']:,}")
