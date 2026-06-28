"""Main Streamlit entry point for the BI decision support dashboard."""

import streamlit as st

from src.data_loader import dataset_profile, load_raw_data
from src.streamlit_helpers import configure_page


configure_page("SME BI Decision Support Dashboard")

st.title("Business Intelligence Decision Support Dashboard")
st.caption("Final year research project prototype using the public UCI Online Retail dataset.")

st.markdown(
    """
This Streamlit web application demonstrates how retail transaction data can be
converted into management KPIs, customer analytics, product insights,
decision-support recommendations, and simple forecasting outputs.

The system uses the attached public UCI Online Retail dataset as a representative
SME-style retail transaction dataset. It does not use private Sri Lankan SME data.
"""
)

try:
    raw_df = load_raw_data(nrows=1000)
    profile = dataset_profile(raw_df)
    st.success("Dataset found and readable.")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Sample rows loaded", f"{profile['rows']:,}")
    col2.metric("Columns", profile["columns"])
    col3.metric("Countries in sample", profile.get("countries", 0))
    col4.metric("Products in sample", profile.get("products", 0))
except Exception as exc:
    st.error(str(exc))
    st.stop()

st.subheader("Dashboard modules")
modules = [
    ("Executive Overview", "High-level KPIs, revenue trend, top countries, and top products."),
    ("Sales Analytics", "Monthly sales, market performance, customer revenue, and cancellation handling."),
    ("Customer Analytics", "RFM analysis, customer segmentation, high-value and low-value customers."),
    ("Product Analytics", "Best-selling, slow-moving, and high-revenue product analysis."),
    ("Decision Support", "Insight cards, alerts, and recommendation-style business logic."),
    ("Forecasting", "Simple monthly sales forecasting and model comparison."),
]

for title, description in modules:
    with st.container(border=True):
        st.markdown(f"**{title}**")
        st.write(description)

st.info("Use the pages in the left sidebar to open each dashboard module.")

