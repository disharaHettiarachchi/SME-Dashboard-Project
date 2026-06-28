import streamlit as st

from src.forecasting import run_forecasting
from src.streamlit_helpers import configure_page, get_dashboard_data, show_data_quality, sidebar_filters
from src.visualizations import forecast_chart, revenue_trend


configure_page("Forecasting")

prepared_all, clean_sales, quality = get_dashboard_data()
filtered = sidebar_filters(clean_sales)
show_data_quality(quality)

st.title("Forecasting / Prediction")
st.caption("Simple monthly sales forecasting with model comparison metrics.")

target = st.sidebar.selectbox("Forecast target", ["Revenue", "Quantity"])
periods = st.sidebar.slider("Future months to forecast", min_value=1, max_value=6, value=3)

try:
    result = run_forecasting(filtered, target=target, periods=periods, test_size=3)
except ValueError as exc:
    st.warning(str(exc))
    st.stop()

col1, col2 = st.columns(2)
col1.metric("Best model", result.best_model)
col2.metric("Monthly observations", len(result.monthly))

st.plotly_chart(forecast_chart(result.monthly, result.future_forecast), use_container_width=True)

left, right = st.columns(2)
with left:
    st.subheader("Model comparison")
    st.dataframe(result.comparison, use_container_width=True)
with right:
    st.subheader("Future forecast")
    st.dataframe(result.future_forecast, use_container_width=True)

with st.expander("Monthly aggregated data"):
    st.dataframe(result.monthly, use_container_width=True)

st.plotly_chart(revenue_trend(result.monthly, value_col="Target"), use_container_width=True)
