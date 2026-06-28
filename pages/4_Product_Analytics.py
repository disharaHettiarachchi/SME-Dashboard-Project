import streamlit as st

from src.kpi_calculator import product_performance, slow_moving_products, top_products
from src.streamlit_helpers import configure_page, get_dashboard_data, show_data_quality, sidebar_filters
from src.utils import format_currency, format_number
from src.visualizations import product_bar


configure_page("Product Analytics")

prepared_all, clean_sales, quality = get_dashboard_data()
filtered = sidebar_filters(clean_sales)
show_data_quality(quality)

st.title("Product Analytics")
st.caption("Best-selling products, slow-moving products, high-revenue items, and product recommendations.")

products = product_performance(filtered)
best_revenue = top_products(filtered, top_n=15)
best_quantity = products.sort_values("Quantity", ascending=False).head(15)
slow = slow_moving_products(filtered, top_n=20)

col1, col2, col3 = st.columns(3)
col1.metric("Products analysed", format_number(products["StockCode"].nunique()))
col2.metric("Highest product revenue", format_currency(products["Revenue"].max()))
col3.metric("Highest product quantity", format_number(products["Quantity"].max()))

left, right = st.columns(2)
with left:
    st.plotly_chart(product_bar(best_revenue, value_col="Revenue", title="High-Revenue Products"), use_container_width=True)
with right:
    st.plotly_chart(product_bar(best_quantity, value_col="Quantity", title="Best-Selling Products by Quantity"), use_container_width=True)

st.subheader("Product-level recommendation logic")
st.markdown(
    """
- High revenue and high quantity: protect stock availability and prioritize promotion.
- High revenue but lower quantity: monitor premium or high-value product opportunities.
- Low revenue and low quantity: review pricing, bundling, promotion, or discontinuation.
"""
)

st.subheader("Slow-moving product candidates")
st.dataframe(slow, use_container_width=True)

with st.expander("Full product performance table"):
    st.dataframe(products, use_container_width=True)

