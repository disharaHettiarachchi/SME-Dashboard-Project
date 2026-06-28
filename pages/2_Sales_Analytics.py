import streamlit as st

from src.kpi_calculator import cancellation_summary, customer_value_table, monthly_sales, sales_by_country, top_products
from src.streamlit_helpers import configure_page, get_dashboard_data, show_data_quality, sidebar_filters
from src.utils import format_currency, format_number, format_percent
from src.visualizations import country_bar, product_bar, quantity_histogram, revenue_trend


configure_page("Sales Analytics")

prepared_all, clean_sales, quality = get_dashboard_data()
filtered = sidebar_filters(clean_sales)
show_data_quality(quality)

st.title("Sales Analytics")
st.caption("Explore revenue, orders, customer contribution, quantity movement, and cancelled transactions.")

monthly_df = monthly_sales(filtered)
country_df = sales_by_country(filtered, top_n=15)
product_df = top_products(filtered, top_n=15)
customer_df = customer_value_table(filtered).head(20)
cancel_summary = cancellation_summary(prepared_all)

col1, col2, col3, col4 = st.columns(4)
col1.metric("Cancelled / return rows", format_number(cancel_summary["cancelled_rows"]))
col2.metric("Cancelled invoices", format_number(cancel_summary["cancelled_invoices"]))
col3.metric("Cancelled value", format_currency(cancel_summary["cancelled_value"]))
col4.metric("Share of rows", format_percent(cancel_summary["cancelled_share_rows"]))

st.plotly_chart(revenue_trend(monthly_df), use_container_width=True)

left, right = st.columns(2)
with left:
    st.plotly_chart(country_bar(country_df), use_container_width=True)
with right:
    st.plotly_chart(product_bar(product_df, value_col="Quantity", title="Top Products by Quantity Sold"), use_container_width=True)

left, right = st.columns(2)
with left:
    st.subheader("Top customers by revenue")
    st.dataframe(customer_df, use_container_width=True)
with right:
    st.plotly_chart(quantity_histogram(filtered), use_container_width=True)

