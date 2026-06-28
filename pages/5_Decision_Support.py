import streamlit as st

from src.decision_support import generate_business_insights
from src.segmentation import build_rfm_table, segment_customers
from src.streamlit_helpers import configure_page, get_dashboard_data, show_data_quality, sidebar_filters


configure_page("Decision Support")

prepared_all, clean_sales, quality = get_dashboard_data()
filtered = sidebar_filters(clean_sales)
show_data_quality(quality)

st.title("Decision Support")
st.caption("Rule-based business insights translated from dashboard analytics.")

rfm = build_rfm_table(filtered)
segmented = None
if not rfm.empty:
    segmented = segment_customers(rfm, n_clusters=4).rfm

insights = generate_business_insights(filtered, segmented)

severity_colors = {
    "High": "[HIGH]",
    "Medium": "[MEDIUM]",
    "Positive": "[POSITIVE]",
    "Info": "[INFO]",
}

for item in insights:
    with st.container(border=True):
        st.subheader(f"{severity_colors.get(item['severity'], '[NOTE]')} {item['title']}")
        st.write(f"**Category:** {item['category']} | **Priority:** {item['severity']}")
        st.write(item["message"])
        st.info(item["recommendation"])

st.subheader("How the decision-support rules work")
st.markdown(
    """
The decision-support page uses transparent rule logic rather than a black-box
recommendation engine. This makes the system easier to explain during viva:

- Sales rules compare recent monthly revenue changes.
- Product rules identify high-revenue and slow-moving products.
- Customer rules use RFM segmentation and customer recency.
- Each insight includes a recommendation so the dashboard supports action, not only reporting.
"""
)
