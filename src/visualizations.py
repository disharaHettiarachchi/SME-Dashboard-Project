"""Plotly visualization helpers used across dashboard pages."""

from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

from src.utils import short_text


COLOR_SEQUENCE = ["#111111", "#FF6B35", "#7F7F7F", "#B8BCC4", "#444444"]


def revenue_trend(monthly_df: pd.DataFrame, value_col: str = "Revenue") -> go.Figure:
    fig = px.line(monthly_df, x="Month", y=value_col, markers=True, title=f"Monthly {value_col} Trend")
    fig.update_traces(line_color="#FF6B35", line_width=3)
    fig.update_layout(template="plotly_white", hovermode="x unified")
    return fig


def country_bar(country_df: pd.DataFrame) -> go.Figure:
    fig = px.bar(
        country_df.sort_values("Revenue"),
        x="Revenue",
        y="Country",
        orientation="h",
        title="Top Countries / Markets by Revenue",
        color_discrete_sequence=["#111111"],
    )
    fig.update_layout(template="plotly_white", yaxis_title="", xaxis_title="Revenue")
    return fig


def product_bar(product_df: pd.DataFrame, value_col: str = "Revenue", title: str = "Top Products") -> go.Figure:
    data = product_df.copy()
    data["Product"] = data["Description"].apply(short_text)
    fig = px.bar(
        data.sort_values(value_col),
        x=value_col,
        y="Product",
        orientation="h",
        title=title,
        color_discrete_sequence=["#111111"],
    )
    fig.update_layout(template="plotly_white", yaxis_title="", xaxis_title=value_col)
    return fig


def customer_scatter(customer_df: pd.DataFrame) -> go.Figure:
    fig = px.scatter(
        customer_df,
        x="Orders",
        y="Revenue",
        size="Quantity",
        hover_name="CustomerID",
        title="Customer Revenue vs Purchase Frequency",
        color_discrete_sequence=["#FF6B35"],
    )
    fig.update_layout(template="plotly_white")
    return fig


def rfm_segment_scatter(rfm_df: pd.DataFrame) -> go.Figure:
    fig = px.scatter(
        rfm_df,
        x="Frequency",
        y="Monetary",
        color="SegmentName",
        size="RFMScore",
        hover_name="CustomerID",
        title="Customer Segments by Frequency and Monetary Value",
        color_discrete_sequence=COLOR_SEQUENCE,
    )
    fig.update_layout(template="plotly_white")
    return fig


def segment_bar(summary_df: pd.DataFrame) -> go.Figure:
    fig = px.bar(
        summary_df.sort_values("TotalRevenue"),
        x="TotalRevenue",
        y="SegmentName",
        orientation="h",
        title="Customer Segment Revenue Contribution",
        color="Customers",
        color_continuous_scale=["#EDEDED", "#FF6B35"],
    )
    fig.update_layout(template="plotly_white", yaxis_title="", xaxis_title="Total Revenue")
    return fig


def forecast_chart(monthly_df: pd.DataFrame, future_df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=monthly_df["Month"],
            y=monthly_df["Target"],
            mode="lines+markers",
            name="Actual",
            line=dict(color="#111111", width=3),
        )
    )
    fig.add_trace(
        go.Scatter(
            x=future_df["Month"],
            y=future_df["Forecast"],
            mode="lines+markers",
            name="Forecast",
            line=dict(color="#FF6B35", width=3, dash="dash"),
        )
    )
    fig.update_layout(template="plotly_white", title="Actual vs Forecasted Monthly Sales", hovermode="x unified")
    return fig


def quantity_histogram(clean_sales: pd.DataFrame) -> go.Figure:
    capped = clean_sales[clean_sales["Quantity"] <= clean_sales["Quantity"].quantile(0.99)]
    fig = px.histogram(capped, x="Quantity", nbins=50, title="Quantity Sold Distribution")
    fig.update_traces(marker_color="#111111")
    fig.update_layout(template="plotly_white")
    return fig

