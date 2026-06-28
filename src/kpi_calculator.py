"""Business KPI calculations for retail decision support."""

from __future__ import annotations

import pandas as pd

from src.utils import safe_divide


def calculate_overview_kpis(clean_sales: pd.DataFrame) -> dict:
    """Calculate high-level executive KPIs."""

    total_revenue = float(clean_sales["Revenue"].sum())
    total_orders = int(clean_sales["InvoiceNo"].nunique())
    total_customers = int(clean_sales["CustomerID"].nunique(dropna=True))
    total_quantity = float(clean_sales["Quantity"].sum())

    return {
        "total_revenue": total_revenue,
        "total_orders": total_orders,
        "total_customers": total_customers,
        "average_order_value": safe_divide(total_revenue, total_orders),
        "total_quantity": total_quantity,
        "unique_products": int(clean_sales["StockCode"].nunique()),
        "countries": int(clean_sales["Country"].nunique()),
        "date_min": clean_sales["InvoiceDate"].min(),
        "date_max": clean_sales["InvoiceDate"].max(),
    }


def monthly_sales(clean_sales: pd.DataFrame) -> pd.DataFrame:
    """Aggregate sales by month."""

    grouped = (
        clean_sales.groupby("Month", as_index=False)
        .agg(
            Revenue=("Revenue", "sum"),
            Quantity=("Quantity", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Customers=("CustomerID", "nunique"),
        )
        .sort_values("Month")
    )
    grouped["AverageOrderValue"] = grouped["Revenue"] / grouped["Orders"].replace(0, pd.NA)
    return grouped


def sales_by_country(clean_sales: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """Rank countries or markets by revenue."""

    return (
        clean_sales.groupby("Country", as_index=False)
        .agg(
            Revenue=("Revenue", "sum"),
            Quantity=("Quantity", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Customers=("CustomerID", "nunique"),
        )
        .sort_values("Revenue", ascending=False)
        .head(top_n)
    )


def top_products(clean_sales: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """Rank products by revenue."""

    grouped = (
        clean_sales.groupby(["StockCode", "Description"], as_index=False)
        .agg(
            Revenue=("Revenue", "sum"),
            Quantity=("Quantity", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Customers=("CustomerID", "nunique"),
            AvgUnitPrice=("UnitPrice", "mean"),
        )
        .sort_values("Revenue", ascending=False)
        .head(top_n)
    )
    return grouped


def product_performance(clean_sales: pd.DataFrame) -> pd.DataFrame:
    """Create a product-level performance table."""

    product_df = (
        clean_sales.groupby(["StockCode", "Description"], as_index=False)
        .agg(
            Revenue=("Revenue", "sum"),
            Quantity=("Quantity", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Customers=("CustomerID", "nunique"),
            AvgUnitPrice=("UnitPrice", "mean"),
            FirstSale=("InvoiceDate", "min"),
            LastSale=("InvoiceDate", "max"),
        )
        .sort_values("Revenue", ascending=False)
    )
    product_df["RevenuePerOrder"] = product_df["Revenue"] / product_df["Orders"].replace(0, pd.NA)
    return product_df


def customer_value_table(clean_sales: pd.DataFrame) -> pd.DataFrame:
    """Create a customer-level sales summary."""

    customer_sales = clean_sales.dropna(subset=["CustomerID"]).copy()
    grouped = (
        customer_sales.groupby("CustomerID", as_index=False)
        .agg(
            Revenue=("Revenue", "sum"),
            Orders=("InvoiceNo", "nunique"),
            Quantity=("Quantity", "sum"),
            FirstPurchase=("InvoiceDate", "min"),
            LastPurchase=("InvoiceDate", "max"),
            Countries=("Country", "nunique"),
        )
        .sort_values("Revenue", ascending=False)
    )
    grouped["AverageOrderValue"] = grouped["Revenue"] / grouped["Orders"].replace(0, pd.NA)
    return grouped


def cancellation_summary(prepared_all: pd.DataFrame) -> dict:
    """Summarize cancelled/negative transactions."""

    cancelled = prepared_all[prepared_all["IsCancelled"]].copy()
    return {
        "cancelled_rows": int(len(cancelled)),
        "cancelled_invoices": int(cancelled["InvoiceNo"].nunique()) if not cancelled.empty else 0,
        "cancelled_quantity": float(cancelled["Quantity"].sum()) if not cancelled.empty else 0.0,
        "cancelled_value": float(cancelled["Revenue"].sum()) if not cancelled.empty else 0.0,
        "cancelled_share_rows": safe_divide(len(cancelled), len(prepared_all)),
    }


def slow_moving_products(clean_sales: pd.DataFrame, top_n: int = 15) -> pd.DataFrame:
    """Identify products with low quantity and older last-sale dates."""

    product_df = product_performance(clean_sales)
    if product_df.empty:
        return product_df

    revenue_cutoff = product_df["Revenue"].quantile(0.25)
    quantity_cutoff = product_df["Quantity"].quantile(0.25)
    slow = product_df[
        (product_df["Revenue"] <= revenue_cutoff)
        & (product_df["Quantity"] <= quantity_cutoff)
    ].sort_values(["LastSale", "Revenue"], ascending=[True, True])
    return slow.head(top_n)

