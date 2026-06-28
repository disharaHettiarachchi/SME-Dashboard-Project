"""Build a self-contained Colab notebook for the BI dashboard project."""

from __future__ import annotations

import json
from pathlib import Path
from textwrap import dedent


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "notebooks" / "BI_Dashboard_Full_Project_Colab.ipynb"


def md(source: str) -> dict:
    return {
        "cell_type": "markdown",
        "metadata": {},
        "source": dedent(source).strip().splitlines(keepends=True),
    }


def code(source: str) -> dict:
    return {
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": dedent(source).strip().splitlines(keepends=True),
    }


cells = [
    md(
        """
        # BI-Based Decision Support Dashboard - Full Colab Notebook

        This notebook contains the full project logic in one place so it can be run directly in Google Colab.

        It covers:

        - Dataset upload / loading
        - Data cleaning and preprocessing
        - Business KPI calculation
        - Sales, customer, and product analytics
        - RFM customer segmentation using K-Means
        - Monthly sales forecasting and model comparison
        - Decision-support insight generation
        - Saving processed outputs and model artifacts
        - Optional single-file Streamlit app export for GitHub and Streamlit Cloud

        Dataset: public UCI Online Retail dataset (`Online Retail.xlsx`).

        Ethical note: do not upload private SME data unless supervisor and ethics approval are obtained.
        """
    ),
    md(
        """
        ## 1. Install Dependencies

        Run this cell first in Colab. If you are running locally and already installed packages, it is safe to skip.
        """
    ),
    code(
        """
        !pip install -q pandas numpy openpyxl plotly scikit-learn joblib streamlit
        """
    ),
    md(
        """
        ## 2. Imports and Project Folders
        """
    ),
    code(
        """
        from pathlib import Path
        import os
        import warnings

        os.environ.setdefault("LOKY_MAX_CPU_COUNT", "4")
        warnings.filterwarnings("ignore", category=UserWarning)

        import joblib
        import numpy as np
        import pandas as pd
        import plotly.express as px
        import plotly.graph_objects as go
        from sklearn.cluster import KMeans
        from sklearn.linear_model import LinearRegression
        from sklearn.metrics import mean_absolute_error, mean_squared_error, silhouette_score
        from sklearn.preprocessing import StandardScaler

        PROJECT_ROOT = Path.cwd()
        RAW_DIR = PROJECT_ROOT / "data" / "raw"
        PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
        MODELS_DIR = PROJECT_ROOT / "models"
        ASSETS_DIR = PROJECT_ROOT / "assets"

        for folder in [RAW_DIR, PROCESSED_DIR, MODELS_DIR, ASSETS_DIR]:
            folder.mkdir(parents=True, exist_ok=True)

        DATASET_NAME = "Online Retail.xlsx"
        DATASET_PATH = RAW_DIR / DATASET_NAME

        print("Project root:", PROJECT_ROOT)
        print("Dataset path should be:", DATASET_PATH)
        """
    ),
    md(
        """
        ## 3. Upload or Locate the Dataset

        In Colab, upload `Online Retail.xlsx` when prompted. If the file is already in `data/raw/`, the upload step is skipped.
        """
    ),
    code(
        """
        def ensure_dataset_available():
            candidates = [
                RAW_DIR / DATASET_NAME,
                PROJECT_ROOT / DATASET_NAME,
                PROJECT_ROOT / "Online_Retail.xlsx",
            ]
            for candidate in candidates:
                if candidate.exists():
                    if candidate != DATASET_PATH:
                        DATASET_PATH.write_bytes(candidate.read_bytes())
                    return DATASET_PATH

            try:
                from google.colab import files
                print("Please upload Online Retail.xlsx")
                uploaded = files.upload()
                for filename in uploaded:
                    if filename.lower().endswith((".xlsx", ".xls", ".csv")):
                        target = DATASET_PATH if filename.lower().endswith((".xlsx", ".xls")) else RAW_DIR / filename
                        target.write_bytes(uploaded[filename])
                        return target
            except Exception as exc:
                print("Upload helper not available:", exc)

            raise FileNotFoundError("Dataset not found. Upload Online Retail.xlsx or place it in data/raw/.")

        DATASET_PATH = ensure_dataset_available()
        print("Using dataset:", DATASET_PATH)
        """
    ),
    md(
        """
        ## 4. Data Loading and Preprocessing Functions
        """
    ),
    code(
        """
        REQUIRED_COLUMNS = [
            "InvoiceNo",
            "StockCode",
            "Description",
            "Quantity",
            "InvoiceDate",
            "UnitPrice",
            "CustomerID",
            "Country",
        ]


        def load_raw_data(path=DATASET_PATH, nrows=None):
            path = Path(path)
            if path.suffix.lower() in [".xlsx", ".xls"]:
                return pd.read_excel(path, nrows=nrows, engine="openpyxl")
            if path.suffix.lower() == ".csv":
                return pd.read_csv(path, nrows=nrows)
            raise ValueError(f"Unsupported dataset format: {path.suffix}")


        def validate_columns(df):
            missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
            if missing:
                raise ValueError(f"Missing required columns: {missing}")


        def clean_identifier(value):
            if pd.isna(value):
                return pd.NA
            try:
                numeric = float(value)
                if np.isfinite(numeric) and numeric.is_integer():
                    return str(int(numeric))
            except (TypeError, ValueError):
                pass
            text = str(value).strip()
            return text if text else pd.NA


        def standardize_raw_data(df):
            data = df.copy()
            data.columns = [str(col).strip() for col in data.columns]
            validate_columns(data)

            data["InvoiceNo"] = data["InvoiceNo"].apply(clean_identifier)
            data["StockCode"] = data["StockCode"].apply(clean_identifier)
            data["CustomerID"] = data["CustomerID"].apply(clean_identifier)
            data["Description"] = (
                data["Description"]
                .astype("string")
                .str.strip()
                .str.replace(r"\\s+", " ", regex=True)
            )
            data["Country"] = data["Country"].astype("string").str.strip()
            data["Quantity"] = pd.to_numeric(data["Quantity"], errors="coerce")
            data["UnitPrice"] = pd.to_numeric(data["UnitPrice"], errors="coerce")
            data["InvoiceDate"] = pd.to_datetime(data["InvoiceDate"], errors="coerce")
            return data


        def prepare_retail_data(raw_df):
            data = standardize_raw_data(raw_df)
            data["IsCancelled"] = (
                data["InvoiceNo"].astype("string").str.startswith("C", na=False)
                | (data["Quantity"] < 0)
            )
            data["IsMissingCustomer"] = data["CustomerID"].isna()
            data["IsZeroOrNegativePrice"] = data["UnitPrice"].fillna(0) <= 0
            data["IsZeroOrNegativeQuantity"] = data["Quantity"].fillna(0) <= 0
            data["Revenue"] = data["Quantity"] * data["UnitPrice"]
            data["InvoiceDateOnly"] = pd.to_datetime(data["InvoiceDate"].dt.date)
            data["Month"] = data["InvoiceDate"].dt.to_period("M").dt.to_timestamp()
            data["Year"] = data["InvoiceDate"].dt.year
            data["YearMonth"] = data["InvoiceDate"].dt.to_period("M").astype("string")

            valid_sales_mask = (
                data["InvoiceDate"].notna()
                & data["Description"].notna()
                & data["Quantity"].gt(0)
                & data["UnitPrice"].gt(0)
                & ~data["IsCancelled"]
            )
            clean_sales = data.loc[valid_sales_mask].copy()
            return data, clean_sales


        def summarize_data_quality(prepared_df):
            return {
                "total_rows": int(len(prepared_df)),
                "missing_customer_rows": int(prepared_df["CustomerID"].isna().sum()),
                "cancelled_or_return_rows": int(prepared_df["IsCancelled"].sum()),
                "zero_or_negative_price_rows": int(prepared_df["IsZeroOrNegativePrice"].sum()),
                "zero_or_negative_quantity_rows": int(prepared_df["IsZeroOrNegativeQuantity"].sum()),
                "valid_sales_rows": int(
                    (
                        prepared_df["Quantity"].gt(0)
                        & prepared_df["UnitPrice"].gt(0)
                        & ~prepared_df["IsCancelled"]
                        & prepared_df["InvoiceDate"].notna()
                    ).sum()
                ),
            }


        def dataset_profile(df):
            return {
                "rows": int(len(df)),
                "columns": int(len(df.columns)),
                "date_min": pd.to_datetime(df["InvoiceDate"], errors="coerce").min(),
                "date_max": pd.to_datetime(df["InvoiceDate"], errors="coerce").max(),
                "countries": int(df["Country"].nunique(dropna=True)),
                "customers": int(df["CustomerID"].nunique(dropna=True)),
                "orders": int(df["InvoiceNo"].nunique(dropna=True)),
                "products": int(df["StockCode"].nunique(dropna=True)),
            }
        """
    ),
    md(
        """
        ## 5. KPI and Analytics Functions
        """
    ),
    code(
        """
        def safe_divide(numerator, denominator, default=0.0):
            if denominator in (0, None) or pd.isna(denominator):
                return default
            return float(numerator) / float(denominator)


        def calculate_overview_kpis(clean_sales):
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
            }


        def monthly_sales(clean_sales):
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


        def sales_by_country(clean_sales, top_n=10):
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


        def product_performance(clean_sales):
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


        def customer_value_table(clean_sales):
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


        def slow_moving_products(clean_sales, top_n=15):
            products = product_performance(clean_sales)
            revenue_cutoff = products["Revenue"].quantile(0.25)
            quantity_cutoff = products["Quantity"].quantile(0.25)
            slow = products[
                (products["Revenue"] <= revenue_cutoff)
                & (products["Quantity"] <= quantity_cutoff)
            ].sort_values(["LastSale", "Revenue"], ascending=[True, True])
            return slow.head(top_n)
        """
    ),
    md(
        """
        ## 6. RFM Customer Segmentation Functions
        """
    ),
    code(
        """
        def score_series(series, high_is_good=True):
            if series.nunique(dropna=True) <= 1:
                return pd.Series([3] * len(series), index=series.index)
            ranked = series.rank(method="first", ascending=high_is_good)
            return pd.qcut(ranked, q=5, labels=[1, 2, 3, 4, 5]).astype(int)


        def build_rfm_table(clean_sales, analysis_date=None):
            customer_sales = clean_sales.dropna(subset=["CustomerID"]).copy()
            if customer_sales.empty:
                return pd.DataFrame()

            max_date = customer_sales["InvoiceDate"].max()
            analysis_date = pd.to_datetime(analysis_date or max_date + pd.Timedelta(days=1))

            rfm = (
                customer_sales.groupby("CustomerID", as_index=False)
                .agg(
                    LastPurchase=("InvoiceDate", "max"),
                    FirstPurchase=("InvoiceDate", "min"),
                    Frequency=("InvoiceNo", "nunique"),
                    Monetary=("Revenue", "sum"),
                    Quantity=("Quantity", "sum"),
                )
            )
            rfm["Recency"] = (analysis_date - rfm["LastPurchase"]).dt.days
            rfm["CustomerAgeDays"] = (analysis_date - rfm["FirstPurchase"]).dt.days.clip(lower=1)
            rfm["AverageOrderValue"] = rfm["Monetary"] / rfm["Frequency"].replace(0, pd.NA)
            rfm["CLVIndicator"] = rfm["AverageOrderValue"] * rfm["Frequency"]
            rfm["RScore"] = score_series(rfm["Recency"], high_is_good=False)
            rfm["FScore"] = score_series(rfm["Frequency"], high_is_good=True)
            rfm["MScore"] = score_series(rfm["Monetary"], high_is_good=True)
            rfm["RFMScore"] = rfm["RScore"] + rfm["FScore"] + rfm["MScore"]
            return rfm.sort_values("Monetary", ascending=False)


        def cluster_label_map(summary):
            ranked = summary.copy()
            ranked["RankScore"] = (
                ranked["Monetary"].rank(ascending=False)
                + ranked["Frequency"].rank(ascending=False)
                + ranked["Recency"].rank(ascending=True)
            )
            ranked = ranked.sort_values("RankScore")
            label_pool = ["Champions", "Loyal Customers", "Potential Loyalists", "At Risk", "Low Value", "New or Occasional"]
            return {
                int(cluster_id): label_pool[index] if index < len(label_pool) else f"Segment {index + 1}"
                for index, cluster_id in enumerate(ranked["Cluster"].tolist())
            }


        def segment_customers(rfm, n_clusters=4, random_state=42):
            if rfm.empty:
                raise ValueError("RFM table is empty.")
            usable_clusters = max(2, min(n_clusters, len(rfm)))

            features = rfm[["Recency", "Frequency", "Monetary"]].copy()
            features["Frequency"] = np.log1p(features["Frequency"])
            features["Monetary"] = np.log1p(features["Monetary"].clip(lower=0))

            scaler = StandardScaler()
            scaled = scaler.fit_transform(features)

            model = KMeans(n_clusters=usable_clusters, random_state=random_state, n_init=10)
            clusters = model.fit_predict(scaled)

            segmented = rfm.copy()
            segmented["Cluster"] = clusters

            summary = (
                segmented.groupby("Cluster", as_index=False)
                .agg(
                    Customers=("CustomerID", "count"),
                    Recency=("Recency", "mean"),
                    Frequency=("Frequency", "mean"),
                    Monetary=("Monetary", "mean"),
                    TotalRevenue=("Monetary", "sum"),
                    AvgOrderValue=("AverageOrderValue", "mean"),
                    AvgRFMScore=("RFMScore", "mean"),
                )
                .sort_values("TotalRevenue", ascending=False)
            )
            labels = cluster_label_map(summary)
            segmented["SegmentName"] = segmented["Cluster"].map(labels)
            summary["SegmentName"] = summary["Cluster"].map(labels)

            metrics = {
                "n_clusters": usable_clusters,
                "customers_segmented": int(len(segmented)),
                "silhouette_score": None,
            }
            if usable_clusters > 1 and len(segmented) > usable_clusters:
                metrics["silhouette_score"] = float(silhouette_score(scaled, clusters))

            return segmented, summary, metrics, model, scaler
        """
    ),
    md(
        """
        ## 7. Forecasting Functions
        """
    ),
    code(
        """
        def aggregate_monthly(clean_sales, target="Revenue"):
            if target not in ["Revenue", "Quantity"]:
                raise ValueError("target must be Revenue or Quantity")
            monthly = (
                clean_sales.groupby("Month", as_index=False)
                .agg(Revenue=("Revenue", "sum"), Quantity=("Quantity", "sum"), Orders=("InvoiceNo", "nunique"))
                .sort_values("Month")
            )
            monthly["Target"] = monthly[target]
            return monthly


        def moving_average_forecast(history, horizon, window=3):
            values = list(history)
            preds = []
            for _ in range(horizon):
                window_values = values[-window:] if len(values) >= window else values
                pred = float(np.mean(window_values))
                preds.append(pred)
                values.append(pred)
            return np.array(preds)


        def linear_trend_forecast(history, horizon):
            history = np.array(history, dtype=float)
            x_train = np.arange(len(history)).reshape(-1, 1)
            model = LinearRegression()
            model.fit(x_train, history)
            x_future = np.arange(len(history), len(history) + horizon).reshape(-1, 1)
            return model.predict(x_future)


        def naive_forecast(history, horizon):
            return np.repeat(float(history[-1]), horizon)


        def mape(actual, predicted):
            actual = np.array(actual, dtype=float)
            predicted = np.array(predicted, dtype=float)
            mask = actual != 0
            if not mask.any():
                return np.nan
            return float(np.mean(np.abs((actual[mask] - predicted[mask]) / actual[mask])) * 100)


        def forecast_metrics(actual, predicted):
            return {
                "MAE": float(mean_absolute_error(actual, predicted)),
                "RMSE": float(np.sqrt(mean_squared_error(actual, predicted))),
                "MAPE": mape(actual, predicted),
            }


        def evaluate_forecast_models(monthly, test_size=3):
            if len(monthly) < 6:
                raise ValueError("At least 6 monthly observations are recommended.")
            test_size = min(test_size, max(1, len(monthly) // 4))
            train = monthly["Target"].iloc[:-test_size].astype(float).to_numpy()
            test = monthly["Target"].iloc[-test_size:].astype(float).to_numpy()

            predictions = {
                "Naive Last Value": naive_forecast(train, len(test)),
                "Moving Average (3 months)": moving_average_forecast(train, len(test), window=3),
                "Linear Trend": linear_trend_forecast(train, len(test)),
            }
            rows = []
            for name, pred in predictions.items():
                rows.append({"Model": name, **forecast_metrics(test, pred)})
            return pd.DataFrame(rows).sort_values("MAE")


        def forecast_future(monthly, model_name, periods=3):
            history = monthly["Target"].astype(float).to_numpy()
            if model_name == "Moving Average (3 months)":
                pred = moving_average_forecast(history, periods, window=3)
            elif model_name == "Linear Trend":
                pred = linear_trend_forecast(history, periods)
            else:
                pred = naive_forecast(history, periods)
            last_month = pd.to_datetime(monthly["Month"].max())
            future_months = pd.date_range(last_month + pd.offsets.MonthBegin(1), periods=periods, freq="MS")
            return pd.DataFrame({"Month": future_months, "Forecast": pred.clip(min=0)})


        def run_forecasting(clean_sales, target="Revenue", periods=3, test_size=3):
            monthly = aggregate_monthly(clean_sales, target=target)
            comparison = evaluate_forecast_models(monthly, test_size=test_size)
            best_model = str(comparison.iloc[0]["Model"])
            future = forecast_future(monthly, best_model, periods=periods)
            return monthly, comparison, future, best_model
        """
    ),
    md(
        """
        ## 8. Visualization and Decision-Support Functions
        """
    ),
    code(
        """
        COLOR_SEQUENCE = ["#111111", "#FF6B35", "#7F7F7F", "#B8BCC4", "#444444"]


        def short_text(value, max_length=45):
            text = str(value)
            return text if len(text) <= max_length else text[: max_length - 3] + "..."


        def revenue_trend_chart(monthly_df, value_col="Revenue", title=None):
            fig = px.line(monthly_df, x="Month", y=value_col, markers=True, title=title or f"Monthly {value_col} Trend")
            fig.update_traces(line_color="#FF6B35", line_width=3)
            fig.update_layout(template="plotly_white", hovermode="x unified")
            return fig


        def country_bar_chart(country_df):
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


        def product_bar_chart(product_df, value_col="Revenue", title="Top Products"):
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


        def rfm_segment_scatter(rfm_df):
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


        def forecast_chart(monthly_df, future_df):
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=monthly_df["Month"], y=monthly_df["Target"], mode="lines+markers", name="Actual", line=dict(color="#111111", width=3)))
            fig.add_trace(go.Scatter(x=future_df["Month"], y=future_df["Forecast"], mode="lines+markers", name="Forecast", line=dict(color="#FF6B35", width=3, dash="dash")))
            fig.update_layout(template="plotly_white", title="Actual vs Forecasted Monthly Sales", hovermode="x unified")
            return fig


        def insight(category, severity, title, message, recommendation):
            return {
                "category": category,
                "severity": severity,
                "title": title,
                "message": message,
                "recommendation": recommendation,
            }


        def generate_business_insights(clean_sales, rfm_df=None):
            insights = []
            monthly = monthly_sales(clean_sales)
            if len(monthly) >= 2:
                latest = monthly.iloc[-1]
                previous = monthly.iloc[-2]
                change = (latest["Revenue"] - previous["Revenue"]) / previous["Revenue"] if previous["Revenue"] else 0
                if change < -0.1:
                    insights.append(insight("Sales", "High", "Recent revenue decline detected", f"Latest monthly revenue is {change:.1%} lower than the previous month.", "Review recent product demand, market performance, and customer activity."))
                elif change > 0.1:
                    insights.append(insight("Sales", "Positive", "Recent revenue growth detected", f"Latest monthly revenue is {change:.1%} higher than the previous month.", "Identify the products and countries driving the increase."))

            products = product_performance(clean_sales)
            if not products.empty:
                top_product = products.iloc[0]
                insights.append(insight("Product", "Positive", "Top product deserves priority", f"{top_product['Description']} is the highest-revenue product.", "Prioritize availability and consider promotion."))

                slow = slow_moving_products(clean_sales, top_n=1)
                if not slow.empty:
                    slow_product = slow.iloc[0]
                    insights.append(insight("Product", "Medium", "Slow-moving product review needed", f"{slow_product['Description']} has low revenue and low quantity movement.", "Review whether the product needs bundling, discounting, or removal."))

            if rfm_df is not None and not rfm_df.empty and "SegmentName" in rfm_df.columns:
                segment_revenue = rfm_df.groupby("SegmentName")["Monetary"].sum().sort_values(ascending=False)
                top_segment = segment_revenue.index[0]
                insights.append(insight("Customer", "Positive", "High-value customer segment identified", f"The {top_segment} segment contributes the largest customer revenue share.", "Use retention or targeted communication strategies for this segment."))

                at_risk = rfm_df[rfm_df["Recency"] > rfm_df["Recency"].quantile(0.75)]
                if not at_risk.empty:
                    insights.append(insight("Customer", "Medium", "Inactive customer group requires attention", f"{len(at_risk):,} customers have relatively high recency values.", "Create a re-engagement campaign for inactive customers."))

            return insights
        """
    ),
    md(
        """
        ## 9. Run the Full Pipeline

        This cell loads the full Excel file, cleans it, calculates KPIs, runs segmentation, and runs forecasting.
        """
    ),
    code(
        """
        raw_df = load_raw_data(DATASET_PATH)
        prepared_all, clean_sales = prepare_retail_data(raw_df)

        profile = dataset_profile(raw_df)
        quality = summarize_data_quality(prepared_all)
        kpis = calculate_overview_kpis(clean_sales)

        monthly_df = monthly_sales(clean_sales)
        country_df = sales_by_country(clean_sales, top_n=10)
        products_df = product_performance(clean_sales)
        top_products_df = products_df.head(10)
        customers_df = customer_value_table(clean_sales)

        rfm_df = build_rfm_table(clean_sales)
        segmented_rfm, segment_summary, segment_metrics, kmeans_model, scaler = segment_customers(rfm_df, n_clusters=4)

        forecast_monthly, forecast_comparison, future_forecast, best_forecast_model = run_forecasting(clean_sales, target="Revenue", periods=3, test_size=3)
        insights = generate_business_insights(clean_sales, segmented_rfm)

        print("Dataset profile")
        print(profile)
        print("\\nData quality")
        print(quality)
        print("\\nKPIs")
        print(kpis)
        print("\\nSegmentation metrics")
        print(segment_metrics)
        print("\\nBest forecast model:", best_forecast_model)
        """
    ),
    md(
        """
        ## 10. View Key Tables
        """
    ),
    code(
        """
        display(pd.DataFrame([kpis]))
        display(country_df)
        display(top_products_df[["StockCode", "Description", "Revenue", "Quantity", "Orders"]])
        display(segment_summary)
        display(forecast_comparison)
        display(future_forecast)
        """
    ),
    md(
        """
        ## 11. Visual Outputs
        """
    ),
    code(
        """
        revenue_trend_chart(monthly_df).show()
        country_bar_chart(country_df).show()
        product_bar_chart(top_products_df, title="Top Products by Revenue").show()
        rfm_segment_scatter(segmented_rfm).show()
        forecast_chart(forecast_monthly, future_forecast).show()
        """
    ),
    md(
        """
        ## 12. Decision-Support Insights
        """
    ),
    code(
        """
        for number, item in enumerate(insights, start=1):
            print(f"{number}. [{item['severity']}] {item['title']}")
            print(f"   Category: {item['category']}")
            print(f"   Finding: {item['message']}")
            print(f"   Recommendation: {item['recommendation']}\\n")
        """
    ),
    md(
        """
        ## 13. Save Processed Data and Models
        """
    ),
    code(
        """
        clean_sales.to_csv(PROCESSED_DIR / "cleaned_online_retail.csv", index=False)
        monthly_df.to_csv(PROCESSED_DIR / "monthly_sales.csv", index=False)
        products_df.to_csv(PROCESSED_DIR / "product_performance.csv", index=False)
        segmented_rfm.to_csv(PROCESSED_DIR / "rfm_segments.csv", index=False)
        segment_summary.to_csv(PROCESSED_DIR / "segment_summary.csv", index=False)
        forecast_comparison.to_csv(PROCESSED_DIR / "forecast_model_comparison_revenue.csv", index=False)
        future_forecast.to_csv(PROCESSED_DIR / "future_forecast_revenue.csv", index=False)

        joblib.dump(
            {
                "model": kmeans_model,
                "scaler": scaler,
                "metrics": segment_metrics,
            },
            MODELS_DIR / "customer_segmentation.joblib",
        )
        joblib.dump(
            {
                "monthly": forecast_monthly,
                "comparison": forecast_comparison,
                "future_forecast": future_forecast,
                "best_model": best_forecast_model,
            },
            MODELS_DIR / "forecast_results.joblib",
        )

        print("Saved outputs to:")
        print(PROCESSED_DIR)
        print(MODELS_DIR)
        """
    ),
    md(
        """
        ## 14. Optional: Download Outputs from Colab
        """
    ),
    code(
        """
        # Run this in Colab if you want to download the generated outputs.
        # from google.colab import files
        # !zip -r bi_dashboard_outputs.zip data/processed models
        # files.download("bi_dashboard_outputs.zip")
        """
    ),
    md(
        """
        ## 15. Optional: Create a Single-File Streamlit App for Deployment

        Streamlit Cloud cannot deploy directly from a notebook. It needs an `app.py` file.

        This cell writes a compact single-file `app.py` from inside the notebook. You can push the generated `app.py`, `requirements.txt`, and `data/raw/Online Retail.xlsx` to GitHub for Streamlit Cloud.
        """
    ),
    code(
        r'''
        app_code = r"""
        from pathlib import Path
        import os
        import numpy as np
        import pandas as pd
        import plotly.express as px
        import plotly.graph_objects as go
        import streamlit as st
        from sklearn.cluster import KMeans
        from sklearn.linear_model import LinearRegression
        from sklearn.metrics import mean_absolute_error, mean_squared_error, silhouette_score
        from sklearn.preprocessing import StandardScaler

        os.environ.setdefault("LOKY_MAX_CPU_COUNT", "4")

        RAW_DIR = Path("data/raw")
        DATASET_PATH = RAW_DIR / "Online Retail.xlsx"

        st.set_page_config(page_title="BI Decision Support Dashboard", layout="wide")

        REQUIRED_COLUMNS = ["InvoiceNo", "StockCode", "Description", "Quantity", "InvoiceDate", "UnitPrice", "CustomerID", "Country"]

        def clean_identifier(value):
            if pd.isna(value):
                return pd.NA
            try:
                numeric = float(value)
                if np.isfinite(numeric) and numeric.is_integer():
                    return str(int(numeric))
            except (TypeError, ValueError):
                pass
            text = str(value).strip()
            return text if text else pd.NA

        @st.cache_data(show_spinner="Loading and cleaning dataset...")
        def load_data():
            if not DATASET_PATH.exists():
                fallback = Path("Online Retail.xlsx")
                if fallback.exists():
                    path = fallback
                else:
                    raise FileNotFoundError("Place Online Retail.xlsx inside data/raw/")
            else:
                path = DATASET_PATH
            raw = pd.read_excel(path, engine="openpyxl")
            data = raw.copy()
            data.columns = [str(col).strip() for col in data.columns]
            missing = [col for col in REQUIRED_COLUMNS if col not in data.columns]
            if missing:
                raise ValueError(f"Missing columns: {missing}")
            data["InvoiceNo"] = data["InvoiceNo"].apply(clean_identifier)
            data["StockCode"] = data["StockCode"].apply(clean_identifier)
            data["CustomerID"] = data["CustomerID"].apply(clean_identifier)
            data["Description"] = data["Description"].astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
            data["Country"] = data["Country"].astype("string").str.strip()
            data["Quantity"] = pd.to_numeric(data["Quantity"], errors="coerce")
            data["UnitPrice"] = pd.to_numeric(data["UnitPrice"], errors="coerce")
            data["InvoiceDate"] = pd.to_datetime(data["InvoiceDate"], errors="coerce")
            data["IsCancelled"] = data["InvoiceNo"].astype("string").str.startswith("C", na=False) | (data["Quantity"] < 0)
            data["Revenue"] = data["Quantity"] * data["UnitPrice"]
            data["Month"] = data["InvoiceDate"].dt.to_period("M").dt.to_timestamp()
            clean = data[(data["InvoiceDate"].notna()) & (data["Quantity"] > 0) & (data["UnitPrice"] > 0) & (~data["IsCancelled"])].copy()
            return data, clean

        def safe_divide(a, b):
            return 0 if b == 0 or pd.isna(b) else float(a) / float(b)

        def monthly_sales(clean):
            return clean.groupby("Month", as_index=False).agg(Revenue=("Revenue", "sum"), Quantity=("Quantity", "sum"), Orders=("InvoiceNo", "nunique"), Customers=("CustomerID", "nunique")).sort_values("Month")

        def product_performance(clean):
            return clean.groupby(["StockCode", "Description"], as_index=False).agg(Revenue=("Revenue", "sum"), Quantity=("Quantity", "sum"), Orders=("InvoiceNo", "nunique"), Customers=("CustomerID", "nunique")).sort_values("Revenue", ascending=False)

        def build_rfm(clean):
            sales = clean.dropna(subset=["CustomerID"]).copy()
            analysis_date = sales["InvoiceDate"].max() + pd.Timedelta(days=1)
            rfm = sales.groupby("CustomerID", as_index=False).agg(LastPurchase=("InvoiceDate", "max"), Frequency=("InvoiceNo", "nunique"), Monetary=("Revenue", "sum"), Quantity=("Quantity", "sum"))
            rfm["Recency"] = (analysis_date - rfm["LastPurchase"]).dt.days
            return rfm

        def segment_customers(rfm, n_clusters=4):
            features = rfm[["Recency", "Frequency", "Monetary"]].copy()
            features["Frequency"] = np.log1p(features["Frequency"])
            features["Monetary"] = np.log1p(features["Monetary"].clip(lower=0))
            scaler = StandardScaler()
            scaled = scaler.fit_transform(features)
            model = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
            rfm = rfm.copy()
            rfm["Cluster"] = model.fit_predict(scaled)
            summary = rfm.groupby("Cluster", as_index=False).agg(Customers=("CustomerID", "count"), Recency=("Recency", "mean"), Frequency=("Frequency", "mean"), Monetary=("Monetary", "mean"), TotalRevenue=("Monetary", "sum")).sort_values("TotalRevenue", ascending=False)
            labels = ["Champions", "Loyal Customers", "Potential Loyalists", "At Risk", "Low Value", "New or Occasional"]
            label_map = {cluster: labels[i] if i < len(labels) else f"Segment {i+1}" for i, cluster in enumerate(summary["Cluster"])}
            rfm["SegmentName"] = rfm["Cluster"].map(label_map)
            summary["SegmentName"] = summary["Cluster"].map(label_map)
            score = silhouette_score(scaled, rfm["Cluster"]) if len(rfm) > n_clusters else None
            return rfm, summary, score

        def forecast(clean, target="Revenue", periods=3):
            monthly = monthly_sales(clean)
            monthly["Target"] = monthly[target]
            test_size = min(3, max(1, len(monthly) // 4))
            train = monthly["Target"].iloc[:-test_size].to_numpy(dtype=float)
            test = monthly["Target"].iloc[-test_size:].to_numpy(dtype=float)
            preds = {}
            preds["Naive Last Value"] = np.repeat(train[-1], len(test))
            preds["Moving Average (3 months)"] = np.repeat(np.mean(train[-3:]), len(test))
            model = LinearRegression().fit(np.arange(len(train)).reshape(-1, 1), train)
            preds["Linear Trend"] = model.predict(np.arange(len(train), len(train) + len(test)).reshape(-1, 1))
            rows = []
            for name, pred in preds.items():
                rows.append({"Model": name, "MAE": mean_absolute_error(test, pred), "RMSE": np.sqrt(mean_squared_error(test, pred))})
            comparison = pd.DataFrame(rows).sort_values("MAE")
            best = comparison.iloc[0]["Model"]
            future_months = pd.date_range(monthly["Month"].max() + pd.offsets.MonthBegin(1), periods=periods, freq="MS")
            if best == "Linear Trend":
                future_pred = model.predict(np.arange(len(monthly), len(monthly) + periods).reshape(-1, 1))
            elif best == "Moving Average (3 months)":
                future_pred = np.repeat(np.mean(monthly["Target"].tail(3)), periods)
            else:
                future_pred = np.repeat(monthly["Target"].iloc[-1], periods)
            future = pd.DataFrame({"Month": future_months, "Forecast": np.clip(future_pred, 0, None)})
            return monthly, comparison, future, best

        prepared, clean = load_data()

        min_date = clean["InvoiceDate"].min().date()
        max_date = clean["InvoiceDate"].max().date()
        st.sidebar.header("Filters")
        date_range = st.sidebar.date_input("Date range", value=(min_date, max_date), min_value=min_date, max_value=max_date)
        countries = sorted(clean["Country"].dropna().unique().tolist())
        selected_countries = st.sidebar.multiselect("Countries", countries, default=[])
        filtered = clean.copy()
        if isinstance(date_range, tuple) and len(date_range) == 2:
            filtered = filtered[(filtered["InvoiceDate"] >= pd.to_datetime(date_range[0])) & (filtered["InvoiceDate"] <= pd.to_datetime(date_range[1]) + pd.Timedelta(days=1))]
        if selected_countries:
            filtered = filtered[filtered["Country"].isin(selected_countries)]

        page = st.sidebar.radio("Page", ["Executive Overview", "Sales Analytics", "Customer Analytics", "Product Analytics", "Decision Support", "Forecasting"])

        st.title("BI-Based Decision Support Dashboard")
        st.caption("Public UCI Online Retail dataset. No private SME data is used.")

        if page == "Executive Overview":
            revenue = filtered["Revenue"].sum()
            orders = filtered["InvoiceNo"].nunique()
            customers = filtered["CustomerID"].nunique()
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total revenue", f"GBP {revenue:,.2f}")
            c2.metric("Orders", f"{orders:,}")
            c3.metric("Customers", f"{customers:,}")
            c4.metric("Average order value", f"GBP {safe_divide(revenue, orders):,.2f}")
            monthly = monthly_sales(filtered)
            st.plotly_chart(px.line(monthly, x="Month", y="Revenue", markers=True, title="Monthly Revenue Trend"), use_container_width=True)
            top_country = filtered.groupby("Country", as_index=False)["Revenue"].sum().sort_values("Revenue", ascending=False).head(10)
            st.plotly_chart(px.bar(top_country.sort_values("Revenue"), x="Revenue", y="Country", orientation="h", title="Top Countries"), use_container_width=True)

        elif page == "Sales Analytics":
            monthly = monthly_sales(filtered)
            st.plotly_chart(px.line(monthly, x="Month", y="Revenue", markers=True, title="Sales Trend"), use_container_width=True)
            st.dataframe(monthly, use_container_width=True)
            cancelled = prepared[prepared["IsCancelled"]]
            st.metric("Cancelled / return rows", f"{len(cancelled):,}")

        elif page == "Customer Analytics":
            rfm = build_rfm(filtered)
            segmented, summary, score = segment_customers(rfm, n_clusters=4)
            st.metric("Customers segmented", f"{len(segmented):,}")
            st.metric("Silhouette score", "N/A" if score is None else f"{score:.3f}")
            st.plotly_chart(px.scatter(segmented, x="Frequency", y="Monetary", color="SegmentName", size="Quantity", hover_name="CustomerID", title="Customer Segments"), use_container_width=True)
            st.dataframe(summary, use_container_width=True)

        elif page == "Product Analytics":
            products = product_performance(filtered)
            top = products.head(15)
            top["Product"] = top["Description"].astype(str).str.slice(0, 45)
            st.plotly_chart(px.bar(top.sort_values("Revenue"), x="Revenue", y="Product", orientation="h", title="Top Products by Revenue"), use_container_width=True)
            st.dataframe(products, use_container_width=True)

        elif page == "Decision Support":
            monthly = monthly_sales(filtered)
            products = product_performance(filtered)
            if len(monthly) >= 2:
                change = (monthly.iloc[-1]["Revenue"] - monthly.iloc[-2]["Revenue"]) / monthly.iloc[-2]["Revenue"]
                if change < -0.1:
                    st.warning(f"Revenue declined by {change:.1%}. Review product and market performance.")
                else:
                    st.success(f"Recent revenue change: {change:.1%}.")
            if not products.empty:
                st.info(f"Top product to prioritize: {products.iloc[0]['Description']}")

        elif page == "Forecasting":
            target = st.sidebar.selectbox("Target", ["Revenue", "Quantity"])
            monthly, comparison, future, best = forecast(filtered, target=target)
            st.metric("Best model", best)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=monthly["Month"], y=monthly["Target"], mode="lines+markers", name="Actual"))
            fig.add_trace(go.Scatter(x=future["Month"], y=future["Forecast"], mode="lines+markers", name="Forecast"))
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(comparison, use_container_width=True)
            st.dataframe(future, use_container_width=True)
        """

        Path("app.py").write_text(app_code.strip() + "\n", encoding="utf-8")
        Path("requirements.txt").write_text(
            "pandas\nnumpy\nopenpyxl\nplotly\nscikit-learn\njoblib\nstreamlit\n",
            encoding="utf-8",
        )
        print("Created app.py and requirements.txt for Streamlit Cloud.")
        print("Push app.py, requirements.txt, and data/raw/Online Retail.xlsx to GitHub.")
        '''
    ),
    md(
        """
        ## 16. Run the Streamlit App in Colab for Preview

        Colab does not host Streamlit apps permanently, but you can preview using a tunnel such as localtunnel or pyngrok. For final hosting, use GitHub + Streamlit Cloud.
        """
    ),
    code(
        """
        # Optional local Colab preview. Run after creating app.py above.
        # !streamlit run app.py --server.port 8501 --server.headless true
        """
    ),
]


notebook = {
    "cells": cells,
    "metadata": {
        "kernelspec": {
            "display_name": "Python 3",
            "language": "python",
            "name": "python3",
        },
        "language_info": {
            "name": "python",
            "version": "3.10",
        },
    },
    "nbformat": 4,
    "nbformat_minor": 5,
}


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(notebook, indent=2), encoding="utf-8")
    print(OUT)


if __name__ == "__main__":
    main()
