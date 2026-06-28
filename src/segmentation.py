"""RFM analysis and customer segmentation."""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path

os.environ.setdefault("LOKY_MAX_CPU_COUNT", "4")

import joblib
import numpy as np
import pandas as pd
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score
from sklearn.preprocessing import StandardScaler


@dataclass
class SegmentationResult:
    rfm: pd.DataFrame
    summary: pd.DataFrame
    metrics: dict
    model: KMeans
    scaler: StandardScaler


def build_rfm_table(clean_sales: pd.DataFrame, analysis_date: pd.Timestamp | None = None) -> pd.DataFrame:
    """Build a customer-level RFM table."""

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

    rfm = add_rfm_scores(rfm)
    return rfm.sort_values("Monetary", ascending=False)


def _score_series(series: pd.Series, high_is_good: bool = True) -> pd.Series:
    """Score a series from 1 to 5 using ranks so duplicate values are safe."""

    if series.nunique(dropna=True) <= 1:
        return pd.Series([3] * len(series), index=series.index)
    ranked = series.rank(method="first", ascending=high_is_good)
    scored = pd.qcut(ranked, q=5, labels=[1, 2, 3, 4, 5])
    return scored.astype(int)


def add_rfm_scores(rfm: pd.DataFrame) -> pd.DataFrame:
    """Add R, F, M, and combined RFM scores."""

    scored = rfm.copy()
    scored["RScore"] = _score_series(scored["Recency"], high_is_good=False)
    scored["FScore"] = _score_series(scored["Frequency"], high_is_good=True)
    scored["MScore"] = _score_series(scored["Monetary"], high_is_good=True)
    scored["RFMScore"] = scored["RScore"] + scored["FScore"] + scored["MScore"]
    return scored


def _cluster_label_map(summary: pd.DataFrame) -> dict:
    """Assign human-readable labels to cluster IDs based on business value."""

    ranked = summary.copy()
    ranked["RankScore"] = (
        ranked["Monetary"].rank(ascending=False)
        + ranked["Frequency"].rank(ascending=False)
        + ranked["Recency"].rank(ascending=True)
    )
    ranked = ranked.sort_values("RankScore")

    label_pool = [
        "Champions",
        "Loyal Customers",
        "Potential Loyalists",
        "At Risk",
        "Low Value",
        "New or Occasional",
    ]
    return {
        int(cluster_id): label_pool[index] if index < len(label_pool) else f"Segment {index + 1}"
        for index, cluster_id in enumerate(ranked["Cluster"].tolist())
    }


def segment_customers(rfm: pd.DataFrame, n_clusters: int = 4, random_state: int = 42) -> SegmentationResult:
    """Cluster customers using scaled log-transformed RFM features."""

    if rfm.empty:
        raise ValueError("RFM table is empty. Customer segmentation requires valid CustomerID values.")

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
    labels = _cluster_label_map(summary)
    segmented["SegmentName"] = segmented["Cluster"].map(labels)
    summary["SegmentName"] = summary["Cluster"].map(labels)

    metrics = {
        "n_clusters": usable_clusters,
        "customers_segmented": int(len(segmented)),
        "silhouette_score": None,
    }
    if usable_clusters > 1 and len(segmented) > usable_clusters:
        metrics["silhouette_score"] = float(silhouette_score(scaled, clusters))

    return SegmentationResult(
        rfm=segmented.sort_values("Monetary", ascending=False),
        summary=summary,
        metrics=metrics,
        model=model,
        scaler=scaler,
    )


def save_segmentation_result(result: SegmentationResult, output_dir: str | Path) -> None:
    """Save segmentation tables and model artifacts."""

    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)
    result.rfm.to_csv(output / "rfm_segments.csv", index=False)
    result.summary.to_csv(output / "segment_summary.csv", index=False)
    joblib.dump({"model": result.model, "scaler": result.scaler, "metrics": result.metrics}, output / "customer_segmentation.joblib")
