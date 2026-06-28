"""Simple monthly sales forecasting models."""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error


@dataclass
class ForecastResult:
    monthly: pd.DataFrame
    comparison: pd.DataFrame
    future_forecast: pd.DataFrame
    best_model: str


def aggregate_monthly(clean_sales: pd.DataFrame, target: str = "Revenue") -> pd.DataFrame:
    """Aggregate revenue or quantity by month."""

    if target not in {"Revenue", "Quantity"}:
        raise ValueError("target must be either 'Revenue' or 'Quantity'")

    monthly = (
        clean_sales.groupby("Month", as_index=False)
        .agg(Revenue=("Revenue", "sum"), Quantity=("Quantity", "sum"), Orders=("InvoiceNo", "nunique"))
        .sort_values("Month")
    )
    monthly["Target"] = monthly[target]
    return monthly


def _moving_average_forecast(history: list[float], horizon: int, window: int = 3) -> np.ndarray:
    """Forecast by repeatedly averaging the latest observations."""

    values = list(history)
    preds = []
    for _ in range(horizon):
        window_values = values[-window:] if len(values) >= window else values
        pred = float(np.mean(window_values))
        preds.append(pred)
        values.append(pred)
    return np.array(preds)


def _linear_trend_forecast(history: np.ndarray, horizon: int) -> np.ndarray:
    """Forecast with a simple linear trend over time index."""

    x_train = np.arange(len(history)).reshape(-1, 1)
    model = LinearRegression()
    model.fit(x_train, history)
    x_future = np.arange(len(history), len(history) + horizon).reshape(-1, 1)
    return model.predict(x_future)


def _naive_forecast(history: np.ndarray, horizon: int) -> np.ndarray:
    """Forecast by repeating the latest actual value."""

    return np.repeat(history[-1], horizon)


def _mape(actual: np.ndarray, predicted: np.ndarray) -> float:
    mask = actual != 0
    if not mask.any():
        return float("nan")
    return float(np.mean(np.abs((actual[mask] - predicted[mask]) / actual[mask])) * 100)


def _metrics(actual: np.ndarray, predicted: np.ndarray) -> dict:
    return {
        "MAE": float(mean_absolute_error(actual, predicted)),
        "RMSE": float(np.sqrt(mean_squared_error(actual, predicted))),
        "MAPE": _mape(actual, predicted),
    }


def evaluate_forecast_models(monthly: pd.DataFrame, test_size: int = 3) -> pd.DataFrame:
    """Compare naive, moving average, and linear trend models."""

    if len(monthly) < 6:
        raise ValueError("At least 6 monthly observations are recommended for forecast comparison.")

    test_size = min(test_size, max(1, len(monthly) // 4))
    train = monthly["Target"].iloc[:-test_size].astype(float).to_numpy()
    test = monthly["Target"].iloc[-test_size:].astype(float).to_numpy()

    predictions = {
        "Naive Last Value": _naive_forecast(train, len(test)),
        "Moving Average (3 months)": _moving_average_forecast(train.tolist(), len(test), window=3),
        "Linear Trend": _linear_trend_forecast(train, len(test)),
    }

    rows = []
    for name, pred in predictions.items():
        values = _metrics(test, pred)
        rows.append({"Model": name, **values})
    return pd.DataFrame(rows).sort_values("MAE")


def forecast_future(monthly: pd.DataFrame, model_name: str, periods: int = 3) -> pd.DataFrame:
    """Forecast future months using the selected model."""

    history = monthly["Target"].astype(float).to_numpy()
    if model_name == "Moving Average (3 months)":
        pred = _moving_average_forecast(history.tolist(), periods, window=3)
    elif model_name == "Linear Trend":
        pred = _linear_trend_forecast(history, periods)
    else:
        pred = _naive_forecast(history, periods)

    last_month = pd.to_datetime(monthly["Month"].max())
    future_months = pd.date_range(last_month + pd.offsets.MonthBegin(1), periods=periods, freq="MS")
    return pd.DataFrame({"Month": future_months, "Forecast": pred.clip(min=0)})


def run_forecasting(clean_sales: pd.DataFrame, target: str = "Revenue", periods: int = 3, test_size: int = 3) -> ForecastResult:
    """Run the full forecasting workflow."""

    monthly = aggregate_monthly(clean_sales, target=target)
    comparison = evaluate_forecast_models(monthly, test_size=test_size)
    best_model = str(comparison.iloc[0]["Model"])
    future = forecast_future(monthly, best_model, periods=periods)
    return ForecastResult(monthly=monthly, comparison=comparison, future_forecast=future, best_model=best_model)

