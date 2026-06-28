# BI-Based Decision Support Dashboard for SME-Style Retail Analysis

This is a final year research project prototype for:

**Design and Development of a Business Intelligence-Based Decision Support Dashboard for Strategic Decision-Making in Small and Medium Enterprises**

The project uses the public **UCI Online Retail** dataset as a representative retail transaction dataset. It does not use or invent private Sri Lankan SME data.

## Main Features

- Streamlit web application
- Executive business KPI dashboard
- Sales analytics by month, country, product, customer, and quantity
- Customer analytics using RFM analysis
- Customer segmentation using K-Means clustering
- Product analytics and slow-moving product detection
- Decision-support insight cards and alerts
- Monthly revenue or quantity forecasting
- Colab-friendly training/preprocessing workflow
- GitHub and Streamlit Cloud deployment-ready structure

## Project Structure

```text
app.py
pages/
src/
  data_loader.py
  preprocessing.py
  kpi_calculator.py
  segmentation.py
  forecasting.py
  visualizations.py
  decision_support.py
  streamlit_helpers.py
  utils.py
data/
  raw/
    Online Retail.xlsx
  processed/
models/
notebooks/
  01_colab_training.ipynb
assets/
docs/
scripts/
  prepare_data.py
  train_models.py
requirements.txt
README.md
.gitignore
```

## Dataset

Place the dataset here:

```text
data/raw/Online Retail.xlsx
```

The app also checks the project root for `Online Retail.xlsx`, but `data/raw/` is the recommended path for GitHub and Streamlit Cloud.

The dataset includes these fields:

- `InvoiceNo`
- `StockCode`
- `Description`
- `Quantity`
- `InvoiceDate`
- `UnitPrice`
- `CustomerID`
- `Country`

## Local Setup

Create and activate a virtual environment if desired:

```bash
python -m venv .venv
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Prepare cleaned data:

```bash
python scripts/prepare_data.py
```

Train segmentation and forecasting artifacts:

```bash
python scripts/train_models.py
```

Run the Streamlit app:

```bash
streamlit run app.py
```

## Google Colab Workflow

Use:

```text
notebooks/01_colab_training.ipynb
```

Recommended Colab flow:

1. Clone your GitHub repository or upload the project files.
2. Ensure `Online Retail.xlsx` is inside `data/raw/`.
3. Run `pip install -r requirements.txt`.
4. Run `python scripts/prepare_data.py`.
5. Run `python scripts/train_models.py`.
6. Download generated files from `data/processed/` and `models/` if needed.

The Streamlit app can still run without pre-trained artifacts because it can compute analytics from the raw dataset.

## Streamlit Cloud Deployment

1. Push this repository to GitHub.
2. Make sure `data/raw/Online Retail.xlsx` is included in the repository.
3. Go to [Streamlit Cloud](https://streamlit.io/cloud).
4. Create a new app from your GitHub repository.
5. Set the main file path to:

```text
app.py
```

6. Streamlit Cloud will install packages from `requirements.txt`.
7. Launch the app.

## Viva Explanation Notes

This system follows a simple BI pipeline:

1. Load public retail transaction data.
2. Clean invalid, cancelled, zero-price, and negative-quantity records.
3. Calculate KPIs such as revenue, orders, customers, AOV, and product sales.
4. Visualize sales, product, country, and customer patterns.
5. Apply RFM analysis and K-Means clustering for customer segmentation.
6. Compare simple forecasting models using monthly revenue or quantity.
7. Translate analytics into decision-support recommendations.

The machine learning components are intentionally explainable and suitable for a final year individual project.

## Important Ethical Note

The project uses a public dataset only. It does not collect private customer information, does not use private Sri Lankan SME data, and does not attempt to identify real people or businesses.

