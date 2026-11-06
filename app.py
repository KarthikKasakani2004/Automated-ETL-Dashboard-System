# ================================================================
# AUTOMATED ETL + INTERACTIVE DASHBOARD SYSTEM
# Complete single-file Streamlit application
# Save as: app.py
# Run: python -m streamlit run app.py
# ================================================================

import io
import re
import warnings
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

warnings.filterwarnings("ignore")
PLOTLY_TEMPLATE = "plotly_dark" if st.get_option("theme.base") == "dark" else "plotly_white"

# Optional export libraries
try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        Image as ReportLabImage,
        PageBreak,
    )
    PDF_AVAILABLE = True
except Exception:
    PDF_AVAILABLE = False

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    PPT_AVAILABLE = True
except Exception:
    PPT_AVAILABLE = False


# ================================================================
# PAGE CONFIGURATION + THEME
# ================================================================

st.set_page_config(
    page_title="Automated ETL & Dashboard System",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY = "#17365D"
BLUE = "#2F75B5"
LIGHT_BLUE = "#EAF3FA"
PALE_BLUE = "#F4F8FC"
DARK = "#263238"
GREY = "#64748B"
GREEN = "#198754"
ORANGE = "#E69500"
RED = "#C0392B"

st.markdown(
    """
    <style>
        /* Theme-aware colors: Streamlit updates these variables in Light/Dark mode. */
        :root {
            color-scheme: light dark;
        }
        .stApp {
            background: var(--background-color);
            color: var(--text-color);
        }
        [data-testid="stSidebar"] {
            background: var(--secondary-background-color);
        }
        [data-testid="stSidebar"] *,
        .stApp h1, .stApp h2, .stApp h3, .stApp h4,
        .stApp p, .stApp label, .stApp li,
        .stApp [data-testid="stMarkdownContainer"] {
            color: var(--text-color);
        }
        .main-title {
            font-size: 2rem;
            font-weight: 800;
            color: var(--text-color);
            line-height: 1.25;
            overflow-wrap: anywhere;
        }
        .subtitle, .small-note {
            color: var(--text-color);
            opacity: .78;
            font-size: .95rem;
        }
        .kpi-card, .section-card,
        div[data-testid="stMetric"] {
            background: var(--secondary-background-color);
            color: var(--text-color);
            border: 1px solid var(--border-color, rgba(128,128,128,.28));
            border-radius: 12px;
            padding: 16px 14px;
            box-shadow: 0 2px 8px rgba(0,0,0,.06);
        }
        .kpi-card { min-height: 105px; }
        .kpi-label {
            color: var(--text-color);
            font-size: .83rem;
            font-weight: 700;
            margin-bottom: 8px;
        }
        .kpi-value {
            color: var(--primary-color);
            font-size: 1.55rem;
            font-weight: 800;
            overflow-wrap: anywhere;
        }
        div[data-testid="stMetric"] {
            padding: 12px;
        }
        /* Ensure form controls remain readable in both themes. */
        input, textarea, [data-baseweb="select"] > div,
        [data-testid="stFileUploader"] section {
            color: var(--text-color) !important;
            background-color: var(--secondary-background-color) !important;
        }
        [data-testid="stFileUploader"] small,
        [data-testid="stFileUploader"] span,
        [data-testid="stFileUploader"] label {
            color: var(--text-color) !important;
        }
        /* Tables and dataframe containers follow the active theme. */
        [data-testid="stDataFrame"], [data-testid="stTable"] {
            color: var(--text-color);
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# ================================================================
# SESSION STATE
# ================================================================

DEFAULT_STATE = {
    "raw_df": None,
    "clean_df": None,
    "file_name": None,
    "cleaning_log": [],
    "quality_report": None,
    "selected_date_col": None,
    "selected_id_col": None,
    "selected_domain": "General",
    "last_findings": [],
    "last_recommendations": [],
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# ================================================================
# HELPER FUNCTIONS
# ================================================================

def safe_name(value):
    """Create a safe filename."""
    value = str(value).strip()
    value = re.sub(r"[^A-Za-z0-9_-]+", "_", value)
    return value[:80] or "dataset"


def read_uploaded_file(uploaded_file):
    """Read CSV, Excel, or JSON files."""
    file_name = uploaded_file.name.lower()

    if file_name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding="latin-1")

    if file_name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)

    if file_name.endswith(".json"):
        return pd.read_json(uploaded_file)

    raise ValueError("Unsupported file type. Upload CSV, Excel, or JSON.")


def detect_domain(df):
    """Infer a broad business domain from column names."""
    columns = " ".join(str(c).lower() for c in df.columns)

    patterns = {
        "Finance": [
            "transaction", "loan", "balance", "credit", "debit",
            "interest", "account", "payment", "salary", "expense",
            "income", "profit", "revenue", "budget", "cost",
        ],
        "Sales / Retail": [
            "sales", "order", "product", "quantity", "discount",
            "customer", "invoice", "unit price", "price", "category",
        ],
        "Supply Chain": [
            "warehouse", "supplier", "shipping", "shipment",
            "inventory", "stock", "delivery", "lead time",
        ],
        "Hospitality": [
            "hotel", "booking", "reservation", "guest", "room",
            "check in", "check out", "occupancy", "review",
        ],
        "Healthcare": [
            "patient", "diagnosis", "doctor", "hospital",
            "medical", "treatment", "admission",
        ],
        "HR / Workforce": [
            "employee", "salary", "department", "attrition",
            "hire", "performance", "job role",
        ],
        "Movies / Entertainment": [
            "movie", "film", "title", "runtime", "vote average",
            "popularity", "release date", "budget",
        ],
    }

    for domain, keywords in patterns.items():
        if any(keyword in columns for keyword in keywords):
            return domain

    return "General"


def infer_column_role(column_name, series):
    """Suggest a data role based on name and observed values."""
    name = str(column_name).lower()

    if any(x in name for x in ["id", "key", "code", "number"]):
        if series.nunique(dropna=True) >= max(2, len(series) * 0.5):
            return "Possible identifier"

    if any(x in name for x in ["date", "time", "timestamp", "year", "month"]):
        return "Date / Time"

    if pd.api.types.is_bool_dtype(series):
        return "Boolean"

    if pd.api.types.is_numeric_dtype(series):
        if any(x in name for x in [
            "amount", "sales", "revenue", "profit", "cost", "price",
            "budget", "balance", "salary", "fee", "loan", "income",
        ]):
            return "Numeric measure"

        if any(x in name for x in [
            "age", "quantity", "count", "rating", "score", "rate",
            "percentage", "percent", "discount", "duration", "days",
        ]):
            return "Numeric measure"

        return "Numeric"

    if series.nunique(dropna=True) <= min(30, max(2, int(len(series) * 0.05))):
        return "Category"

    return "Text"


def suggest_primary_key(df):
    """Suggest columns that are unique and non-null."""
    suggestions = []

    for column in df.columns:
        series = df[column]
        if (
            series.notna().all()
            and series.nunique(dropna=True) == len(df)
        ):
            name = str(column).lower()
            score = 2 if any(x in name for x in ["id", "key", "code"]) else 1
            suggestions.append((column, score))

    suggestions.sort(key=lambda item: item[1], reverse=True)
    return [item[0] for item in suggestions]


def get_column_recommendations(df):
    """Create cleaning suggestions for every column."""
    rows = []

    for column in df.columns:
        series = df[column]
        missing_count = int(series.isna().sum())
        missing_pct = (missing_count / len(df) * 100) if len(df) else 0
        role = infer_column_role(column, series)

        if missing_count == 0:
            missing_action = "No missing values"
        elif pd.api.types.is_numeric_dtype(series):
            missing_action = "Review: median/mean, group-based fill, or keep missing"
        elif role == "Category":
            missing_action = "Review: mode, 'Unknown', or keep missing"
        else:
            missing_action = "Review: 'Unknown', valid source lookup, or keep missing"

        rows.append({
            "Column": column,
            "Detected Type": str(series.dtype),
            "Suggested Role": role,
            "Missing Count": missing_count,
            "Missing %": round(missing_pct, 2),
            "Unique Values": int(series.nunique(dropna=True)),
            "Suggested Missing-Value Action": missing_action,
        })

    return pd.DataFrame(rows)


def create_quality_report(df):
    """Calculate data-quality checks without automatically deleting data."""
    rows = []

    for column in df.columns:
        series = df[column]
        missing = int(series.isna().sum())
        duplicate_values = int(series.duplicated(keep=False).sum())
        rows.append({
            "Column": column,
            "Missing Values": missing,
            "Missing %": round(missing / len(df) * 100, 2) if len(df) else 0,
            "Duplicate Values": duplicate_values,
            "Unique Values": int(series.nunique(dropna=True)),
            "Data Type": str(series.dtype),
        })

    report = pd.DataFrame(rows)

    duplicate_rows = int(df.duplicated().sum())
    duplicate_rows_by_id = None

    candidate_ids = [
        c for c in df.columns
        if any(x in str(c).lower() for x in ["id", "key"])
    ]

    if candidate_ids:
        duplicate_rows_by_id = {
            c: int(df[c].duplicated().sum())
            for c in candidate_ids
        }

    return {
        "column_report": report,
        "duplicate_rows": duplicate_rows,
        "duplicate_id_counts": duplicate_rows_by_id,
        "total_missing": int(df.isna().sum().sum()),
        "rows": len(df),
        "columns": len(df.columns),
    }


def format_number(value, currency=False):
    """Readable KPI formatting."""
    if pd.isna(value):
        return "N/A"

    try:
        value = float(value)
    except Exception:
        return str(value)

    prefix = "₹" if currency else ""

    absolute = abs(value)
    if absolute >= 1_000_000_000:
        return f"{prefix}{value / 1_000_000_000:,.2f} B"
    if absolute >= 1_000_000:
        return f"{prefix}{value / 1_000_000:,.2f} M"
    if absolute >= 100_000:
        return f"{prefix}{value / 100_000:,.2f} L"
    if absolute >= 1_000:
        return f"{prefix}{value / 1_000:,.2f} K"

    return f"{prefix}{value:,.2f}" if currency else f"{value:,.2f}"


def find_column(df, keywords, numeric_only=False):
    """Find a likely column using keyword matching."""
    for column in df.columns:
        name = str(column).lower()
        if any(keyword in name for keyword in keywords):
            if numeric_only and not pd.api.types.is_numeric_dtype(df[column]):
                continue
            return column
    return None


def get_numeric_columns(df):
    return df.select_dtypes(include=np.number).columns.tolist()


def get_categorical_columns(df):
    return [
        c for c in df.columns
        if (
            pd.api.types.is_object_dtype(df[c])
            or pd.api.types.is_categorical_dtype(df[c])
            or pd.api.types.is_bool_dtype(df[c])
        )
        and df[c].nunique(dropna=True) > 1
    ]


def get_date_columns(df):
    date_columns = []

    for column in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[column]):
            date_columns.append(column)
        elif any(x in str(column).lower() for x in ["date", "time", "timestamp"]):
            parsed = pd.to_datetime(df[column], errors="coerce")
            if parsed.notna().sum() > 0:
                date_columns.append(column)

    return date_columns


def prepare_dates(df):
    """Convert likely date columns to datetime where parsing is reasonable."""
    result = df.copy()

    for column in result.columns:
        name = str(column).lower()

        if (
            pd.api.types.is_datetime64_any_dtype(result[column])
            or any(x in name for x in ["date", "timestamp", "datetime"])
        ):
            converted = pd.to_datetime(result[column], errors="coerce")
            if converted.notna().sum() > 0:
                result[column] = converted

    return result


def generate_kpis(df, domain):
    """Return business-oriented KPIs; keep data-quality counts separate."""
    kpis = []
    if df.empty:
        return [("Records Analyzed", "0")]

    amount_col = find_column(df, ["revenue", "sales", "transaction amount", "amount", "total", "income", "budget"], numeric_only=True)
    profit_col = find_column(df, ["profit", "net income", "net_profit"], numeric_only=True)
    cost_col = find_column(df, ["cost", "expense", "cogs"], numeric_only=True)
    quantity_col = find_column(df, ["quantity", "units", "items"], numeric_only=True)
    rating_col = find_column(df, ["rating", "vote average", "score"], numeric_only=True)
    date_cols = get_date_columns(df)
    id_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["order_id", "transaction_id", "customer_id", "booking_id", "invoice_id"])), None)

    if amount_col:
        kpis.append((f"Total {amount_col}", format_number(df[amount_col].sum(), currency=True)))
        kpis.append((f"Average {amount_col}", format_number(df[amount_col].mean(), currency=True)))
        if id_col:
            kpis.append(("Average Value per Record", format_number(df[amount_col].sum() / max(df[id_col].nunique(), 1), currency=True)))
    if profit_col:
        kpis.append((f"Total {profit_col}", format_number(df[profit_col].sum(), currency=True)))
        if amount_col and df[amount_col].sum() != 0:
            kpis.append(("Profit Margin", f"{df[profit_col].sum()/df[amount_col].sum()*100:.2f}%"))
    elif amount_col and cost_col:
        profit = df[amount_col].sum() - df[cost_col].sum()
        kpis.append(("Estimated Profit", format_number(profit, currency=True)))
        if df[amount_col].sum() != 0:
            kpis.append(("Estimated Margin", f"{profit/df[amount_col].sum()*100:.2f}%"))
    if quantity_col:
        kpis.append((f"Total {quantity_col}", format_number(df[quantity_col].sum())))
    if rating_col:
        kpis.append((f"Average {rating_col}", f"{df[rating_col].mean():.2f}"))
    if date_cols:
        dates = pd.to_datetime(df[date_cols[0]], errors="coerce").dropna()
        if not dates.empty:
            kpis.append(("Analysis Period", f"{dates.min():%Y-%m} – {dates.max():%Y-%m}"))
    if not kpis:
        numeric = get_numeric_columns(df)
        if numeric:
            c = numeric[0]
            kpis.extend([(f"Average {c}", format_number(df[c].mean())), (f"Maximum {c}", format_number(df[c].max())), (f"Minimum {c}", format_number(df[c].min()))])
        kpis.append(("Records Analyzed", format_number(len(df))))
    # Deduplicate labels and never insert data-quality metrics as business KPIs.
    unique=[]
    seen=set()
    for item in kpis:
        if item[0] not in seen:
            unique.append(item); seen.add(item[0])
    return unique[:8]


def create_findings(df):
    """Generate data-grounded findings and actionable recommendations."""
    findings, recommendations = [], []
    if df.empty:
        return ["No records are available for analysis."], ["Upload a non-empty dataset before interpreting results."]

    n, m = len(df), len(df.columns)
    missing_by_col = df.isna().sum().sort_values(ascending=False)
    missing_total = int(missing_by_col.sum())
    duplicate_rows = int(df.duplicated().sum())
    findings.append(f"The working dataset contains {n:,} records across {m:,} columns.")
    if missing_total:
        top = missing_by_col[missing_by_col > 0].head(3)
        detail = "; ".join(f"{c}: {v:,} ({v/n:.1%})" for c,v in top.items())
        findings.append(f"{missing_total:,} missing cells ({missing_total/(n*m):.2%} of all cells); highest affected columns: {detail}.")
        recommendations.append("Review missingness by column and business meaning; use a justified imputation, source lookup, or documented exclusion rather than blanket filling.")
    else:
        findings.append("No missing cells were detected in the current working dataset.")
    if duplicate_rows:
        findings.append(f"{duplicate_rows:,} fully duplicated rows ({duplicate_rows/n:.2%} of records) were found.")
        recommendations.append("Inspect duplicate records and confirm whether they are repeated events or accidental copies before removal.")
    else:
        findings.append("No fully duplicated rows were detected.")

    id_candidates=[c for c in df.columns if any(k in str(c).lower() for k in ["id", "key"])]
    for c in id_candidates[:3]:
        dup=int(df[c].duplicated(keep=False).sum())
        null=int(df[c].isna().sum())
        if dup or null:
            findings.append(f"Identifier candidate '{c}' has {dup:,} rows with repeated values and {null:,} missing values; uniqueness is not confirmed.")
            recommendations.append(f"Validate whether '{c}' should uniquely identify rows; if not, define a composite key or document valid repeats.")

    numeric=get_numeric_columns(df)
    for c in numeric:
        vals=pd.to_numeric(df[c], errors="coerce").dropna()
        if len(vals)>=2:
            if any(k in str(c).lower() for k in ["revenue","sales","amount","profit","cost","price","budget","balance","loan"]):
                findings.append(f"'{c}' ranges from {vals.min():,.2f} to {vals.max():,.2f}, with mean {vals.mean():,.2f} and median {vals.median():,.2f}.")
                if vals.skew() > 1:
                    findings.append(f"'{c}' is positively skewed (skewness {vals.skew():.2f}); a small number of high values may influence the mean.")
                    recommendations.append(f"Compare median and mean for '{c}' and review the highest-value records before making typical-value claims.")
                break
    cats=get_categorical_columns(df)
    for c in cats:
        counts=df[c].fillna("Missing").value_counts()
        if len(counts):
            findings.append(f"'{c}' has {df[c].nunique(dropna=True):,} distinct non-missing categories; the most frequent is '{counts.index[0]}' ({counts.iloc[0]:,} records, {counts.iloc[0]/n:.1%}).")
            recommendations.append(f"Compare key outcomes across '{c}' categories to identify meaningful differences; avoid interpreting record count alone as performance.")
            break
    dates=get_date_columns(df)
    if dates:
        d=pd.to_datetime(df[dates[0]], errors="coerce").dropna()
        if not d.empty:
            findings.append(f"'{dates[0]}' spans {d.min().date()} through {d.max().date()} ({d.nunique():,} distinct date values parsed).")
            recommendations.append(f"Use '{dates[0]}' for trend analysis after confirming it represents the relevant business event date and checking for gaps.")
    if not recommendations:
        recommendations.append("Confirm field definitions with the data owner, validate the measures against source totals, and choose analysis questions aligned to the dataset's purpose.")
    return findings, list(dict.fromkeys(recommendations))



def build_report_charts(df):
    """Build compact, data-supported figures for PDF/PPT export."""
    figs=[]
    nums=get_numeric_columns(df); cats=get_categorical_columns(df); dates=get_date_columns(df)
    amount=find_column(df,["revenue","sales","transaction amount","amount","total","income"],numeric_only=True)
    if dates and amount:
        tmp=df[[dates[0],amount]].copy(); tmp[dates[0]]=pd.to_datetime(tmp[dates[0]],errors="coerce"); tmp=tmp.dropna()
        if not tmp.empty:
            tmp["Period"]=tmp[dates[0]].dt.to_period("M").astype(str)
            g=tmp.groupby("Period")[amount].sum().reset_index()
            figs.append(px.line(g,x="Period",y=amount,markers=True,title=f"{amount} over time",template=PLOTLY_TEMPLATE))
    if cats and amount:
        c=cats[0]; g=df.groupby(c,dropna=False)[amount].sum().reset_index().nlargest(12,amount)
        figs.append(px.bar(g,x=amount,y=c,orientation="h",title=f"{amount} by {c}",template=PLOTLY_TEMPLATE))
    elif nums:
        c=amount or nums[0]
        figs.append(px.histogram(df,x=c,nbins=25,title=f"Distribution of {c}",template=PLOTLY_TEMPLATE))
    if len(nums)>=2:
        figs.append(px.imshow(df[nums].corr(numeric_only=True),text_auto=".2f",zmin=-1,zmax=1,title="Numeric correlation",template=PLOTLY_TEMPLATE))
    return figs

def add_chart_to_ppt(slide, fig, left, top, width, height):
    """Export a Plotly figure to a slide image when Kaleido is available."""
    try:
        image_bytes = fig.to_image(format="png", width=1200, height=650, scale=1)
        image_stream = io.BytesIO(image_bytes)
        slide.shapes.add_picture(
            image_stream,
            Inches(left),
            Inches(top),
            width=Inches(width),
            height=Inches(height),
        )
        return True
    except Exception:
        return False


# ================================================================
# SIDEBAR: WORKFLOW MENU
# ================================================================

st.sidebar.markdown("## 📊 ETL Workspace")
st.sidebar.caption("Raw data → Cleaning → Quality → Analysis → Dashboard → Reports")

menu_items = [
    "Upload Data",
    "Data Overview",
    "Data Cleaning",
    "Data Quality",
    "Explore & Analyze",
    "Visualize",
    "Dashboard",
    "Findings & Recommendations",
    "Export Reports",
]

selected_menu = st.sidebar.radio(
    "Workflow",
    menu_items,
    index=0,
)

st.sidebar.divider()
st.sidebar.markdown("### Current Dataset")

if st.session_state.raw_df is not None:
    st.sidebar.success(st.session_state.file_name or "Dataset loaded")
    st.sidebar.write(f"Rows: **{len(st.session_state.raw_df):,}**")
    st.sidebar.write(f"Columns: **{len(st.session_state.raw_df.columns):,}**")
    st.sidebar.write(f"Domain: **{st.session_state.selected_domain}**")
else:
    st.sidebar.info("Upload a dataset to begin.")


# ================================================================
# HEADER
# ================================================================

st.markdown(
    """
    <div class="main-title">Automated ETL & Interactive Dashboard System</div>
    <div class="subtitle">
        A data analyst workflow for uploading, cleaning, validating,
        exploring, visualizing, and reporting business data.
    </div>
    """,
    unsafe_allow_html=True,
)
st.write("")


# ================================================================
# 1. UPLOAD DATA
# ================================================================

if selected_menu == "Upload Data":
    st.header("Upload Raw Dataset")

    st.write(
        "Upload a CSV, Excel, or JSON file. The application will inspect "
        "the columns and suggest suitable cleaning and analysis steps."
    )

    uploaded_file = st.file_uploader(
        "Choose your dataset",
        type=["csv", "xlsx", "xls", "json"],
    )

    if uploaded_file is not None:
        try:
            raw_df = read_uploaded_file(uploaded_file)

            if raw_df.empty:
                st.error("The uploaded file contains no records.")
            else:
                st.session_state.raw_df = raw_df.copy()
                st.session_state.clean_df = raw_df.copy()
                st.session_state.file_name = uploaded_file.name
                st.session_state.selected_domain = detect_domain(raw_df)
                st.session_state.cleaning_log = []
                st.session_state.quality_report = None

                st.success("Dataset uploaded successfully.")
                st.subheader("Raw Dataset Preview")
                st.dataframe(raw_df.head(20), use_container_width=True)

                col1, col2, col3 = st.columns(3)
                col1.metric("Rows", f"{len(raw_df):,}")
                col2.metric("Columns", f"{len(raw_df.columns):,}")
                col3.metric("Detected Domain", st.session_state.selected_domain)

                st.info(
                    "Next step: open **2. Data Overview** to inspect the dataset, "
                    "then continue through the workflow."
                )

        except Exception as error:
            st.error(f"Unable to read the file: {error}")


# ================================================================
# CHECK DATASET
# ================================================================

if selected_menu != "Upload Data" and st.session_state.raw_df is None:
    st.warning("Please upload a dataset first using **1. Upload Data**.")

elif selected_menu != "Upload Data":

    raw_df = st.session_state.raw_df.copy()
    clean_df = (
        st.session_state.clean_df.copy()
        if st.session_state.clean_df is not None
        else raw_df.copy()
    )

    # ============================================================
    # 2. DATA OVERVIEW
    # ============================================================

    if selected_menu == "Data Overview":
        st.header("Data Overview")

        st.subheader("Dataset Summary")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Rows", f"{len(raw_df):,}")
        col2.metric("Columns", f"{len(raw_df.columns):,}")
        col3.metric("Missing Cells", f"{raw_df.isna().sum().sum():,}")
        col4.metric("Duplicate Rows", f"{raw_df.duplicated().sum():,}")

        st.subheader("Column Information")

        overview = pd.DataFrame({
            "Column": raw_df.columns,
            "Data Type": [str(raw_df[c].dtype) for c in raw_df.columns],
            "Non-Null Count": [int(raw_df[c].notna().sum()) for c in raw_df.columns],
            "Missing Count": [int(raw_df[c].isna().sum()) for c in raw_df.columns],
            "Unique Values": [int(raw_df[c].nunique(dropna=True)) for c in raw_df.columns],
            "Suggested Role": [
                infer_column_role(c, raw_df[c]) for c in raw_df.columns
            ],
        })

        st.dataframe(overview, use_container_width=True)

        st.subheader("Raw Data Preview")
        st.dataframe(raw_df.head(50), use_container_width=True)

        st.subheader("Statistical Summary")
        st.dataframe(raw_df.describe(include="all").transpose(), use_container_width=True)


    # ============================================================
    # 3. DATA CLEANING
    # ============================================================

    elif selected_menu == "Data Cleaning":
        st.header("Data Cleaning")

        st.warning(
            "Cleaning changes are applied to the working dataset. "
            "The original raw dataset remains available."
        )

        st.subheader("Column-wise Cleaning Recommendations")
        recommendations_df = get_column_recommendations(clean_df)
        st.dataframe(recommendations_df, use_container_width=True)

        st.subheader("Primary Key Suggestions")
        possible_keys = suggest_primary_key(clean_df)

        if possible_keys:
            st.success(
                "Potential unique key column(s): " + ", ".join(map(str, possible_keys))
            )
        else:
            st.info(
                "No fully unique, non-null column was detected. "
                "A composite key may be required."
            )

        with st.expander("Inspect possible identifier columns"):
            identifier_columns = [
                c for c in clean_df.columns
                if any(x in str(c).lower() for x in ["id", "key", "code"])
            ]

            if identifier_columns:
                for column in identifier_columns:
                    st.write(
                        f"**{column}** — unique: "
                        f"{clean_df[column].nunique(dropna=True):,} / "
                        f"{len(clean_df):,}; missing: "
                        f"{clean_df[column].isna().sum():,}"
                    )
            else:
                st.write("No obvious ID/key column names were found.")

        st.subheader("Cleaning Actions")

        # Convert likely date columns
        date_candidates = [
            c for c in clean_df.columns
            if any(x in str(c).lower() for x in ["date", "timestamp", "datetime"])
        ]

        if date_candidates:
            selected_date = st.selectbox(
                "Select a date column to convert",
                ["Do not convert"] + date_candidates,
            )

            if selected_date != "Do not convert":
                if st.button("Convert selected column to Date/Time"):
                    converted = pd.to_datetime(
                        clean_df[selected_date],
                        errors="coerce",
                    )
                    clean_df[selected_date] = converted
                    st.session_state.clean_df = clean_df
                    st.session_state.cleaning_log.append(
                        f"Converted '{selected_date}' to datetime."
                    )
                    st.success("Date conversion completed.")
                    st.rerun()

        # Numeric conversion
        with st.expander("Convert a column to numeric"):
            selected_numeric_conversion = st.selectbox(
                "Column to convert",
                ["Select column"] + list(clean_df.columns),
                key="numeric_conversion_column",
            )

            if selected_numeric_conversion != "Select column":
                if st.button("Convert selected column to numeric"):
                    clean_df[selected_numeric_conversion] = pd.to_numeric(
                        clean_df[selected_numeric_conversion],
                        errors="coerce",
                    )
                    st.session_state.clean_df = clean_df
                    st.session_state.cleaning_log.append(
                        f"Converted '{selected_numeric_conversion}' to numeric."
                    )
                    st.success("Numeric conversion completed.")
                    st.rerun()

        # Missing values
        st.subheader("Handle Missing Values")

        missing_columns = [
            c for c in clean_df.columns if clean_df[c].isna().any()
        ]

        if not missing_columns:
            st.success("No missing values are present in the working dataset.")
        else:
            selected_missing_column = st.selectbox(
                "Select a column with missing values",
                missing_columns,
                key="missing_column",
            )

            series = clean_df[selected_missing_column]
            st.write(
                f"Missing values: **{series.isna().sum():,}** "
                f"({series.isna().mean() * 100:.2f}%)"
            )

            options = ["Keep missing values", "Drop rows with missing value"]

            if pd.api.types.is_numeric_dtype(series):
                options += ["Fill with mean", "Fill with median", "Fill with zero"]
            else:
                options += ["Fill with mode", "Fill with 'Unknown'"]

            missing_action = st.selectbox(
                "Choose an action",
                options,
                key="missing_action",
            )

            if st.button("Apply Missing-Value Action"):
                if missing_action == "Drop rows with missing value":
                    before = len(clean_df)
                    clean_df = clean_df.dropna(subset=[selected_missing_column])
                    action_text = (
                        f"Dropped {before - len(clean_df):,} rows with missing "
                        f"'{selected_missing_column}'."
                    )

                elif missing_action == "Fill with mean":
                    fill_value = clean_df[selected_missing_column].mean()
                    clean_df[selected_missing_column] = clean_df[
                        selected_missing_column
                    ].fillna(fill_value)
                    action_text = (
                        f"Filled missing '{selected_missing_column}' values "
                        f"with mean ({fill_value})."
                    )

                elif missing_action == "Fill with median":
                    fill_value = clean_df[selected_missing_column].median()
                    clean_df[selected_missing_column] = clean_df[
                        selected_missing_column
                    ].fillna(fill_value)
                    action_text = (
                        f"Filled missing '{selected_missing_column}' values "
                        f"with median ({fill_value})."
                    )

                elif missing_action == "Fill with zero":
                    clean_df[selected_missing_column] = clean_df[
                        selected_missing_column
                    ].fillna(0)
                    action_text = (
                        f"Filled missing '{selected_missing_column}' values with 0."
                    )

                elif missing_action == "Fill with mode":
                    mode_values = clean_df[selected_missing_column].mode(dropna=True)
                    if len(mode_values):
                        fill_value = mode_values.iloc[0]
                        clean_df[selected_missing_column] = clean_df[
                            selected_missing_column
                        ].fillna(fill_value)
                        action_text = (
                            f"Filled missing '{selected_missing_column}' values "
                            f"with mode ({fill_value})."
                        )
                    else:
                        action_text = "No mode was available; no change was made."

                elif missing_action == "Fill with 'Unknown'":
                    clean_df[selected_missing_column] = clean_df[
                        selected_missing_column
                    ].fillna("Unknown")
                    action_text = (
                        f"Filled missing '{selected_missing_column}' values with Unknown."
                    )

                else:
                    action_text = "Missing values were retained."

                st.session_state.clean_df = clean_df
                st.session_state.cleaning_log.append(action_text)
                st.success(action_text)
                st.rerun()

        # Duplicate rows
        st.subheader("Duplicate Row Handling")
        duplicate_count = int(clean_df.duplicated().sum())

        st.write(f"Fully duplicated rows: **{duplicate_count:,}**")

        if duplicate_count > 0:
            if st.button("Remove Fully Duplicated Rows"):
                before = len(clean_df)
                clean_df = clean_df.drop_duplicates()
                removed = before - len(clean_df)
                st.session_state.clean_df = clean_df
                st.session_state.cleaning_log.append(
                    f"Removed {removed:,} fully duplicated rows."
                )
                st.success(f"Removed {removed:,} fully duplicated rows.")
                st.rerun()
        else:
            st.info("No fully duplicated rows were detected.")

        # Identifier duplicates
        st.subheader("Identifier Duplicate Check")
        possible_id_columns = [
            c for c in clean_df.columns
            if any(x in str(c).lower() for x in ["id", "key"])
        ]

        if possible_id_columns:
            id_column = st.selectbox(
                "Select identifier column",
                possible_id_columns,
                key="id_duplicate_column",
            )

            duplicate_id_count = int(clean_df[id_column].duplicated().sum())
            st.write(f"Duplicate values in **{id_column}**: {duplicate_id_count:,}")

            st.caption(
                "Do not remove duplicate identifier values automatically. "
                "First confirm whether the identifier should be unique and "
                "whether repeated values represent valid transactions or records."
            )

        # Outlier review
        st.subheader("Outlier Review")
        numeric_columns = get_numeric_columns(clean_df)

        if numeric_columns:
            outlier_column = st.selectbox(
                "Select numeric column",
                numeric_columns,
                key="outlier_column",
            )

            values = clean_df[outlier_column].dropna()

            if len(values) >= 4:
                q1 = values.quantile(0.25)
                q3 = values.quantile(0.75)
                iqr = q3 - q1
                lower_bound = q1 - 1.5 * iqr
                upper_bound = q3 + 1.5 * iqr

                outlier_count = int(
                    ((clean_df[outlier_column] < lower_bound) |
                     (clean_df[outlier_column] > upper_bound)).sum()
                )

                st.write(f"Potential outliers using IQR: **{outlier_count:,}**")
                st.write(f"Lower bound: `{lower_bound:,.3f}`")
                st.write(f"Upper bound: `{upper_bound:,.3f}`")
                st.warning(
                    "Potential outliers are flagged for review only. "
                    "They are not automatically deleted because extreme values "
                    "may represent valid business events."
                )

        # Cleaning log
        st.subheader("Cleaning Log")
        if st.session_state.cleaning_log:
            for i, log_item in enumerate(st.session_state.cleaning_log, 1):
                st.write(f"{i}. {log_item}")
        else:
            st.info("No cleaning actions have been applied yet.")

        st.subheader("Current Cleaned Dataset")
        st.dataframe(clean_df.head(50), use_container_width=True)

        csv_bytes = clean_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Cleaned Dataset (CSV)",
            data=csv_bytes,
            file_name=f"{safe_name(st.session_state.file_name)}_cleaned.csv",
            mime="text/csv",
        )


    # ============================================================
    # 4. DATA QUALITY
    # ============================================================

    elif selected_menu == "Data Quality":
        st.header("Data Quality Validation")

        quality = create_quality_report(clean_df)
        st.session_state.quality_report = quality

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Rows", f"{quality['rows']:,}")
        col2.metric("Columns", f"{quality['columns']:,}")
        col3.metric("Missing Cells", f"{quality['total_missing']:,}")
        col4.metric("Duplicate Rows", f"{quality['duplicate_rows']:,}")

        st.subheader("Column-level Quality Report")
        st.dataframe(quality["column_report"], use_container_width=True)

        st.subheader("Quality Checks")

        checks = [
            ("Missing values", quality["total_missing"] == 0),
            ("Fully duplicated rows", quality["duplicate_rows"] == 0),
            ("Empty dataset", quality["rows"] > 0),
        ]

        for check_name, passed in checks:
            if passed:
                st.success(f"PASS — {check_name}")
            else:
                st.warning(f"REVIEW — {check_name}")

        if quality["duplicate_id_counts"]:
            st.subheader("Possible Identifier Duplicate Counts")
            st.dataframe(
                pd.DataFrame(
                    [
                        {"Identifier Column": column, "Duplicate Values": count}
                        for column, count in quality["duplicate_id_counts"].items()
                    ]
                ),
                use_container_width=True,
            )

        st.caption(
            "A quality check flags potential issues. Whether a value is invalid "
            "depends on the dataset's business rules and context."
        )

        quality_csv = quality["column_report"].to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Quality Report (CSV)",
            data=quality_csv,
            file_name="data_quality_report.csv",
            mime="text/csv",
        )


    # ============================================================
    # 5. EXPLORE & ANALYZE
    # ============================================================

    elif selected_menu == "Explore & Analyze":
        st.header("Explore & Analyze")

        st.subheader("Numeric Summary")
        numeric_columns = get_numeric_columns(clean_df)

        if numeric_columns:
            st.dataframe(
                clean_df[numeric_columns].describe().transpose(),
                use_container_width=True,
            )
        else:
            st.info("No numeric columns are available.")

        st.subheader("Category Summary")
        categorical_columns = get_categorical_columns(clean_df)

        if categorical_columns:
            selected_category = st.selectbox(
                "Select category column",
                categorical_columns,
                key="category_summary",
            )

            category_summary = (
                clean_df[selected_category]
                .fillna("Missing")
                .value_counts()
                .head(20)
                .rename_axis(selected_category)
                .reset_index(name="Record Count")
            )

            st.dataframe(category_summary, use_container_width=True)

            fig = px.bar(
                category_summary,
                x=selected_category,
                y="Record Count",
                title=f"Top Categories — {selected_category}",
                color="Record Count",
                color_continuous_scale="Blues",
            )
            fig.update_layout(
                template=PLOTLY_TEMPLATE,
                margin=dict(l=15, r=15, t=55, b=15),
            )
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Correlation Analysis")

        if len(numeric_columns) >= 2:
            correlation = clean_df[numeric_columns].corr(numeric_only=True)

            fig = px.imshow(
                correlation,
                text_auto=".2f",
                color_continuous_scale="RdBu_r",
                zmin=-1,
                zmax=1,
                title="Numeric Column Correlation Heatmap",
            )
            fig.update_layout(template=PLOTLY_TEMPLATE)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("At least two numeric columns are required for correlation analysis.")

        st.subheader("Custom Analysis")

        if numeric_columns:
            metric_column = st.selectbox(
                "Numeric metric",
                numeric_columns,
                key="custom_metric",
            )

            operation = st.selectbox(
                "Aggregation",
                ["Sum", "Mean", "Median", "Minimum", "Maximum", "Count"],
                key="custom_operation",
            )

            if operation == "Sum":
                result = clean_df[metric_column].sum()
            elif operation == "Mean":
                result = clean_df[metric_column].mean()
            elif operation == "Median":
                result = clean_df[metric_column].median()
            elif operation == "Minimum":
                result = clean_df[metric_column].min()
            elif operation == "Maximum":
                result = clean_df[metric_column].max()
            else:
                result = clean_df[metric_column].count()

            st.metric(f"{operation} of {metric_column}", format_number(result))


    # ============================================================
    # 6. VISUALIZE
    # ============================================================

    elif selected_menu == "Visualize":
        st.header("Visualize Data")

        numeric_columns = get_numeric_columns(clean_df)
        categorical_columns = get_categorical_columns(clean_df)
        date_columns = get_date_columns(clean_df)
        chart_df_source = clean_df.copy()
        # For numeric-only datasets, create temporary quantile bands so categorical charts remain usable.
        chart_categories = list(categorical_columns)
        if not chart_categories and numeric_columns:
            for _c in numeric_columns[:3]:
                try:
                    _band = pd.qcut(chart_df_source[_c], q=min(5, chart_df_source[_c].nunique()), duplicates="drop")
                    _name = f"{_c} (value bands)"
                    chart_df_source[_name] = _band.astype(str)
                    chart_categories.append(_name)
                except Exception:
                    pass

        chart_type = st.selectbox(
            "Choose chart type",
            [
                "Bar Chart",
                "Column Chart",
                "Line Chart",
                "Pie Chart",
                "Donut Chart",
                "Scatter Plot",
                "Histogram",
                "Box Plot",
                "Treemap",
                "Heatmap",
                "Gauge",
                "Map",
            ],
        )

        fig = None

        if chart_type in ["Bar Chart", "Column Chart"]:
            if chart_categories and numeric_columns:
                category_col = st.selectbox("Category", chart_categories)
                metric_col = st.selectbox("Metric", numeric_columns)
                aggregation = st.selectbox("Aggregation", ["Sum", "Mean", "Count"])

                grouped = chart_df_source.groupby(
                    category_col,
                    dropna=False,
                )[metric_col]

                if aggregation == "Sum":
                    chart_df = grouped.sum().reset_index()
                elif aggregation == "Mean":
                    chart_df = grouped.mean().reset_index()
                else:
                    chart_df = grouped.count().reset_index()

                chart_df = chart_df.sort_values(metric_col, ascending=False).head(30)

                if chart_type == "Bar Chart":
                    fig = px.bar(
                        chart_df,
                        x=metric_col,
                        y=category_col,
                        orientation="h",
                        title=f"{aggregation} of {metric_col} by {category_col}",
                        color=metric_col,
                        color_continuous_scale="Blues",
                    )
                    fig.update_layout(yaxis={"categoryorder": "total ascending"})
                else:
                    fig = px.bar(
                        chart_df,
                        x=category_col,
                        y=metric_col,
                        title=f"{aggregation} of {metric_col} by {category_col}",
                        color=metric_col,
                        color_continuous_scale="Blues",
                    )
            else:
                st.info("A numeric column is required. If no category exists, use the generated numeric value bands or choose Histogram/Scatter Plot.")

        elif chart_type == "Line Chart":
            if date_columns and numeric_columns:
                date_col = st.selectbox("Date", date_columns)
                metric_col = st.selectbox("Metric", numeric_columns)
                aggregation = st.selectbox("Aggregation", ["Sum", "Mean", "Count"])

                temp = clean_df[[date_col, metric_col]].copy()
                temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
                temp = temp.dropna(subset=[date_col])
                temp["Period"] = temp[date_col].dt.to_period("M").astype(str)

                if aggregation == "Sum":
                    chart_df = temp.groupby("Period")[metric_col].sum().reset_index()
                elif aggregation == "Mean":
                    chart_df = temp.groupby("Period")[metric_col].mean().reset_index()
                else:
                    chart_df = temp.groupby("Period")[metric_col].count().reset_index()

                fig = px.line(
                    chart_df,
                    x="Period",
                    y=metric_col,
                    markers=True,
                    title=f"{aggregation} of {metric_col} over time",
                )
            else:
                st.info("A date column and numeric column are required.")

        elif chart_type in ["Pie Chart", "Donut Chart"]:
            if chart_categories:
                category_col = st.selectbox("Category", chart_categories)
                top_n = st.slider("Number of categories", 3, 15, 8)

                chart_df = (
                    chart_df_source[category_col]
                    .fillna("Missing")
                    .value_counts()
                    .head(top_n)
                    .rename_axis(category_col)
                    .reset_index(name="Record Count")
                )

                fig = px.pie(
                    chart_df,
                    names=category_col,
                    values="Record Count",
                    hole=0.45 if chart_type == "Donut Chart" else 0,
                    title=f"Record Distribution by {category_col}",
                )
            else:
                st.info("No category was detected. Numeric value bands are generated when possible; otherwise choose Histogram or Box Plot.")

        elif chart_type == "Scatter Plot":
            if len(numeric_columns) >= 2:
                x_col = st.selectbox("X-axis", numeric_columns)
                y_col = st.selectbox(
                    "Y-axis",
                    [c for c in numeric_columns if c != x_col],
                )

                color_col = st.selectbox(
                    "Color grouping (optional)",
                    ["None"] + categorical_columns,
                )

                fig = px.scatter(
                    clean_df,
                    x=x_col,
                    y=y_col,
                    color=None if color_col == "None" else color_col,
                    title=f"{y_col} vs {x_col}",
                    trendline=None,
                )
            else:
                st.info("At least two numeric columns are required.")

        elif chart_type == "Histogram":
            if numeric_columns:
                metric_col = st.selectbox("Numeric column", numeric_columns)
                bins = st.slider("Bins", 5, 100, 30)

                fig = px.histogram(
                    clean_df,
                    x=metric_col,
                    nbins=bins,
                    title=f"Distribution of {metric_col}",
                )
            else:
                st.info("A numeric column is required.")

        elif chart_type == "Box Plot":
            if numeric_columns:
                metric_col = st.selectbox("Numeric column", numeric_columns)
                category_col = st.selectbox(
                    "Group by (optional)",
                    ["None"] + categorical_columns,
                )

                fig = px.box(
                    clean_df,
                    x=None if category_col == "None" else category_col,
                    y=metric_col,
                    title=f"Distribution and Potential Outliers — {metric_col}",
                )
            else:
                st.info("A numeric column is required.")

        elif chart_type == "Treemap":
            if chart_categories and numeric_columns:
                category_col = st.selectbox("Category", chart_categories)
                metric_col = st.selectbox("Metric", numeric_columns)

                chart_df = (
                    chart_df_source.groupby(category_col, dropna=False)[metric_col]
                    .sum()
                    .reset_index()
                    .sort_values(metric_col, ascending=False)
                    .head(30)
                )

                fig = px.treemap(
                    chart_df,
                    path=[category_col],
                    values=metric_col,
                    title=f"{metric_col} by {category_col}",
                )
            else:
                st.info("A numeric column is required. Numeric-only datasets use generated value bands as treemap categories when possible.")

        elif chart_type == "Heatmap":
            if len(numeric_columns) >= 2:
                corr = clean_df[numeric_columns].corr(numeric_only=True)
                fig = px.imshow(
                    corr,
                    text_auto=".2f",
                    color_continuous_scale="RdBu_r",
                    zmin=-1,
                    zmax=1,
                    title="Correlation Heatmap",
                )
            else:
                st.info("At least two numeric columns are required.")

        elif chart_type == "Gauge":
            if numeric_columns:
                metric_col = st.selectbox("Metric", numeric_columns)
                value = float(clean_df[metric_col].mean())

                fig = go.Figure(
                    go.Indicator(
                        mode="gauge+number",
                        value=value,
                        title={"text": f"Average {metric_col}"},
                        gauge={
                            "axis": {"visible": True},
                            "bar": {"color": BLUE},
                            "bgcolor": LIGHT_BLUE,
                        },
                    )
                )
            else:
                st.info("A numeric column is required.")

        elif chart_type == "Map":
            location_candidates = [
                c for c in clean_df.columns
                if any(x in str(c).lower() for x in [
                    "latitude", "longitude", "lat", "lon", "long"
                ])
            ]

            latitude_col = find_column(
                clean_df,
                ["latitude", "lat"],
                numeric_only=True,
            )
            longitude_col = find_column(
                clean_df,
                ["longitude", "long", "lon"],
                numeric_only=True,
            )

            if latitude_col and longitude_col:
                map_df = clean_df.dropna(subset=[latitude_col, longitude_col]).copy()
                fig = px.scatter_mapbox(
                    map_df,
                    lat=latitude_col,
                    lon=longitude_col,
                    zoom=2,
                    title="Records by Geographic Coordinates",
                    mapbox_style="open-street-map",
                )
            else:
                geo_col = find_column(clean_df, ["country", " 국가", "state", "province", "region"])
                metric_col = find_column(clean_df, ["sales", "revenue", "amount", "count", "population"], numeric_only=True)
                if geo_col and metric_col and any(k in str(geo_col).lower() for k in ["country", "state", "province"]):
                    geo_df = clean_df.groupby(geo_col, dropna=False)[metric_col].sum().reset_index()
                    st.info("Using geographic names for a choropleth. Country/state names must be recognized by Plotly; city names may require geocoding.")
                    fig = px.choropleth(geo_df, locations=geo_col, locationmode="country names", color=metric_col, title=f"{metric_col} by {geo_col}", color_continuous_scale="Blues") if "country" in str(geo_col).lower() else None
                    if fig is None:
                        st.warning("State/province map needs a supported geographic code or coordinates. Use a bar chart for this field instead.")
                else:
                    st.info("No usable geographic fields were detected. A map needs latitude/longitude or recognized country codes/names; this is a dataset limitation, not a chart error.")

        if fig is not None:
            fig.update_layout(
                template=PLOTLY_TEMPLATE,
                margin=dict(l=20, r=20, t=60, b=20),
                title_x=0.02,
            )
            st.plotly_chart(fig, use_container_width=True)

            try:
                image_bytes = fig.to_image(format="png", width=1400, height=750)
                st.download_button(
                    "Download Chart as PNG",
                    data=image_bytes,
                    file_name="analysis_chart.png",
                    mime="image/png",
                )
            except Exception:
                st.caption(
                    "PNG export requires a working Kaleido installation. "
                    "The interactive chart is available above."
                )


    # ============================================================
    # 7. DASHBOARD
    # ============================================================

    elif selected_menu == "Dashboard":
        st.header("Interactive Dashboard")

        domain = detect_domain(clean_df)
        date_columns = get_date_columns(clean_df)
        numeric_columns = get_numeric_columns(clean_df)
        categorical_columns = get_categorical_columns(clean_df)

        dashboard_title = st.text_input(
            "Dashboard title",
            value=f"{domain} Analytics Dashboard",
            max_chars=120,
        )

        st.markdown(
            f"""
            <div style="
                background: var(--secondary-background-color);
                border: 1px solid #D9E5F0;
                border-radius: 12px;
                padding: 18px 20px;
                margin: 8px 0 16px 0;
            ">
                <div style="
                    font-size: 1.55rem;
                    font-weight: 800;
                    color: var(--text-color);
                    overflow-wrap: anywhere;
                ">{dashboard_title}</div>
                <div style="color:var(--text-color);opacity:.78;font-size:0.85rem;">
                    Interactive business performance overview
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        # Dashboard filters: prioritize important business fields and use numeric
        # range filters as a fallback, so datasets with few categories still get
        # multiple useful filters. Filter options are based on the full dataset.
        filtered_df = clean_df.copy()
        filter_columns = []
        all_dates = get_date_columns(clean_df)
        all_cats = get_categorical_columns(clean_df)

        preferred_words = [
            "region", "country", "state", "city", "category", "segment",
            "status", "type", "department", "payment", "method", "channel",
            "supplier", "warehouse", "customer", "product", "gender",
            "age group", "year", "month", "language", "genre", "rating",
        ]
        preferred_cats = sorted(
            all_cats,
            key=lambda c: (
                next((i for i, word in enumerate(preferred_words)
                      if word in str(c).lower()), len(preferred_words)),
                str(c).lower(),
            ),
        )

        # Date filter first when present.
        if all_dates:
            filter_columns.append(("Date", all_dates[0]))

        # Categorical filters: allow useful fields with up to 250 values.
        # Avoid very high-cardinality free text and likely unique identifiers.
        for column in preferred_cats:
            name = str(column).lower()
            cardinality = int(clean_df[column].nunique(dropna=True))
            likely_id = any(token in name for token in [" id", "_id", "key", "uuid", "email", "phone"])
            likely_free_text = any(token in name for token in ["description", "comment", "review text", "address"])
            if 2 <= cardinality <= 250 and not likely_id and not likely_free_text:
                filter_columns.append(("Category", column))

        # Low-cardinality numeric columns (e.g. year, rating, status code).
        for column in numeric_columns:
            name = str(column).lower()
            if any(token in name for token in ["id", "key", "phone", "zip", "postal"]):
                continue
            cardinality = int(clean_df[column].nunique(dropna=True))
            if 2 <= cardinality <= 20:
                filter_columns.append(("Category", column))

        # Add numeric range filters if there are fewer than four categorical/date
        # filters. These are especially useful for budget, revenue, score, age,
        # quantity, popularity, and other numeric measures.
        already_selected = {column for _, column in filter_columns}
        numeric_priority = [
            "amount", "revenue", "sales", "budget", "profit", "cost",
            "price", "salary", "loan", "balance", "age", "score",
            "rating", "popularity", "runtime", "quantity", "count",
        ]
        range_candidates = sorted(
            [c for c in numeric_columns if c not in already_selected
             and clean_df[c].nunique(dropna=True) > 20
             and not any(token in str(c).lower() for token in ["id", "key", "phone", "zip", "postal"])],
            key=lambda c: (
                next((i for i, word in enumerate(numeric_priority)
                      if word in str(c).lower()), len(numeric_priority)),
                str(c).lower(),
            ),
        )
        for column in range_candidates:
            filter_columns.append(("Numeric", column))

        unique_filters = []
        seen_filter_cols = set()
        for filter_type, column in filter_columns:
            if column not in seen_filter_cols:
                seen_filter_cols.add(column)
                unique_filters.append((filter_type, column))
        unique_filters = unique_filters[:5]

        if unique_filters:
            st.subheader("Dashboard Filters")
            filter_cols = st.columns(len(unique_filters))
            for i, (filter_type, column) in enumerate(unique_filters):
                with filter_cols[i]:
                    if filter_type == "Date":
                        date_series = pd.to_datetime(clean_df[column], errors="coerce")
                        valid_dates = date_series.dropna()
                        if not valid_dates.empty:
                            min_date = valid_dates.min().date()
                            max_date = valid_dates.max().date()
                            selected_range = st.date_input(
                                str(column), value=(min_date, max_date),
                                min_value=min_date, max_value=max_date,
                                key=f"dashboard_date_{column}",
                            )
                            if isinstance(selected_range, (tuple, list)) and len(selected_range) == 2:
                                start_date, end_date = selected_range
                                date_mask = date_series.dt.date.between(start_date, end_date)
                                filtered_df = filtered_df[date_mask.fillna(False)]
                    elif filter_type == "Numeric":
                        numeric_series = pd.to_numeric(clean_df[column], errors="coerce")
                        valid_values = numeric_series.dropna()
                        if not valid_values.empty:
                            min_value = float(valid_values.min())
                            max_value = float(valid_values.max())
                            if min_value < max_value:
                                selected_range = st.slider(
                                    str(column), min_value=min_value, max_value=max_value,
                                    value=(min_value, max_value), key=f"dashboard_range_{column}",
                                )
                                mask = numeric_series.between(selected_range[0], selected_range[1])
                                filtered_df = filtered_df[mask.fillna(False)]
                            else:
                                st.caption(f"{column}: one distinct value")
                    else:
                        options = sorted(clean_df[column].dropna().astype(str).unique().tolist())
                        selected_values = st.multiselect(
                            str(column), options=options, default=options,
                            key=f"dashboard_filter_{column}",
                            help="Select one or more values. All values are selected by default.",
                        )
                        if selected_values and len(selected_values) < len(options):
                            filtered_df = filtered_df[filtered_df[column].astype(str).isin(selected_values)]
                        elif not selected_values:
                            filtered_df = filtered_df.iloc[0:0]
        else:
            st.info("No suitable columns were detected for interactive dashboard filters.")

        st.caption(
            f"Showing {len(filtered_df):,} of {len(clean_df):,} records."
        )

        # Business KPI cards
        kpis = generate_kpis(filtered_df, domain)

        st.subheader("Key Performance Indicators")

        # Keep horizontal KPI cards; show up to 6 at a time
        visible_kpis = kpis[:6]
        kpi_cols = st.columns(len(visible_kpis))

        for i, (label, value) in enumerate(visible_kpis):
            with kpi_cols[i]:
                st.markdown(
                    f"""
                    <div class="kpi-card">
                        <div class="kpi-label">{label}</div>
                        <div class="kpi-value">{value}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.write("")

        # Charts
        st.subheader("Dashboard Visualizations")

        chart_figures = []

        amount_col = find_column(
            filtered_df,
            ["revenue", "sales", "transaction amount", "amount", "total", "income"],
            numeric_only=True,
        )
        profit_col = find_column(
            filtered_df,
            ["profit", "net income"],
            numeric_only=True,
        )
        date_col = date_columns[0] if date_columns else None

        # Chart 1: Time trend
        if date_col and amount_col:
            temp = filtered_df[[date_col, amount_col]].copy()
            temp[date_col] = pd.to_datetime(temp[date_col], errors="coerce")
            temp = temp.dropna(subset=[date_col])
            temp["Month"] = temp[date_col].dt.to_period("M").astype(str)
            trend_df = temp.groupby("Month")[amount_col].sum().reset_index()

            fig = px.line(
                trend_df,
                x="Month",
                y=amount_col,
                markers=True,
                title=f"{amount_col} Trend",
            )
            fig.update_layout(template=PLOTLY_TEMPLATE)
            chart_figures.append(fig)

        # Chart 2: Category performance
        if categorical_columns and amount_col:
            category_col = categorical_columns[0]
            category_df = (
                filtered_df.groupby(category_col, dropna=False)[amount_col]
                .sum()
                .reset_index()
                .sort_values(amount_col, ascending=False)
                .head(15)
            )

            fig = px.bar(
                category_df,
                x=amount_col,
                y=category_col,
                orientation="h",
                color=amount_col,
                color_continuous_scale="Blues",
                title=f"{amount_col} by {category_col}",
            )
            fig.update_layout(
                template=PLOTLY_TEMPLATE,
                yaxis={"categoryorder": "total ascending"},
            )
            chart_figures.append(fig)

        # Chart 3: Profit or second numeric measure
        if profit_col and categorical_columns:
            category_col = categorical_columns[0]
            profit_df = (
                filtered_df.groupby(category_col, dropna=False)[profit_col]
                .sum()
                .reset_index()
                .sort_values(profit_col, ascending=False)
                .head(15)
            )

            fig = px.bar(
                profit_df,
                x=category_col,
                y=profit_col,
                title=f"{profit_col} by {category_col}",
                color=profit_col,
                color_continuous_scale="Teal",
            )
            fig.update_layout(template=PLOTLY_TEMPLATE)
            chart_figures.append(fig)

        # Chart 4: Distribution pie/donut
        if categorical_columns:
            category_col = categorical_columns[-1]
            distribution = (
                filtered_df[category_col]
                .fillna("Missing")
                .value_counts()
                .head(8)
                .rename_axis(category_col)
                .reset_index(name="Record Count")
            )

            fig = px.pie(
                distribution,
                names=category_col,
                values="Record Count",
                hole=0.45,
                title=f"Record Distribution by {category_col}",
            )
            fig.update_layout(template=PLOTLY_TEMPLATE)
            chart_figures.append(fig)

        # Chart 5: Numeric distribution
        if numeric_columns:
            distribution_col = amount_col or numeric_columns[0]
            fig = px.histogram(
                filtered_df,
                x=distribution_col,
                nbins=30,
                title=f"Distribution of {distribution_col}",
            )
            fig.update_layout(template=PLOTLY_TEMPLATE)
            chart_figures.append(fig)

        # Chart 6: Correlation heatmap
        if len(numeric_columns) >= 2:
            correlation = filtered_df[numeric_columns].corr(numeric_only=True)

            fig = px.imshow(
                correlation,
                text_auto=".2f",
                color_continuous_scale="RdBu_r",
                zmin=-1,
                zmax=1,
                title="Numeric Correlation Heatmap",
            )
            fig.update_layout(template=PLOTLY_TEMPLATE)
            chart_figures.append(fig)

        # Chart 7: Scatter plot
        if len(numeric_columns) >= 2:
            x_col = numeric_columns[0]
            y_col = numeric_columns[1]

            fig = px.scatter(
                filtered_df,
                x=x_col,
                y=y_col,
                title=f"{y_col} vs {x_col}",
                opacity=0.7,
            )
            fig.update_layout(template=PLOTLY_TEMPLATE)
            chart_figures.append(fig)

        # Chart 8: Geographic map, only if coordinates exist
        latitude_col = find_column(
            filtered_df,
            ["latitude", "lat"],
            numeric_only=True,
        )
        longitude_col = find_column(
            filtered_df,
            ["longitude", "long", "lon"],
            numeric_only=True,
        )

        if latitude_col and longitude_col:
            map_df = filtered_df.dropna(subset=[latitude_col, longitude_col])

            fig = px.scatter_mapbox(
                map_df,
                lat=latitude_col,
                lon=longitude_col,
                zoom=2,
                title="Geographic Distribution",
                mapbox_style="open-street-map",
            )
            chart_figures.append(fig)

        if not chart_figures:
            st.info(
                "The dataset does not yet contain enough suitable columns "
                "for dashboard charts. Check the data types and column values."
            )
        else:
            # Responsive two-column layout
            for i in range(0, len(chart_figures), 2):
                left, right = st.columns(2)

                with left:
                    st.plotly_chart(
                        chart_figures[i],
                        use_container_width=True,
                        key=f"dashboard_chart_{i}",
                    )

                if i + 1 < len(chart_figures):
                    with right:
                        st.plotly_chart(
                            chart_figures[i + 1],
                            use_container_width=True,
                            key=f"dashboard_chart_{i+1}",
                        )

        st.subheader("Filtered Data")
        st.dataframe(filtered_df.head(100), use_container_width=True)

        filtered_csv = filtered_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download Dashboard Data (CSV)",
            data=filtered_csv,
            file_name="dashboard_filtered_data.csv",
            mime="text/csv",
        )


    # ============================================================
    # 8. FINDINGS & RECOMMENDATIONS
    # ============================================================

    elif selected_menu == "Findings & Recommendations":
        st.header("Findings & Recommendations")

        findings, recommendations = create_findings(clean_df)

        st.session_state.last_findings = findings
        st.session_state.last_recommendations = recommendations

        st.subheader("Key Findings")

        for i, finding in enumerate(findings, 1):
            st.markdown(f"**{i}.** {finding}")

        st.subheader("Recommendations")

        for i, recommendation in enumerate(recommendations, 1):
            st.markdown(f"**{i}.** {recommendation}")

        st.subheader("Domain Context")
        st.write(f"Detected domain: **{detect_domain(clean_df)}**")

        st.caption(
            "These are automatically generated observations. "
            "Validate them against business definitions and stakeholder context "
            "before presenting them as final conclusions."
        )


    # ============================================================
    # 9. EXPORT REPORTS
    # ============================================================

    elif selected_menu == "Export Reports":
        st.header("Export Reports")

        findings, recommendations = create_findings(clean_df)
        kpis = generate_kpis(clean_df, detect_domain(clean_df))
        quality = create_quality_report(clean_df)

        st.subheader("Report Contents")
        st.write("The report includes:")
        st.write("- Project title and dataset overview")
        st.write("- Raw-to-cleaning workflow and cleaning log")
        st.write("- Data-quality summary")
        st.write("- KPI summary")
        st.write("- Findings and recommendations")
        st.write("- Dashboard charts where chart-image export is available")
        st.write("- Final conclusion")

        # --------------------------------------------------------
        # PDF EXPORT
        # --------------------------------------------------------

        st.subheader("PDF Report")

        if PDF_AVAILABLE:
            if st.button("Generate PDF Report"):
                pdf_buffer = io.BytesIO()

                styles = getSampleStyleSheet()
                styles.add(
                    ParagraphStyle(
                        name="ReportTitle",
                        parent=styles["Title"],
                        textColor=colors.HexColor(NAVY),
                        alignment=TA_CENTER,
                        fontSize=22,
                        leading=28,
                        spaceAfter=16,
                    )
                )
                styles.add(
                    ParagraphStyle(
                        name="SectionHeading",
                        parent=styles["Heading2"],
                        textColor=colors.HexColor(NAVY),
                        spaceBefore=12,
                        spaceAfter=8,
                    )
                )

                document = SimpleDocTemplate(
                    pdf_buffer,
                    pagesize=A4,
                    rightMargin=0.55 * inch,
                    leftMargin=0.55 * inch,
                    topMargin=0.55 * inch,
                    bottomMargin=0.55 * inch,
                )

                story = []

                story.append(Paragraph(
                    "Automated ETL & Dashboard Analysis Report",
                    styles["ReportTitle"],
                ))
                story.append(Paragraph(
                    f"Dataset: {st.session_state.file_name or 'Uploaded Dataset'}",
                    styles["Normal"],
                ))
                story.append(Paragraph(
                    f"Domain: {detect_domain(clean_df)}",
                    styles["Normal"],
                ))
                story.append(Paragraph(
                    f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}",
                    styles["Normal"],
                ))
                story.append(Spacer(1, 14))

                story.append(Paragraph("1. Dataset Overview", styles["SectionHeading"]))

                overview_data = [
                    ["Metric", "Value"],
                    ["Rows", f"{len(raw_df):,}"],
                    ["Columns", f"{len(raw_df.columns):,}"],
                    ["Rows after cleaning", f"{len(clean_df):,}"],
                    ["Missing cells after cleaning", f"{clean_df.isna().sum().sum():,}"],
                    ["Duplicate rows after cleaning", f"{clean_df.duplicated().sum():,}"],
                ]

                overview_table = Table(overview_data, colWidths=[3.3 * inch, 2.5 * inch])
                overview_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9E5F0")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                        colors.white,
                        colors.HexColor("#F4F8FC"),
                    ]),
                    ("PADDING", (0, 0), (-1, -1), 7),
                ]))
                story.append(overview_table)

                story.append(Paragraph("2. ETL Process", styles["SectionHeading"]))
                etl_steps = [
                    "Raw data uploaded and inspected.",
                    "Column names, data types, missing values, and duplicates reviewed.",
                    "Cleaning actions applied only after user selection.",
                    "Data quality checks performed on the working dataset.",
                    "Exploratory analysis and visualizations generated.",
                    "Interactive dashboard prepared with available filters and KPIs.",
                ]

                for step in etl_steps:
                    story.append(Paragraph(f"• {step}", styles["Normal"]))

                if st.session_state.cleaning_log:
                    story.append(Paragraph("Cleaning Log", styles["Heading3"]))
                    for log_item in st.session_state.cleaning_log:
                        story.append(Paragraph(f"• {log_item}", styles["Normal"]))
                else:
                    story.append(Paragraph(
                        "No manual cleaning actions were recorded.",
                        styles["Normal"],
                    ))

                story.append(Paragraph("3. Data Quality", styles["SectionHeading"]))

                quality_data = [["Column", "Type", "Missing", "Unique"]]

                for _, row in quality["column_report"].iterrows():
                    quality_data.append([
                        str(row["Column"])[:30],
                        str(row["Data Type"])[:20],
                        str(row["Missing Values"]),
                        str(row["Unique Values"]),
                    ])

                quality_table = Table(
                    quality_data,
                    repeatRows=1,
                    colWidths=[2.0 * inch, 1.6 * inch, 1.1 * inch, 1.1 * inch],
                )
                quality_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9E5F0")),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("PADDING", (0, 0), (-1, -1), 5),
                ]))
                story.append(quality_table)

                story.append(Paragraph("4. KPI Summary", styles["SectionHeading"]))

                kpi_data = [["KPI", "Value"]]
                for label, value in kpis:
                    kpi_data.append([str(label), str(value)])

                kpi_table = Table(kpi_data, colWidths=[3.3 * inch, 2.5 * inch])
                kpi_table.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#D9E5F0")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [
                        colors.white,
                        colors.HexColor("#F4F8FC"),
                    ]),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]))
                story.append(kpi_table)

                story.append(Paragraph("5. Key Findings", styles["SectionHeading"]))
                for i, finding in enumerate(findings, 1):
                    story.append(Paragraph(f"{i}. {finding}", styles["Normal"]))

                story.append(Paragraph("6. Recommendations", styles["SectionHeading"]))
                for i, recommendation in enumerate(recommendations, 1):
                    story.append(Paragraph(
                        f"{i}. {recommendation}",
                        styles["Normal"],
                    ))

                story.append(Paragraph("7. Visual Analysis", styles["SectionHeading"]))
                for report_fig in build_report_charts(clean_df):
                    try:
                        chart_bytes = report_fig.to_image(format="png", width=1100, height=520, scale=1)
                        story.append(ReportLabImage(io.BytesIO(chart_bytes), width=6.8*inch, height=3.2*inch))
                        story.append(Spacer(1, 8))
                    except Exception:
                        story.append(Paragraph(f"Chart: {report_fig.layout.title.text} (image export unavailable; chart data remains summarized above).", styles["Normal"]))

                story.append(Paragraph("8. Conclusion", styles["SectionHeading"]))
                story.append(Paragraph(
                    "The dataset was processed through a structured analyst workflow "
                    "covering data inspection, cleaning, quality validation, analysis, "
                    "visualization, and reporting. Final business decisions should be "
                    "based on validated definitions and stakeholder context.",
                    styles["Normal"],
                ))

                document.build(story)
                pdf_buffer.seek(0)

                st.download_button(
                    "Download PDF Report",
                    data=pdf_buffer.getvalue(),
                    file_name=f"{safe_name(st.session_state.file_name)}_analysis_report.pdf",
                    mime="application/pdf",
                )
        else:
            st.error(
                "PDF export library is unavailable. Install reportlab using "
                "`python -m pip install reportlab`."
            )

        # --------------------------------------------------------
        # PPT EXPORT
        # --------------------------------------------------------

        st.subheader("PowerPoint Analyst Presentation")

        if PPT_AVAILABLE:
            if st.button("Generate PowerPoint Presentation"):
                prs = Presentation()
                prs.slide_width = Inches(13.333)
                prs.slide_height = Inches(7.5)

                title_layout = prs.slide_layouts[0]
                content_layout = prs.slide_layouts[5]

                # Slide 1: Title
                slide = prs.slides.add_slide(title_layout)
                slide.shapes.title.text = "Data Analysis Project"
                slide.placeholders[1].text = (
                    f"{st.session_state.file_name or 'Dataset'}\n"
                    f"{detect_domain(clean_df)} | Automated ETL & Dashboard"
                )

                # Slide 2: Project workflow
                slide = prs.slides.add_slide(content_layout)
                slide.shapes.title.text = "Project Workflow"

                workflow_text = (
                    "1. Business problem and objectives\n"
                    "2. Raw dataset inspection\n"
                    "3. Data cleaning and transformation\n"
                    "4. Data quality validation\n"
                    "5. Exploratory data analysis\n"
                    "6. KPI development\n"
                    "7. Interactive dashboard\n"
                    "8. Findings, recommendations, and reporting"
                )

                textbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(1.4), Inches(11.5), Inches(4.8)
                )
                textbox.text_frame.text = workflow_text

                # Slide 3: Dataset overview
                slide = prs.slides.add_slide(content_layout)
                slide.shapes.title.text = "Dataset Overview"

                overview_text = (
                    f"Dataset: {st.session_state.file_name or 'Uploaded Dataset'}\n"
                    f"Domain: {detect_domain(clean_df)}\n"
                    f"Raw rows: {len(raw_df):,}\n"
                    f"Raw columns: {len(raw_df.columns):,}\n"
                    f"Cleaned rows: {len(clean_df):,}\n"
                    f"Missing cells after cleaning: {clean_df.isna().sum().sum():,}\n"
                    f"Duplicate rows after cleaning: {clean_df.duplicated().sum():,}"
                )

                textbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(1.4), Inches(11.5), Inches(4.8)
                )
                textbox.text_frame.text = overview_text

                # Slide 4: Cleaning and quality
                slide = prs.slides.add_slide(content_layout)
                slide.shapes.title.text = "Data Cleaning & Quality"

                cleaning_text = (
                    "Cleaning and validation summary\n\n"
                    f"Missing cells: {clean_df.isna().sum().sum():,}\n"
                    f"Duplicate rows: {clean_df.duplicated().sum():,}\n"
                    f"Columns: {len(clean_df.columns):,}\n\n"
                    "Cleaning actions:\n"
                )

                if st.session_state.cleaning_log:
                    cleaning_text += "\n".join(
                        f"• {item}" for item in st.session_state.cleaning_log
                    )
                else:
                    cleaning_text += "No manual cleaning actions recorded."

                textbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(1.4), Inches(11.5), Inches(5.2)
                )
                textbox.text_frame.text = cleaning_text

                # Slide 5: KPI summary
                slide = prs.slides.add_slide(content_layout)
                slide.shapes.title.text = "Key Performance Indicators"

                kpi_text = "\n".join(
                    f"• {label}: {value}" for label, value in kpis[:12]
                )

                textbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(1.4), Inches(11.5), Inches(5.2)
                )
                textbox.text_frame.text = kpi_text

                # Slide 6: Findings
                slide = prs.slides.add_slide(content_layout)
                slide.shapes.title.text = "Key Findings"

                finding_text = "\n".join(
                    f"• {finding}" for finding in findings
                )

                textbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(1.4), Inches(11.5), Inches(5.2)
                )
                textbox.text_frame.text = finding_text

                # Slide 7: Recommendations
                slide = prs.slides.add_slide(content_layout)
                slide.shapes.title.text = "Recommendations"

                recommendation_text = "\n".join(
                    f"• {recommendation}" for recommendation in recommendations
                )

                textbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(1.4), Inches(11.5), Inches(5.2)
                )
                textbox.text_frame.text = recommendation_text

                # Chart slides: add exported visualizations when Kaleido is available.
                for chart_index, report_fig in enumerate(build_report_charts(clean_df), 1):
                    slide = prs.slides.add_slide(content_layout)
                    slide.shapes.title.text = f"Data Visualization {chart_index}"
                    added = add_chart_to_ppt(slide, report_fig, 0.7, 1.25, 11.9, 5.7)
                    if not added:
                        box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.5), Inches(1.0))
                        box.text_frame.text = f"{report_fig.layout.title.text} — chart image export unavailable; install/verify Kaleido."

                # Final conclusion slide
                slide = prs.slides.add_slide(content_layout)
                slide.shapes.title.text = "Conclusion"

                conclusion_text = (
                    "The project followed a structured data analyst workflow "
                    "from raw data inspection through cleaning, quality checks, "
                    "analysis, dashboarding, and reporting.\n\n"
                    "The dashboard provides an interactive overview of the "
                    "available data. Business conclusions should be validated "
                    "with domain knowledge and stakeholder requirements."
                )

                textbox = slide.shapes.add_textbox(
                    Inches(0.8), Inches(1.4), Inches(11.5), Inches(4.8)
                )
                textbox.text_frame.text = conclusion_text

                ppt_buffer = io.BytesIO()
                prs.save(ppt_buffer)
                ppt_buffer.seek(0)

                st.download_button(
                    "Download PowerPoint Presentation",
                    data=ppt_buffer.getvalue(),
                    file_name=f"{safe_name(st.session_state.file_name)}_analysis_presentation.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
        else:
            st.error(
                "PowerPoint export library is unavailable. Install python-pptx using "
                "`python -m pip install python-pptx`."
            )


# ================================================================
# FOOTER
# ================================================================

st.sidebar.divider()
st.sidebar.caption("Automated ETL & Dashboard System")
st.sidebar.caption("Built with Python, Pandas, Streamlit, and Plotly")