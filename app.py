import io
import re
import json
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# -----------------------------
# App Config
# -----------------------------
st.set_page_config(page_title="Virtual CI Specialist - Module 1", layout="wide")

# -----------------------------
# Helpers
# -----------------------------
def normalize_colname(c: str) -> str:
    return re.sub(r"\s+", " ", str(c).strip())

def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # Standardize column names
    df = df.copy()
    df.columns = [normalize_colname(c) for c in df.columns]
    # Trim cell strings and normalize whitespace
    for c in df.columns:
        if pd.api.types.is_string_dtype(df[c]):
            df[c] = df[c].astype(str).str.strip()
            df[c] = df[c].str.replace(r"\s+", " ", regex=True)
    return df

def guess_columns(df: pd.DataFrame):
    cols = { "loss_minutes": None, "category": None, "reason": None, "date": None }
    lc = {c: c.lower() for c in df.columns}

    # Loss minutes
    for c, l in lc.items():
        if ("loss" in l or "downtime" in l or "idle" in l) and ("min" in l or "minutes" in l or "duration" in l or "time" in l):
            cols["loss_minutes"] = c
            break
    if cols["loss_minutes"] is None:
        for c, l in lc.items():
            if l in ["loss minutes", "loss(min)", "loss_min", "loss (mins)"]:
                cols["loss_minutes"] = c
                break

    # Category (Department / 6M / Area / Line)
    for c, l in lc.items():
        if any(k in l for k in ["department", "dept", "6m", "pillar", "category", "line", "machine", "area", "section"]):
            cols["category"] = c
            break

    # Reason
    for c, l in lc.items():
        if any(k in l for k in ["reason", "cause", "issue", "fault", "problem", "loss name", "loss type"]):
            cols["reason"] = c
            break

    # Date
    for c in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            cols["date"] = c
            break
    if cols["date"] is None:
        for c, l in lc.items():
            if any(k in l for k in ["date", "day", "timestamp", "time", "dt"]):
                cols["date"] = c
                break

    return cols

def ensure_numeric_minutes(series: pd.Series) -> pd.Series:
    # Try numeric conversion
    s = pd.to_numeric(series, errors="coerce")
    # If too many NaN, try extracting numbers from strings (e.g., "15 min")
    if s.isna().mean() > 0.5:
        # Extract first number in the string
        extracted = series.astype(str).str.extract(r"([0-9]+(?:\.[0-9]+)?)")[0]
        s = pd.to_numeric(extracted, errors="coerce")
    # Replace negatives with abs, nulls with 0
    s = s.fillna(0)
    s = s.clip(lower=0)
    return s

def safe_to_datetime(series: pd.Series) -> pd.Series:
    try:
        return pd.to_datetime(series, errors="coerce")
    except Exception:
        return pd.to_datetime(pd.Series([None]*len(series)))

def kpi_card(label: str, value: str, help_text: str = ""):
    st.metric(label, value, help=help_text)

# -----------------------------
# UI
# -----------------------------
st.title("🧪 Virtual CI Specialist — Module 1 (Real Data-Ready)")
st.caption("Upload your factory data (CSV/Excel). Map columns. Auto-clean. See OAE, Pareto, Trend. Download cleaned data + your mapping.")

with st.sidebar:
    st.header("1) Upload Data")
    file = st.file_uploader("Choose CSV or Excel", type=["csv", "xlsx", "xls"])

    st.header("2) Planned Time")
    planned_minutes = st.number_input(
        "Planned Available Minutes (per dataset scope)",
        min_value=0,
        value=0,
        help="Used to compute OAE = (1 - Total Loss Minutes / Planned Minutes) × 100. If 0, OAE won't be shown."
    )

    st.header("3) Display")
    top_n = st.slider("Pareto Top N", 5, 30, 10)
    show_category = st.checkbox("Show Pareto by Category", value=True)
    show_reason = st.checkbox("Show Pareto by Reason", value=True)

if not file:
    st.info("👈 Upload a CSV/Excel to begin. You can export results and mapping afterwards.")
    st.stop()

# Load
try:
    if file.name.lower().endswith(".csv"):
        df_raw = pd.read_csv(file, encoding="utf-8", engine="python")
    else:
        df_raw = pd.read_excel(file, engine="openpyxl")
except Exception as e:
    st.error(f"Could not read file: {e}")
    st.stop()

# Clean
df = clean_dataframe(df_raw)

# Guess columns, then let user confirm/override
guesses = guess_columns(df)
st.subheader("Map Your Columns (Flexible — choose what matches your file)")
col1, col2, col3, col4 = st.columns(4)

with col1:
    loss_col = st.selectbox(
        "Loss Minutes Column (required)",
        options=["-- Select --"] + list(df.columns),
        index=(["-- Select --"] + list(df.columns)).index(guesses["loss_minutes"]) if guesses["loss_minutes"] in df.columns else 0,
        help="Any numeric/text column that represents loss/downtime minutes. We will clean/convert it."
    )

with col2:
    cat_col = st.selectbox(
        "Category / Dept / 6M (optional)",
        options=["(none)"] + list(df.columns),
        index=(["(none)"] + list(df.columns)).index(guesses["category"]) if guesses["category"] in df.columns else 0
    )

with col3:
    reason_col = st.selectbox(
        "Reason / Cause (optional)",
        options=["(none)"] + list(df.columns),
        index=(["(none)"] + list(df.columns)).index(guesses["reason"]) if guesses["reason"] in df.columns else 0
    )

with col4:
    date_col = st.selectbox(
        "Date (optional)",
        options=["(none)"] + list(df.columns),
        index=(["(none)"] + list(df.columns)).index(guesses["date"]) if guesses["date"] in df.columns else 0
    )

if loss_col == "-- Select --":
    st.error("Please select the Loss Minutes column.")
    st.stop()

# Prepare working df
work = df.copy()

# Minutes numeric
work["_loss_minutes"] = ensure_numeric_minutes(work[loss_col])

# Optional columns
if cat_col != "(none)":
    work["_category"] = work[cat_col].astype(str).str.strip()
else:
    work["_category"] = "(unspecified)"

if reason_col != "(none)":
    work["_reason"] = work[reason_col].astype(str).str.strip()
else:
    work["_reason"] = "(unspecified)"

if date_col != "(none)":
    work["_date"] = safe_to_datetime(work[date_col])
else:
    work["_date"] = pd.NaT

# Basic stats
total_loss = float(work["_loss_minutes"].sum())
records = len(work)

k1, k2, k3 = st.columns(3)
with k1:
    kpi_card("Total Loss Minutes", f"{total_loss:,.0f}", "Sum of cleaned loss minutes")
with k2:
    kpi_card("Records", f"{records:,}")
with k3:
    if planned_minutes and planned_minutes > 0:
        oae = max(0.0, (1 - total_loss / planned_minutes) * 100)
        kpi_card("OAE %", f"{oae:,.2f}%", "OAE = (1 - Loss/Planned) × 100")
    else:
        kpi_card("OAE %", "—", "Enter Planned Minutes in the sidebar")

st.divider()

# Pareto(s)
charts = []
if show_category and "_category" in work.columns:
    pareto_cat = work.groupby("_category")["_loss_minutes"].sum().sort_values(ascending=False).head(top_n).reset_index()
    fig_cat = px.bar(pareto_cat, x="_category", y="_loss_minutes", title=f"Pareto by Category (Top {top_n})", text_auto=True)
    fig_cat.update_layout(xaxis_title="", yaxis_title="Loss Minutes")
    charts.append(("Category Pareto", fig_cat))

if show_reason and "_reason" in work.columns:
    pareto_reason = work.groupby("_reason")["_loss_minutes"].sum().sort_values(ascending=False).head(top_n).reset_index()
    fig_reason = px.bar(pareto_reason, x="_reason", y="_loss_minutes", title=f"Pareto by Reason (Top {top_n})", text_auto=True)
    fig_reason.update_layout(xaxis_title="", yaxis_title="Loss Minutes")
    charts.append(("Reason Pareto", fig_reason))

if charts:
    for title, fig in charts:
        st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Add Category/Reason mappings to see Pareto charts.")

# Trend (if date provided)
if work["_date"].notna().any():
    trend = work.dropna(subset=["_date"]).copy()
    trend["_date"] = trend["_date"].dt.date
    trend_grouped = trend.groupby("_date")["_loss_minutes"].sum().reset_index()
    fig_trend = px.line(trend_grouped, x="_date", y="_loss_minutes", markers=True, title="Loss Minutes Trend by Date")
    fig_trend.update_layout(xaxis_title="Date", yaxis_title="Loss Minutes")
    st.plotly_chart(fig_trend, use_container_width=True)
else:
    st.caption("Tip: Map your Date column to unlock the trend chart.")

st.divider()

# Preview & Downloads
with st.expander("Preview Cleaned Data (first 200 rows)"):
    st.dataframe(work.head(200), use_container_width=True)

mapping = {
    "loss_minutes": loss_col,
    "category": None if cat_col == "(none)" else cat_col,
    "reason": None if reason_col == "(none)" else reason_col,
    "date": None if date_col == "(none)" else date_col,
    "planned_minutes": planned_minutes
}

# Download buttons
clean_csv = work.rename(columns={
    "_loss_minutes": "LossMinutes_Cleaned",
    "_category": "Category_Mapped",
    "_reason": "Reason_Mapped",
    "_date": "Date_Parsed"
})

csv_buf = io.StringIO()
clean_csv.to_csv(csv_buf, index=False)
st.download_button("⬇️ Download Cleaned CSV", data=csv_buf.getvalue(), file_name="cleaned_losses.csv", mime="text/csv")

st.download_button("⬇️ Download Mapping JSON", data=json.dumps(mapping, indent=2), file_name="mapping.json", mime="application/json")

st.success("Ready. Upload any real Excel/CSV. Map columns once, get OAE, Pareto, Trend, and downloads.")
