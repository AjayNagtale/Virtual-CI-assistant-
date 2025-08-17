import streamlit as st
import pandas as pd
import io
import plotly.express as px

# ──────────────────────────────────────────────────────────────────────────────
# Config
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Virtual CI Specialist", layout="wide")
st.title("📊 Virtual CI Specialist — Multi-Sheet Excel Ingest")

st.caption(
    "Upload a single Excel workbook with one or more sheets. "
    "The app will auto-detect columns (Department, Date, Loss Minutes, Reason), "
    "clean data, combine valid sheets, and build visuals."
)

# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────

def sanitize(s: str) -> str:
    """Normalize a column name for easier matching."""
    return (
        str(s)
        .strip()
        .lower()
        .replace("\n", " ")
        .replace("\t", " ")
        .replace("-", " ")
        .replace("/", " ")
        .replace("\\", " ")
        .replace("(", " ")
        .replace(")", " ")
        .replace("[", " ")
        .replace("]", " ")
        .replace(".", " ")
        .replace(",", " ")
        .replace("  ", " ")
        .strip()
    )

ALIAS_MAP = {
    "department": {
        "exact": [
            "department", "dept", "area", "line", "cell", "process", "work center",
            "workcenter", "machine", "asset"
        ],
        "contains_any": ["dept", "area", "line", "cell", "process", "work center", "workcenter", "machine", "asset"],
    },
    "date": {
        "exact": ["date", "dt", "event date", "recorded date", "day", "shift date", "production date"],
        "contains_any": ["date", "shift", "day"],
    },
    "loss_minutes": {
        "exact": [
            "loss minutes", "loss_min", "loss_minute", "downtime minutes", "downtime min",
            "downtime", "minutes", "min", "loss time (min)", "loss time min", "loss time"
        ],
        "contains_any": ["minute", "min", "downtime", "loss time"],
    },
    "reason": {
        "exact": ["reason", "loss reason", "issue", "category", "cause", "problem", "failure mode"],
        "contains_any": ["reason", "issue", "category", "cause", "problem", "failure mode"],
    },
}

def auto_map_columns(df: pd.DataFrame):
    """
    Try to map messy column names to the required four: department, date, loss_minutes, reason.
    Returns mapping dict {required: original_col_name} or raises ValueError.
    """
    normalized = {sanitize(c): c for c in df.columns}
    found = {}

    def find_for(key):
        # 1) exact list
        for label in ALIAS_MAP[key]["exact"]:
            lbl = sanitize(label)
            if lbl in normalized:
                return normalized[lbl]
        # 2) contains_any heuristic
        for lbl in normalized:
            if any(tok in lbl for tok in ALIAS_MAP[key]["contains_any"]):
                return normalized[lbl]
        return None

    for required in ["department", "date", "loss_minutes", "reason"]:
        col = find_for(required)
        if col is not None:
            found[required] = col

    missing = [k for k in ["department", "date", "loss_minutes", "reason"] if k not in found]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    return found

def clean_and_standardize(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """
    Map columns, coerce types, and return a standardized frame with
    columns: department, date, loss_minutes, reason, sheet.
    """
    mapping = auto_map_columns(df)

    out = pd.DataFrame()
    out["department"] = df[mapping["department"]].astype(str).str.strip()
    out["reason"] = df[mapping["reason"]].astype(str).str.strip()

    # dates
    out["date"] = pd.to_datetime(df[mapping["date"]], errors="coerce")

    # numeric loss minutes
    # try to coerce common strings like "12 min", "12.5", etc.
    lm_raw = df[mapping["loss_minutes"]].copy()
    if lm_raw.dtype == object:
        lm_raw = lm_raw.astype(str).str.replace("[^0-9.]", "", regex=True)
    out["loss_minutes"] = pd.to_numeric(lm_raw, errors="coerce")

    # drop incomplete rows
    out = out.dropna(subset=["department", "reason", "date", "loss_minutes"])

    # logical cleanup
    out = out[out["loss_minutes"] >= 0]

    # attach sheet source
    out["sheet"] = sheet_name

    return out

def calc_oae(df: pd.DataFrame, minutes_per_period: float) -> float:
    """
    Very simple OAE proxy:
    OAE% = ((minutes_per_period - total_loss_minutes) / minutes_per_period) * 100
    computed over the whole dataset (sum of loss).
    """
    total_loss = df["loss_minutes"].sum()
    if minutes_per_period <= 0:
        return 0.0
    oae = max(0.0, min(100.0, (minutes_per_period - total_loss) / minutes_per_period * 100.0))
    return oae

# ──────────────────────────────────────────────────────────────────────────────
# Sidebar controls
# ──────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    minutes_per_period = st.number_input(
        "Minutes per analysis period",
        min_value=1,
        max_value=1000000,
        value=1440,
        step=1,
        help="Use 1440 for a full day, or enter your planned available minutes for the scope you are analyzing."
    )
    group_granularity = st.selectbox(
        "Trend granularity",
        options=["Day", "Week", "Month"],
        index=1
    )

# ──────────────────────────────────────────────────────────────────────────────
# File upload
# ──────────────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader("📥 Upload an Excel workbook (.xlsx) with one or more sheets", type=["xlsx"])

if not uploaded:
    st.info("Upload a workbook to begin.")
    st.stop()

# Read the Excel file and list sheets
try:
    xls = pd.ExcelFile(uploaded)
    sheet_names = xls.sheet_names
except Exception as e:
    st.error(f"Could not read Excel file: {e}")
    st.stop()

process_mode = st.radio(
    "Select sheets to process",
    options=[f"All sheets ({len(sheet_names)})", "Choose sheets"],
    horizontal=True
)

if process_mode == "Choose sheets":
    chosen = st.multiselect("Pick sheets to include", options=sheet_names, default=sheet_names)
else:
    chosen = sheet_names

if not chosen:
    st.warning("No sheets selected.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# Process each sheet, combine, and report
# ──────────────────────────────────────────────────────────────────────────────
frames = []
skipped = []

for s in chosen:
    try:
        df_s = pd.read_excel(xls, sheet_name=s)
        if df_s.empty or len(df_s.columns) == 0:
            skipped.append((s, "Empty or no columns"))
            continue
        cleaned = clean_and_standardize(df_s, s)
        if cleaned.empty:
            skipped.append((s, "No valid rows after cleaning"))
            continue
        frames.append(cleaned)
    except ValueError as ve:
        skipped.append((s, f"Column mapping error: {ve}"))
    except Exception as e:
        skipped.append((s, f"Read/clean error: {e}"))

if not frames:
    st.error("No valid data found in the selected sheets. See the skip report below.")
    if skipped:
        with st.expander("Skip report"):
            for s, msg in skipped:
                st.write(f"• **{s}** — {msg}")
    st.stop()

df = pd.concat(frames, ignore_index=True)

# Show summary and any skipped sheets
valid_rows = len(df)
st.success(f"✅ Loaded **{valid_rows}** rows from **{len(frames)}** sheet(s).")
if skipped:
    with st.expander("ℹ️ Sheets skipped or partially invalid"):
        for s, msg in skipped:
            st.write(f"• **{s}** — {msg}")

# ──────────────────────────────────────────────────────────────────────────────
# Preview + Download Cleaned Data
# ──────────────────────────────────────────────────────────────────────────────
st.subheader("🧹 Cleaned & Standardized Data (first 200 rows)")
st.dataframe(df.head(200), use_container_width=True)

csv_bytes = df.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Download cleaned data (CSV)",
    data=csv_bytes,
    file_name="cleaned_ci_data.csv",
    mime="text/csv"
)

# ──────────────────────────────────────────────────────────────────────────────
# KPI: OAE (simple proxy)
# ──────────────────────────────────────────────────────────────────────────────
oae = calc_oae(df, minutes_per_period)
kpi1, kpi2 = st.columns(2)
kpi1.metric("OAE % (proxy)", f"{oae:.2f}%")
kpi2.metric("Total Loss Minutes", f"{df['loss_minutes'].sum():,.0f}")

# ──────────────────────────────────────────────────────────────────────────────
# Visuals: Pareto, Trend, Department breakdown
# ──────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📌 Pareto — Loss Minutes by Reason")
pareto = (
    df.groupby("reason", as_index=False)["loss_minutes"]
    .sum()
    .sort_values("loss_minutes", ascending=False)
)
fig_pareto = px.bar(
    pareto, x="reason", y="loss_minutes", text="loss_minutes",
    title="Top Loss Reasons (Pareto)"
)
fig_pareto.update_xaxes(tickangle=45)
st.plotly_chart(fig_pareto, use_container_width=True)

st.subheader("📈 Trend")
df_trend = df.copy()
if group_granularity == "Day":
    df_trend["period"] = df_trend["date"].dt.to_period("D").dt.to_timestamp()
elif group_granularity == "Week":
    df_trend["period"] = df_trend["date"].dt.to_period("W").dt.start_time
else:  # Month
    df_trend["period"] = df_trend["date"].dt.to_period("M").dt.to_timestamp()

trend = (
    df_trend.groupby("period", as_index=False)["loss_minutes"]
    .sum()
    .sort_values("period")
)
fig_trend = px.line(trend, x="period", y="loss_minutes", markers=True, title=f"Loss Minutes Trend by {group_granularity}")
st.plotly_chart(fig_trend, use_container_width=True)

st.subheader("🏭 Department Breakdown")
dept = (
    df.groupby("department", as_index=False)["loss_minutes"]
    .sum()
    .sort_values("loss_minutes", ascending=False)
)
fig_dept = px.bar(dept, x="department", y="loss_minutes", text="loss_minutes", title="Loss Minutes by Department")
fig_dept.update_xaxes(tickangle=45)
st.plotly_chart(fig_dept, use_container_width=True)

st.caption("Source columns mapped automatically from each sheet. Standard columns: department, date, loss_minutes, reason. "
           "Adjust minutes per period in the sidebar to reflect your planned available time window.")
