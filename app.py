import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Virtual CI Assistant (MVP)", layout="wide")

st.title("Virtual CI Assistant — MVP")
st.caption("Upload your Excel. Map columns once. Get loss insights fast. (Multi-sheet supported)")

# -------------------------
# Helpers
# -------------------------
CANON = {
    "department": ["department", "dept", "area", "line", "cell", "workcenter", "work center"],
    "date": ["date", "dt", "timestamp", "day", "calendar date", "prod date", "production date"],
    "loss_minutes": ["loss minutes", "loss (min)", "loss_min", "loss mins", "downtime min", "dt minutes", "loss in minutes", "minutes lost"],
    "reason": ["reason", "cause", "category", "loss reason", "dt reason", "issue"],
}

def normalize(s: str) -> str:
    return str(s).strip().lower().replace("_", " ").replace("-", " ")

def auto_pick_column(cols, candidates):
    cols_n = [normalize(c) for c in cols]
    for cand in candidates:
        n = normalize(cand)
        for i, c in enumerate(cols_n):
            if n == c:
                return cols[i]
        # loose contains
        for i, c in enumerate(cols_n):
            if n in c or c in n:
                return cols[i]
    return None

def try_parse_date(series: pd.Series):
    try:
        return pd.to_datetime(series, errors="coerce", dayfirst=True, infer_datetime_format=True)
    except Exception:
        return pd.to_datetime(series, errors="coerce")

def coerce_numeric(series: pd.Series):
    return pd.to_numeric(series, errors="coerce")

def build_download(dataframes: dict, filename="ci_results.xlsx"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for name, df in dataframes.items():
            # Truncate very long sheet names to avoid Excel errors
            safe_name = str(name)[:31] if name else "Sheet1"
            df.to_excel(writer, index=False, sheet_name=safe_name)
    output.seek(0)
    return output, filename

# -------------------------
# Sidebar: Inputs
# -------------------------
with st.sidebar:
    st.header("Inputs")
    excel_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    st.markdown("---")
    st.subheader("Optional: OAE settings")
    total_available_minutes = st.number_input(
        "Available minutes for the chosen period (e.g., planned minutes)",
        min_value=0, value=0, step=1, help="If set > 0, the app will compute OAE% = 1 - (Loss/Available)."
    )

    st.markdown("---")
    st.subheader("Help")
    st.write("If the app cannot find your columns, use the mapping dropdowns that appear after upload.")
    st.write("Multi-sheet files are supported; pick which sheets to include below the upload.")

# -------------------------
# If no file, provide sample data
# -------------------------
def sample_df():
    data = {
        "Department": ["Line A","Line A","Line B","Line B","Line C"],
        "Date": ["2025-08-01","2025-08-01","2025-08-02","2025-08-02","2025-08-03"],
        "Loss Minutes": [45, 30, 60, 15, 0],
        "Reason": ["Setup","Breakdown","Quality","Waiting","No Loss"],
    }
    return pd.DataFrame(data)

# -------------------------
# Load & choose sheets
# -------------------------
dfs = []
sheet_names = []

if excel_file:
    try:
        xl = pd.ExcelFile(excel_file)
        sheet_names = xl.sheet_names
        if len(sheet_names) == 1:
            use_sheets = [sheet_names[0]]
        else:
            st.success(f"Detected {len(sheet_names)} sheets: {', '.join(sheet_names)}")
            use_sheets = st.multiselect("Select sheets to include", sheet_names, default=sheet_names)
        for s in use_sheets:
            df_s = xl.parse(s)
            df_s["__source_sheet__"] = s
            dfs.append(df_s)
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
else:
    st.info("No Excel uploaded. Using sample data.")
    df_s = sample_df()
    df_s["__source_sheet__"] = "Sample"
    dfs.append(df_s)

if not dfs:
    st.stop()

raw = pd.concat(dfs, ignore_index=True)

# -------------------------
# Column detection + mapping UI
# -------------------------
st.subheader("Step 1 — Map your columns (only if needed)")
cols = list(raw.columns)
col_map_detected = {}
for key, candidates in CANON.items():
    picked = auto_pick_column(cols, candidates)
    col_map_detected[key] = picked

with st.expander("Column mapping", expanded=any(v is None for v in col_map_detected.values())):
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        col_department = st.selectbox("Department column", options=["<none>"]+cols,
                                      index=(cols.index(col_map_detected["department"]) + 1) if col_map_detected["department"] in cols else 0)
    with c2:
        col_date = st.selectbox("Date column", options=["<none>"]+cols,
                                index=(cols.index(col_map_detected["date"]) + 1) if col_map_detected["date"] in cols else 0)
    with c3:
        col_loss = st.selectbox("Loss Minutes column", options=["<none>"]+cols,
                                index=(cols.index(col_map_detected["loss_minutes"]) + 1) if col_map_detected["loss_minutes"] in cols else 0)
    with c4:
        col_reason = st.selectbox("Reason column", options=["<none>"]+cols,
                                  index=(cols.index(col_map_detected["reason"]) + 1) if col_map_detected["reason"] in cols else 0)

# Validate selection
missing = []
if col_department == "<none>": missing.append("Department")
if col_date == "<none>": missing.append("Date")
if col_loss == "<none>": missing.append("Loss Minutes")
if col_reason == "<none>": missing.append("Reason")

if missing:
    st.warning(f"Please map these columns to proceed: {', '.join(missing)}")
    st.stop()

# -------------------------
# Cleaning
# -------------------------
df = raw[[col_department, col_date, col_loss, col_reason, "__source_sheet__"]].copy()
df.columns = ["Department", "Date", "Loss Minutes", "Reason", "Source Sheet"]

# Trim strings
for c in ["Department", "Reason", "Source Sheet"]:
    df[c] = df[c].astype(str).str.strip()

# Parse dates
df["Date"] = try_parse_date(df["Date"])
df = df.dropna(subset=["Date"])

# Coerce loss minutes
df["Loss Minutes"] = coerce_numeric(df["Loss Minutes"]).fillna(0).clip(lower=0)

# -------------------------
# Metrics & visuals
# -------------------------
st.subheader("Step 2 — Results")

total_loss = df["Loss Minutes"].sum()
st.metric("Total Loss Minutes (selected data)", f"{int(total_loss):,}")

if total_available_minutes and total_available_minutes > 0:
    oae = 1 - (total_loss / total_available_minutes)
    oae_pct = max(0.0, min(1.0, oae)) * 100
    st.metric("OAE % (simple)", f"{oae_pct:,.2f}%")
else:
    st.info("Provide 'Available minutes' in the sidebar to compute a simple OAE%.")

# Loss by Department
dept_agg = (
    df.groupby("Department", as_index=False)["Loss Minutes"]
    .sum()
    .sort_values("Loss Minutes", ascending=False)
)
if not dept_agg.empty:
    fig_dept = px.bar(dept_agg, x="Department", y="Loss Minutes", title="Loss Minutes by Department")
    st.plotly_chart(fig_dept, use_container_width=True)

# Top Reasons
reason_agg = (
    df.groupby("Reason", as_index=False)["Loss Minutes"]
    .sum()
    .sort_values("Loss Minutes", ascending=False)
)
top_n = st.slider("Show top N reasons", 3, min(20, max(3, len(reason_agg))), 5)
if not reason_agg.empty:
    fig_reason = px.bar(reason_agg.head(top_n), x="Reason", y="Loss Minutes", title=f"Top {top_n} Loss Reasons")
    st.plotly_chart(fig_reason, use_container_width=True)

# Trend
df["Month"] = df["Date"].dt.to_period("M").astype(str)
trend = df.groupby(["Month"], as_index=False)["Loss Minutes"].sum()
if not trend.empty:
    fig_trend = px.line(trend, x="Month", y="Loss Minutes", title="Loss Minutes Trend by Month", markers=True)
    st.plotly_chart(fig_trend, use_container_width=True)

# Data preview
with st.expander("See cleaned data"):
    st.dataframe(df, use_container_width=True, hide_index=True)

# Download
results = {
    "Cleaned Data": df,
    "Loss by Department": dept_agg,
    "Loss by Reason": reason_agg,
    "Trend by Month": trend,
}
dl_bytes, dl_name = build_download(results, filename=f"ci_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
st.download_button("Download results Excel", data=dl_bytes, file_name=dl_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Note: This MVP does not store your data. It processes in-session only. Persistence/OneDrive sync is on the roadmap.")
        
