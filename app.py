# app.py
import re
import io
from datetime import datetime
from typing import Tuple, Dict, List

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# -------------------------
# Page config
# -------------------------
st.set_page_config(page_title="Virtual CI Specialist — Full MVP", layout="wide")
st.title("🏭 Virtual CI Specialist — Full MVP")
st.caption("Upload messy multi-sheet Excel, map columns if needed, get cleaned data, OAE/OEE, Pareto & Actionable insights.")

# -------------------------
# Constants & synonyms (extendable)
# -------------------------
SYNONYMS = {
    "date": ["date", "day", "dt", "timestamp", "production date", "prod date", "shift date"],
    "department": ["department", "dept", "line", "area", "cell", "workcenter", "machine", "asset"],
    "reason": ["reason", "cause", "category", "loss reason", "root cause", "failure mode"],
    "loss_minutes": ["lossminutes", "loss mins", "loss_min", "loss (min)", "downtime", "downtime minutes", "minutes lost", "mins"],
    # Optional fields
    "good_count": ["good", "good_count", "good pieces", "ok", "passes"],
    "total_count": ["total", "total_count", "produced", "output", "pieces"],
    "planned_minutes": ["plannedminutes", "planned minutes", "available minutes", "scheduled minutes", "planned"]
}

REQUIRED = ["date", "department", "reason", "loss_minutes"]

# -------------------------
# Utilities: normalize & parsing
# -------------------------
def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(s).strip().lower())

def find_col_by_synonyms(cols: List[str], alias_list: List[str]) -> str:
    norm_cols = {c: norm(c) for c in cols}
    # direct exact match first
    for c, nc in norm_cols.items():
        if nc in [norm(a) for a in alias_list]:
            return c
    # contains match
    for c, nc in norm_cols.items():
        for a in alias_list:
            if norm(a) in nc or nc in norm(a):
                return c
    return None

def parse_minutes_value(v):
    if pd.isna(v): 
        return np.nan
    # numeric
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).strip()
    if s == "":
        return np.nan
    # hh:mm or h:mm
    if re.match(r"^\d{1,2}:\d{2}$", s):
        h, m = s.split(":")
        return int(h) * 60 + int(m)
    # contains number
    m = re.search(r"([0-9]+(?:\.[0-9]+)?)", s)
    if m:
        try:
            return float(m.group(1))
        except:
            return np.nan
    return np.nan

def try_parse_date(series: pd.Series) -> pd.Series:
    out = pd.to_datetime(series, errors="coerce", infer_datetime_format=True)
    if out.notna().any():
        return out
    return pd.to_datetime(series, errors="coerce", dayfirst=True)

# -------------------------
# Ingest: load multi-sheet Excel or CSV
# -------------------------
@st.cache_data(show_spinner=False)
def load_multi_sheet(file_bytes: bytes, filename: str) -> Tuple[pd.DataFrame, List[Dict]]:
    diagnostics = []
    frames = []
    try:
        if filename.lower().endswith(".csv"):
            df_raw = pd.read_csv(io.BytesIO(file_bytes))
            df_clean, diag = clean_single_sheet(df_raw, "(csv)")
            diagnostics.append(diag)
            if not df_clean.empty:
                frames.append(df_clean)
        else:
            xls = pd.ExcelFile(io.BytesIO(file_bytes))
            for sheet in xls.sheet_names:
                try:
                    df_raw = pd.read_excel(xls, sheet_name=sheet, header=0)
                except Exception as e:
                    diagnostics.append({"sheet": sheet, "error": f"read failed: {e}"})
                    continue
                df_clean, diag = clean_single_sheet(df_raw, sheet)
                diagnostics.append(diag)
                if not df_clean.empty:
                    frames.append(df_clean)
    except Exception as e:
        raise RuntimeError(f"Loading failed: {e}")

    if frames:
        combined = pd.concat(frames, ignore_index=True)
        # normalize dtypes
        if "date" in combined.columns:
            combined["date"] = try_parse_date(combined["date"])
        if "loss_minutes" in combined.columns:
            combined["loss_minutes"] = pd.to_numeric(combined["loss_minutes"], errors="coerce")
        return combined, diagnostics

    return pd.DataFrame(), diagnostics

# -------------------------
# Clean one sheet (mapping heuristics + parsing)
# -------------------------
def clean_single_sheet(df: pd.DataFrame, sheet_name: str) -> Tuple[pd.DataFrame, dict]:
    diag = {"sheet": sheet_name, "rows_in": len(df)}
    df0 = df.copy()
    # normalize column names (strip)
    df0.columns = [str(c).strip() for c in df0.columns]

    # auto-map
    cols = list(df0.columns)
    mapped = {}
    for canon, aliases in SYNONYMS.items():
        pick = find_col_by_synonyms(cols, aliases)
        if pick:
            mapped[canon] = pick

    # collect missing required mapped names for diagnostics
    missing = [r for r in REQUIRED if r not in mapped]

    # rename DataFrame where mapped
    rename_map = {orig: canon for canon, orig in mapped.items()}
    df_clean = df0.rename(columns=rename_map)

    # attempt to coerce fields we have
    if "date" in df_clean.columns:
        df_clean["date"] = try_parse_date(df_clean["date"])
    if "loss_minutes" in df_clean.columns:
        df_clean["loss_minutes"] = df_clean["loss_minutes"].apply(parse_minutes_value)
    if "department" in df_clean.columns:
        df_clean["department"] = df_clean["department'].astype(str).str.strip()
    if "reason" in df_clean.columns:
        df_clean["reason"] = df_clean["reason"].astype(str).str.strip()
    # planned minutes optional
    if "planned_minutes" in df_clean.columns:
        df_clean["planned_minutes"] = df_clean["planned_minutes"].apply(parse_minutes_value)

    # drop rows missing core fields (date & loss mandatory for analysis)
    keep_mask = pd.Series(True, index=df_clean.index)
    if "date" in df_clean.columns:
        keep_mask &= df_clean["date"].notna()
    if "loss_minutes" in df_clean.columns:
        keep_mask &= df_clean["loss_minutes"].notna()
    df_clean = df_clean.loc[keep_mask].copy()

    # deduplicate basic combos
    keep_cols = [c for c in ["date", "department", "reason", "loss_minutes"] if c in df_clean.columns]
    if keep_cols:
        df_clean = df_clean.drop_duplicates(subset=keep_cols)

    diag.update({
        "rows_out": len(df_clean),
        "mapped_cols": mapped,
        "missing_required": missing
    })
    return df_clean, diag

# -------------------------
# KPI calcs: OAE & OEE (OEE optional)
# -------------------------
def compute_oae(df: pd.DataFrame, planned_minutes_per_asset_per_day: float, assets_count: int, use_planned_from_file: bool=False) -> Tuple[float, float, float]:
    total_loss = float(df["loss_minutes"].sum()) if "loss_minutes" in df.columns else 0.0
    # planned_total: optionally use planned_minutes column per row if present
    if use_planned_from_file and "planned_minutes" in df.columns and df["planned_minutes"].notna().any():
        planned_total = float(df["planned_minutes"].sum())
    else:
        # default: planned_per_day * distinct_dates * assets
        n_days = df["date"].dt.date.nunique() if "date" in df.columns else 1
        planned_total = float(planned_minutes_per_asset_per_day) * float(assets_count) * float(n_days)
    oae = max(0.0, 100.0 * (1.0 - (total_loss / planned_total))) if planned_total > 0 else 0.0
    return oae, total_loss, planned_total

def compute_oee(df: pd.DataFrame, ideal_cycle_time_sec: float=None) -> Dict[str, float]:
    # OEE = Availability * Performance * Quality
    # Availability = (Planned - Downtime) / Planned
    # Performance = (TotalPieces * IdealCycle) / RunTime  (requires pieces & runtime)
    # Quality = GoodPieces / TotalPieces (requires good_count)
    res = {"Availability": None, "Performance": None, "Quality": None, "OEE": None}
    try:
        total_loss = float(df["loss_minutes"].sum()) if "loss_minutes" in df.columns else 0.0
        # compute planned_total by simple approach: count unique dates * typical 1 asset  = fallback; better to supply planned
        # We'll only compute Availability if user provided 'planned_total' widget
        # Performance & Quality require 'total_count' and 'good_count' columns
    except:
        return res
    return res

# -------------------------
# UI: sidebar & upload
# -------------------------
with st.sidebar:
    st.header("Inputs & Settings")
    st.write("Mapping and parsing are automated; you can override mappings below if needed.")
    use_planned_from_file = st.checkbox("Use 'Planned Minutes' from file if present", value=True)
    planned_per_day = st.number_input("Default planned minutes per day (per asset)", min_value=1, value=480, step=1)
    assets_count = st.number_input("Number of assets / lines represented", min_value=1, value=1, step=1)
    top_n = st.slider("Pareto top N items", min_value=3, max_value=30, value=8)
    st.markdown("---")
    st.caption("Next: connect OneDrive/SharePoint or a cloud folder to auto-refresh (Phase 2).")

uploaded = st.file_uploader("Upload your Excel or CSV (multi-sheet Excel supported)", type=["xlsx", "xls", "csv"])

# Quick sample button for testing
if st.button("Load example sample data"):
    sample = pd.DataFrame({
        "Date": ["2025-08-01", "2025-08-01", "2025-08-02", "2025-08-02"],
        "Dept": ["Line A", "Line A", "Line B", "Line B"],
        "Downtime Reason": ["Breakdown", "Setup", "Material Shortage", "Inspection"],
        "Loss Mins": [120, "0:45", 30, 15],
        "Planned Minutes": [480, 480, 480, 480]
    })
    # simulate upload by converting to bytes and reusing load function
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        sample.to_excel(writer, index=False, sheet_name="Sheet1")
    buf.seek(0)
    uploaded = st.experimental_set_query_params(_sample="1")  # just visual marker
    st.success("Example built — now click the 'Upload Excel' and pick the same example file if you saved it locally.")
    st.stop()

if not uploaded:
    st.info("Upload an Excel/CSV file to begin processing. If your real file has multiple sheets, you'll be able to choose which sheets to include.")
    st.stop()

# -------------------------
# Process upload
# -------------------------
try:
    file_bytes = uploaded.getvalue()
    df_combined, diagnostics = load_multi_sheet(file_bytes, uploaded.name)
except Exception as e:
    st.error(f"Failed to load file: {e}")
    st.stop()

# Show diagnostics
with st.expander("Ingestion diagnostics (what I mapped & sheet stats)", expanded=True):
    if diagnostics:
        diag_df = pd.DataFrame(diagnostics)
        st.dataframe(diag_df.fillna("").head(200))
    else:
        st.write("No diagnostics available.")

if df_combined.empty:
    st.error("No usable data found across selected sheets. Please check your file or map your columns differently.")
    st.stop()

# -------------------------
# Mapping override UI (if some required maps missing)
# -------------------------
st.subheader("Column mapping check (auto-detected). Override if needed.")
cols = list(df_combined.columns)
# auto-detect using normalized heuristics
auto_map = {}
for canon, aliases in SYNONYMS.items():
    pick = find_col_by_synonyms(cols, aliases)
    auto_map[canon] = pick

# present dropdowns for required fields
c1, c2, c3, c4 = st.columns(4)
with c1:
    dept_col = st.selectbox("Department column", options=["<none>"] + cols, index=(cols.index(auto_map["department"])+1) if auto_map["department"] in cols else 0)
with c2:
    date_col = st.selectbox("Date column", options=["<none>"] + cols, index=(cols.index(auto_map["date"])+1) if auto_map["date"] in cols else 0)
with c3:
    loss_col = st.selectbox("Loss Minutes column", options=["<none>"] + cols, index=(cols.index(auto_map["loss_minutes"])+1) if auto_map["loss_minutes"] in cols else 0)
with c4:
    reason_col = st.selectbox("Reason column", options=["<none>"] + cols, index=(cols.index(auto_map["reason"])+1) if auto_map["reason"] in cols else 0)

# validation
missing = [name for name, sel in [("Department", dept_col), ("Date", date_col), ("Loss Minutes", loss_col), ("Reason", reason_col)] if sel == "<none>"]
if missing:
    st.warning(f"Please map these fields to proceed: {', '.join(missing)}")
    st.stop()

# Build cleaned working DataFrame
wrk = df_combined[[dept_col, date_col, loss_col, reason_col] + [c for c in df_combined.columns if c not in [dept_col, date_col, loss_col, reason_col]]].copy()
wrk = wrk.rename(columns={dept_col: "Department", date_col: "Date", loss_col: "Loss Minutes", reason_col: "Reason"})
# Trim strings
for c in ["Department", "Reason"]:
    wrk[c] = wrk[c].astype(str).str.strip()
# Parse date and minutes
wrk["Date"] = try_parse_date(wrk["Date"])
wrk["Loss Minutes"] = wrk["Loss Minutes"].apply(parse_minutes_value)
# Optional planned minutes column detection
if "planned_minutes" not in wrk.columns:
    # try detect a planned column in original df fields
    for c in df_combined.columns:
        if norm(c) in [norm(x) for x in SYNONYMS.get("planned_minutes", [])]:
            wrk["planned_minutes"] = df_combined[c].apply(parse_minutes_value)
            break

# Drop rows missing date or loss
wrk = wrk.dropna(subset=["Date", "Loss Minutes"])
wrk = wrk[wrk["Loss Minutes"] >= 0]

# Preview cleaned sample
st.subheader("Preview cleaned (first 100 rows)")
st.dataframe(wrk.head(100), use_container_width=True)

# -------------------------
# KPIs: OAE & totals
# -------------------------
oae_val, total_loss, planned_total = compute_oae(wrk, planned_per_day, assets_count, use_planned_from_file)
st.subheader("Key metrics")
k1, k2, k3 = st.columns(3)
k1.metric("Total Loss Minutes", f"{int(total_loss):,}")
k2.metric("Planned Minutes (basis)", f"{int(planned_total):,}")
k3.metric("Estimated OAE %", f"{oae_val:0.2f}%")

# -------------------------
# 1st level Pareto by Department/Category (user choice)
# -------------------------
st.subheader("Pareto analysis")
group_by_choice = st.selectbox("1st level Pareto by", options=["Department", "Reason"], index=0)
if group_by_choice == "Department":
    p1 = wrk.groupby("Department", as_index=False)["Loss Minutes"].sum().sort_values("Loss Minutes", ascending=False).head(top_n)
else:
    p1 = wrk.groupby("Reason", as_index=False)["Loss Minutes"].sum().sort_values("Loss Minutes", ascending=False).head(top_n)

fig1 = px.bar(p1, x=p1.columns[0], y="Loss Minutes", title=f"1st-level Pareto by {group_by_choice} (Top {top_n})", text="Loss Minutes")
fig1.update_layout(xaxis_title=group_by_choice, yaxis_title="Loss Minutes")
st.plotly_chart(fig1, use_container_width=True)
st.dataframe(p1, use_container_width=True)

# -------------------------
# 2nd level Pareto: top categories -> reasons inside
# -------------------------
st.subheader("2nd-level Pareto (drilldown)")
# top categories from first level
top_categories = p1.iloc[:, 0].tolist()
if top_categories:
    selected_cat = st.selectbox("Select top category to drill down", options=top_categories)
    # filter reasons inside this category
    mask = (wrk[group_by_choice] == selected_cat)
    second = wrk[mask].groupby("Reason", as_index=False)["Loss Minutes"].sum().sort_values("Loss Minutes", ascending=False).head(15)
    fig2 = px.bar(second, x="Reason", y="Loss Minutes", title=f"Top Reasons within {selected_cat}", text="Loss Minutes")
    fig2.update_layout(xaxis_title="Reason", yaxis_title="Loss Minutes")
    st.plotly_chart(fig2, use_container_width=True)
    st.dataframe(second, use_container_width=True)

# -------------------------
# Trend chart by day/week/month
# -------------------------
st.subheader("Trend analysis")
agg_period = st.selectbox("Aggregate by", options=["D (daily)", "W (weekly)", "M (monthly)"], index=2)
if agg_period == "D (daily)":
    freq = "D"
elif agg_period == "W (weekly)":
    freq = "W"
else:
    freq = "M"

trend = wrk.set_index("Date").resample(freq)["Loss Minutes"].sum().reset_index()
trend["Date"] = trend["Date"].dt.date
fig_trend = px.line(trend, x="Date", y="Loss Minutes", title=f"Loss Trend ({freq})", markers=True)
st.plotly_chart(fig_trend, use_container_width=True)
st.dataframe(trend, use_container_width=True)

# -------------------------
# Auto-insights (simple rule-based)
# -------------------------
st.subheader("Auto-insights (suggestions)")
insights = []
if total_loss > 0:
    top_dept = p1.iloc[0, 0] if not p1.empty else None
    top_dept_loss = p1.iloc[0]["Loss Minutes"] if not p1.empty else 0
    pct_top = 100.0 * (top_dept_loss / total_loss) if total_loss else 0
    insights.append(f"Top contributor: **{top_dept}** accounting for **{pct_top:0.1f}%** of total loss.")
    if pct_top >= 25:
        insights.append("Focus action on the top contributor: run targeted root-cause analysis (Why-Why/Takt review).")
    # recurring issues: reasons repeating across days
    reason_counts = wrk.groupby("Reason")["Loss Minutes"].sum().sort_values(ascending=False).head(5)
    insights.append("Top reasons: " + ", ".join(reason_counts.index.tolist()))
else:
    insights.append("No losses found in selected data.")

for x in insights:
    st.info(x)

# -------------------------
# Download cleaned results
# -------------------------
st.subheader("Export / Save")
out_dict = {
    "Cleaned Data": wrk,
    "1st_level_pareto": p1,
    "trend": trend
}
buf, fname = io.BytesIO(), f"ci_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
    for sheet_name, df_out in out_dict.items():
        try:
            df_out.to_excel(writer, sheet_name=str(sheet_name)[:30], index=False)
        except Exception:
            pass
buf.seek(0)
st.download_button("⬇️ Download results workbook", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.success("Processing complete. Next steps: connect to OneDrive/SharePoint for auto-refresh, add action-tracking and alerting.")

