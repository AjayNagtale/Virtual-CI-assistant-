# app.py - Virtual CI Specialist (Full MVP)
import io
import re
from datetime import datetime
from typing import List, Dict, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

# -------------------------
# Page config
# -------------------------
st.set_page_config(page_title="Virtual CI Specialist — Full MVP", layout="wide")
st.title("🏭 Virtual CI Specialist — Full MVP")
st.caption("Upload multi-sheet Excel / CSV, auto-clean, map columns, compute OAE/OEE basics, show Pareto & trends, download results.")

# -------------------------
# Synonyms & helpers
# -------------------------
SYNONYMS = {
    "date": ["date", "day", "dt", "timestamp", "prod date", "productiondate", "shiftdate"],
    "department": ["department", "dept", "line", "area", "cell", "workcenter", "work center", "machine", "asset"],
    "reason": ["reason", "cause", "category", "lossreason", "downtimereason", "rootcause", "failuremode"],
    "loss_minutes": ["lossminutes", "loss mins", "loss_min", "loss (min)", "downtime", "downtime minutes", "minuteslost", "mins"],
    "planned_minutes": ["plannedminutes", "planned minutes", "available minutes", "scheduled minutes", "planned"]
}

REQUIRED = ["date", "department", "reason", "loss_minutes"]

def norm_col(c: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(c).strip().lower())

def find_col(cols: List[str], aliases: List[str]) -> str:
    """Return first matching original column name or None."""
    normalized = {c: norm_col(c) for c in cols}
    alias_norm = [norm_col(a) for a in aliases]
    # exact match
    for orig, n in normalized.items():
        if n in alias_norm:
            return orig
    # contains or partial match
    for orig, n in normalized.items():
        for a in alias_norm:
            if a in n or n in a:
                return orig
    return None

def parse_minutes(v):
    if pd.isna(v):
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).strip()
    if s == "":
        return np.nan
    # hh:mm format
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if m:
        h = int(m.group(1)); mm = int(m.group(2))
        return h * 60 + mm
    # find first number in string
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
# Load all sheets and clean per sheet
# -------------------------
@st.cache_data(show_spinner=False)
def load_all_sheets(file_bytes: bytes, filename: str) -> Tuple[pd.DataFrame, List[Dict]]:
    diagnostics = []
    frames = []
    try:
        if filename.lower().endswith(".csv"):
            df_raw = pd.read_csv(io.BytesIO(file_bytes))
            df_clean, diag = clean_single_sheet(df_raw, "(csv)")
            diagnostics.append(diag)
            if not df_clean.empty:
                df_clean["source_sheet"] = "(csv)"
                frames.append(df_clean)
        else:
            xls = pd.ExcelFile(io.BytesIO(file_bytes))
            for s in xls.sheet_names:
                try:
                    df_raw = pd.read_excel(xls, sheet_name=s, header=0)
                except Exception as e:
                    diagnostics.append({"sheet": s, "error": f"read failed: {e}"})
                    continue
                df_clean, diag = clean_single_sheet(df_raw, s)
                diagnostics.append(diag)
                if not df_clean.empty:
                    df_clean["source_sheet"] = s
                    frames.append(df_clean)
    except Exception as e:
        return pd.DataFrame(), [{"error": f"load failed: {e}"}]

    if frames:
        combined = pd.concat(frames, ignore_index=True, sort=False)
        # ensure datetime and numeric types
        if "date" in combined.columns:
            combined["date"] = try_parse_date(combined["date"])
        if "loss_minutes" in combined.columns:
            combined["loss_minutes"] = pd.to_numeric(combined["loss_minutes"], errors="coerce")
        # drop rows missing mandatory fields
        combined = combined.dropna(subset=[c for c in ["date", "loss_minutes"] if c in combined.columns])
        return combined, diagnostics
    return pd.DataFrame(), diagnostics

def clean_single_sheet(df: pd.DataFrame, sheet_name: str) -> Tuple[pd.DataFrame, Dict]:
    diag = {"sheet": sheet_name, "rows_in": len(df)}
    df0 = df.copy()
    # normalize column strings (strip)
    df0.columns = [str(c).strip() for c in df0.columns]
    cols = list(df0.columns)
    mapped = {}
    for canon, aliases in SYNONYMS.items():
        pick = find_col(cols, aliases)
        if pick:
            mapped[canon] = pick
    missing = [r for r in REQUIRED if r not in mapped]
    # rename
    rename_map = {orig: canon for canon, orig in mapped.items()}
    df_ren = df0.rename(columns=rename_map)
    # parse and clean
    if "date" in df_ren.columns:
        df_ren["date"] = try_parse_date(df_ren["date"])
    if "loss_minutes" in df_ren.columns:
        df_ren["loss_minutes"] = df_ren["loss_minutes"].apply(parse_minutes)
    if "department" in df_ren.columns:
        df_ren["department"] = df_ren["department"].astype(str).str.strip()
    if "reason" in df_ren.columns:
        df_ren["reason"] = df_ren["reason"].astype(str).str.strip()
    if "planned_minutes" in df_ren.columns:
        df_ren["planned_minutes"] = df_ren["planned_minutes"].apply(parse_minutes)
    # drop rows missing date or minutes
    keep = pd.Series(True, index=df_ren.index)
    if "date" in df_ren.columns:
        keep &= df_ren["date"].notna()
    if "loss_minutes" in df_ren.columns:
        keep &= df_ren["loss_minutes"].notna()
    df_clean = df_ren.loc[keep].copy()
    # dedupe on core columns if exist
    core = [c for c in ["date", "department", "reason", "loss_minutes"] if c in df_clean.columns]
    if core:
        df_clean = df_clean.drop_duplicates(subset=core)
    diag.update({"rows_out": len(df_clean), "mapped": mapped, "missing_required": missing})
    return df_clean, diag

# -------------------------
# OAE & simple OEE helpers
# -------------------------
def compute_oae(df: pd.DataFrame, planned_per_day: float, assets_count: int, use_planned_from_file: bool) -> Tuple[float, float, float]:
    total_loss = float(df["loss_minutes"].sum()) if "loss_minutes" in df.columns else 0.0
    if use_planned_from_file and "planned_minutes" in df.columns and df["planned_minutes"].notna().any():
        planned_total = float(df["planned_minutes"].sum())
    else:
        n_days = df["date"].dt.date.nunique() if "date" in df.columns else 1
        planned_total = float(planned_per_day) * float(assets_count) * float(n_days)
    oae = max(0.0, 100.0 * (1.0 - (total_loss / planned_total))) if planned_total > 0 else 0.0
    return oae, total_loss, planned_total

def compute_simple_oee_placeholders(df: pd.DataFrame) -> Dict[str, float]:
    # placeholder: if you supply good_count & total_count we can compute quality/performance
    res = {"Availability": None, "Performance": None, "Quality": None, "OEE": None}
    # Availability can be derived if planned_minutes or planned_total known; currently rely on compute_oae
    return res

# -------------------------
# Sidebar controls
# -------------------------
with st.sidebar:
    st.header("Settings")
    use_planned_from_file = st.checkbox("Use 'Planned Minutes' from file if present", value=True)
    planned_per_day = st.number_input("Default planned minutes per day (per asset)", min_value=1, value=480, step=1)
    assets_count = st.number_input("Assets / lines in this dataset", min_value=1, value=1, step=1)
    top_n = st.slider("Pareto Top N", 3, 30, 8)
    st.markdown("---")
    st.caption("Tip: Upload your Excel. If the app cannot auto-map, use the mapping dropdowns shown.")

# -------------------------
# Upload UI
# -------------------------
uploaded = st.file_uploader("Upload Excel (.xlsx/.xls) or CSV (multi-sheet Excel supported)", type=["xlsx", "xls", "csv"])
if not uploaded:
    st.info("Upload your workbook (multi-sheet supported). Use the sample if you want to test.")
    st.stop()

# -------------------------
# Load & process
# -------------------------
file_bytes = uploaded.getvalue()
df_combined, diagnostics = load_all_sheets(file_bytes, uploaded.name)

with st.expander("Ingestion diagnostics (auto-mapping & rows)", expanded=True):
    if diagnostics:
        st.dataframe(pd.DataFrame(diagnostics).fillna("").head(200))
    else:
        st.write("No diagnostics available.")

if df_combined.empty:
    st.error("No usable data found across sheets. Check file or column names.")
    st.stop()

# -------------------------
# Column mapping UI (override)
# -------------------------
st.subheader("Column mapping (auto-detected — override if needed)")
cols = list(df_combined.columns)
auto_map = {}
for canon, aliases in SYNONYMS.items():
    pick = find_col(cols, aliases)
    auto_map[canon] = pick

c1, c2, c3, c4 = st.columns(4)
with c1:
    dept_col = st.selectbox("Department column", options=["<none>"] + cols, index=(cols.index(auto_map["department"]) + 1) if auto_map["department"] in cols else 0)
with c2:
    date_col = st.selectbox("Date column", options=["<none>"] + cols, index=(cols.index(auto_map["date"]) + 1) if auto_map["date"] in cols else 0)
with c3:
    loss_col = st.selectbox("Loss Minutes column", options=["<none>"] + cols, index=(cols.index(auto_map["loss_minutes"]) + 1) if auto_map["loss_minutes"] in cols else 0)
with c4:
    reason_col = st.selectbox("Reason column", options=["<none>"] + cols, index=(cols.index(auto_map["reason"]) + 1) if auto_map["reason"] in cols else 0)

missing_map = [n for n, sel in [("Department", dept_col), ("Date", date_col), ("Loss Minutes", loss_col), ("Reason", reason_col)] if sel == "<none>"]
if missing_map:
    st.warning(f"Please map these required fields to proceed: {', '.join(missing_map)}")
    st.stop()

# Build working df
wrk = df_combined.copy()
# keep mapped columns plus other optional columns
selected_cols = [dept_col, date_col, loss_col, reason_col]
# ensure they exist
for c in selected_cols:
    if c not in wrk.columns:
        wrk[c] = None
wrk = wrk[selected_cols + [c for c in wrk.columns if c not in selected_cols]]

# Rename for standard use
wrk = wrk.rename(columns={dept_col: "Department", date_col: "Date", loss_col: "Loss Minutes", reason_col: "Reason"})

# Basic cleaning
wrk["Department"] = wrk["Department"].astype(str).str.strip()
wrk["Reason"] = wrk["Reason"].astype(str).str.strip()
wrk["Date"] = try_parse_date(wrk["Date"])
wrk["Loss Minutes"] = wrk["Loss Minutes"].apply(parse_minutes)
# optional planned column detection
if "planned_minutes" not in wrk.columns:
    for c in df_combined.columns:
        if norm_col(c) in [norm_col(x) for x in SYNONYMS.get("planned_minutes", [])]:
            wrk["planned_minutes"] = df_combined[c].apply(parse_minutes)
            break

# drop invalid rows
wrk = wrk.dropna(subset=["Date", "Loss Minutes"])
wrk = wrk[wrk["Loss Minutes"] >= 0]

# Preview
st.subheader("Preview cleaned data (first 100 rows)")
st.dataframe(wrk.head(100), use_container_width=True)

# -------------------------
# KPIs
# -------------------------
oae_val, total_loss, planned_total = compute_oae(wrk, planned_per_day, assets_count, use_planned_from_file)
st.subheader("Key metrics")
k1, k2, k3 = st.columns(3)
k1.metric("Total Loss Minutes", f"{int(total_loss):,}")
k2.metric("Planned Minutes (basis)", f"{int(planned_total):,}")
k3.metric("Estimated OAE %", f"{oae_val:0.2f}%")

# -------------------------
# 1st level Pareto
# -------------------------
st.subheader("Pareto analysis (1st level)")
pareto_by = st.selectbox("1st level Pareto by", options=["Department", "Reason"], index=0)
if pareto_by == "Department":
    p1 = wrk.groupby("Department", as_index=False)["Loss Minutes"].sum().sort_values("Loss Minutes", ascending=False).head(top_n)
else:
    p1 = wrk.groupby("Reason", as_index=False)["Loss Minutes"].sum().sort_values("Loss Minutes", ascending=False).head(top_n)

fig1 = px.bar(p1, x=p1.columns[0], y="Loss Minutes", title=f"1st-level Pareto by {pareto_by} (Top {top_n})", text="Loss Minutes")
fig1.update_layout(xaxis_title=pareto_by, yaxis_title="Loss Minutes")
st.plotly_chart(fig1, use_container_width=True)
st.dataframe(p1, use_container_width=True)

# -------------------------
# 2nd level Pareto (drilldown)
# -------------------------
st.subheader("2nd-level Pareto (drill-down)")
top_cats = p1.iloc[:, 0].tolist() if not p1.empty else []
if top_cats:
    selected_cat = st.selectbox("Select category to drill", options=top_cats)
    if pareto_by == "Department":
        mask = wrk["Department"] == selected_cat
    else:
        mask = wrk["Reason"] == selected_cat
    second = wrk[mask].groupby("Reason", as_index=False)["Loss Minutes"].sum().sort_values("Loss Minutes", ascending=False).head(20)
    fig2 = px.bar(second, x="Reason", y="Loss Minutes", title=f"Top Reasons inside {selected_cat}", text="Loss Minutes")
    st.plotly_chart(fig2, use_container_width=True)
    st.dataframe(second, use_container_width=True)
else:
    st.info("No top categories found to drill down.")

# -------------------------
# Trend analysis
# -------------------------
st.subheader("Trend analysis")
agg_choice = st.selectbox("Aggregate by", options=["D (daily)", "W (weekly)", "M (monthly)"], index=2)
freq = "M" if agg_choice.startswith("M") else ("W" if agg_choice.startswith("W") else "D")
trend = wrk.set_index("Date").resample(freq)["Loss Minutes"].sum().reset_index().sort_values("Date")
trend["Date"] = trend["Date"].dt.date
fig_trend = px.line(trend, x="Date", y="Loss Minutes", title=f"Loss Trend ({freq})", markers=True)
st.plotly_chart(fig_trend, use_container_width=True)
st.dataframe(trend, use_container_width=True)

# -------------------------
# Auto insights (basic)
# -------------------------
st.subheader("Auto-insights (basic)")
ins = []
if total_loss > 0:
    top_row = p1.iloc[0] if not p1.empty else None
    if top_row is not None:
        top_name = top_row.iloc[0]
        top_loss = float(top_row["Loss Minutes"])
        pct = 100.0 * top_loss / total_loss
        ins.append(f"Top contributor: **{top_name}** with {top_loss:.0f} minutes ({pct:0.1f}% of total).")
        if pct >= 25:
            ins.append("Recommendation: Focus a Kaizen on this top contributor (Why-Why and countermeasures).")
    reasons_top = wrk.groupby("Reason")["Loss Minutes"].sum().sort_values(ascending=False).head(5)
    ins.append("Top reasons: " + ", ".join(reasons_top.index.astype(str).tolist()))
else:
    ins.append("No loss minutes detected in selected data.")

for i in ins:
    st.info(i)

# -------------------------
# Export
# -------------------------
st.subheader("Export cleaned results")
out = {
    "Cleaned Data": wrk,
    "Pareto_1": p1,
    "Trend": trend
}
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
    for name, dfout in out.items():
        try:
            dfout.to_excel(writer, sheet_name=str(name)[:31], index=False)
        except Exception:
            pass
buf.seek(0)
fname = f"ci_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
st.download_button("⬇️ Download results workbook", data=buf, file_name=fname, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.success("Done. Next steps: connect to OneDrive/SharePoint for auto-refresh, add action tracking & automated alerts.")
