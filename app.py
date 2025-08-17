import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import plotly.express as px

st.set_page_config(page_title="Virtual CI Specialist", layout="wide")
st.title("📊 Virtual CI Specialist — Multi-Sheet Excel Ingest (Robust)")

st.caption(
    "Upload one Excel workbook (.xlsx). The app auto-detects header rows (even in templates), "
    "maps columns (Department, Date, Loss Minutes, Reason), cleans types, merges valid sheets, "
    "and visualizes OAE proxy, Pareto, trends, and department breakdown."
)

# ----------------------------- Utilities -------------------------------------

def sanitize(s: str) -> str:
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

ALIAS = {
    "department": {
        "exact": [
            "department","dept","area","line","line name","cell","work cell","workcell",
            "process","work center","workcenter","machine","asset","value stream","section","shop"
        ],
        "contains": ["dept","area","line","cell","process","work center","workcenter","machine","asset","stream","section","shop"]
    },
    "date": {
        "exact": ["date","event date","shift date","production date","recorded date","day","timestamp","time"],
        "contains": ["date","shift","day","time","stamp"]
    },
    "loss_minutes": {
        "exact": [
            "loss minutes","loss min","downtime minutes","downtime (min)","downtime min","minutes lost",
            "loss (min)","duration (min)","time lost (min)","stop time (min)","stoppage minutes","dt (min)"
        ],
        "contains": ["min","minute","downtime","loss time","duration","stop","stoppage","time lost"]
    },
    "reason": {
        "exact": ["reason","downtime reason","loss reason","issue","category","sub category","cause","problem","failure mode","reason code"],
        "contains": ["reason","issue","category","cause","problem","failure","code"]
    },
}

def score_header_row(row_vals: list[str]) -> int:
    toks = [sanitize(x) for x in row_vals if str(x).strip() != ""]
    bag = " | ".join(toks)
    score = 0
    # Reward presence of any alias tokens
    for key in ALIAS:
        ex = any(sanitize(a) in toks for a in ALIAS[key]["exact"])
        ct = any(any(tok in t for tok in ALIAS[key]["contains"]) for t in toks)
        if ex or ct:
            score += 1
    # Reward generic hints
    if re.search(r"\bmin\b|\bminute\b|downtime", bag):
        score += 1
    if re.search(r"\bdate\b|\bshift\b|\bday\b|\btime\b", bag):
        score += 1
    return score

def sniff_header_and_reframe(xls: pd.ExcelFile, sheet: str) -> pd.DataFrame:
    """Find a probable header row within first 30 rows; return a DataFrame with proper columns."""
    df0 = pd.read_excel(xls, sheet_name=sheet, header=None, dtype=object)
    if df0.empty:
        return pd.DataFrame()
    max_rows_to_scan = min(30, len(df0))
    best_idx, best_score = None, -1
    for r in range(max_rows_to_scan):
        row_vals = df0.iloc[r].tolist()
        sc = score_header_row(row_vals)
        if sc > best_score:
            best_idx, best_score = r, sc
    # If nothing scored, fallback to first non-empty row
    if best_idx is None:
        best_idx = 0
    # Build framed df
    header = df0.iloc[best_idx].astype(str).tolist()
    df = df0.iloc[best_idx + 1:].copy()
    df.columns = header
    # Drop fully-empty columns and rows
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    return df

def auto_map_columns(df: pd.DataFrame) -> dict:
    normalized = {sanitize(c): c for c in df.columns}
    found = {}
    def find_for(key):
        # exact match first
        for label in ALIAS[key]["exact"]:
            lbl = sanitize(label)
            if lbl in normalized:
                return normalized[lbl]
        # contains heuristic
        for k_norm, orig in normalized.items():
            if any(tok in k_norm for tok in ALIAS[key]["contains"]):
                return orig
        return None
    for req in ["department","date","loss_minutes","reason"]:
        col = find_for(req)
        if col is not None:
            found[req] = col
    return found

def detect_date_column(df: pd.DataFrame) -> str | None:
    best_col, best_hits = None, -1
    for c in df.columns:
        try:
            parsed = pd.to_datetime(df[c], errors="coerce")
            hits = parsed.notna().sum()
            if hits > best_hits and hits > 0:
                best_col, best_hits = c, hits
        except Exception:
            continue
    return best_col

def coerce_to_minutes(series: pd.Series) -> pd.Series:
    s = series.copy()

    # If already numeric, maybe Excel day fraction (<= 1)
    if pd.api.types.is_numeric_dtype(s):
        # treat very small decimals as Excel days -> minutes
        # (but only if a good chunk are between 0 and 1)
        mask = s.between(0, 1, inclusive="neither")
        frac_ratio = (mask & s.notna()).sum() / max(1, s.notna().sum())
        if frac_ratio > 0.5:
            return s.astype(float) * 24.0 * 60.0
        return pd.to_numeric(s, errors="coerce")

    # Strings: strip units and handle hh:mm[:ss]
    s = s.astype(str).str.strip()

    # Handle hh:mm or hh:mm:ss
    time_like = s.str.match(r"^\s*\d{1,2}:\d{2}(:\d{2})?\s*$", na=False)
    if time_like.any():
        mins = pd.to_timedelta(s.where(time_like, np.nan), errors="coerce") / pd.Timedelta(minutes=1)
        # others: fall through
        s = s.mask(time_like, mins)

    # Remove non-numeric except dot
    s = s.where(s.apply(lambda x: isinstance(x, (int, float))), s.astype(str).str.replace(r"[^0-9.\-]", "", regex=True))
    return pd.to_numeric(s, errors="coerce")

def best_guess_mapping(df: pd.DataFrame) -> dict:
    mapping = {}

    # Date
    date_col = detect_date_column(df)
    if date_col: mapping["date"] = date_col

    # Loss minutes: header hint first, else numeric density
    cand = [c for c in df.columns if any(tok in sanitize(c) for tok in ["min","minute","downtime","loss","duration","stop","stoppage","time"])]
    loss_col = cand[0] if cand else None
    if loss_col is None:
        # pick the numeric-ish column with highest non-null after coercion
        best_c, best_hits = None, -1
        for c in df.columns:
            hits = pd.to_numeric(df[c], errors="coerce").notna().sum()
            if hits > best_hits:
                best_c, best_hits = c, hits
        loss_col = best_c
    if loss_col: mapping["loss_minutes"] = loss_col

    # Reason: header hint else most text-like (high unique, mid-length)
    cand = [c for c in df.columns if any(tok in sanitize(c) for tok in ["reason","issue","category","cause","problem","failure","code","comment","description","desc"])]
    reason_col = cand[0] if cand else None
    if reason_col is None:
        text_cols = []
        for c in df.columns:
            if df[c].dtype == object:
                nunq = df[c].nunique(dropna=True)
                text_cols.append((c, nunq))
        if text_cols:
            reason_col = sorted(text_cols, key=lambda x: -x[1])[0][0]
    if reason_col: mapping["reason"] = reason_col

    # Department: header hint else lowest-cardinality text (but >1)
    cand = [c for c in df.columns if any(tok in sanitize(c) for tok in ["dept","department","area","line","cell","workcenter","work center","machine","asset","section","shop","stream"])]
    dept_col = cand[0] if cand else None
    if dept_col is None:
        text_cols = []
        for c in df.columns:
            if df[c].dtype == object:
                nunq = df[c].nunique(dropna=True)
                if nunq > 1:
                    text_cols.append((c, nunq))
        if text_cols:
            dept_col = sorted(text_cols, key=lambda x: x[1])[0][0]
    if dept_col: mapping["department"] = dept_col

    return mapping

def clean_standardize(df_in: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    # Try auto map; if incomplete, try best-guess
    mapping = auto_map_columns(df_in)
    needed = {"department","date","loss_minutes","reason"}
    if not needed.issubset(mapping.keys()):
        mapping = {**mapping, **{k:v for k,v in best_guess_mapping(df_in).items() if k not in mapping}}
    if not needed.issubset(mapping.keys()):
        missing = list(needed - set(mapping.keys()))
        raise ValueError(f"Missing required columns after mapping: {missing}")

    df = pd.DataFrame()
    df["department"] = df_in[mapping["department"]].astype(str).str.strip()
    df["reason"]     = df_in[mapping["reason"]].astype(str).str.strip()

    # Dates
    df["date"] = pd.to_datetime(df_in[mapping["date"]], errors="coerce")

    # Loss minutes (robust)
    df["loss_minutes"] = coerce_to_minutes(df_in[mapping["loss_minutes"]])

    # Drop bad rows
    df = df.dropna(subset=["department","reason","date","loss_minutes"])
    df = df[df["loss_minutes"] >= 0]
    df["sheet"] = sheet_name
    return df

def calc_oae(df: pd.DataFrame, denom_minutes: float) -> float:
    loss = df["loss_minutes"].sum()
    if denom_minutes <= 0:
        return 0.0
    return float(np.clip((denom_minutes - loss) / denom_minutes * 100.0, 0, 100))

# ----------------------------- Sidebar ---------------------------------------

with st.sidebar:
    st.header("⚙️ Settings")
    minutes_per_period = st.number_input(
        "Minutes per analysis period",
        min_value=1, max_value=1_000_000, value=1440, step=1,
        help="Use 1440 for a day (per asset). This is a simple OAE proxy for now."
    )
    gran = st.selectbox("Trend granularity", ["Day","Week","Month"], index=1)

# ----------------------------- Upload ----------------------------------------

upl = st.file_uploader("📥 Upload an Excel workbook (.xlsx).", type=["xlsx"])
if not upl:
    st.info("Upload your workbook to begin.")
    st.stop()

try:
    xls = pd.ExcelFile(upl)
    sheets = xls.sheet_names
except Exception as e:
    st.error(f"Could not read Excel: {e}")
    st.stop()

mode = st.radio("Select sheets to process", [f"All sheets ({len(sheets)})", "Choose sheets"], horizontal=True)
chosen = sheets if mode.startswith("All") else st.multiselect("Pick sheets", options=sheets, default=sheets)

if not chosen:
    st.warning("No sheets selected.")
    st.stop()

# ----------------------------- Process ---------------------------------------

frames, skipped = [], []
for s in chosen:
    try:
        df_s = sniff_header_and_reframe(xls, s)
        if df_s.empty or len(df_s.columns) == 0:
            skipped.append((s, "Empty or no columns"))
            continue
        cleaned = clean_standardize(df_s, s)
        if cleaned.empty:
            skipped.append((s, "No valid rows after cleaning"))
            continue
        frames.append(cleaned)
    except ValueError as ve:
        skipped.append((s, f"Mapping error: {ve}"))
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
st.success(f"✅ Loaded **{len(df):,}** rows from **{len(frames)}** sheet(s).")
if skipped:
    with st.expander("ℹ️ Sheets skipped or partially invalid"):
        for s, msg in skipped:
            st.write(f"• **{s}** — {msg}")

# Preview + download
st.subheader("🧹 Cleaned data (first 200 rows)")
st.dataframe(df.head(200), use_container_width=True)
st.download_button("⬇️ Download cleaned data (CSV)", df.to_csv(index=False).encode("utf-8"),
                   file_name="cleaned_ci_data.csv", mime="text/csv")

# KPIs
oae = calc_oae(df, minutes_per_period)
col1, col2 = st.columns(2)
col1.metric("OAE % (proxy)", f"{oae:.2f}%")
col2.metric("Total Loss Minutes", f"{df['loss_minutes'].sum():,.0f}")

# Pareto
st.markdown("---")
st.subheader("📌 Pareto — Loss Minutes by Reason")
pareto = (df.groupby("reason", as_index=False)["loss_minutes"]
          .sum().sort_values("loss_minutes", ascending=False))
fig_p = px.bar(pareto, x="reason", y="loss_minutes", text="loss_minutes", title="Top Loss Reasons (Pareto)")
fig_p.update_xaxes(tickangle=45)
st.plotly_chart(fig_p, use_container_width=True)

# Trend
st.subheader("📈 Trend")
df_tr = df.copy()
if gran == "Day":
    df_tr["period"] = df_tr["date"].dt.to_period("D").dt.to_timestamp()
elif gran == "Week":
    df_tr["period"] = df_tr["date"].dt.to_period("W").dt.start_time
else:
    df_tr["period"] = df_tr["date"].dt.to_period("M").dt.to_timestamp()
trend = df_tr.groupby("period", as_index=False)["loss_minutes"].sum().sort_values("period")
fig_t = px.line(trend, x="period", y="loss_minutes", markers=True, title=f"Loss Minutes Trend by {gran}")
st.plotly_chart(fig_t, use_container_width=True)

# Department breakdown
st.subheader("🏭 Department Breakdown")
dept = (df.groupby("department", as_index=False)["loss_minutes"]
        .sum().sort_values("loss_minutes", ascending=False))
fig_d = px.bar(dept, x="department", y="loss_minutes", text="loss_minutes", title="Loss Minutes by Department")
fig_d.update_xaxes(tickangle=45)
st.plotly_chart(fig_d, use_container_width=True)

st.caption("Header sniffing + smart mapping enabled. Standardized columns: department, date, loss_minutes, reason.")
