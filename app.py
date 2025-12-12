
# app.py
# Plotly-free Streamlit dashboard for Budget vs Actual expense analytics
# Uses Streamlit native charts + Pandas only.

import io
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Budget & Expense Analytics (No Plotly)", page_icon="📊", layout="wide")

CUSTOM_CSS = """
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
div[data-testid="stMetric"] {background: rgba(28, 32, 42, 0.40); border: 1px solid rgba(255,255,255,0.06);
                            padding: 12px 14px; border-radius: 16px;}
.small-note {opacity:0.75; font-size: 12px;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

COMPANY_MAP = {1000: "Green", 3000: "Solar", 4000: "EPC"}
MONTH_MAP = {
    "April": 4, "May": 5, "June": 6, "July": 7,
    "Aug": 8, "August": 8,
    "Sep": 9, "September": 9,
    "Oct": 10, "October": 10,
    "Nov": 11, "November": 11,
    "Dec": 12, "December": 12,
    "Jan": 1, "January": 1,
    "Feb": 2, "February": 2,
    "Mar": 3, "March": 3,
}

def inr(x: float) -> str:
    try:
        return f"₹{x:,.0f}"
    except Exception:
        return str(x)

def pct(x: float) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    return f"{x*100:,.1f}%"

def fy_month_start(mname: str, fy_start_year: int = 2025) -> pd.Timestamp:
    m = MONTH_MAP.get(str(mname).strip(), None)
    if m is None:
        return pd.NaT
    year = fy_start_year if m >= 4 else fy_start_year + 1
    return pd.Timestamp(year=year, month=m, day=1)

def detect_header_row(excel_bytes: bytes, sheet: str, needle: str = "Month", search_rows: int = 60) -> int:
    raw = pd.read_excel(io.BytesIO(excel_bytes), sheet_name=sheet, header=None, engine="openpyxl")
    top = raw.iloc[: min(search_rows, len(raw))]
    mask = top.apply(lambda r: r.astype(str).str.contains(needle, case=False, na=False).any(), axis=1)
    idx = np.where(mask.values)[0]
    return int(idx[0]) if len(idx) else 0

@st.cache_data(show_spinner=False)
def load_data(excel_bytes: bytes):
    # Actual sheet (header row auto-detect)
    h = detect_header_row(excel_bytes, "Actual", needle="Month")
    actual = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Actual", header=h, engine="openpyxl")
    actual = actual.loc[:, ~actual.columns.astype(str).str.contains(r"^Unnamed")]
    actual = actual.dropna(how="all")

    # Budget sheet (known header row = 2 in this file)
    budget = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Budget", header=2, engine="openpyxl")
    budget = budget.loc[:, ~budget.columns.astype(str).str.contains(r"^Unnamed")]
    budget = budget.dropna(how="all")

    return actual, budget

def clean_actual(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Act. Costs"] = pd.to_numeric(df.get("Act. Costs"), errors="coerce").fillna(0.0)
    df["Plant/HO"] = df.get("Plant/HO").replace({"Ho": "HO"})
    df["CompanyName"] = df.get("Company").map(COMPANY_MAP).fillna(df.get("Company").astype(str))
    df["MonthStart"] = df.get("Month").apply(fy_month_start)
    df["GL Code"] = df.get("GL").astype(str)
    df["Cost Centre"] = df.get("Cost Centers").astype(str)
    return df

def clean_budget(df: pd.DataFrame):
    df = df.copy()
    if "Sr. No." in df.columns:
        df = df[df["Sr. No."].notna()].copy()  # remove totals/blank rows

    df["CompanyName"] = df["Company"].astype(str)
    df["GL Code"] = df["GL Code"].astype(str)
    df["Cost Centre"] = df["Cost Centre"].astype(str)
    if "Plant/Ho" in df.columns:
        df["Plant/Ho"] = df["Plant/Ho"].replace({"Ho": "HO"})

    month_cols = [c for c in df.columns if isinstance(c, (pd.Timestamp, datetime))]
    long = df.melt(
        id_vars=[c for c in df.columns if c not in month_cols],
        value_vars=month_cols,
        var_name="MonthStart",
        value_name="BudgetCost",
    )
    long["BudgetCost"] = pd.to_numeric(long["BudgetCost"], errors="coerce").fillna(0.0)

    map_cols = [
        "CompanyName", "Cost Centre", "GL Code",
        "Department", "Plant/Ho", "Nature", "MIS MAPPING", "CC Maping", "Plant"
    ]
    map_cols = [c for c in map_cols if c in df.columns]
    mapping = df[map_cols].drop_duplicates(subset=["CompanyName", "Cost Centre", "GL Code"], keep="first")

    return long, mapping

def build_combo(actual_enriched: pd.DataFrame, budget_long: pd.DataFrame) -> pd.DataFrame:
    keys = ["CompanyName", "Cost Centre", "GL Code", "MonthStart"]
    act = actual_enriched.groupby(keys, as_index=False)["Act. Costs"].sum()
    bud = budget_long.groupby(keys, as_index=False)["BudgetCost"].sum()

    combo = act.merge(bud, on=keys, how="outer")
    combo["Act. Costs"] = combo["Act. Costs"].fillna(0.0)
    combo["BudgetCost"] = combo["BudgetCost"].fillna(0.0)
    combo["Variance"] = combo["Act. Costs"] - combo["BudgetCost"]
    combo["VarPct"] = np.where(combo["BudgetCost"].abs() > 1e-9, combo["Variance"] / combo["BudgetCost"], np.nan)
    return combo

# --------------------
# Sidebar / File loading
# --------------------
with st.sidebar:
    st.title("📊 Expense Dashboard")
    st.caption("No-Plotly build (uses Streamlit charts)")

    uploaded = st.file_uploader("Upload Excel", type=["xlsx"], help="Upload CC Report excel")
    default_path = "CC Report YTD Oct - 2025_v2_101125.xlsx"

    if uploaded is None:
        try:
            excel_bytes = open(default_path, "rb").read()
            st.success("Using packaged file")
        except Exception:
            st.warning("Upload the Excel file to continue.")
            st.stop()
    else:
        excel_bytes = uploaded.getvalue()

actual_raw, budget_raw = load_data(excel_bytes)
actual = clean_actual(actual_raw)
budget_long, budget_map = clean_budget(budget_raw)

# constrain budget to actual months
min_m, max_m = actual["MonthStart"].min(), actual["MonthStart"].max()
budget_long = budget_long[(budget_long["MonthStart"] >= min_m) & (budget_long["MonthStart"] <= max_m)]

actual_enriched = actual.merge(budget_map, on=["CompanyName", "Cost Centre", "GL Code"], how="left")
combo = build_combo(actual_enriched, budget_long)
combo = combo.merge(budget_map, on=["CompanyName", "Cost Centre", "GL Code"], how="left")

# --------------------
# Filters
# --------------------
with st.sidebar:
    page = st.radio("Navigate", ["Executive Summary", "Variance Explorer", "Trend & Mix", "Transactions", "Data Quality"], index=0)

    months = sorted(combo["MonthStart"].dropna().unique())
    month_sel = st.slider("Month range", min_value=min(months), max_value=max(months),
                          value=(min(months), max(months)), format="MMM YYYY")

    company_sel = st.multiselect("Company", sorted(combo["CompanyName"].dropna().unique()),
                                 default=sorted(combo["CompanyName"].dropna().unique()))
    dept_sel = st.multiselect("Department", sorted(combo.get("Department", pd.Series(dtype=str)).dropna().unique()))
    ccmap_sel = st.multiselect("Category (CC Maping)", sorted(combo.get("CC Maping", pd.Series(dtype=str)).dropna().unique()))
    nature_sel = st.multiselect("Nature", sorted(combo.get("Nature", pd.Series(dtype=str)).dropna().unique()))

    breakdown_dim = st.selectbox("Breakdown dimension", ["CC Maping", "Department", "Nature", "GL Code", "Cost Centre"], index=0)

# Apply filters
cf = combo.copy()
cf = cf[(cf["MonthStart"] >= month_sel[0]) & (cf["MonthStart"] <= month_sel[1])]
cf = cf[cf["CompanyName"].isin(company_sel)]
if dept_sel and "Department" in cf.columns:
    cf = cf[cf["Department"].isin(dept_sel)]
if ccmap_sel and "CC Maping" in cf.columns:
    cf = cf[cf["CC Maping"].isin(ccmap_sel)]
if nature_sel and "Nature" in cf.columns:
    cf = cf[cf["Nature"].isin(nature_sel)]

A = float(cf["Act. Costs"].sum())
B = float(cf["BudgetCost"].sum())
V = float(cf["Variance"].sum())
VP = (V / B) if abs(B) > 1e-9 else np.nan

# --------------------
# Pages
# --------------------
if page == "Executive Summary":
    st.title("Executive Summary")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Actual", inr(A))
    c2.metric("Budget", inr(B))
    c3.metric("Variance (A - B)", inr(V), delta=pct(VP))
    c4.metric("Budget Utilization", pct(A / B) if abs(B) > 1e-9 else "—")

    st.divider()
    m = cf.groupby("MonthStart", as_index=False)[["Act. Costs", "BudgetCost"]].sum().sort_values("MonthStart").set_index("MonthStart")
    st.subheader("Monthly Spend: Actual vs Budget")
    st.line_chart(m[["BudgetCost", "Act. Costs"]])

    st.divider()
    dim = breakdown_dim
    if dim in cf.columns:
        top = cf.groupby(dim, as_index=False)[["Act. Costs", "BudgetCost", "Variance"]].sum() \
                .sort_values("Act. Costs", ascending=False).head(20)
        st.subheader(f"Top 20 by Actual ({dim})")
        st.bar_chart(top.set_index(dim)["Act. Costs"])
        show = top.copy()
        show["Act. Costs"] = show["Act. Costs"].map(inr)
        show["BudgetCost"] = show["BudgetCost"].map(inr)
        show["Variance"] = show["Variance"].map(inr)
        st.dataframe(show, use_container_width=True, height=360)
    else:
        st.info("Selected breakdown dimension not available.")

elif page == "Variance Explorer":
    st.title("Variance Explorer")
    dim = breakdown_dim
    if dim not in cf.columns:
        st.warning("Selected breakdown dimension isn't available.")
        st.stop()

    g = cf.groupby(dim, as_index=False)[["Act. Costs", "BudgetCost", "Variance"]].sum()
    g["VarPct"] = np.where(g["BudgetCost"].abs() > 1e-9, g["Variance"] / g["BudgetCost"], np.nan)

    col1, col2 = st.columns([0.6, 0.4])
    with col1:
        show_n = st.slider("Top N", 5, 50, 20)
        order = st.selectbox("Sort", ["Overspend", "Underspend", "Absolute variance"], index=0)

        if order == "Overspend":
            view = g.sort_values("Variance", ascending=False).head(show_n)
        elif order == "Underspend":
            view = g.sort_values("Variance", ascending=True).head(show_n)
        else:
            view = g.assign(AbsVar=g["Variance"].abs()).sort_values("AbsVar", ascending=False).head(show_n)

        disp = view.copy()
        disp["Act. Costs"] = disp["Act. Costs"].map(inr)
        disp["BudgetCost"] = disp["BudgetCost"].map(inr)
        disp["Variance"] = disp["Variance"].map(inr)
        disp["VarPct"] = view["VarPct"].map(pct)
        st.dataframe(disp, use_container_width=True, height=420)

        st.subheader("Variance bar")
        st.bar_chart(view.set_index(dim)["Variance"])

    with col2:
        st.subheader("Variance heatmap (table styling)")
        pivot = cf.pivot_table(index=dim, columns="MonthStart", values="Variance", aggfunc="sum", fill_value=0)
        top_rows = pivot.abs().sum(axis=1).sort_values(ascending=False).head(25).index
        pivot = pivot.loc[top_rows]
        st.dataframe(pivot.style.background_gradient(cmap="RdBu", axis=None), use_container_width=True, height=520)

elif page == "Trend & Mix":
    st.title("Trend & Mix")
    dim = breakdown_dim
    if dim not in cf.columns:
        st.warning("Selected breakdown dimension isn't available.")
        st.stop()

    t = cf.groupby(["MonthStart", dim], as_index=False)["Act. Costs"].sum()
    topcats = t.groupby(dim, as_index=False)["Act. Costs"].sum().sort_values("Act. Costs", ascending=False).head(12)[dim].tolist()
    t[dim] = np.where(t[dim].isin(topcats), t[dim], "Others")
    t = t.groupby(["MonthStart", dim], as_index=False)["Act. Costs"].sum()

    pv = t.pivot_table(index="MonthStart", columns=dim, values="Act. Costs", aggfunc="sum", fill_value=0).sort_index()
    st.subheader(f"Actual mix over time (Top 12 {dim} + Others)")
    st.area_chart(pv)

    st.subheader("Anomaly signals (z-score)")
    zdim = st.selectbox("Anomaly dimension", ["CC Maping", "Department", "GL Code"], index=0)
    if zdim in cf.columns:
        z = cf.groupby(["MonthStart", zdim], as_index=False)["Act. Costs"].sum()
        z["z"] = z.groupby(zdim)["Act. Costs"].transform(lambda s: (s - s.mean()) / (s.std(ddof=0) + 1e-9))
        anomalies = z[z["z"].abs() >= 3].sort_values("z", key=lambda s: s.abs(), ascending=False).head(40)
        if anomalies.empty:
            st.info("No anomalies found at |z| ≥ 3 for the selected filters.")
        else:
            disp = anomalies.copy()
            disp["Act. Costs"] = disp["Act. Costs"].map(inr)
            st.dataframe(disp, use_container_width=True, height=300)

elif page == "Transactions":
    st.title("Transactions")

    act = actual_enriched.copy()
    act = act[(act["MonthStart"] >= month_sel[0]) & (act["MonthStart"] <= month_sel[1])]
    act = act[act["CompanyName"].isin(company_sel)]
    if dept_sel and "Department" in act.columns:
        act = act[act["Department"].isin(dept_sel)]
    if ccmap_sel and "CC Maping" in act.columns:
        act = act[act["CC Maping"].isin(ccmap_sel)]
    if nature_sel and "Nature" in act.columns:
        act = act[act["Nature"].isin(nature_sel)]

    q = st.text_input("Search (GL, Cost Centre, Short text)")
    if q:
        mask = (
            act["GL"].astype(str).str.contains(q, case=False, na=False)
            | act["Cost Centers"].astype(str).str.contains(q, case=False, na=False)
            | act["Short text"].astype(str).str.contains(q, case=False, na=False)
        )
        act = act[mask]

    show_cols = [
        "Month", "MonthStart", "CompanyName", "Plant", "Cost Centers", "Department",
        "Plant/HO", "GL", "Short text", "Nature", "Group", "MIS MAPPING", "CC Maping", "Act. Costs"
    ]
    show_cols = [c for c in show_cols if c in act.columns]
    disp = act[show_cols].copy()
    disp["Act. Costs"] = disp["Act. Costs"].map(inr)
    st.dataframe(disp, use_container_width=True, height=560)

    st.download_button(
        "Download filtered transactions (CSV)",
        data=act[show_cols].to_csv(index=False).encode("utf-8"),
        file_name="transactions_filtered.csv",
        mime="text/csv",
    )

elif page == "Data Quality":
    st.title("Data Quality & Coverage")

    act = actual_enriched.copy()
    act = act[(act["MonthStart"] >= month_sel[0]) & (act["MonthStart"] <= month_sel[1])]

    total_rows = len(act)
    mapped_ccmap = act["CC Maping"].notna().mean() if "CC Maping" in act.columns else 0
    mapped_mis = act["MIS MAPPING"].notna().mean() if "MIS MAPPING" in act.columns else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Actual rows", f"{total_rows:,}")
    c2.metric("Mapped CC Maping", pct(mapped_ccmap))
    c3.metric("Mapped MIS MAPPING", pct(mapped_mis))

    st.divider()

    act_keys = act[["CompanyName", "Cost Centre", "GL Code"]].drop_duplicates()
    bud_keys = budget_map[["CompanyName", "Cost Centre", "GL Code"]].drop_duplicates()
    unmatched = act_keys.merge(bud_keys, on=["CompanyName", "Cost Centre", "GL Code"], how="left", indicator=True)
    unmatched = unmatched[unmatched["_merge"] == "left_only"].drop(columns=["_merge"])

    st.subheader("Actual lines without matching Budget mapping")
    st.dataframe(unmatched.head(200), use_container_width=True, height=520)

    st.download_button(
        "Download unmatched actual keys (CSV)",
        data=unmatched.to_csv(index=False).encode("utf-8"),
        file_name="unmatched_actual_keys.csv",
        mime="text/csv",
    )

st.markdown("<div class='small-note'>No-Plotly build. If you fix Plotly installation, you can switch back to richer visuals later.</div>", unsafe_allow_html=True)
