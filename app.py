# app.py
# High-end Streamlit dashboard for Budget vs Actual (YTD) expense analytics
# Works with the provided Excel file: CC Report YTD Oct - 2025_v2_101125.xlsx

import io
from datetime import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------
# Page config + styling
# ----------------------------
st.set_page_config(
    page_title="Budget & Expense Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

CUSTOM_CSS = """
<style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    div[data-testid="stMetric"] {background: rgba(28, 32, 42, 0.40); border: 1px solid rgba(255,255,255,0.06);
                                padding: 12px 14px; border-radius: 16px;}
    div[data-testid="stMetric"] label {opacity: 0.85;}
    .kpi-row {gap: 0.75rem;}
    .small-note {opacity:0.75; font-size: 12px;}
    .section-title {font-size: 18px; font-weight: 700; margin-top: 0.25rem;}
    .pill {display:inline-block; padding: 0.18rem 0.55rem; border-radius: 999px; background: rgba(59,130,246,0.15);
           border: 1px solid rgba(59,130,246,0.25); font-size: 12px; margin-right: 6px;}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------------------
# Helpers
# ----------------------------
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
    # Actual sheet: find row containing Month/Period...
    h = detect_header_row(excel_bytes, "Actual", needle="Month")
    actual = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Actual", header=h, engine="openpyxl")
    actual = actual.loc[:, ~actual.columns.astype(str).str.contains(r"^Unnamed")]
    actual = actual.dropna(how="all")

    # Budget sheet is structured with header at row index 2 (as observed)
    budget = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="Budget", header=2, engine="openpyxl")
    budget = budget.loc[:, ~budget.columns.astype(str).str.contains(r"^Unnamed")]
    budget = budget.dropna(how="all")

    # Optional helper sheets
    try:
        comp = pd.read_excel(io.BytesIO(excel_bytes), sheet_name="comperison", header=None, engine="openpyxl")
    except Exception:
        comp = None

    return actual, budget, comp


def clean_actual(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["Act. Costs"] = pd.to_numeric(df.get("Act. Costs"), errors="coerce").fillna(0.0)
    df["Plant/HO"] = df.get("Plant/HO").replace({"Ho": "HO"})
    df["CompanyName"] = df.get("Company").map(COMPANY_MAP).fillna(df.get("Company").astype(str))
    df["MonthStart"] = df.get("Month").apply(fy_month_start)

    df["GL Code"] = df.get("GL").astype(str)
    df["Cost Centre"] = df.get("Cost Centers").astype(str)

    # Normalize text fields
    for c in ["Department", "Nature", "Group", "Short text", "Plant/HO"]:
        if c in df.columns:
            df[c] = df[c].astype(str).replace({"nan": np.nan})

    return df


def clean_budget(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Returns (budget_long, mapping_table).

    mapping_table is a distinct mapping of (Company, Cost Centre, GL Code) -> category fields.
    """
    df = df.copy()

    # Drop grand total / stray rows: keep only numbered rows
    if "Sr. No." in df.columns:
        df = df[df["Sr. No."].notna()].copy()

    # Standardize key columns
    df["CompanyName"] = df["Company"].astype(str)
    df["GL Code"] = df["GL Code"].astype(str)
    df["Cost Centre"] = df["Cost Centre"].astype(str)
    if "Plant/Ho" in df.columns:
        df["Plant/Ho"] = df["Plant/Ho"].replace({"Ho": "HO"})

    # Identify month columns (datetime)
    month_cols = [c for c in df.columns if isinstance(c, (pd.Timestamp, datetime))]

    long = df.melt(
        id_vars=[c for c in df.columns if c not in month_cols],
        value_vars=month_cols,
        var_name="MonthStart",
        value_name="BudgetCost",
    )
    long["BudgetCost"] = pd.to_numeric(long["BudgetCost"], errors="coerce").fillna(0.0)

    # Mapping table for enriching actual rows
    map_cols = [
        "CompanyName", "Cost Centre", "GL Code",
        "Department", "Plant/Ho", "Nature", "MIS MAPPING", "CC Maping", "Plant", "Particulars", "Particulars.1",
    ]
    map_cols = [c for c in map_cols if c in df.columns]
    mapping = df[map_cols].drop_duplicates(subset=["CompanyName", "Cost Centre", "GL Code"], keep="first")

    return long, mapping


def build_combo(actual_long: pd.DataFrame, budget_long: pd.DataFrame) -> pd.DataFrame:
    keys = ["CompanyName", "Cost Centre", "GL Code", "MonthStart"]

    act_agg = actual_long.groupby(keys, as_index=False)["Act. Costs"].sum()
    bud_agg = budget_long.groupby(keys, as_index=False)["BudgetCost"].sum()

    combo = act_agg.merge(bud_agg, on=keys, how="outer")
    combo["Act. Costs"] = combo["Act. Costs"].fillna(0.0)
    combo["BudgetCost"] = combo["BudgetCost"].fillna(0.0)
    combo["Variance"] = combo["Act. Costs"] - combo["BudgetCost"]
    combo["VarPct"] = np.where(combo["BudgetCost"].abs() > 1e-9, combo["Variance"] / combo["BudgetCost"], np.nan)

    return combo


def apply_filters(df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    out = df.copy()
    for col, selected in filters.items():
        if selected is None or selected == [] or selected == "(All)":
            continue
        if col not in out.columns:
            continue
        out = out[out[col].isin(selected)]
    return out


def top_n(df: pd.DataFrame, by: str, metric: str, n: int = 12, asc: bool = False) -> pd.DataFrame:
    g = df.groupby(by, as_index=False)[metric].sum().sort_values(metric, ascending=asc)
    return g.head(n)


# ----------------------------
# Load
# ----------------------------
with st.sidebar:
    st.title("📊 Expense Dashboard")
    st.caption("Budget vs Actual • Drilldowns • Anomalies")

    uploaded = st.file_uploader("Upload Excel (optional)", type=["xlsx"], help="If not provided, the packaged sample file is used.")

    default_path = "CC Report YTD Oct - 2025_v2_101125.xlsx"
    if uploaded is None:
        try:
            excel_bytes = open(default_path, "rb").read()
            st.success("Using packaged file")
        except Exception:
            st.warning("Please upload the Excel file.")
            st.stop()
    else:
        excel_bytes = uploaded.getvalue()

actual_raw, budget_raw, comp_raw = load_data(excel_bytes)
actual = clean_actual(actual_raw)
budget_long, budget_map = clean_budget(budget_raw)

# constrain budget months to actual month range
min_m, max_m = actual["MonthStart"].min(), actual["MonthStart"].max()
budget_long = budget_long[(budget_long["MonthStart"] >= min_m) & (budget_long["MonthStart"] <= max_m)]

# enrich actual with budget mapping (more reliable categories)
actual_enriched = actual.merge(
    budget_map,
    how="left",
    on=["CompanyName", "Cost Centre", "GL Code"],
    suffixes=("", "_bud"),
)

combo = build_combo(actual_enriched, budget_long)

# Attach categories to combo via mapping
combo = combo.merge(
    budget_map,
    how="left",
    on=["CompanyName", "Cost Centre", "GL Code"],
)

# ----------------------------
# Sidebar filters
# ----------------------------
with st.sidebar:
    st.divider()

    page = st.radio(
        "Navigate",
        ["Executive Summary", "Variance Explorer", "Trend & Mix", "Transactions", "Data Quality"],
        index=0,
    )

    # Common filters
    months = sorted([m for m in combo["MonthStart"].dropna().unique()])
    month_sel = st.slider(
        "Month range",
        min_value=min(months),
        max_value=max(months),
        value=(min(months), max(months)),
        format="MMM YYYY",
    )

    company_sel = st.multiselect("Company", sorted(combo["CompanyName"].dropna().unique()), default=sorted(combo["CompanyName"].dropna().unique()))

    # Category filters from budget mapping
    dept_sel = st.multiselect("Department", sorted(combo.get("Department", pd.Series(dtype=str)).dropna().unique()))
    ccmap_sel = st.multiselect("Category (CC Maping)", sorted(combo.get("CC Maping", pd.Series(dtype=str)).dropna().unique()))
    nature_sel = st.multiselect("Nature", sorted(combo.get("Nature", pd.Series(dtype=str)).dropna().unique()))

    st.divider()
    breakdown_dim = st.selectbox(
        "Breakdown dimension",
        ["CC Maping", "Department", "Nature", "GL Code", "Cost Centre"],
        index=0,
    )

# Apply filters to combo
combo_f = combo.copy()
combo_f = combo_f[(combo_f["MonthStart"] >= month_sel[0]) & (combo_f["MonthStart"] <= month_sel[1])]
combo_f = combo_f[combo_f["CompanyName"].isin(company_sel)]
if dept_sel:
    combo_f = combo_f[combo_f["Department"].isin(dept_sel)]
if ccmap_sel:
    combo_f = combo_f[combo_f["CC Maping"].isin(ccmap_sel)]
if nature_sel:
    combo_f = combo_f[combo_f["Nature"].isin(nature_sel)]

# ----------------------------
# KPI calculations
# ----------------------------
A = float(combo_f["Act. Costs"].sum())
B = float(combo_f["BudgetCost"].sum())
V = float(combo_f["Variance"].sum())
VP = (V / B) if abs(B) > 1e-9 else np.nan

# ----------------------------
# Executive Summary
# ----------------------------
if page == "Executive Summary":
    st.title("Executive Summary")
    st.caption("YTD view based on selected filters. Use sidebar to drill down.")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Actual", inr(A))
    c2.metric("Budget", inr(B))
    c3.metric("Variance (A - B)", inr(V), delta=pct(VP))
    c4.metric("Budget Utilization", pct(A / B) if abs(B) > 1e-9 else "—")

    st.markdown(
        f"<span class='pill'>Months: {month_sel[0].strftime('%b %Y')} → {month_sel[1].strftime('%b %Y')}</span>"
        f"<span class='pill'>GLs: {combo_f['GL Code'].nunique()}</span>"
        f"<span class='pill'>Cost Centers: {combo_f['Cost Centre'].nunique()}</span>",
        unsafe_allow_html=True,
    )

    st.divider()

    # Monthly trend
    m = combo_f.groupby("MonthStart", as_index=False)[["Act. Costs", "BudgetCost", "Variance"]].sum().sort_values("MonthStart")
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=m["MonthStart"], y=m["BudgetCost"], mode="lines+markers", name="Budget"))
    fig.add_trace(go.Scatter(x=m["MonthStart"], y=m["Act. Costs"], mode="lines+markers", name="Actual"))
    fig.update_layout(
        title="Monthly Spend: Actual vs Budget",
        xaxis_title="Month",
        yaxis_title="Amount (₹)",
        height=420,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=10, r=10, t=45, b=10),
    )
    st.plotly_chart(fig, use_container_width=True)

    # Top contributors
    left, right = st.columns([1.2, 0.8])
    with left:
        dim = breakdown_dim
        if dim not in combo_f.columns:
            st.info("Selected breakdown dimension is not available for current data.")
        else:
            top = combo_f.groupby(dim, as_index=False)[["Act. Costs", "BudgetCost", "Variance"]].sum()
            top = top.sort_values("Act. Costs", ascending=False).head(15)
            fig2 = px.bar(
                top,
                x="Act. Costs",
                y=dim,
                orientation="h",
                title=f"Top 15 by Actual ({dim})",
                color="Variance",
                color_continuous_scale="RdBu",
                height=520,
            )
            fig2.update_layout(margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig2, use_container_width=True)

    with right:
        # Waterfall variance by dimension
        dim = breakdown_dim
        if dim in combo_f.columns:
            wf = combo_f.groupby(dim, as_index=False)["Variance"].sum().sort_values("Variance", ascending=False)
            wf = pd.concat([wf.head(8), wf.tail(8)]).drop_duplicates(subset=[dim])
            wf = wf.sort_values("Variance", ascending=False)
            fig3 = go.Figure(
                go.Waterfall(
                    name="Variance",
                    orientation="v",
                    x=wf[dim].astype(str),
                    y=wf["Variance"],
                    connector={"line": {"color": "rgba(180,180,180,0.35)"}},
                )
            )
            fig3.update_layout(title=f"Variance Waterfall (Top/Bottom {dim})", height=520, margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("Waterfall needs a valid breakdown dimension.")

# ----------------------------
# Variance Explorer
# ----------------------------
elif page == "Variance Explorer":
    st.title("Variance Explorer")
    st.caption("Find overspends/underspends quickly and drill to GL/Cost Centre.")

    dim = breakdown_dim
    if dim not in combo_f.columns:
        st.warning("Selected breakdown dimension isn't available.")
        st.stop()

    # Variance table
    g = combo_f.groupby(dim, as_index=False)[["Act. Costs", "BudgetCost", "Variance"]].sum()
    g["VarPct"] = np.where(g["BudgetCost"].abs() > 1e-9, g["Variance"] / g["BudgetCost"], np.nan)

    colA, colB = st.columns([0.62, 0.38])

    with colA:
        st.subheader("Variance ranking")
        show_n = st.slider("Top N", 5, 50, 20)
        order = st.selectbox("Sort", ["Overspend (highest variance)", "Underspend (lowest variance)", "Absolute variance"], index=0)

        if order.startswith("Over"):
            view = g.sort_values("Variance", ascending=False).head(show_n)
        elif order.startswith("Under"):
            view = g.sort_values("Variance", ascending=True).head(show_n)
        else:
            view = g.assign(AbsVar=g["Variance"].abs()).sort_values("AbsVar", ascending=False).head(show_n)

        view_disp = view.copy()
        for c in ["Act. Costs", "BudgetCost", "Variance"]:
            view_disp[c] = view_disp[c].map(inr)
        view_disp["VarPct"] = view["VarPct"].map(pct)

        st.dataframe(view_disp, use_container_width=True, height=420)

        # Chart
        fig = px.bar(
            view,
            x="Variance",
            y=dim,
            orientation="h",
            color="VarPct",
            color_continuous_scale="RdBu",
            title=f"{order} — {dim}",
            height=520,
        )
        fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with colB:
        st.subheader("Variance heatmap")
        if "MonthStart" in combo_f.columns:
            pivot = combo_f.pivot_table(index=dim, columns="MonthStart", values="Variance", aggfunc="sum", fill_value=0)
            # keep only top rows by abs variance
            top_rows = pivot.abs().sum(axis=1).sort_values(ascending=False).head(25).index
            pivot = pivot.loc[top_rows]
            fig_h = px.imshow(
                pivot,
                aspect="auto",
                color_continuous_scale="RdBu",
                origin="lower",
                title=f"Top 25 {dim} by |variance| (month-wise)",
                height=600,
            )
            fig_h.update_layout(margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig_h, use_container_width=True)
        else:
            st.info("Heatmap unavailable.")

# ----------------------------
# Trend & Mix
# ----------------------------
elif page == "Trend & Mix":
    st.title("Trend & Mix")

    # Stacked mix by selected dimension over time
    dim = breakdown_dim
    if dim not in combo_f.columns:
        st.warning("Selected breakdown dimension isn't available.")
        st.stop()

    t = combo_f.groupby(["MonthStart", dim], as_index=False)["Act. Costs"].sum()
    # keep top N categories overall, bucket rest
    topcats = (
        t.groupby(dim, as_index=False)["Act. Costs"].sum().sort_values("Act. Costs", ascending=False).head(12)[dim].tolist()
    )
    t[dim] = np.where(t[dim].isin(topcats), t[dim], "Others")
    t = t.groupby(["MonthStart", dim], as_index=False)["Act. Costs"].sum()

    fig = px.area(
        t,
        x="MonthStart",
        y="Act. Costs",
        color=dim,
        title=f"Actual mix over time (Top 12 {dim} + Others)",
        height=460,
    )
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10), legend=dict(orientation="h", y=1.02, x=1, xanchor="right"))
    st.plotly_chart(fig, use_container_width=True)

    # Treemap of Actual
    tree_dim2 = st.selectbox("Treemap hierarchy", ["Department", "CC Maping", "GL Code"], index=0)
    if tree_dim2 in combo_f.columns:
        treedf = combo_f.groupby([tree_dim2], as_index=False)["Act. Costs"].sum()
        treedf = treedf[treedf[tree_dim2].notna()]
        fig2 = px.treemap(treedf, path=[tree_dim2], values="Act. Costs", title=f"Actual Spend Treemap ({tree_dim2})", height=520)
        fig2.update_layout(margin=dict(l=10, r=10, t=45, b=10))
        st.plotly_chart(fig2, use_container_width=True)

    # Simple anomaly detection
    st.subheader("Anomaly signals (z-score)")
    zdim = st.selectbox("Anomaly dimension", ["CC Maping", "Department", "GL Code"], index=0)
    if zdim in combo_f.columns:
        z = combo_f.groupby(["MonthStart", zdim], as_index=False)["Act. Costs"].sum()
        # z-score within each category
        z["z"] = z.groupby(zdim)["Act. Costs"].transform(lambda s: (s - s.mean()) / (s.std(ddof=0) + 1e-9))
        anomalies = z[z["z"].abs() >= 3].sort_values("z", key=lambda s: s.abs(), ascending=False).head(40)
        if anomalies.empty:
            st.info("No anomalies found at |z| ≥ 3 for the selected filters.")
        else:
            disp = anomalies.copy()
            disp["Act. Costs"] = disp["Act. Costs"].map(inr)
            st.dataframe(disp, use_container_width=True, height=280)

            fig3 = px.scatter(
                z,
                x="MonthStart",
                y="Act. Costs",
                color=zdim,
                size=z["Act. Costs"].clip(lower=1),
                hover_data=["z"],
                title=f"Monthly spend scatter ({zdim}) — anomalies highlighted in table",
                height=520,
            )
            fig3.update_layout(margin=dict(l=10, r=10, t=45, b=10))
            st.plotly_chart(fig3, use_container_width=True)

# ----------------------------
# Transactions
# ----------------------------
elif page == "Transactions":
    st.title("Transactions")
    st.caption("Raw Actual transactions (filtered). Useful for audit and drill-through.")

    # apply same filters on actual_enriched
    act = actual_enriched.copy()
    act = act[(act["MonthStart"] >= month_sel[0]) & (act["MonthStart"] <= month_sel[1])]
    act = act[act["CompanyName"].isin(company_sel)]
    if dept_sel and "Department" in act.columns:
        act = act[act["Department"].isin(dept_sel)]
    if ccmap_sel and "CC Maping" in act.columns:
        act = act[act["CC Maping"].isin(ccmap_sel)]
    if nature_sel and "Nature" in act.columns:
        act = act[act["Nature"].isin(nature_sel)]

    # quick search
    q = st.text_input("Search (GL, Cost Centre, Short text)")
    if q:
        ql = q.lower().strip()
        mask = (
            act["GL"].astype(str).str.contains(ql, case=False, na=False)
            | act["Cost Centers"].astype(str).str.contains(ql, case=False, na=False)
            | act["Short text"].astype(str).str.contains(ql, case=False, na=False)
        )
        act = act[mask]

    show_cols = [
        "Month", "MonthStart", "CompanyName", "Plant", "Cost Centers", "Department", "Plant/HO", "GL", "Short text",
        "Nature", "Group", "MIS MAPPING", "CC Maping", "Act. Costs",
    ]
    show_cols = [c for c in show_cols if c in act.columns]

    act_disp = act[show_cols].copy()
    act_disp["Act. Costs"] = act_disp["Act. Costs"].map(inr)

    st.dataframe(act_disp, use_container_width=True, height=540)

    # download
    csv = act[show_cols].to_csv(index=False).encode("utf-8")
    st.download_button("Download filtered transactions (CSV)", data=csv, file_name="transactions_filtered.csv", mime="text/csv")

# ----------------------------
# Data Quality
# ----------------------------
elif page == "Data Quality":
    st.title("Data Quality & Coverage")

    # Mapping coverage
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

    # Unmatched keys between actual and budget
    act_keys = act[["CompanyName", "Cost Centre", "GL Code"]].drop_duplicates()
    bud_keys = budget_map[["CompanyName", "Cost Centre", "GL Code"]].drop_duplicates()

    unmatched_act = act_keys.merge(bud_keys, on=["CompanyName", "Cost Centre", "GL Code"], how="left", indicator=True)
    unmatched_act = unmatched_act[unmatched_act["_merge"] == "left_only"].drop(columns=["_merge"])

    st.subheader("Actual lines without matching Budget mapping")
    st.caption("These may indicate new GLs/Cost Centres or missing mapping in Budget sheet.")

    st.dataframe(unmatched_act.head(200), use_container_width=True, height=420)

    st.download_button(
        "Download unmatched actual keys (CSV)",
        data=unmatched_act.to_csv(index=False).encode("utf-8"),
        file_name="unmatched_actual_keys.csv",
        mime="text/csv",
    )

# Footer
st.markdown("<div class='small-note'>Built with Streamlit • Plotly • Pandas. Tip: Use the sidebar filters to slice by company, department, and category.</div>", unsafe_allow_html=True)
