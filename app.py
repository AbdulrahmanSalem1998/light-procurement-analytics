"""Light – Procurement Analytics.

Streamlit front-end. Business logic lives in the ``light`` package:
``light.pipeline`` (cleaning), ``light.reports`` (Excel exports),
``light.config`` (entity/sector, FX and category settings).

Run:  streamlit run app.py
"""
import pandas as pd
import plotly.express as px
import streamlit as st

from light import pipeline
from light.pipeline import (
    add_kpi_metrics,
    add_recall_status,
    clean_dataframe,
    clean_kpi_df,
    clean_recall_df,
)
from light.reports import to_excel_bytes, to_open_po_excel, to_summary_report_excel
from light.theme import BRAND_SCALE, C_MID, C_PRIMARY, CSS, PIE_COLORS, style_fig

st.set_page_config(
    page_title="Light – Procurement Analytics",
    layout="wide",
    page_icon="⚡",
    initial_sidebar_state="expanded",
)

st.markdown(CSS, unsafe_allow_html=True)

st.markdown("""
<div class="app-header">
  <div class="avatar">⚡</div>
  <div class="header-text">
    <h1>Light</h1>
    <p>Procurement Analytics &nbsp;·&nbsp; Data Cleaning &amp; Business Intelligence</p>
  </div>
</div>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load_data(file_bytes, file_name, header_marker="entity"):
    """Cached wrapper around pipeline.load_data (auto-detects the header row)."""
    return pipeline.load_data(file_bytes, file_name, header_marker=header_marker)


tab_clean, tab_dash, tab_open, tab_summary = st.tabs([
    "  🧹  Data Cleaning  ",
    "  📊  Analytics Dashboard  ",
    "  📋  Open PO Tracker  ",
    "  📈  PO Summary Report  ",
])


# ══════════════════════════════════════════════════════════════
# TAB 1 – DATA CLEANING
# ══════════════════════════════════════════════════════════════
with tab_clean:

    st.markdown('<div class="section-title">Automated Cleaning Pipeline</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="step-grid">
      <div class="step-item">
        <span class="step-icon">🔍</span>
        <div><span class="step-title">Step 1 — Auto-detect Header</span>
        <span class="step-desc">Skips system report metadata rows and finds the true header automatically.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🔠</span>
        <div><span class="step-title">Step 2 — Capitalize Entity</span>
        <span class="step-desc">Standardizes entity codes to uppercase (e.g. fab → FAB).</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🏢</span>
        <div><span class="step-title">Step 3 — Add Sector</span>
        <span class="step-desc">Maps each entity to its business sector (E&C, Industrial, Real Estate, etc.).</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🗑️</span>
        <div><span class="step-title">Step 4 — Remove Canceled POs</span>
        <span class="step-desc">Drops all rows where PO Status = Canceled from the dataset.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">💰</span>
        <div><span class="step-title">Step 5 — Total Amount in SAR</span>
        <span class="step-desc">Converts all currencies to SAR using standard exchange rates.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">📅</span>
        <div><span class="step-title">Step 6 — PO Cycle Time (Days)</span>
        <span class="step-desc">Converts cycle time to whole days. Values under 1 day are set to 1.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">📂</span>
        <div><span class="step-title">Step 7 — Main Category</span>
        <span class="step-desc">Adds Main Category from the Standard Vendor Categorization file.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🏷️</span>
        <div><span class="step-title">Step 8 — Segment</span>
        <span class="step-desc">Adds Segment (category group) from the Standard Vendor Categorization file.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🔖</span>
        <div><span class="step-title">Step 9 — Subcategory</span>
        <span class="step-desc">Adds Subcategory from the Standard Vendor Categorization file.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🗓️</span>
        <div><span class="step-title">Step 10 — Date Parts (Day / Month / Year)</span>
        <span class="step-desc">Extracts Day, Month, and Year from Create Date as separate numeric columns.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🏷️</span>
        <div><span class="step-title">Step 11 — PO Category</span>
        <span class="step-desc">Flags each PO as Standard PO (has a linked PR) or Direct PO (no PR).</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🔁</span>
        <div><span class="step-title">Step 12 — Recall Status</span>
        <span class="step-desc">Cross-checks each PO against the Recall log and adds Recall Status, Action Date, and Days to Action Recall.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">⏱️</span>
        <div><span class="step-title">Step 13 — PR-to-PO Cycle Time</span>
        <span class="step-desc">Matches each PO to the PR-PO KPI report (by PR, then PO) and adds 3 cycle-time metrics in working days (Sun–Thu).</span></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    cu1, cu2, cu3 = st.columns(3)
    with cu1:
        upload_clean = st.file_uploader(
            "Upload your PO report (Excel or CSV)",
            type=["xlsx", "xls", "csv"], key="clean_upload"
        )
    with cu2:
        upload_recall = st.file_uploader(
            "Upload Recall PO / Change History (optional)",
            type=["xlsx", "xls", "csv"], key="recall_upload"
        )
    with cu3:
        upload_kpi = st.file_uploader(
            "Upload PR-to-PO KPI report (optional)",
            type=["xlsx", "xls", "csv"], key="kpi_upload"
        )

    if upload_clean:
        with st.spinner("Cleaning data…"):
            raw_df     = load_data(upload_clean.read(), upload_clean.name)
            cleaned_df = clean_dataframe(raw_df)

            recall_df = None
            if upload_recall:
                raw_recall_df = load_data(upload_recall.read(), upload_recall.name, header_marker="sr#")
                recall_df     = clean_recall_df(raw_recall_df)
                cleaned_df    = add_recall_status(cleaned_df, recall_df)

            kpi_df = None
            if upload_kpi:
                raw_kpi_df = load_data(upload_kpi.read(), upload_kpi.name, header_marker="pr number")
                kpi_df     = clean_kpi_df(raw_kpi_df)
                cleaned_df = add_kpi_metrics(cleaned_df, kpi_df)

        removed = len(raw_df) - len(cleaned_df)
        new_cols = (9 + ("PO Category" in cleaned_df.columns)
                      + 3 * ("Recall Status" in cleaned_df.columns)
                      + 3 * (kpi_df is not None))

        st.markdown("---")
        st.markdown('<div class="section-title">Cleaning Summary</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Original Rows",         f"{len(raw_df):,}")
        c2.metric("Rows After Cleaning",    f"{len(cleaned_df):,}")
        c3.metric("Canceled Rows Removed",  f"{removed:,}")
        c4.metric("New Columns Added",      f"{new_cols}")

        if recall_df is not None and "Recall Status" in cleaned_df.columns:
            st.markdown("---")
            st.markdown('<div class="section-title">Recall Status Breakdown</div>', unsafe_allow_html=True)
            rc1, rc2, rc3 = st.columns(3)
            status_counts = cleaned_df["Recall Status"].value_counts()
            rc1.metric("Request change", f"{status_counts.get('Request change', 0):,}")
            rc2.metric("Cancel",         f"{status_counts.get('Cancel', 0):,}")
            rc3.metric("Not recalled",   f"{status_counts.get('Not recalled', 0):,}")

        st.markdown("---")
        with st.expander("Preview Cleaned Data (first 100 rows)", expanded=True):
            date_col_config = {
                c: st.column_config.DateColumn(c, format="M/D/YYYY")
                for c in cleaned_df.columns if pd.api.types.is_datetime64_any_dtype(cleaned_df[c])
            }
            st.dataframe(cleaned_df.head(100), use_container_width=True, column_config=date_col_config)

        excel_bytes = to_excel_bytes(cleaned_df)
        st.download_button(
            label="⬇️  Download Cleaned File (.xlsx)",
            data=excel_bytes,
            file_name=f"PO Report_{pd.Timestamp.today().strftime('%Y-%m-%d')}_cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ══════════════════════════════════════════════════════════════
# TAB 2 – ANALYTICS DASHBOARD
# ══════════════════════════════════════════════════════════════
with tab_dash:

    uploaded_file = st.file_uploader(
        "Upload your cleaned PO report (Excel or CSV)",
        type=["xlsx", "xls", "csv"], key="analytics_upload"
    )

    if uploaded_file:
        with st.spinner("Loading data…"):
            df = load_data(uploaded_file.read(), uploaded_file.name)
            df.columns = df.columns.str.strip()
            df = df.dropna(axis=1, how="all")
            df = df.loc[:, ~df.columns.str.startswith("Column")]
            if "Create Date" in df.columns:
                df["Create Date"] = pd.to_datetime(df["Create Date"], errors="coerce")

        st.success(f"**{len(df):,}** rows · **{len(df.columns)}** columns loaded")

        amt_col = "Total Amount (SAR)" if "Total Amount (SAR)" in df.columns else "Total amount"

        # ── Sidebar filters ──────────────────────────────────────
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🎛️ Filters")

        if "Entity" in df.columns:
            opts = ["All"] + sorted(df["Entity"].dropna().unique().tolist())
            sel = st.sidebar.selectbox("Entity", opts)
            if sel != "All": df = df[df["Entity"] == sel]

        if "Sector" in df.columns:
            opts = ["All"] + sorted(df["Sector"].dropna().unique().tolist())
            sel = st.sidebar.selectbox("Sector", opts)
            if sel != "All": df = df[df["Sector"] == sel]

        if "PO status" in df.columns:
            opts = ["All"] + sorted(df["PO status"].dropna().unique().tolist())
            sel = st.sidebar.selectbox("PO Status", opts)
            if sel != "All": df = df[df["PO status"] == sel]

        if "Main Category" in df.columns:
            opts = ["All"] + sorted(df["Main Category"].dropna().unique().tolist())
            sel = st.sidebar.selectbox("Main Category", opts)
            if sel != "All": df = df[df["Main Category"] == sel]

        if "Segment" in df.columns:
            opts = ["All"] + sorted(df["Segment"].dropna().unique().tolist())
            sel = st.sidebar.selectbox("Segment", opts)
            if sel != "All": df = df[df["Segment"] == sel]

        if "Year" in df.columns:
            yr_vals = sorted(df["Year"].dropna().unique().tolist())
            opts = ["All"] + [str(int(y)) for y in yr_vals]
            sel = st.sidebar.selectbox("Year", opts)
            if sel != "All": df = df[df["Year"] == int(sel)]

        MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        if "Month" in df.columns:
            opts = ["All"] + [m for m in MONTH_ORDER if m in df["Month"].dropna().unique().tolist()]
            sel = st.sidebar.selectbox("Month", opts)
            if sel != "All": df = df[df["Month"] == sel]

        if "Day" in df.columns:
            opts = ["All"] + sorted([str(int(d)) for d in df["Day"].dropna().unique().tolist()], key=lambda x: int(x))
            sel = st.sidebar.selectbox("Day", opts)
            if sel != "All": df = df[df["Day"] == int(sel)]

        if "Create Date" in df.columns:
            mn, mx = df["Create Date"].min(), df["Create Date"].max()
            if pd.notna(mn) and pd.notna(mx):
                dr = st.sidebar.date_input("Date Range", (mn.date(), mx.date()),
                                           min_value=mn.date(), max_value=mx.date())
                if len(dr) == 2:
                    df = df[(df["Create Date"].dt.date >= dr[0]) &
                            (df["Create Date"].dt.date <= dr[1])]

        # ── KPIs ─────────────────────────────────────────────────
        st.markdown('<div class="section-title">📌 Key Metrics</div>', unsafe_allow_html=True)
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Total POs", f"{len(df):,}")
        if amt_col in df.columns:
            total_spend = df[amt_col].sum()
            k2.metric("Total Spend (SAR)", f"{total_spend/1_000_000:.2f}M" if total_spend >= 1_000_000 else f"{total_spend:,.0f}")
        if "Vendor name" in df.columns:
            k3.metric("Unique Vendors", f"{df['Vendor name'].nunique():,}")
        if "PO Cycle Time" in df.columns:
            ct = pd.to_numeric(df["PO Cycle Time"], errors="coerce")
            if ct.notna().any():
                k4.metric("Avg. Cycle Time", f"{ct.mean():.1f} days")
        if "Bid exemption" in df.columns:
            exempted = df["Bid exemption"].sum() if df["Bid exemption"].dtype != object else (df["Bid exemption"] == 1).sum()
            rate = exempted / len(df) * 100 if len(df) > 0 else 0
            k5.metric("Bid Exemption Rate", f"{rate:.1f}%")

        st.markdown("---")

        # ── Row 1: PO Status | PO Type Count ─────────────────────
        st.markdown('<div class="section-title">📊 Distribution & Volume</div>', unsafe_allow_html=True)
        r1a, r1b = st.columns(2)

        with r1a:
            if "PO status" in df.columns:
                sc = df["PO status"].value_counts().reset_index()
                sc.columns = ["Status", "Count"]
                fig = px.pie(sc, names="Status", values="Count",
                             title="PO Status Distribution", hole=0.5,
                             color_discrete_sequence=PIE_COLORS)
                fig.update_traces(textposition="outside", textinfo="percent+label",
                                  pull=[0.03]*len(sc))
                st.plotly_chart(style_fig(fig), use_container_width=True)

        with r1b:
            if "PO type" in df.columns:
                pt = df["PO type"].value_counts().reset_index()
                pt.columns = ["PO Type", "Count"]
                fig = px.bar(pt, x="PO Type", y="Count",
                             title="PO Count by Type",
                             color="PO Type",
                             color_discrete_sequence=PIE_COLORS,
                             text="Count")
                fig.update_traces(texttemplate="%{text:,}", textposition="outside")
                fig.update_layout(showlegend=False, xaxis_tickangle=-20)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        # ── Row 2: Monthly Spend Trend | Spend by Year ───────────
        st.markdown("---")
        st.markdown('<div class="section-title">📈 Spend Trends</div>', unsafe_allow_html=True)
        r2a, r2b = st.columns(2)

        with r2a:
            if "Create Date" in df.columns and amt_col in df.columns:
                df["_YM"] = df["Create Date"].dt.to_period("M").astype(str)
                monthly = df.groupby("_YM")[amt_col].sum().reset_index().sort_values("_YM")
                monthly.columns = ["Month", "Amount (SAR)"]
                fig = px.area(monthly, x="Month", y="Amount (SAR)",
                              title="Monthly Spend Trend (SAR)",
                              color_discrete_sequence=[C_PRIMARY])
                fig.update_traces(line_width=2.5, fill="tozeroy",
                                  fillcolor="rgba(124,58,237,0.1)")
                fig.update_xaxes(tickangle=45)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        with r2b:
            if "Year" in df.columns and amt_col in df.columns:
                yr = df.groupby("Year")[amt_col].sum().reset_index()
                yr.columns = ["Year", "Amount (SAR)"]
                yr["Year"] = yr["Year"].astype(str)
                fig = px.bar(yr, x="Year", y="Amount (SAR)",
                             title="Spend by Year",
                             color="Amount (SAR)",
                             color_continuous_scale=BRAND_SCALE,
                             text="Amount (SAR)")
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                fig.update_layout(coloraxis_showscale=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        # ── Row 3: Spend by Sector | Top 10 Vendors ─────────────
        st.markdown("---")
        st.markdown('<div class="section-title">🏢 Sector & Vendor Analysis</div>', unsafe_allow_html=True)
        r3a, r3b = st.columns(2)

        with r3a:
            if "Sector" in df.columns and amt_col in df.columns:
                sec = df.groupby("Sector")[amt_col].sum().sort_values(ascending=False).reset_index()
                sec.columns = ["Sector", "Amount (SAR)"]
                fig = px.bar(sec, x="Sector", y="Amount (SAR)",
                             title="Spend by Sector",
                             color="Sector",
                             color_discrete_sequence=PIE_COLORS,
                             text="Amount (SAR)")
                fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
                fig.update_layout(showlegend=False, xaxis_tickangle=-20)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        with r3b:
            if amt_col in df.columns and "Vendor name" in df.columns:
                tv = (df.groupby("Vendor name")[amt_col].sum()
                      .sort_values(ascending=False).head(10).reset_index())
                tv.columns = ["Vendor", "Amount (SAR)"]
                fig = px.bar(tv, x="Amount (SAR)", y="Vendor", orientation="h",
                             title="Top 10 Vendors by Spend",
                             color="Amount (SAR)", color_continuous_scale=BRAND_SCALE)
                fig.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        # ── Row 4: Segment Spend | Main Category Spend ───────────
        st.markdown("---")
        st.markdown('<div class="section-title">📂 Category Intelligence</div>', unsafe_allow_html=True)
        r4a, r4b = st.columns(2)

        with r4a:
            if "Segment" in df.columns and amt_col in df.columns:
                seg = df.groupby("Segment")[amt_col].sum().sort_values(ascending=False).reset_index()
                seg.columns = ["Segment", "Amount (SAR)"]
                fig = px.bar(seg, x="Amount (SAR)", y="Segment", orientation="h",
                             title="Spend by Segment",
                             color="Amount (SAR)", color_continuous_scale=BRAND_SCALE)
                fig.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        with r4b:
            if "Main Category" in df.columns and amt_col in df.columns:
                mc = df.groupby("Main Category")[amt_col].sum().sort_values(ascending=False).reset_index()
                mc.columns = ["Category", "Amount (SAR)"]
                fig = px.pie(mc, names="Category", values="Amount (SAR)",
                             title="Spend Share by Main Category", hole=0.45,
                             color_discrete_sequence=PIE_COLORS)
                fig.update_traces(textposition="outside", textinfo="percent+label")
                st.plotly_chart(style_fig(fig), use_container_width=True)

        # ── Row 5: Segment PO Count | Subcategory Spend ──────────
        r5a, r5b = st.columns(2)

        with r5a:
            if "Segment" in df.columns:
                seg_cnt = df["Segment"].value_counts().reset_index()
                seg_cnt.columns = ["Segment", "PO Count"]
                fig = px.bar(seg_cnt, x="PO Count", y="Segment", orientation="h",
                             title="PO Count by Segment",
                             color="PO Count", color_continuous_scale=BRAND_SCALE,
                             text="PO Count")
                fig.update_traces(texttemplate="%{text:,}", textposition="outside")
                fig.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        with r5b:
            if "Subcategory" in df.columns and amt_col in df.columns:
                sub = df.groupby("Subcategory")[amt_col].sum().sort_values(ascending=False).head(10).reset_index()
                sub.columns = ["Subcategory", "Amount (SAR)"]
                fig = px.bar(sub, x="Amount (SAR)", y="Subcategory", orientation="h",
                             title="Top 10 Subcategories by Spend",
                             color="Amount (SAR)", color_continuous_scale=BRAND_SCALE)
                fig.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        # ── Row 6: Bid Exemption | Approval Status ───────────────
        st.markdown("---")
        st.markdown('<div class="section-title">✅ Compliance & Approval</div>', unsafe_allow_html=True)
        r6a, r6b = st.columns(2)

        with r6a:
            if "Bid exemption" in df.columns:
                bid = df["Bid exemption"].value_counts().reset_index()
                bid.columns = ["Bid Exemption", "Count"]
                bid["Bid Exemption"] = bid["Bid Exemption"].map({0: "No Exemption", 1: "Exempted"})
                fig = px.pie(bid, names="Bid Exemption", values="Count",
                             title="Bid Exemption Breakdown", hole=0.5,
                             color_discrete_sequence=[C_PRIMARY, C_MID])
                fig.update_traces(textposition="outside", textinfo="percent+label", pull=[0.03, 0])
                st.plotly_chart(style_fig(fig), use_container_width=True)

        with r6b:
            if "Approval status" in df.columns:
                ap = df["Approval status"].value_counts().reset_index()
                ap.columns = ["Approval Status", "Count"]
                fig = px.bar(ap, x="Approval Status", y="Count",
                             title="Approval Status Breakdown",
                             color="Approval Status",
                             color_discrete_sequence=PIE_COLORS,
                             text="Count")
                fig.update_traces(texttemplate="%{text:,}", textposition="outside")
                fig.update_layout(showlegend=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        # ── Row 7: Cycle Time by Category | Monthly PO Volume ────
        st.markdown("---")
        st.markdown('<div class="section-title">⏱️ Cycle Time & Activity</div>', unsafe_allow_html=True)
        r7a, r7b = st.columns(2)

        with r7a:
            if "PO Cycle Time" in df.columns and "Main Category" in df.columns:
                ct_df = df[["Main Category", "PO Cycle Time"]].copy()
                ct_df["PO Cycle Time"] = pd.to_numeric(ct_df["PO Cycle Time"], errors="coerce")
                ct_df = ct_df.dropna()
                avg_ct = ct_df.groupby("Main Category")["PO Cycle Time"].mean().sort_values(ascending=False).reset_index()
                avg_ct.columns = ["Category", "Avg Cycle Time (Days)"]
                fig = px.bar(avg_ct, x="Avg Cycle Time (Days)", y="Category", orientation="h",
                             title="Avg PO Cycle Time by Category",
                             color="Avg Cycle Time (Days)", color_continuous_scale=BRAND_SCALE)
                fig.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        with r7b:
            if "Create Date" in df.columns:
                df["_YM2"] = df["Create Date"].dt.to_period("M").astype(str)
                vol = df.groupby("_YM2").size().reset_index()
                vol.columns = ["Month", "PO Count"]
                fig = px.bar(vol, x="Month", y="PO Count",
                             title="Monthly PO Volume",
                             color="PO Count", color_continuous_scale=BRAND_SCALE)
                fig.update_layout(coloraxis_showscale=False, xaxis_tickangle=45)
                st.plotly_chart(style_fig(fig), use_container_width=True)

        # ── Filtered Data + Download ──────────────────────────────
        st.markdown("---")
        st.markdown('<div class="section-title">📋 Filtered Data</div>', unsafe_allow_html=True)

        col_info, col_dl = st.columns([3, 1])
        with col_info:
            st.markdown(f"Showing **{len(df):,}** rows · **{len(df.columns)}** columns")
        with col_dl:
            dl_bytes = to_excel_bytes(df)
            st.download_button(
                label="⬇️  Download Filtered Data",
                data=dl_bytes,
                file_name=f"PO Report_{pd.Timestamp.today().strftime('%Y-%m-%d')}_filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with st.expander("View Data Table", expanded=False):
            st.dataframe(df, use_container_width=True, height=400)

    else:
        st.markdown("""
        <div class="section-card" style="text-align:center; padding:56px 24px;">
          <div style="font-size:3.5rem; margin-bottom:16px;">📊</div>
          <h3 style="color:#4c1d95; margin:0 0 8px 0; font-weight:700;">Upload a PO Report to Get Started</h3>
          <p style="color:#7c3aed; margin:0 0 6px 0; font-size:0.95rem;">Dashboard renders instantly on upload.</p>
          <p style="color:#a78bfa; margin:0; font-size:0.85rem;">Supports Excel (.xlsx, .xls) and CSV · Headers auto-detected</p>
        </div>
        """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# TAB 3 – OPEN PO TRACKER
# ══════════════════════════════════════════════════════════════
with tab_open:

    st.markdown('<div class="section-title">📋 Open PO Tracker — Export by Entity</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="step-grid">
      <div class="step-item">
        <span class="step-icon">🔓</span>
        <div><span class="step-title">Open POs Only</span>
        <span class="step-desc">Filters rows where PO Status = "Open order" automatically.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🗂️</span>
        <div><span class="step-title">One Sheet per Entity</span>
        <span class="step-desc">Each entity (BLD, FAB, DEV…) gets its own tab in the Excel file.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">📝</span>
        <div><span class="step-title">Open PO Reason</span>
        <span class="step-desc">Dropdown list with 8 standard reasons for users to fill in.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">📆</span>
        <div><span class="step-title">Expected Closure Date</span>
        <span class="step-desc">Date validation column for users to enter the expected closure date.</span></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    upload_open = st.file_uploader(
        "Upload your cleaned PO report (Excel or CSV)",
        type=["xlsx", "xls", "csv"], key="open_upload"
    )

    if upload_open:
        with st.spinner("Processing Open POs…"):
            open_raw = load_data(upload_open.read(), upload_open.name)
            open_raw.columns = open_raw.columns.str.strip()
            open_raw = open_raw.dropna(axis=1, how="all")

            if "Sector" not in open_raw.columns:
                open_raw = clean_dataframe(open_raw)

            open_pos = open_raw[open_raw["PO status"].str.strip() == "Open order"].copy() if "PO status" in open_raw.columns else open_raw.copy()

        if open_pos.empty:
            st.warning("No Open POs found in the uploaded file.")
        else:
            amt_col = "Total Amount (SAR)" if "Total Amount (SAR)" in open_pos.columns else "Total amount"

            st.markdown('<div class="section-title">🎛️ Filters</div>', unsafe_allow_html=True)
            fc1, fc2, fc3 = st.columns(3)

            with fc1:
                if "Entity" in open_pos.columns:
                    all_entities = sorted(open_pos["Entity"].dropna().unique().tolist())
                    sel_entities = st.multiselect("Entity", all_entities, default=all_entities,
                                                  placeholder="Select entities…")
                else:
                    sel_entities = []

            with fc2:
                if "Year" in open_pos.columns:
                    all_years = sorted(open_pos["Year"].dropna().unique().tolist())
                    year_opts = [str(int(y)) for y in all_years]
                    sel_years = st.multiselect("Year", year_opts, default=year_opts,
                                               placeholder="Select years…")
                else:
                    sel_years = []

            MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
            with fc3:
                if "Month" in open_pos.columns:
                    all_months = [m for m in MONTH_ORDER if m in open_pos["Month"].dropna().unique().tolist()]
                    sel_months = st.multiselect("Month", all_months, default=all_months,
                                                placeholder="Select months…")
                else:
                    sel_months = []

            filtered_pos = open_pos.copy()
            if sel_entities:
                filtered_pos = filtered_pos[filtered_pos["Entity"].isin(sel_entities)]
            if sel_years and "Year" in filtered_pos.columns:
                filtered_pos = filtered_pos[filtered_pos["Year"].astype(str).isin(sel_years)]
            if sel_months and "Month" in filtered_pos.columns:
                filtered_pos = filtered_pos[filtered_pos["Month"].isin(sel_months)]

            if filtered_pos.empty:
                st.warning("No Open POs match the selected filters.")
            else:
                st.markdown("---")
                st.markdown('<div class="section-title">📌 Open PO Summary</div>', unsafe_allow_html=True)

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total Open POs", f"{len(filtered_pos):,}")
                if amt_col in filtered_pos.columns:
                    total = filtered_pos[amt_col].sum()
                    m2.metric("Open Spend (SAR)", f"{total/1_000_000:.2f}M" if total >= 1_000_000 else f"{total:,.0f}")
                if "Entity" in filtered_pos.columns:
                    m3.metric("Entities", f"{filtered_pos['Entity'].nunique():,}")
                if "Vendor name" in filtered_pos.columns:
                    m4.metric("Vendors", f"{filtered_pos['Vendor name'].nunique():,}")

                st.markdown("---")
                st.markdown('<div class="section-title">🗂️ Column Selection</div>', unsafe_allow_html=True)
                st.caption("Leave empty to export all columns. Select specific columns to export only those.")

                tracker_cols = ["Open PO Reason", "Expected Closure Date", "Comment"]
                selectable_cols = [c for c in filtered_pos.columns if c not in tracker_cols]

                sel_cols = st.multiselect(
                    "Select columns to include (leave empty = all columns)",
                    options=selectable_cols,
                    default=[],
                    placeholder="Select columns…"
                )

                if sel_cols:
                    final_cols = sel_cols + [c for c in tracker_cols if c in filtered_pos.columns]
                    export_df = filtered_pos[final_cols]
                else:
                    export_df = filtered_pos

                st.markdown("---")
                open_bytes = to_open_po_excel(export_df)
                st.download_button(
                    label=f"⬇️  Download Open PO Tracker ({len(export_df):,} POs · {export_df['Entity'].nunique() if 'Entity' in export_df.columns else 0} Entities · {len(export_df.columns)} Columns)",
                    data=open_bytes,
                    file_name=f"Open PO Report_{pd.Timestamp.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    else:
        st.markdown("""
        <div class="section-card" style="text-align:center; padding:56px 24px;">
          <div style="font-size:3.5rem; margin-bottom:16px;">📋</div>
          <h3 style="color:#4c1d95; margin:0 0 8px 0; font-weight:700;">Upload a Cleaned PO Report</h3>
          <p style="color:#7c3aed; margin:0 0 6px 0; font-size:0.95rem;">Upload your cleaned file to generate the Open PO Tracker.</p>
          <p style="color:#a78bfa; margin:0; font-size:0.85rem;">Open POs will be split by entity into separate sheets automatically.</p>
        </div>
        """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# TAB 4 – PO SUMMARY REPORT
# ══════════════════════════════════════════════════════════════
with tab_summary:

    st.markdown('<div class="section-title">📈 PO Summary Report — Entity × Year Breakdown</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="step-grid">
      <div class="step-item">
        <span class="step-icon">📊</span>
        <div><span class="step-title">Summary Sheet</span>
        <span class="step-desc">KPI cards showing Total POs, Spend, Entities & Vendors overview.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">📈</span>
        <div><span class="step-title">Sector Breakdown</span>
        <span class="step-desc">Sector overview table with Total PO, Open PO, Spend & % of Total.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">✅</span>
        <div><span class="step-title">Status Analysis</span>
        <span class="step-desc">PO Status & Approval Status breakdown with detailed KPI cards.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">📋</span>
        <div><span class="step-title">Entity × Year Sheets</span>
        <span class="step-desc">One sheet per entity × year combination with approval status cards.</span></div>
      </div>
      <div class="step-item">
        <span class="step-icon">🗂️</span>
        <div><span class="step-title">Cleaned Data Sheet</span>
        <span class="step-desc">Complete processed data exported with all columns & formatting.</span></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    upload_summary = st.file_uploader(
        "Upload your cleaned PO report (Excel or CSV)",
        type=["xlsx", "xls", "csv"], key="summary_upload"
    )

    if upload_summary:
        with st.spinner("Processing data…"):
            sum_raw = load_data(upload_summary.read(), upload_summary.name)
            sum_raw.columns = sum_raw.columns.str.strip()
            sum_raw = sum_raw.dropna(axis=1, how="all")
            if "Sector" not in sum_raw.columns:
                sum_raw = clean_dataframe(sum_raw)
            sum_pos = sum_raw[sum_raw["PO status"].str.strip() == "Open order"].copy() if "PO status" in sum_raw.columns else sum_raw.copy()

        if sum_pos.empty:
            st.warning("No Open POs found in the uploaded file.")
        else:
            amt_col_s = "Total Amount (SAR)" if "Total Amount (SAR)" in sum_pos.columns else "Total amount"

            st.markdown('<div class="section-title">🎛️ Filters</div>', unsafe_allow_html=True)
            sf1, sf2, sf3 = st.columns(3)

            with sf1:
                all_entities = sorted(sum_pos["Entity"].dropna().unique().tolist()) if "Entity" in sum_pos.columns else []
                sel_entities = st.multiselect("Entity", all_entities, default=all_entities, key="sum_entity")

            MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
            with sf2:
                all_years = sorted([str(int(y)) for y in sum_pos["Year"].dropna().unique().tolist()]) if "Year" in sum_pos.columns else []
                sel_years = st.multiselect("Year", all_years, default=all_years, key="sum_year")

            with sf3:
                all_months = [m for m in MONTH_ORDER if m in sum_pos.get("Month", pd.Series()).dropna().unique().tolist()] if "Month" in sum_pos.columns else []
                sel_months = st.multiselect("Month", all_months, default=all_months, key="sum_month")

            filtered_sum = sum_pos.copy()
            if sel_entities: filtered_sum = filtered_sum[filtered_sum["Entity"].isin(sel_entities)]
            if sel_years and "Year" in filtered_sum.columns:
                filtered_sum = filtered_sum[filtered_sum["Year"].astype(str).isin(sel_years)]
            if sel_months and "Month" in filtered_sum.columns:
                filtered_sum = filtered_sum[filtered_sum["Month"].isin(sel_months)]

            if filtered_sum.empty:
                st.warning("No data matches the selected filters.")
            else:
                st.markdown("---")
                st.markdown('<div class="section-title">📌 Executive Summary</div>', unsafe_allow_html=True)
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("Total Open POs", f"{len(filtered_sum):,}")
                if amt_col_s in filtered_sum.columns:
                    total = filtered_sum[amt_col_s].sum()
                    k2.metric("Total Spend (SAR)", f"{total/1_000_000:.2f}M" if total >= 1_000_000 else f"{total:,.0f}")
                if "Entity" in filtered_sum.columns:
                    k3.metric("Entities", f"{filtered_sum['Entity'].nunique():,}")
                if "Year" in filtered_sum.columns:
                    k4.metric("Years", f"{filtered_sum['Year'].nunique():,}")

                st.markdown("---")
                st.markdown('<div class="section-title">🏢 Sector Breakdown</div>', unsafe_allow_html=True)
                if "Sector" in filtered_sum.columns:
                    sec_grp = filtered_sum.groupby("Sector").agg(PO_Count=("Sector","count")).reset_index()
                    if amt_col_s in filtered_sum.columns:
                        sec_spend = filtered_sum.groupby("Sector")[amt_col_s].sum().reset_index()
                        sec_grp = sec_grp.merge(sec_spend, on="Sector")
                        sec_grp.columns = ["Sector","PO Count","Spend (SAR)"]
                        sec_grp["% of Total"] = (sec_grp["PO Count"] / sec_grp["PO Count"].sum() * 100).round(1)
                    st.dataframe(sec_grp, use_container_width=True, hide_index=True)

                st.markdown("---")
                st.markdown('<div class="section-title">📋 Entity × Year Matrix</div>', unsafe_allow_html=True)
                if "Entity" in filtered_sum.columns and "Year" in filtered_sum.columns:
                    matrix = filtered_sum.groupby(["Entity","Year"]).size().reset_index()
                    matrix.columns = ["Entity","Year","PO Count"]
                    matrix["Year"] = matrix["Year"].astype(str)
                    pivot = matrix.pivot(index="Entity", columns="Year", values="PO Count").fillna(0).astype(int)
                    st.dataframe(pivot, use_container_width=True)

                st.markdown("---")
                st.markdown('<div class="section-title">🗂️ Column Selection</div>', unsafe_allow_html=True)
                st.caption("Leave empty to export all columns. Select specific columns to export only those.")

                tracker_cols_sum = ["Open PO Reason", "Expected Closure Date", "Comment"]
                selectable_cols_sum = [c for c in filtered_sum.columns if c not in tracker_cols_sum]

                sel_cols_sum = st.multiselect(
                    "Select columns to include (leave empty = all columns)",
                    options=selectable_cols_sum,
                    default=[],
                    placeholder="Select columns…",
                    key="sum_col_select"
                )

                if sel_cols_sum:
                    final_cols_sum = sel_cols_sum + [c for c in tracker_cols_sum if c in filtered_sum.columns]
                    export_sum_df = filtered_sum[final_cols_sum]
                else:
                    export_sum_df = filtered_sum

                st.markdown("---")

                sheets_count = export_sum_df.groupby(["Entity","Year"]).ngroups if "Entity" in export_sum_df.columns and "Year" in export_sum_df.columns else 0

                if sel_cols_sum:
                    full_export_cols = sel_cols_sum + [c for c in tracker_cols_sum if c in sum_raw.columns]
                    full_export_df = sum_raw[[c for c in full_export_cols if c in sum_raw.columns]]
                else:
                    full_export_df = sum_raw

                with st.spinner("Building Excel report…"):
                    # Filter full_export_df based on selected filters
                    filtered_full_export = full_export_df.copy()
                    if sel_entities and "Entity" in filtered_full_export.columns:
                        filtered_full_export = filtered_full_export[filtered_full_export["Entity"].isin(sel_entities)]
                    if sel_years and "Year" in filtered_full_export.columns:
                        filtered_full_export = filtered_full_export[filtered_full_export["Year"].astype(str).isin(sel_years)]
                    if sel_months and "Month" in filtered_full_export.columns:
                        filtered_full_export = filtered_full_export[filtered_full_export["Month"].isin(sel_months)]
                    
                    # Filter open POs based on the same filters
                    open_export_df = filtered_full_export[filtered_full_export["PO status"].str.strip() == "Open order"].copy() if "PO status" in filtered_full_export.columns else filtered_full_export
                    
                    # Get filtered cleaned data (apply same filters)
                    filtered_cleaned_data = sum_raw.copy()
                    if sel_entities and "Entity" in filtered_cleaned_data.columns:
                        filtered_cleaned_data = filtered_cleaned_data[filtered_cleaned_data["Entity"].isin(sel_entities)]
                    if sel_years and "Year" in filtered_cleaned_data.columns:
                        filtered_cleaned_data = filtered_cleaned_data[filtered_cleaned_data["Year"].astype(str).isin(sel_years)]
                    if sel_months and "Month" in filtered_cleaned_data.columns:
                        filtered_cleaned_data = filtered_cleaned_data[filtered_cleaned_data["Month"].isin(sel_months)]
                    
                    sum_bytes = to_summary_report_excel(filtered_full_export, open_df=open_export_df, cleaned_data_df=filtered_cleaned_data)
                
                st.download_button(
                    label=f"⬇️  Download PO Summary Report ({len(open_export_df):,} Open POs · {sheets_count + 2} Sheets · {len(filtered_full_export.columns)} Columns)",
                    data=sum_bytes,
                    file_name=f"PO Summary Report_{pd.Timestamp.today().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                
                st.info("📊 **Report Structure (Filtered):**\n"
                        "• **Summary**: KPI cards, sector breakdown & status analysis\n"
                        "• **Entity × Year Sheets**: One per combination with approval status cards\n"
                        f"• **Cleaned Data**: Filtered dataset ({len(filtered_cleaned_data):,} rows × {len(filtered_cleaned_data.columns)} columns)\n\n"
                        f"**Applied Filters:** Entity: {len(sel_entities)} | Year: {len(sel_years)} | Month: {len(sel_months)}")
    else:
        st.markdown("""
        <div class="section-card" style="text-align:center; padding:56px 24px;">
          <div style="font-size:3.5rem; margin-bottom:16px;">📈</div>
          <h3 style="color:#4c1d95; margin:0 0 8px 0; font-weight:700;">Upload a Cleaned PO Report</h3>
          <p style="color:#7c3aed; margin:0 0 6px 0; font-size:0.95rem;">Generates a complete summary report with Summary sheet, Entity×Year breakdown, and Cleaned Data.</p>
          <p style="color:#a78bfa; margin:0; font-size:0.85rem;">Multiple sheets with professional formatting & complete data export.</p>
        </div>
        """, unsafe_allow_html=True)
