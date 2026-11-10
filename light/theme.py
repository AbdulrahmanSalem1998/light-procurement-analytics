"""Visual theme: brand colours, Streamlit CSS and the Plotly chart style."""
from __future__ import annotations

# ── Brand Colors ───────────────────────────────────────────────
C_PRIMARY   = "#7c3aed"   # violet-600
C_SEC       = "#ec4899"   # pink-500
C_ACCENT    = "#06b6d4"   # cyan-500
C_DARK      = "#0f0a1e"   # near black
C_DARK2     = "#1a1035"   # card dark
C_MID       = "#6d28d9"   # violet-700
C_LIGHT     = "#ede9fe"   # violet-100
C_MUTED     = "#a78bfa"   # violet-400
C_MUTED2    = "#c4b5fd"   # violet-300
C_BG        = "#f5f3ff"   # violet-50
C_CARD      = "#ffffff"

BRAND_SCALE = ["#0f0a1e", "#6d28d9", "#7c3aed", "#a78bfa", "#ec4899", "#f472b6"]
PIE_COLORS  = ["#7c3aed", "#ec4899", "#06b6d4", "#f59e0b", "#10b981", "#6d28d9", "#f472b6", "#a78bfa"]

# ── Streamlit CSS (injected once by app.py) ───────────────────
CSS = f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

  html, body, [class*="css"] {{ font-family: 'Inter', sans-serif; }}
  .stApp {{ background: linear-gradient(135deg, #f5f3ff 0%, #fdf4ff 50%, #f0f9ff 100%) !important; }}

  /* ── Sidebar ── */
  [data-testid="stSidebar"] {{
      background: linear-gradient(180deg, {C_DARK} 0%, {C_DARK2} 60%, #2d1b69 100%) !important;
      border-right: 1px solid rgba(124,58,237,0.3);
  }}
  [data-testid="stSidebar"] * {{ color: {C_MUTED2} !important; }}
  [data-testid="stSidebar"] h1,
  [data-testid="stSidebar"] h2,
  [data-testid="stSidebar"] h3 {{ color: #fff !important; }}
  [data-testid="stSidebar"] .stSelectbox > div > div {{
      background: rgba(124,58,237,0.15) !important;
      border: 1px solid rgba(124,58,237,0.4) !important;
      color: #fff !important;
      border-radius: 10px;
  }}
  [data-testid="stSidebar"] hr {{ border-color: rgba(124,58,237,0.2) !important; }}

  /* ── Header ── */
  .app-header {{
      background: linear-gradient(135deg, {C_DARK} 0%, #3b0764 40%, #6d28d9 70%, {C_SEC} 100%);
      padding: 24px 36px;
      border-radius: 20px;
      margin-bottom: 28px;
      display: flex;
      align-items: center;
      gap: 20px;
      box-shadow: 0 8px 32px rgba(124,58,237,0.35), 0 2px 8px rgba(236,72,153,0.2);
      position: relative;
      overflow: hidden;
  }}
  .app-header::before {{
      content: '';
      position: absolute;
      top: -60px; right: -60px;
      width: 220px; height: 220px;
      background: radial-gradient(circle, rgba(236,72,153,0.25) 0%, transparent 70%);
      pointer-events: none;
  }}
  .app-header::after {{
      content: '';
      position: absolute;
      bottom: -40px; left: 30%;
      width: 180px; height: 180px;
      background: radial-gradient(circle, rgba(6,182,212,0.15) 0%, transparent 70%);
      pointer-events: none;
  }}
  .app-header .header-text h1 {{
      color: #ffffff;
      margin: 0;
      font-size: 2rem;
      font-weight: 800;
      letter-spacing: -0.5px;
      background: linear-gradient(90deg, #fff 0%, {C_MUTED2} 100%);
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
  }}
  .app-header .header-text p {{
      color: rgba(196,181,253,0.85);
      margin: 6px 0 0 0;
      font-size: 0.88rem;
      font-weight: 400;
      letter-spacing: 0.3px;
  }}
  .avatar {{
      width: 56px; height: 56px;
      border-radius: 16px;
      background: linear-gradient(135deg, {C_PRIMARY} 0%, {C_SEC} 100%);
      display: flex; align-items: center; justify-content: center;
      font-size: 1.5rem;
      box-shadow: 0 4px 16px rgba(124,58,237,0.4);
      flex-shrink: 0;
      position: relative;
      z-index: 1;
  }}

  /* ── KPI Cards ── */
  [data-testid="stMetric"] {{
      background: {C_CARD} !important;
      border: 1px solid rgba(124,58,237,0.12) !important;
      border-top: 3px solid transparent !important;
      border-radius: 16px !important;
      padding: 20px 22px !important;
      box-shadow: 0 4px 16px rgba(124,58,237,0.08) !important;
      transition: all 0.25s ease !important;
      background-image: linear-gradient({C_CARD}, {C_CARD}),
                        linear-gradient(135deg, {C_PRIMARY}, {C_SEC}) !important;
      background-origin: border-box !important;
      background-clip: padding-box, border-box !important;
  }}
  [data-testid="stMetric"]:hover {{
      box-shadow: 0 8px 28px rgba(124,58,237,0.18) !important;
      transform: translateY(-2px);
  }}
  [data-testid="stMetricLabel"] {{
      color: {C_MUTED} !important;
      font-size: 0.75rem !important;
      font-weight: 600 !important;
      text-transform: uppercase;
      letter-spacing: 0.8px;
  }}
  [data-testid="stMetricValue"] {{
      color: {C_DARK} !important;
      font-size: 1.7rem !important;
      font-weight: 800 !important;
      background: linear-gradient(135deg, {C_PRIMARY}, {C_SEC});
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
  }}

  /* ── Tabs ── */
  .stTabs [data-baseweb="tab-list"] {{
      background: {C_DARK};
      border-radius: 14px 14px 0 0;
      padding: 8px 12px 0 12px;
      gap: 6px;
  }}
  .stTabs [data-baseweb="tab"] {{
      color: {C_MUTED2} !important;
      background: transparent !important;
      border-radius: 10px 10px 0 0;
      padding: 10px 28px;
      font-weight: 600;
      font-size: 0.9rem;
      letter-spacing: 0.2px;
      border: none !important;
      transition: all 0.2s;
  }}
  .stTabs [aria-selected="true"] {{
      background: linear-gradient(135deg, {C_PRIMARY}, {C_SEC}) !important;
      color: #fff !important;
  }}
  .stTabs [data-baseweb="tab-panel"] {{
      background: linear-gradient(135deg, #f5f3ff 0%, #fdf4ff 100%);
      border: 1px solid rgba(124,58,237,0.12);
      border-top: none;
      border-radius: 0 0 14px 14px;
      padding: 28px 24px;
  }}

  /* ── Section title ── */
  .section-card {{
      background: {C_CARD};
      border: 1px solid rgba(124,58,237,0.1);
      border-radius: 16px;
      padding: 20px 24px;
      margin-bottom: 20px;
      box-shadow: 0 2px 12px rgba(124,58,237,0.06);
  }}
  .section-title {{
      font-size: 0.8rem;
      font-weight: 700;
      background: linear-gradient(90deg, {C_PRIMARY}, {C_SEC});
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      background-clip: text;
      text-transform: uppercase;
      letter-spacing: 1px;
      margin-bottom: 16px;
      padding-bottom: 10px;
      border-bottom: 2px solid {C_LIGHT};
  }}

  /* ── Step cards ── */
  .step-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
      gap: 12px;
      margin-bottom: 20px;
  }}
  .step-item {{
      background: {C_CARD};
      border: 1px solid rgba(124,58,237,0.1);
      border-left: 4px solid {C_PRIMARY};
      border-radius: 12px;
      padding: 14px 18px;
      display: flex;
      align-items: flex-start;
      gap: 12px;
      font-size: 0.875rem;
      color: #1e1b4b;
      box-shadow: 0 2px 8px rgba(124,58,237,0.05);
      transition: all 0.2s;
  }}
  .step-item:hover {{
      border-left-color: {C_SEC};
      box-shadow: 0 4px 16px rgba(124,58,237,0.12);
      transform: translateY(-1px);
  }}
  .step-icon {{ font-size: 1.1rem; flex-shrink: 0; margin-top: 1px; }}
  .step-title {{ font-weight: 700; color: {C_MID}; display: block; margin-bottom: 3px; }}
  .step-desc  {{ color: #6b7280; font-size: 0.8rem; line-height: 1.4; }}

  /* ── Download button ── */
  .stDownloadButton > button {{
      background: linear-gradient(135deg, {C_PRIMARY}, {C_SEC}) !important;
      color: white !important;
      border: none !important;
      border-radius: 12px !important;
      font-weight: 700 !important;
      padding: 12px 24px !important;
      font-size: 0.95rem !important;
      box-shadow: 0 4px 16px rgba(124,58,237,0.35) !important;
      transition: all 0.25s !important;
      letter-spacing: 0.2px !important;
  }}
  .stDownloadButton > button:hover {{
      box-shadow: 0 6px 24px rgba(124,58,237,0.5) !important;
      transform: translateY(-2px) !important;
  }}

  /* ── File uploader ── */
  [data-testid="stFileUploader"] {{
      background: {C_CARD};
      border: 2px dashed rgba(124,58,237,0.3);
      border-radius: 14px;
      padding: 8px;
      transition: all 0.2s;
  }}
  [data-testid="stFileUploader"]:hover {{
      border-color: {C_PRIMARY};
      background: {C_LIGHT};
  }}

  /* ── Expander ── */
  [data-testid="stExpander"] {{
      background: {C_CARD} !important;
      border: 1px solid rgba(124,58,237,0.1) !important;
      border-radius: 14px !important;
  }}

  /* ── Divider ── */
  hr {{ border-color: rgba(124,58,237,0.12) !important; margin: 24px 0; }}

  /* ── Alert ── */
  [data-testid="stAlert"] {{ border-radius: 12px !important; }}
</style>
"""

def style_fig(fig):
    fig.update_layout(
        paper_bgcolor=C_CARD, plot_bgcolor=C_CARD,
        font=dict(family="Inter", size=12, color="#4b5563"),
        title_font=dict(family="Inter", size=14, color=C_MID, weight=700),
        legend=dict(bgcolor=C_CARD, bordercolor=C_LIGHT, borderwidth=1,
                    font=dict(size=11)),
        margin=dict(t=48, b=20, l=8, r=8),
        hoverlabel=dict(bgcolor=C_DARK, font_color="#fff",
                        font_family="Inter", bordercolor=C_PRIMARY),
    )
    fig.update_xaxes(gridcolor="#ede9fe", linecolor="#ede9fe", tickfont=dict(size=11))
    fig.update_yaxes(gridcolor="#ede9fe", linecolor="#ede9fe", tickfont=dict(size=11))
    return fig
