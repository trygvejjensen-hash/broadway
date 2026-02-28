import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import os
from datetime import datetime

# ─── Page Config ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Broadway · TrendVision",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── TrendVision Style CSS (Light Theme) ───────────────────────────────────
PURPLE = "#7C3AED"
PURPLE_LIGHT = "#EDE9FE"
PURPLE_MID = "#A78BFA"
BLUE = "#3B82F6"
GREEN = "#10B981"
RED = "#EF4444"
AMBER = "#F59E0B"
GRAY_50 = "#F9FAFB"
GRAY_100 = "#F3F4F6"
GRAY_200 = "#E5E7EB"
GRAY_300 = "#D1D5DB"
GRAY_500 = "#6B7280"
GRAY_700 = "#374151"
GRAY_900 = "#111827"

CHART_TEMPLATE = "plotly_white"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    .stApp {{background-color: {GRAY_50}; font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;}}
    #MainMenu {{visibility: hidden;}}
    footer {{visibility: hidden;}}
    header[data-testid="stHeader"] {{background-color: white; border-bottom: 1px solid {GRAY_200};}}
    [data-testid="stSidebar"] {{background-color: white; border-right: 1px solid {GRAY_200};}}
    [data-testid="stSidebar"] .stMarkdown h2 {{color: {GRAY_900}; font-weight: 700;}}
    [data-testid="stSidebar"] hr {{border-color: {GRAY_200};}}
    .tv-kpi {{background: white; border: 1px solid {GRAY_200}; border-radius: 12px; padding: 20px 24px; transition: all 0.2s ease;}}
    .tv-kpi:hover {{border-color: {PURPLE}; box-shadow: 0 4px 12px rgba(124, 58, 237, 0.08);}}
    .tv-kpi-label {{font-size: 0.75rem; font-weight: 500; color: {GRAY_500}; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 4px;}}
    .tv-kpi-value {{font-size: 1.75rem; font-weight: 700; color: {GRAY_900}; line-height: 1.2;}}
    .tv-kpi-delta {{display: inline-flex; align-items: center; gap: 3px; font-size: 0.8rem; font-weight: 500; margin-top: 6px; padding: 2px 8px; border-radius: 20px;}}
    .tv-delta-up {{color: #059669; background: #D1FAE5;}}
    .tv-delta-down {{color: #DC2626; background: #FEE2E2;}}
    .tv-section {{font-size: 1.1rem; font-weight: 600; color: {GRAY_900}; margin: 2rem 0 1rem 0; display: flex; align-items: center; gap: 8px;}}
    .tv-section-badge {{background: {PURPLE_LIGHT}; color: {PURPLE}; font-size: 0.7rem; font-weight: 600; padding: 3px 10px; border-radius: 20px; text-transform: uppercase; letter-spacing: 0.05em;}}
    .tv-alert {{background: white; border: 1px solid {GRAY_200}; border-left: 4px solid {RED}; border-radius: 8px; padding: 12px 16px; margin-bottom: 8px; font-size: 0.85rem; color: {GRAY_700};}}
    .tv-alert strong {{color: {GRAY_900};}}
    .tv-alert-warn {{border-left-color: {AMBER};}}
    .tv-brand-badge {{display: inline-block; background: {PURPLE_LIGHT}; color: {PURPLE}; font-size: 0.75rem; font-weight: 600; padding: 2px 8px; border-radius: 6px; margin-right: 6px;}}
    .tv-welcome {{background: white; border: 2px dashed {GRAY_300}; border-radius: 16px; padding: 3rem 2rem; text-align: center;}}
    .tv-welcome h3 {{color: {GRAY_900}; font-weight: 700;}}
    .tv-welcome p {{color: {GRAY_500};}}
    .stTabs [data-baseweb="tab-list"] {{gap: 0; background: white; border-bottom: 1px solid {GRAY_200}; padding: 0 4px;}}
    .stTabs [data-baseweb="tab"] {{color: {GRAY_500}; font-weight: 500; font-size: 0.9rem; padding: 12px 20px; border-bottom: 2px solid transparent;}}
    .stTabs [aria-selected="true"] {{color: {PURPLE} !important; border-bottom: 2px solid {PURPLE} !important; font-weight: 600;}}
    [data-testid="stMetric"] {{background: white; border: 1px solid {GRAY_200}; border-radius: 12px; padding: 16px 20px;}}
    [data-testid="stMetricLabel"] {{color: {GRAY_500}; font-size: 0.8rem;}}
    [data-testid="stMetricValue"] {{color: {GRAY_900}; font-weight: 700;}}
    [data-testid="stDataFrame"] {{border: 1px solid {GRAY_200}; border-radius: 12px; overflow: hidden;}}
</style>
""", unsafe_allow_html=True)


# ─── Data Loading ───────────────────────────────────────────────────────────
DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "data", "broadway_data.xlsm")

# Columns that store 0-1 fractions and need ×100 conversion
FRACTION_COLS = [
    "Qualified product page %",
    "Products with >20 reviews %",
    "Products with brand authorization %",
    "Products with promotion %",
    "Products with free shipping %",
    "Products in collaborations % (last 30d)",
    "Open collaboration product % (last 30d)",
    "Products with samples % (last 30d)",
    "High-GMV creators %",
    "Low PPS creators %",
    "LIVE with engagement tools %",
    "LIVE with promotions %",
    "30-Day Seller Fault Cancellation Rate",
    "60-Day Negative Review Rate (NRR)",
    "60-Day Non-Buyer Fault R&R Rate",
    "60-Day IM Dissatisfaction Rate",
]

# All numeric columns in Raw sheet
RAW_NUMERIC_COLS = [
    "shop performance score", "product satisfaction", "Fulfillment and Logistics", "Customer Service",
    "Shop Ranking", "# of PID",
    "60-Day Negative Review Rate (NRR)", "60-Day Non-Buyer Fault R&R Rate",
    "30-Day Seller Fault Cancellation Rate", "30-Day On-Time Delivery Rate",
    "60-Day IM Dissatisfaction Rate", "60-Day After-ale Handling Time",
    "Qualified product page %", "Products with >20 reviews %",
    "Avg. product rating", "Out-of-stock products",
    "Avg. qualified hero products", "Avg. live listings",
    "Products with brand authorization %", "New active products (last 30d)",
    "Low PPS creators %", "Products in collaborations % (last 30d)",
    "Open collaboration product % (last 30d)", "Products with samples % (last 30d)",
    "Product impressions to creators (last 30d)",
    "Creators added products from invitations (last 30d)",
    "Creators received samples (last 15d)", "Creators posted content",
    "High-GMV creators %",
    "Videos with new violations", "Number of videos", "Avg. video views", "Avg GPM per video",
    "LIVE streams with violations", "Number of LIVE streams", "Avg. LIVE duration (hr)",
    "Avg. LIVE impressions per hour", "Avg. GPM per LIVE",
    "LIVE with engagement tools %", "LIVE with promotions %",
    "Products with promotion %", "Products with free shipping %",
    "Campaigns joined", "Shop ads cost",
    "Days with low AHR (<200)", "Days with SPS < 3.5",
    "Recent shop violations",
    "# of lisitng quality not good",
    "Product opportunity submissions (last 30d)",
    "Recipients read (last 30d)",
    "Days to first sale", "Days to T3",
    "Return and refund rate",
]

# Partner Raw numeric columns
PARTNER_NUMERIC_COLS = [
    "GMV", "Items sold", "LIVE GMV", "Video GMV", "Product card GMV",
    "Ads GMV", "Ads cost", "Ads SKU orders", "Ads cost per order",
    "Affiliate GMV", "Avg customers", "Refunds", "Impressions",
    "Avg visitors", "Avg conversion rate",
]


@st.cache_data(ttl=300)
def load_data(file):
    """Load all relevant sheets from the Broadway Excel file."""
    try:
        raw = pd.read_excel(file, sheet_name="Raw", engine="openpyxl")
        try:
            partner_raw = pd.read_excel(file, sheet_name="Partner Raw", engine="openpyxl")
        except Exception:
            partner_raw = pd.DataFrame()
        try:
            am_mapping = pd.read_excel(file, sheet_name="AM Seller Mapping", engine="openpyxl")
        except Exception:
            am_mapping = pd.DataFrame()
        try:
            video_raw = pd.read_excel(file, sheet_name="Partner Video Raw", engine="openpyxl")
        except Exception:
            video_raw = pd.DataFrame()
        try:
            campaigns = pd.read_excel(file, sheet_name="Campaign Date", engine="openpyxl")
        except Exception:
            campaigns = pd.DataFrame()
        return {"raw": raw, "partner_raw": partner_raw, "am_mapping": am_mapping, "video_raw": video_raw, "campaigns": campaigns}
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None


def clean_raw_data(df: pd.DataFrame) -> pd.DataFrame:
    """Clean Raw sheet: convert dashes → NaN, fix types, convert fractions → percentages."""
    df = df.copy()
    # String columns
    str_cols = ["AM", "AD", "Brand", "Stage", "Account Type", "Priority", "TT Category Team",
                "Shop Code", "Top Brand Badge", "VoC diagnosis", "Shop design",
                "Factors below benchmark-Assortment", "Factors below benchmark-Content",
                "Factors below benchmark-Empowerment",
                "Subscription and Save", "co-funded program toggle on",
                "feb restock registered", "March glow up campaign registered", "pre-built automated plan"]
    for col in str_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).replace({"nan": pd.NA, "-": pd.NA, "0": pd.NA})
    if "report date" in df.columns:
        df["report date"] = pd.to_datetime(df["report date"], errors="coerce")

    # Convert numeric columns (dash → NaN)
    for col in RAW_NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # ── Convert fraction columns (0-1) → percentage (0-100) ──
    for col in FRACTION_COLS:
        if col in df.columns:
            df[col] = df[col] * 100.0

    # ── Normalize OTD Rate (mixed format: some 0-1, some already 0-100) ──
    if "30-Day On-Time Delivery Rate" in df.columns:
        otd = df["30-Day On-Time Delivery Rate"]
        # Values ≤ 1 are fractions; values > 1 are already percentages
        mask_frac = otd.notna() & (otd <= 1.0)
        df.loc[mask_frac, "30-Day On-Time Delivery Rate"] = otd[mask_frac] * 100.0

    return df


def clean_partner_raw(df: pd.DataFrame) -> pd.DataFrame:
    """Clean Partner Raw: convert numeric cols, normalize date column."""
    if df.empty:
        return df
    df = df.copy()
    # Standardize date column name
    if "Report Date" in df.columns:
        df = df.rename(columns={"Report Date": "report date"})
    if "report date" in df.columns:
        df["report date"] = pd.to_datetime(df["report date"], errors="coerce")
    # Rename Shop name → Brand for joining
    if "Shop name" in df.columns:
        df = df.rename(columns={"Shop name": "Brand"})
    for col in PARTNER_NUMERIC_COLS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    # Avg conversion rate: check if it's fraction or already %
    if "Avg conversion rate" in df.columns:
        cr = df["Avg conversion rate"]
        # If max < 1, it's fractions → convert
        if cr.dropna().max() <= 1.0:
            df["Avg conversion rate"] = cr * 100.0
    return df


def merge_partner_data(raw_df, partner_df):
    """Merge Partner Raw GMV data into Raw data on Brand + report date."""
    if partner_df.empty:
        return raw_df
    # Select only the columns we need from partner
    partner_cols = ["Brand", "report date"] + [c for c in PARTNER_NUMERIC_COLS if c in partner_df.columns]
    pr_subset = partner_df[partner_cols].copy()
    # Merge
    merged = raw_df.merge(pr_subset, on=["Brand", "report date"], how="left", suffixes=("", "_pr"))
    return merged


# ─── Helpers ────────────────────────────────────────────────────────────────
def safe_mean(s):
    v = s.dropna()
    return v.mean() if len(v) > 0 else None

def safe_sum(s):
    v = s.dropna()
    return v.sum() if len(v) > 0 else 0

def safe_median(s):
    v = s.dropna()
    return v.median() if len(v) > 0 else None

def fmt_pct(val):
    return "—" if val is None or pd.isna(val) else f"{val:.1f}%"

def fmt_num(val, d=1):
    if val is None or pd.isna(val): return "—"
    if abs(val) >= 1_000_000: return f"{val/1_000_000:.{d}f}M"
    if abs(val) >= 1_000: return f"{val/1_000:.{d}f}K"
    return f"{val:.{d}f}"

def fmt_dollar(val, d=0):
    if val is None or pd.isna(val): return "—"
    if abs(val) >= 1_000_000: return f"${val/1_000_000:.1f}M"
    if abs(val) >= 1_000: return f"${val/1_000:.{d}f}K"
    return f"${val:,.{d}f}"

def delta_val(curr, prev):
    if curr is not None and prev is not None and not pd.isna(curr) and not pd.isna(prev):
        return curr - prev
    return None

def kpi_card_html(label, value, delta=None, direction="up"):
    delta_html = ""
    if delta is not None and not pd.isna(delta):
        is_up = delta > 0
        is_good = (is_up and direction == "up") or (not is_up and direction == "down")
        cls = "tv-delta-up" if is_good else "tv-delta-down"
        arrow = "↑" if is_up else "↓"
        delta_html = f'<span class="tv-kpi-delta {cls}">{arrow} {abs(delta):.1f}</span>'
    return f'<div class="tv-kpi"><div class="tv-kpi-label">{label}</div><div class="tv-kpi-value">{value}</div>{delta_html}</div>'

def section_header(icon, title, badge=None):
    badge_html = f'<span class="tv-section-badge">{badge}</span>' if badge else ""
    st.markdown(f'<div class="tv-section">{icon} {title} {badge_html}</div>', unsafe_allow_html=True)

def chart_layout(fig, h=380):
    """Apply TrendVision light theme to a plotly figure."""
    fig.update_layout(
        template=CHART_TEMPLATE,
        paper_bgcolor="white",
        plot_bgcolor="white",
        height=h,
        margin=dict(t=40, b=30, l=50, r=20),
        font=dict(family="Inter, sans-serif", color=GRAY_700, size=12),
        title_font=dict(size=14, color=GRAY_900, family="Inter, sans-serif"),
        legend=dict(font=dict(size=11)),
        xaxis=dict(gridcolor=GRAY_100, linecolor=GRAY_200),
        yaxis=dict(gridcolor=GRAY_100, linecolor=GRAY_200),
    )
    return fig


# ─── Sidebar ────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown('<div style="display:flex; align-items:center; gap:10px; margin-bottom:8px;"><div style="background:#7C3AED; color:white; font-weight:800; font-size:0.8rem; padding:6px 12px; border-radius:8px; letter-spacing:0.5px;">⚡ TrendVision</div></div>', unsafe_allow_html=True)
    st.caption("Broadway Portfolio Manager")
    st.markdown("---")

    st.markdown("**📁 Data Source**")
    upload_mode = st.radio("Load data from:", ["Upload file", "Use default file"], index=0, label_visibility="collapsed")
    uploaded_file = None
    if upload_mode == "Upload file":
        uploaded_file = st.file_uploader("Drop Broadway .xlsm", type=["xlsm", "xlsx", "xls"], label_visibility="collapsed")
        if uploaded_file:
            st.success(f"✅ {uploaded_file.name}")
    else:
        if os.path.exists(DEFAULT_FILE):
            st.info("Using pre-loaded data")
        else:
            st.warning("No default file. Upload one.")
    st.markdown("---")

# ─── Load Data ──────────────────────────────────────────────────────────────
data_source = None
if upload_mode == "Upload file" and uploaded_file is not None:
    data_source = uploaded_file
elif upload_mode == "Use default file" and os.path.exists(DEFAULT_FILE):
    data_source = DEFAULT_FILE

if data_source is None:
    st.markdown("# ⚡ Broadway Portfolio Dashboard")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(
            '<div class="tv-welcome"><h3>Welcome to TrendVision Broadway</h3><p>Upload your Broadway Tool export (.xlsm) using the sidebar to get started.</p><p style="font-size:0.8rem;">The dashboard will parse your data and show portfolio metrics, trends, and action items.</p></div>',
            unsafe_allow_html=True,
        )
    st.stop()

data = load_data(data_source)
if data is None:
    st.error("Failed to load data.")
    st.stop()

raw = clean_raw_data(data["raw"])
partner_raw = clean_partner_raw(data["partner_raw"])
raw = merge_partner_data(raw, partner_raw)
am_mapping = data["am_mapping"]

# ─── Sidebar Filters ────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("**👤 Account Manager**")
    am_list = sorted([str(x) for x in raw["AM"].dropna().unique().tolist()]) if "AM" in raw.columns else []
    selected_am = st.selectbox("Select your name", ["All AMs"] + am_list, index=0, label_visibility="collapsed")

    # Brand selector — filtered by selected AM
    st.markdown("**🏷️ Brand**")
    if "Brand" in raw.columns:
        if selected_am != "All AMs" and "AM" in raw.columns:
            brand_pool = raw[raw["AM"] == selected_am]["Brand"].dropna().unique().tolist()
        else:
            brand_pool = raw["Brand"].dropna().unique().tolist()
        brand_list = sorted([str(x) for x in brand_pool])
        selected_brand = st.selectbox("Select a brand", ["All Brands"] + brand_list, index=0, label_visibility="collapsed")
    else:
        selected_brand = "All Brands"
    st.markdown("---")

    st.markdown("**📅 Report Week**")
    if "report date" in raw.columns and raw["report date"].notna().any():
        available_weeks = sorted(raw["report date"].dropna().unique())
        week_labels = [d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d) for d in available_weeks]
        selected_week_label = st.selectbox("Week", ["Latest"] + week_labels, index=0, label_visibility="collapsed")
        if selected_week_label == "Latest":
            selected_week = available_weeks[-1]
        else:
            selected_week = available_weeks[week_labels.index(selected_week_label)]
    else:
        selected_week = None
        available_weeks = []
    st.markdown("---")

    st.markdown("**🔍 Filters**")
    if "Account Type" in raw.columns:
        acct_types = sorted([str(x) for x in raw["Account Type"].dropna().unique().tolist()])
        selected_acct_type = st.multiselect("Account Type", acct_types, default=acct_types)
    else:
        selected_acct_type = None
    if "Priority" in raw.columns:
        priorities = sorted([str(x) for x in raw["Priority"].dropna().unique().tolist()])
        selected_priorities = st.multiselect("Priority", priorities, default=priorities)
    else:
        selected_priorities = None
    if "TT Category Team" in raw.columns:
        categories = sorted([str(x) for x in raw["TT Category Team"].dropna().unique().tolist()])
        selected_categories = st.multiselect("Category", categories, default=categories)
    else:
        selected_categories = None

    st.markdown("---")
    st.caption(f"Updated {datetime.now().strftime('%b %d, %Y · %H:%M')}")

# ─── Apply Filters ──────────────────────────────────────────────────────────
df = raw.copy()
if selected_week is not None and "report date" in df.columns:
    df_current = df[df["report date"] == selected_week].copy()
    if len(available_weeks) >= 2:
        prev_idx = list(available_weeks).index(selected_week) - 1
        df_prev = df[df["report date"] == available_weeks[prev_idx]].copy() if prev_idx >= 0 else pd.DataFrame()
    else:
        df_prev = pd.DataFrame()
else:
    df_current = df.copy()
    df_prev = pd.DataFrame()

if selected_am != "All AMs" and "AM" in df_current.columns:
    df_current = df_current[df_current["AM"] == selected_am]
    if not df_prev.empty:
        df_prev = df_prev[df_prev["AM"] == selected_am]
if selected_brand != "All Brands" and "Brand" in df_current.columns:
    df_current = df_current[df_current["Brand"] == selected_brand]
    if not df_prev.empty:
        df_prev = df_prev[df_prev["Brand"] == selected_brand]
if selected_acct_type is not None and "Account Type" in df_current.columns:
    df_current = df_current[df_current["Account Type"].isin(selected_acct_type)]
if selected_priorities is not None and "Priority" in df_current.columns:
    df_current = df_current[df_current["Priority"].isin(selected_priorities)]
if selected_categories is not None and "TT Category Team" in df_current.columns:
    df_current = df_current[df_current["TT Category Team"].isin(selected_categories)]

# Trend data (all weeks, with AM + Brand filter)
df_trend = raw.copy()
if selected_am != "All AMs" and "AM" in df_trend.columns:
    df_trend = df_trend[df_trend["AM"] == selected_am]
if selected_brand != "All Brands" and "Brand" in df_trend.columns:
    df_trend = df_trend[df_trend["Brand"] == selected_brand]


# ─── Header ─────────────────────────────────────────────────────────────────
am_text = selected_am if selected_am != "All AMs" else "All Account Managers"
brand_text = selected_brand if selected_brand != "All Brands" else ""
week_text = selected_week.strftime("%b %d, %Y") if hasattr(selected_week, "strftime") else str(selected_week)
subtitle_parts = [am_text]
if brand_text:
    subtitle_parts.append(brand_text)
subtitle_parts.append(week_text)
subtitle = " · ".join(subtitle_parts)

st.markdown(
    f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:4px;">'
    f'<div><h1 style="margin:0;font-weight:800;color:{GRAY_900};font-size:1.6rem;">Broadway Portfolio</h1>'
    f'<p style="margin:0;color:{GRAY_500};font-size:0.9rem;">{subtitle}</p></div>'
    f'<div style="display:flex;gap:8px;"><span style="background:{PURPLE_LIGHT};color:{PURPLE};padding:6px 14px;border-radius:8px;font-size:0.8rem;font-weight:600;">⚡ TrendVision</span></div>'
    f'</div>',
    unsafe_allow_html=True,
)

# ─── Top KPI Cards ──────────────────────────────────────────────────────────
n_brands = df_current["Brand"].nunique() if "Brand" in df_current.columns else 0
n_brands_prev = df_prev["Brand"].nunique() if not df_prev.empty and "Brand" in df_prev.columns else None
avg_sps = safe_mean(df_current.get("shop performance score", pd.Series()))
avg_sps_prev = safe_mean(df_prev.get("shop performance score", pd.Series())) if not df_prev.empty else None
avg_rating = safe_mean(df_current.get("Avg. product rating", pd.Series()))
avg_rating_prev = safe_mean(df_prev.get("Avg. product rating", pd.Series())) if not df_prev.empty else None
total_gmv = safe_sum(df_current.get("GMV", pd.Series()))
total_gmv_prev = safe_sum(df_prev.get("GMV", pd.Series())) if not df_prev.empty else None
total_items = safe_sum(df_current.get("Items sold", pd.Series()))
total_items_prev = safe_sum(df_prev.get("Items sold", pd.Series())) if not df_prev.empty else None
avg_nrr = safe_mean(df_current.get("60-Day Negative Review Rate (NRR)", pd.Series()))
avg_nrr_prev = safe_mean(df_prev.get("60-Day Negative Review Rate (NRR)", pd.Series())) if not df_prev.empty else None
avg_otd = safe_mean(df_current.get("30-Day On-Time Delivery Rate", pd.Series()))
avg_otd_prev = safe_mean(df_prev.get("30-Day On-Time Delivery Rate", pd.Series())) if not df_prev.empty else None
total_videos = safe_sum(df_current.get("Number of videos", pd.Series()))
total_live = safe_sum(df_current.get("Number of LIVE streams", pd.Series()))

cols = st.columns(8)
kpi_data = [
    ("Active Brands", str(int(n_brands)), delta_val(n_brands, n_brands_prev), "up"),
    ("Total GMV", fmt_dollar(total_gmv), delta_val(total_gmv, total_gmv_prev), "up"),
    ("Items Sold", fmt_num(total_items, 0), delta_val(total_items, total_items_prev), "up"),
    ("Avg SPS", fmt_num(avg_sps) if avg_sps else "—", delta_val(avg_sps, avg_sps_prev), "up"),
    ("Avg Rating", fmt_num(avg_rating) if avg_rating else "—", delta_val(avg_rating, avg_rating_prev), "up"),
    ("Avg NRR", fmt_pct(avg_nrr), delta_val(avg_nrr, avg_nrr_prev), "down"),
    ("Avg OTD", fmt_pct(avg_otd), delta_val(avg_otd, avg_otd_prev), "up"),
    ("Content", f"{fmt_num(total_videos, 0)}V / {fmt_num(total_live, 0)}L", None, "up"),
]
for i, (label, val, delta, direction) in enumerate(kpi_data):
    with cols[i]:
        st.markdown(kpi_card_html(label, val, delta, direction), unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# TABBED NAVIGATION
# ═══════════════════════════════════════════════════════════════════════════
tab_overview, tab_health, tab_content, tab_live, tab_actions, tab_data = st.tabs([
    "📊 Overview", "🏥 Shop Health", "🎬 Content & Creators", "📡 LIVE", "🚨 Action Items", "📥 Data"
])


# ═══════════════════════════════════════════════════════════════════════════
# TAB 1: OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════
with tab_overview:
    # ── GMV & Revenue KPIs ──
    section_header("💰", "GMV & Revenue", "Weekly")
    gk1, gk2, gk3, gk4, gk5, gk6 = st.columns(6)
    with gk1:
        st.metric("Total GMV", fmt_dollar(total_gmv),
                  delta=f"{fmt_dollar(delta_val(total_gmv, total_gmv_prev))}" if total_gmv_prev else None)
    with gk2:
        v = safe_sum(df_current.get("Video GMV", pd.Series()))
        st.metric("Video GMV", fmt_dollar(v))
    with gk3:
        v = safe_sum(df_current.get("LIVE GMV", pd.Series()))
        st.metric("LIVE GMV", fmt_dollar(v))
    with gk4:
        v = safe_sum(df_current.get("Affiliate GMV", pd.Series()))
        st.metric("Affiliate GMV", fmt_dollar(v))
    with gk5:
        v = safe_sum(df_current.get("Ads GMV", pd.Series()))
        st.metric("Ads GMV", fmt_dollar(v))
    with gk6:
        v = safe_sum(df_current.get("Product card GMV", pd.Series()))
        st.metric("Product Card GMV", fmt_dollar(v))

    # ── SPS Sub-Scores ──
    section_header("⭐", "SPS Breakdown", "Weekly")
    ss1, ss2, ss3, ss4 = st.columns(4)
    with ss1:
        v = safe_mean(df_current.get("shop performance score", pd.Series()))
        vp = safe_mean(df_prev.get("shop performance score", pd.Series())) if not df_prev.empty else None
        st.metric("Overall SPS", fmt_num(v) if v else "—",
                  delta=f"{delta_val(v, vp):.2f}" if delta_val(v, vp) is not None else None)
    with ss2:
        v = safe_mean(df_current.get("product satisfaction", pd.Series()))
        vp = safe_mean(df_prev.get("product satisfaction", pd.Series())) if not df_prev.empty else None
        st.metric("Product Satisfaction", fmt_num(v) if v else "—",
                  delta=f"{delta_val(v, vp):.2f}" if delta_val(v, vp) is not None else None)
    with ss3:
        v = safe_mean(df_current.get("Fulfillment and Logistics", pd.Series()))
        vp = safe_mean(df_prev.get("Fulfillment and Logistics", pd.Series())) if not df_prev.empty else None
        st.metric("Fulfillment & Logistics", fmt_num(v) if v else "—",
                  delta=f"{delta_val(v, vp):.2f}" if delta_val(v, vp) is not None else None)
    with ss4:
        v = safe_mean(df_current.get("Customer Service", pd.Series()))
        vp = safe_mean(df_prev.get("Customer Service", pd.Series())) if not df_prev.empty else None
        st.metric("Customer Service", fmt_num(v) if v else "—",
                  delta=f"{delta_val(v, vp):.2f}" if delta_val(v, vp) is not None else None)

    # ── Performance Trends ──
    section_header("📈", "Performance Trends", "Weekly")

    if "report date" in df_trend.columns and df_trend["report date"].notna().any():
        # Build trend aggregation
        agg_dict = {}
        for col_name, agg_func in [
            ("shop performance score", "mean"), ("product satisfaction", "mean"),
            ("Fulfillment and Logistics", "mean"), ("Customer Service", "mean"),
            ("Avg. product rating", "mean"),
            ("60-Day Negative Review Rate (NRR)", "mean"),
            ("30-Day On-Time Delivery Rate", "mean"),
            ("Number of videos", "sum"), ("Number of LIVE streams", "sum"),
            ("Avg. video views", "mean"), ("Avg GPM per video", "mean"),
            ("Avg. GPM per LIVE", "mean"), ("Brand", "nunique"),
            ("GMV", "sum"), ("Items sold", "sum"), ("Video GMV", "sum"),
            ("LIVE GMV", "sum"), ("Affiliate GMV", "sum"), ("Ads GMV", "sum"),
            ("Avg visitors", "sum"), ("Avg customers", "sum"),
            ("Avg conversion rate", "mean"),
            ("Qualified product page %", "mean"),
            ("Products with brand authorization %", "mean"),
            ("Products with promotion %", "mean"),
        ]:
            if col_name in df_trend.columns:
                agg_dict[col_name] = agg_func

        trend_agg = df_trend.groupby("report date").agg(agg_dict).reset_index()
        trend_agg = trend_agg.rename(columns={"report date": "Week", "Brand": "Brands"})
        trend_agg = trend_agg.sort_values("Week")

        # Row 1: GMV trend + SPS trend
        c1, c2 = st.columns(2)
        with c1:
            if "GMV" in trend_agg.columns:
                fig = go.Figure()
                fig.add_trace(go.Bar(x=trend_agg["Week"], y=trend_agg["GMV"], name="Total GMV",
                    marker_color=PURPLE, opacity=0.8))
                if "Video GMV" in trend_agg.columns:
                    fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["Video GMV"],
                        name="Video GMV", line=dict(color=BLUE, width=2), mode="lines+markers"))
                if "LIVE GMV" in trend_agg.columns:
                    fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["LIVE GMV"],
                        name="LIVE GMV", line=dict(color=GREEN, width=2), mode="lines+markers"))
                fig.update_layout(title="GMV Trend (Weekly)", yaxis_title="$",
                    legend=dict(orientation="h", y=1.08, x=0))
                st.plotly_chart(chart_layout(fig), use_container_width=True)

        with c2:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["shop performance score"],
                mode="lines+markers", line=dict(color=PURPLE, width=2.5), marker=dict(size=7, color=PURPLE),
                fill="tozeroy", fillcolor="rgba(124,58,237,0.06)", name="Overall SPS"))
            if "product satisfaction" in trend_agg.columns:
                fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["product satisfaction"],
                    mode="lines+markers", line=dict(color=GREEN, width=1.5, dash="dot"), name="Product Sat."))
            if "Fulfillment and Logistics" in trend_agg.columns:
                fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["Fulfillment and Logistics"],
                    mode="lines+markers", line=dict(color=BLUE, width=1.5, dash="dot"), name="Fulfillment"))
            if "Customer Service" in trend_agg.columns:
                fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["Customer Service"],
                    mode="lines+markers", line=dict(color=AMBER, width=1.5, dash="dot"), name="Cust. Service"))
            fig.add_hline(y=3.5, line_dash="dash", line_color=RED, annotation_text="Benchmark 3.5", annotation_position="top left")
            fig.update_layout(title="SPS & Sub-Scores Trend", legend=dict(orientation="h", y=1.08, x=0))
            st.plotly_chart(chart_layout(fig), use_container_width=True)

        # Row 2: Rating + NRR/OTD
        c3, c4 = st.columns(2)
        with c3:
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["Avg. product rating"],
                mode="lines+markers", line=dict(color=GREEN, width=2.5), marker=dict(size=7, color=GREEN)))
            fig.add_hline(y=4.0, line_dash="dash", line_color=AMBER, annotation_text="Target 4.0")
            fig.update_layout(title="Average Product Rating")
            st.plotly_chart(chart_layout(fig), use_container_width=True)

        with c4:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            if "60-Day Negative Review Rate (NRR)" in trend_agg.columns:
                fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["60-Day Negative Review Rate (NRR)"],
                    mode="lines+markers", line=dict(color=RED, width=2.5), name="NRR %"), secondary_y=False)
            if "30-Day On-Time Delivery Rate" in trend_agg.columns:
                fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["30-Day On-Time Delivery Rate"],
                    mode="lines+markers", line=dict(color=GREEN, width=2), name="OTD %"), secondary_y=True)
            fig.update_layout(title="NRR & OTD Rate Trends", legend=dict(orientation="h", y=1.08, x=0))
            fig.update_yaxes(title_text="NRR %", secondary_y=False)
            fig.update_yaxes(title_text="OTD %", secondary_y=True)
            st.plotly_chart(chart_layout(fig), use_container_width=True)

        # Row 3: Content Volume + Conversion/Traffic
        c5, c6 = st.columns(2)
        with c5:
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            if "Number of videos" in trend_agg.columns:
                fig.add_trace(go.Bar(x=trend_agg["Week"], y=trend_agg["Number of videos"], name="Videos",
                    marker_color=PURPLE, opacity=0.8), secondary_y=False)
            if "Number of LIVE streams" in trend_agg.columns:
                fig.add_trace(go.Bar(x=trend_agg["Week"], y=trend_agg["Number of LIVE streams"], name="LIVE",
                    marker_color=PURPLE_MID, opacity=0.6), secondary_y=False)
            if "Avg. video views" in trend_agg.columns:
                fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["Avg. video views"], name="Avg Views",
                    line=dict(color=AMBER, width=2), mode="lines+markers"), secondary_y=True)
            fig.update_layout(title="Content Volume & Views", barmode="group",
                legend=dict(orientation="h", y=1.08, x=0))
            fig.update_yaxes(title_text="Count", secondary_y=False)
            fig.update_yaxes(title_text="Avg Views", secondary_y=True)
            st.plotly_chart(chart_layout(fig), use_container_width=True)

        with c6:
            if "Avg visitors" in trend_agg.columns:
                fig = make_subplots(specs=[[{"secondary_y": True}]])
                fig.add_trace(go.Bar(x=trend_agg["Week"], y=trend_agg["Avg visitors"], name="Visitors",
                    marker_color=BLUE, opacity=0.7), secondary_y=False)
                if "Avg customers" in trend_agg.columns:
                    fig.add_trace(go.Bar(x=trend_agg["Week"], y=trend_agg["Avg customers"], name="Customers",
                        marker_color=GREEN, opacity=0.7), secondary_y=False)
                if "Avg conversion rate" in trend_agg.columns:
                    fig.add_trace(go.Scatter(x=trend_agg["Week"], y=trend_agg["Avg conversion rate"],
                        name="Conv. Rate %", line=dict(color=AMBER, width=2), mode="lines+markers"), secondary_y=True)
                fig.update_layout(title="Traffic & Conversion", barmode="group",
                    legend=dict(orientation="h", y=1.08, x=0))
                fig.update_yaxes(title_text="Count", secondary_y=False)
                fig.update_yaxes(title_text="Conv Rate %", secondary_y=True)
                st.plotly_chart(chart_layout(fig), use_container_width=True)

    # ── Portfolio Distribution ──
    section_header("🗂️", "Portfolio Distribution")
    d1, d2, d3 = st.columns(3)

    with d1:
        if "Priority" in df_current.columns:
            priority_counts = df_current["Priority"].value_counts().reset_index()
            priority_counts.columns = ["Priority", "Count"]
            fig = px.pie(priority_counts, values="Count", names="Priority", hole=0.55,
                color_discrete_sequence=[PURPLE, PURPLE_MID, BLUE, GREEN, AMBER])
            fig.update_layout(title="By Priority")
            fig.update_traces(textinfo="label+value", textfont_size=11)
            st.plotly_chart(chart_layout(fig, 320), use_container_width=True)

    with d2:
        if "Account Type" in df_current.columns:
            at_counts = df_current["Account Type"].value_counts().reset_index()
            at_counts.columns = ["Account Type", "Count"]
            fig = px.pie(at_counts, values="Count", names="Account Type", hole=0.55,
                color_discrete_sequence=[BLUE, PURPLE, GREEN, AMBER, RED])
            fig.update_layout(title="By Account Type")
            fig.update_traces(textinfo="label+value", textfont_size=11)
            st.plotly_chart(chart_layout(fig, 320), use_container_width=True)

    with d3:
        if "TT Category Team" in df_current.columns:
            cat_counts = df_current["TT Category Team"].value_counts().head(8).reset_index()
            cat_counts.columns = ["Category", "Count"]
            fig = px.bar(cat_counts, x="Count", y="Category", orientation="h",
                color="Count", color_continuous_scale=["#EDE9FE", PURPLE])
            fig.update_layout(title="Top Categories", showlegend=False, coloraxis_showscale=False)
            st.plotly_chart(chart_layout(fig, 320), use_container_width=True)

    # ── GMV by Brand (Top 15) ──
    section_header("🏆", "Top Brands by GMV")
    if "GMV" in df_current.columns:
        top_gmv = df_current[["Brand", "GMV"]].dropna().query("GMV > 0").sort_values("GMV", ascending=True).tail(15)
        if not top_gmv.empty:
            fig = px.bar(top_gmv, y="Brand", x="GMV", orientation="h",
                color="GMV", color_continuous_scale=[PURPLE_LIGHT, PURPLE])
            fig.update_layout(title="Top 15 Brands by GMV", coloraxis_showscale=False, showlegend=False,
                xaxis_title="GMV ($)")
            st.plotly_chart(chart_layout(fig, 420), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 2: SHOP HEALTH
# ═══════════════════════════════════════════════════════════════════════════
with tab_health:
    # ── SPS Sub-Score KPIs ──
    section_header("⭐", "SPS Components")
    sk1, sk2, sk3, sk4, sk5, sk6 = st.columns(6)
    with sk1:
        v = safe_mean(df_current.get("shop performance score", pd.Series()))
        st.metric("Overall SPS", fmt_num(v) if v else "—")
    with sk2:
        v = safe_mean(df_current.get("product satisfaction", pd.Series()))
        st.metric("Product Satisfaction", fmt_num(v) if v else "—")
    with sk3:
        v = safe_mean(df_current.get("Fulfillment and Logistics", pd.Series()))
        st.metric("Fulfillment", fmt_num(v) if v else "—")
    with sk4:
        v = safe_mean(df_current.get("Customer Service", pd.Series()))
        st.metric("Customer Service", fmt_num(v) if v else "—")
    with sk5:
        v = safe_mean(df_current.get("30-Day On-Time Delivery Rate", pd.Series()))
        st.metric("OTD Rate", fmt_pct(v))
    with sk6:
        v = safe_mean(df_current.get("60-Day Negative Review Rate (NRR)", pd.Series()))
        st.metric("NRR", fmt_pct(v))

    # ── Brand Health Table ──
    section_header("🏥", "Brand Health Table", f"{len(df_current)} brands")

    health_cols = ["Brand", "Priority", "shop performance score", "product satisfaction",
        "Fulfillment and Logistics", "Customer Service",
        "Shop Ranking", "GMV", "Items sold",
        "Avg. product rating", "60-Day Negative Review Rate (NRR)",
        "30-Day On-Time Delivery Rate", "Qualified product page %",
        "Products with brand authorization %",
        "Out-of-stock products", "Recent shop violations",
        "Days with SPS < 3.5", "Number of videos", "Number of LIVE streams"]
    available_health_cols = [c for c in health_cols if c in df_current.columns]

    if available_health_cols:
        hdf = df_current[available_health_cols].copy().sort_values(
            by="shop performance score" if "shop performance score" in df_current.columns else available_health_cols[0], ascending=True)
        rename = {"shop performance score": "SPS", "product satisfaction": "Prod Sat",
            "Fulfillment and Logistics": "Fulfill", "Customer Service": "Cust Svc",
            "Shop Ranking": "Rank", "Avg. product rating": "Rating",
            "60-Day Negative Review Rate (NRR)": "NRR %", "30-Day On-Time Delivery Rate": "OTD %",
            "Qualified product page %": "Qual Pages %",
            "Products with brand authorization %": "Brand Auth %",
            "Out-of-stock products": "OOS",
            "Recent shop violations": "Violations", "Days with SPS < 3.5": "Days<3.5",
            "Number of videos": "Videos", "Number of LIVE streams": "LIVE",
            "Items sold": "Items"}
        hdf = hdf.rename(columns=rename)

        def style_health(row):
            styles = [""] * len(row)
            c = row.index.tolist()
            if "SPS" in c:
                v = row["SPS"]
                i = c.index("SPS")
                if pd.notna(v):
                    if v < 3.0: styles[i] = f"background-color: #FEE2E2; color: {RED}"
                    elif v < 3.5: styles[i] = f"background-color: #FEF3C7; color: {AMBER}"
                    elif v >= 4.5: styles[i] = f"background-color: #D1FAE5; color: #059669"
            if "Violations" in c:
                v = row["Violations"]
                i = c.index("Violations")
                if pd.notna(v) and v > 0: styles[i] = f"background-color: #FEE2E2; color: {RED}"
            if "OOS" in c:
                v = row["OOS"]
                i = c.index("OOS")
                if pd.notna(v) and v > 5: styles[i] = f"background-color: #FEF3C7; color: {AMBER}"
            if "NRR %" in c:
                v = row["NRR %"]
                i = c.index("NRR %")
                if pd.notna(v) and v > 5: styles[i] = f"background-color: #FEE2E2; color: {RED}"
                elif pd.notna(v) and v > 3: styles[i] = f"background-color: #FEF3C7; color: {AMBER}"
            return styles

        styled = hdf.style.apply(style_health, axis=1).format(precision=1, na_rep="—")
        st.dataframe(styled, use_container_width=True, height=min(500, 35 * len(hdf) + 40), hide_index=True)

    # ── SPS Distribution + Radar ──
    section_header("📊", "SPS Distribution & Health Radar")
    if "shop performance score" in df_current.columns:
        sps_vals = df_current["shop performance score"].dropna()
        sps_vals = sps_vals[sps_vals > 0]
        c1, c2 = st.columns(2)
        with c1:
            fig = px.histogram(sps_vals, nbins=20, color_discrete_sequence=[PURPLE])
            fig.update_layout(title="SPS Distribution Across Brands", xaxis_title="SPS", yaxis_title="Brand Count")
            fig.add_vrect(x0=0, x1=3.5, fillcolor=RED, opacity=0.05, line_width=0, annotation_text="Below benchmark", annotation_position="top left")
            st.plotly_chart(chart_layout(fig), use_container_width=True)

        with c2:
            # Radar chart — all values now in correct scales after clean_raw_data
            radar_metrics = {
                "Overall SPS": (safe_mean(df_current.get("shop performance score", pd.Series())), 5),
                "Product Sat.": (safe_mean(df_current.get("product satisfaction", pd.Series())), 5),
                "Fulfillment": (safe_mean(df_current.get("Fulfillment and Logistics", pd.Series())), 5),
                "Cust. Service": (safe_mean(df_current.get("Customer Service", pd.Series())), 5),
                "Rating": (safe_mean(df_current.get("Avg. product rating", pd.Series())), 5),
                "OTD Rate": (safe_mean(df_current.get("30-Day On-Time Delivery Rate", pd.Series())), 100),
                "Qual Pages": (safe_mean(df_current.get("Qualified product page %", pd.Series())), 100),
                "Brand Auth": (safe_mean(df_current.get("Products with brand authorization %", pd.Series())), 100),
                "Promo %": (safe_mean(df_current.get("Products with promotion %", pd.Series())), 100),
            }
            labels = list(radar_metrics.keys())
            vals = [(v[0] / v[1] * 100) if v[0] is not None else 0 for v in radar_metrics.values()]

            fig = go.Figure(go.Scatterpolar(r=vals + [vals[0]], theta=labels + [labels[0]],
                fill="toself", fillcolor="rgba(124,58,237,0.12)", line=dict(color=PURPLE, width=2),
                marker=dict(size=6, color=PURPLE)))
            fig.update_layout(
                title="Portfolio Health Radar",
                polar=dict(
                    radialaxis=dict(visible=True, range=[0, 100], gridcolor=GRAY_200, linecolor=GRAY_200),
                    angularaxis=dict(gridcolor=GRAY_200, linecolor=GRAY_200)
                ),
            )
            st.plotly_chart(chart_layout(fig), use_container_width=True)

    # ── Compliance & Quality ──
    section_header("🛡️", "Compliance & Quality Flags")
    cq1, cq2, cq3, cq4 = st.columns(4)
    with cq1:
        v = safe_sum(df_current.get("Recent shop violations", pd.Series()))
        st.metric("Shop Violations", fmt_num(v, 0))
    with cq2:
        v = safe_sum(df_current.get("Videos with new violations", pd.Series()))
        st.metric("Video Violations", fmt_num(v, 0))
    with cq3:
        v = safe_sum(df_current.get("LIVE streams with violations", pd.Series()))
        st.metric("LIVE Violations", fmt_num(v, 0))
    with cq4:
        v = safe_sum(df_current.get("# of lisitng quality not good", pd.Series()))
        st.metric("Low Quality Listings", fmt_num(v, 0))

    # ── Milestones & Programs ──
    section_header("🎯", "Milestones & Programs")
    mp1, mp2, mp3, mp4 = st.columns(4)
    with mp1:
        if "Top Brand Badge" in df_current.columns:
            badge_count = df_current["Top Brand Badge"].dropna().count()
            st.metric("Top Brand Badges", str(int(badge_count)))
    with mp2:
        v = safe_median(df_current.get("Days to first sale", pd.Series()))
        st.metric("Median Days to First Sale", fmt_num(v, 0) if v else "—")
    with mp3:
        v = safe_median(df_current.get("Days to T3", pd.Series()))
        st.metric("Median Days to T3", fmt_num(v, 0) if v else "—")
    with mp4:
        if "Subscription and Save" in df_current.columns:
            sub_y = df_current["Subscription and Save"].str.upper().eq("Y").sum()
            st.metric("Sub & Save Enrolled", str(int(sub_y)))

    # ── SPS Heatmap by Brand Over Time ──
    section_header("🔥", "SPS Heatmap Over Time")
    if "report date" in raw.columns and "Brand" in raw.columns:
        hm_df = df_trend.copy() if selected_am != "All AMs" else raw.copy()
        if len(hm_df) > 0:
            pivot = hm_df.pivot_table(index="Brand", columns="report date", values="shop performance score", aggfunc="mean")
            pivot = pivot.dropna(how="all")
            if pivot.shape[0] > 25:
                pivot = pivot.iloc[:25]
            if not pivot.empty:
                fig = px.imshow(pivot, color_continuous_scale=["#FEE2E2", "#FEF3C7", "#D1FAE5"],
                    aspect="auto", labels=dict(color="SPS"))
                fig.update_layout(title="SPS by Brand Over Time", xaxis_title="Week", yaxis_title="")
                st.plotly_chart(chart_layout(fig, h=max(300, pivot.shape[0] * 22)), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 3: CONTENT & CREATORS
# ═══════════════════════════════════════════════════════════════════════════
with tab_content:
    # ── Creator KPIs ──
    section_header("🎨", "Creator & Content Metrics", "Weekly")
    ck1, ck2, ck3, ck4, ck5, ck6 = st.columns(6)
    with ck1:
        st.metric("Collab Rate", fmt_pct(safe_mean(df_current.get("Products in collaborations % (last 30d)", pd.Series()))))
    with ck2:
        st.metric("Sample Rate", fmt_pct(safe_mean(df_current.get("Products with samples % (last 30d)", pd.Series()))))
    with ck3:
        st.metric("High-GMV Creators", fmt_pct(safe_mean(df_current.get("High-GMV creators %", pd.Series()))))
    with ck4:
        st.metric("Low PPS Creators", fmt_pct(safe_mean(df_current.get("Low PPS creators %", pd.Series()))))
    with ck5:
        st.metric("Open Collab %", fmt_pct(safe_mean(df_current.get("Open collaboration product % (last 30d)", pd.Series()))))
    with ck6:
        st.metric("Total Videos", fmt_num(total_videos, 0))

    # ── Creator Engagement Funnel ──
    section_header("🔄", "Creator Engagement Funnel", "Last 30 Days")

    funnel_raw = {
        "Creators Added Products": safe_sum(df_current.get("Creators added products from invitations (last 30d)", pd.Series())),
        "Samples Received": safe_sum(df_current.get("Creators received samples (last 15d)", pd.Series())),
        "Content Posted": safe_sum(df_current.get("Creators posted content", pd.Series())),
        "Product Impressions": safe_sum(df_current.get("Product impressions to creators (last 30d)", pd.Series())),
    }

    if any(v > 0 for v in funnel_raw.values()):
        fc1, fc2 = st.columns([5, 3])
        with fc1:
            # Build real funnel chart using plotly shapes
            funnel_items = list(funnel_raw.items())
            n_stages = len(funnel_items)
            f_colors = [PURPLE, "#6D28D9", PURPLE_MID, BLUE]
            # Inset fractions from each side at each boundary (creates narrowing shape)
            inset_fracs = [0.0, 0.12, 0.24, 0.34, 0.42]

            fig = go.Figure()
            for idx in range(n_stages):
                stage_label, stage_val = funnel_items[idx]
                y_top = (n_stages - idx)
                y_bot = (n_stages - idx - 0.88)
                xl_top = inset_fracs[idx]
                xr_top = 1 - inset_fracs[idx]
                xl_bot = inset_fracs[idx + 1]
                xr_bot = 1 - inset_fracs[idx + 1]

                # Trapezoid trace
                fig.add_trace(go.Scatter(
                    x=[xl_top, xr_top, xr_bot, xl_bot, xl_top],
                    y=[y_top, y_top, y_bot, y_bot, y_top],
                    fill="toself", fillcolor=f_colors[idx],
                    line=dict(color="white", width=2),
                    mode="lines", showlegend=False,
                    hoverinfo="text",
                    hovertext=f"{stage_label}: {fmt_num(stage_val, 0)}",
                ))

                # Value label (center of trapezoid)
                y_mid = (y_top + y_bot) / 2
                fig.add_annotation(
                    x=0.5, y=y_mid + 0.08,
                    text=f"<b>{fmt_num(stage_val, 0)}</b>",
                    showarrow=False,
                    font=dict(color="white", size=20, family="Inter, sans-serif"),
                )
                fig.add_annotation(
                    x=0.5, y=y_mid - 0.15,
                    text=stage_label,
                    showarrow=False,
                    font=dict(color="rgba(255,255,255,0.85)", size=11, family="Inter, sans-serif"),
                )

                # Conversion rate annotation (right side)
                if idx > 0 and funnel_items[idx - 1][1] > 0:
                    conv_rate = stage_val / funnel_items[idx - 1][1] * 100
                    fig.add_annotation(
                        x=1 - inset_fracs[idx] + 0.03,
                        y=(n_stages - idx) + 0.04,
                        text=f"<b>↓ {conv_rate:.1f}%</b>",
                        showarrow=False, xanchor="left",
                        font=dict(color=GRAY_500, size=10, family="Inter, sans-serif"),
                    )

            fig.update_xaxes(visible=False, range=[-0.05, 1.25])
            fig.update_yaxes(visible=False, range=[-0.15, n_stages + 0.3])
            fig.update_layout(
                paper_bgcolor="white", plot_bgcolor="white",
                height=420, margin=dict(t=20, b=20, l=10, r=70),
                font=dict(family="Inter, sans-serif"),
            )
            st.plotly_chart(fig, use_container_width=True)

        with fc2:
            st.markdown("**📊 Stage Conversion**")
            short_labels = ["Added", "Samples", "Content", "Impressions"]
            for i in range(1, n_stages):
                prev_val = funnel_items[i - 1][1]
                curr_val = funnel_items[i][1]
                rate = (curr_val / prev_val * 100) if prev_val > 0 else None
                st.metric(f"{short_labels[i-1]} → {short_labels[i]}", f"{rate:.1f}%" if rate is not None else "—")

            # Overall conversion
            first_v = funnel_items[0][1]
            last_v = funnel_items[-1][1]
            if first_v > 0:
                total_conv = last_v / first_v * 100
                st.markdown(
                    f'<div style="background:{PURPLE_LIGHT};padding:14px 16px;border-radius:10px;text-align:center;margin-top:16px;">'
                    f'<div style="color:{GRAY_500};font-size:0.75rem;text-transform:uppercase;letter-spacing:0.05em;">Overall Conversion</div>'
                    f'<div style="color:{PURPLE};font-size:1.5rem;font-weight:700;margin-top:4px;">{total_conv:.1f}%</div>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
    else:
        st.info("No creator funnel data available for this period.")

    # ── Content Trend ──
    section_header("📈", "Content Trends Over Time")
    if "report date" in df_trend.columns:
        ct_agg = df_trend.groupby("report date").agg({
            col: func for col, func in [
                ("Number of videos", "sum"), ("Number of LIVE streams", "sum"),
                ("Avg. video views", "mean"), ("Avg GPM per video", "mean"),
                ("Avg. GPM per LIVE", "mean"),
                ("Products in collaborations % (last 30d)", "mean"),
                ("Products with samples % (last 30d)", "mean"),
            ] if col in df_trend.columns
        }).reset_index().sort_values("report date")

        ct1, ct2 = st.columns(2)
        with ct1:
            if "Products in collaborations % (last 30d)" in ct_agg.columns:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=ct_agg["report date"], y=ct_agg["Products in collaborations % (last 30d)"],
                    mode="lines+markers", line=dict(color=PURPLE, width=2.5), name="Collab Rate %"))
                if "Products with samples % (last 30d)" in ct_agg.columns:
                    fig.add_trace(go.Scatter(x=ct_agg["report date"], y=ct_agg["Products with samples % (last 30d)"],
                        mode="lines+markers", line=dict(color=BLUE, width=2), name="Sample Rate %"))
                fig.update_layout(title="Collaboration & Sampling Rate Trends",
                    yaxis_title="%", legend=dict(orientation="h", y=1.08, x=0))
                st.plotly_chart(chart_layout(fig), use_container_width=True)

        with ct2:
            if "Avg GPM per video" in ct_agg.columns:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=ct_agg["report date"], y=ct_agg["Avg GPM per video"],
                    mode="lines+markers", line=dict(color=PURPLE, width=2.5), name="GPM/Video"))
                if "Avg. GPM per LIVE" in ct_agg.columns:
                    fig.add_trace(go.Scatter(x=ct_agg["report date"], y=ct_agg["Avg. GPM per LIVE"],
                        mode="lines+markers", line=dict(color=GREEN, width=2), name="GPM/LIVE"))
                fig.update_layout(title="GPM Trends", yaxis_title="$",
                    legend=dict(orientation="h", y=1.08, x=0))
                st.plotly_chart(chart_layout(fig), use_container_width=True)

    # ── GPM by Brand ──
    section_header("💰", "GPM Performance by Brand")
    g1, g2 = st.columns(2)
    with g1:
        if "Avg GPM per video" in df_current.columns:
            gv = df_current[["Brand", "Avg GPM per video"]].dropna().query("`Avg GPM per video` > 0").sort_values("Avg GPM per video", ascending=True).tail(15)
            if not gv.empty:
                fig = px.bar(gv, y="Brand", x="Avg GPM per video", orientation="h",
                    color="Avg GPM per video", color_continuous_scale=[PURPLE_LIGHT, PURPLE])
                fig.update_layout(title="GPM per Video (Top 15)", coloraxis_showscale=False, showlegend=False)
                st.plotly_chart(chart_layout(fig, 400), use_container_width=True)

    with g2:
        if "Avg. GPM per LIVE" in df_current.columns:
            gl = df_current[["Brand", "Avg. GPM per LIVE"]].dropna().query("`Avg. GPM per LIVE` > 0").sort_values("Avg. GPM per LIVE", ascending=True).tail(15)
            if not gl.empty:
                fig = px.bar(gl, y="Brand", x="Avg. GPM per LIVE", orientation="h",
                    color="Avg. GPM per LIVE", color_continuous_scale=[PURPLE_LIGHT, PURPLE])
                fig.update_layout(title="GPM per LIVE (Top 15)", coloraxis_showscale=False, showlegend=False)
                st.plotly_chart(chart_layout(fig, 400), use_container_width=True)

    # ── Video Performance Scatter ──
    section_header("🎯", "Video Volume vs. Views")
    if "Number of videos" in df_current.columns and "Avg. video views" in df_current.columns:
        scatter_df = df_current[["Brand", "Number of videos", "Avg. video views"]].dropna()
        if "Avg GPM per video" in df_current.columns:
            scatter_df = scatter_df.join(df_current["Avg GPM per video"])
        scatter_df = scatter_df[(scatter_df["Number of videos"] > 0) & (scatter_df["Avg. video views"] > 0)]
        if not scatter_df.empty and "Avg GPM per video" in scatter_df.columns:
            fig = px.scatter(scatter_df, x="Number of videos", y="Avg. video views", size="Avg GPM per video",
                color="Avg GPM per video", hover_name="Brand", color_continuous_scale=[GRAY_300, PURPLE],
                size_max=30)
            fig.update_layout(title="Videos × Views × GPM", coloraxis_showscale=False)
            st.plotly_chart(chart_layout(fig, 420), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 4: LIVE
# ═══════════════════════════════════════════════════════════════════════════
with tab_live:
    section_header("📡", "LIVE Performance")

    l1, l2, l3, l4, l5, l6 = st.columns(6)
    with l1:
        st.metric("Total LIVE Streams", fmt_num(total_live, 0))
    with l2:
        v = safe_mean(df_current.get("Avg. LIVE duration (hr)", pd.Series()))
        st.metric("Avg Duration", f"{fmt_num(v)} hr" if v else "—")
    with l3:
        st.metric("Avg Impressions/hr", fmt_num(safe_mean(df_current.get("Avg. LIVE impressions per hour", pd.Series()))))
    with l4:
        st.metric("Avg GPM/LIVE", fmt_dollar(safe_mean(df_current.get("Avg. GPM per LIVE", pd.Series()))))
    with l5:
        st.metric("Engagement Tool %", fmt_pct(safe_mean(df_current.get("LIVE with engagement tools %", pd.Series()))))
    with l6:
        st.metric("Promotion %", fmt_pct(safe_mean(df_current.get("LIVE with promotions %", pd.Series()))))

    # ── LIVE GMV Section ──
    section_header("💰", "LIVE GMV")
    lg1, lg2, lg3 = st.columns(3)
    with lg1:
        v = safe_sum(df_current.get("LIVE GMV", pd.Series()))
        st.metric("Total LIVE GMV", fmt_dollar(v))
    with lg2:
        v = safe_sum(df_current.get("Video GMV", pd.Series()))
        st.metric("Total Video GMV", fmt_dollar(v))
    with lg3:
        if "LIVE GMV" in df_current.columns and "GMV" in df_current.columns:
            live_gmv = safe_sum(df_current["LIVE GMV"])
            total = safe_sum(df_current["GMV"])
            pct = (live_gmv / total * 100) if total > 0 else 0
            st.metric("LIVE GMV Share", f"{pct:.1f}%")

    # LIVE scatter: Duration vs GPM
    section_header("📊", "LIVE Duration vs. GPM")
    if "Avg. LIVE duration (hr)" in df_current.columns and "Avg. GPM per LIVE" in df_current.columns:
        live_scatter = df_current[["Brand", "Avg. LIVE duration (hr)", "Avg. GPM per LIVE", "Number of LIVE streams"]].dropna()
        live_scatter = live_scatter[(live_scatter["Avg. GPM per LIVE"] > 0) & (live_scatter["Number of LIVE streams"] > 0)]
        if not live_scatter.empty:
            fig = px.scatter(live_scatter, x="Avg. LIVE duration (hr)", y="Avg. GPM per LIVE",
                size="Number of LIVE streams", hover_name="Brand",
                color="Number of LIVE streams", color_continuous_scale=[PURPLE_LIGHT, PURPLE], size_max=30)
            fig.update_layout(title="Duration vs GPM (bubble = stream count)", coloraxis_showscale=False)
            st.plotly_chart(chart_layout(fig, 420), use_container_width=True)

    # ── LIVE Trends ──
    section_header("📈", "LIVE Trends Over Time")
    if "report date" in df_trend.columns:
        live_trend = df_trend.groupby("report date").agg({
            col: func for col, func in [
                ("Number of LIVE streams", "sum"),
                ("Avg. LIVE duration (hr)", "mean"),
                ("Avg. GPM per LIVE", "mean"),
                ("Avg. LIVE impressions per hour", "mean"),
                ("LIVE with engagement tools %", "mean"),
                ("LIVE with promotions %", "mean"),
                ("LIVE GMV", "sum"),
            ] if col in df_trend.columns
        }).reset_index().sort_values("report date")

        lt1, lt2 = st.columns(2)
        with lt1:
            if "Number of LIVE streams" in live_trend.columns:
                fig = make_subplots(specs=[[{"secondary_y": True}]])
                fig.add_trace(go.Bar(x=live_trend["report date"], y=live_trend["Number of LIVE streams"],
                    name="Streams", marker_color=PURPLE, opacity=0.8), secondary_y=False)
                if "LIVE GMV" in live_trend.columns:
                    fig.add_trace(go.Scatter(x=live_trend["report date"], y=live_trend["LIVE GMV"],
                        name="LIVE GMV", line=dict(color=GREEN, width=2), mode="lines+markers"), secondary_y=True)
                fig.update_layout(title="LIVE Streams & GMV Trend",
                    legend=dict(orientation="h", y=1.08, x=0))
                fig.update_yaxes(title_text="Streams", secondary_y=False)
                fig.update_yaxes(title_text="GMV ($)", secondary_y=True)
                st.plotly_chart(chart_layout(fig), use_container_width=True)

        with lt2:
            if "LIVE with engagement tools %" in live_trend.columns:
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=live_trend["report date"], y=live_trend["LIVE with engagement tools %"],
                    mode="lines+markers", line=dict(color=PURPLE, width=2.5), name="Engagement Tool %"))
                if "LIVE with promotions %" in live_trend.columns:
                    fig.add_trace(go.Scatter(x=live_trend["report date"], y=live_trend["LIVE with promotions %"],
                        mode="lines+markers", line=dict(color=AMBER, width=2), name="Promotions %"))
                fig.update_layout(title="LIVE Engagement & Promotion Adoption",
                    yaxis_title="%", legend=dict(orientation="h", y=1.08, x=0))
                st.plotly_chart(chart_layout(fig), use_container_width=True)

    # LIVE engagement treemap
    section_header("🌳", "LIVE Engagement Breakdown")
    if "LIVE with engagement tools %" in df_current.columns and "Number of LIVE streams" in df_current.columns:
        tree_df = df_current[["Brand", "LIVE with engagement tools %", "LIVE with promotions %", "Number of LIVE streams"]].dropna()
        tree_df = tree_df[tree_df["Number of LIVE streams"] > 0]
        if not tree_df.empty:
            fig = px.treemap(tree_df, path=["Brand"], values="Number of LIVE streams",
                color="LIVE with engagement tools %", color_continuous_scale=[GRAY_100, PURPLE],
                hover_data=["LIVE with promotions %"])
            fig.update_layout(title="LIVE Streams by Brand (color = engagement tool %)")
            st.plotly_chart(chart_layout(fig, 450), use_container_width=True)

    # ── LIVE Violations ──
    if "LIVE streams with violations" in df_current.columns:
        violations_df = df_current[["Brand", "LIVE streams with violations"]].dropna()
        violations_df = violations_df[violations_df["LIVE streams with violations"] > 0]
        if not violations_df.empty:
            section_header("⚠️", "LIVE Violations by Brand")
            fig = px.bar(violations_df.sort_values("LIVE streams with violations", ascending=True),
                y="Brand", x="LIVE streams with violations", orientation="h",
                color_discrete_sequence=[RED])
            fig.update_layout(title="Brands with LIVE Violations")
            st.plotly_chart(chart_layout(fig, max(250, len(violations_df) * 25)), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 5: ACTION ITEMS
# ═══════════════════════════════════════════════════════════════════════════
with tab_actions:
    section_header("🚨", "Action Items & Flags")
    alerts = []

    # SPS alerts
    if "shop performance score" in df_current.columns:
        for _, r in df_current[df_current["shop performance score"].notna() & (df_current["shop performance score"] < 3.5)][["Brand", "shop performance score"]].iterrows():
            alerts.append(("high", r["Brand"], f"SPS below 3.5 → currently {r['shop performance score']:.2f}", "Shop Health"))

    # Violation alerts
    if "Recent shop violations" in df_current.columns:
        for _, r in df_current[df_current["Recent shop violations"].notna() & (df_current["Recent shop violations"] > 0)][["Brand", "Recent shop violations"]].iterrows():
            alerts.append(("high", r["Brand"], f"{int(r['Recent shop violations'])} shop violation(s)", "Compliance"))

    if "Videos with new violations" in df_current.columns:
        for _, r in df_current[df_current["Videos with new violations"].notna() & (df_current["Videos with new violations"] > 0)][["Brand", "Videos with new violations"]].iterrows():
            alerts.append(("high", r["Brand"], f"{int(r['Videos with new violations'])} video violation(s)", "Compliance"))

    if "LIVE streams with violations" in df_current.columns:
        for _, r in df_current[df_current["LIVE streams with violations"].notna() & (df_current["LIVE streams with violations"] > 0)][["Brand", "LIVE streams with violations"]].iterrows():
            alerts.append(("high", r["Brand"], f"{int(r['LIVE streams with violations'])} LIVE violation(s)", "Compliance"))

    # Low AHR
    if "Days with low AHR (<200)" in df_current.columns:
        for _, r in df_current[df_current["Days with low AHR (<200)"].notna() & (df_current["Days with low AHR (<200)"] > 5)][["Brand", "Days with low AHR (<200)"]].iterrows():
            alerts.append(("high", r["Brand"], f"{int(r['Days with low AHR (<200)'])} days with AHR < 200", "Shop Health"))

    # NRR alerts (now in percentage form, so >5% is concerning)
    if "60-Day Negative Review Rate (NRR)" in df_current.columns:
        for _, r in df_current[df_current["60-Day Negative Review Rate (NRR)"].notna() & (df_current["60-Day Negative Review Rate (NRR)"] > 5)][["Brand", "60-Day Negative Review Rate (NRR)"]].iterrows():
            alerts.append(("warn", r["Brand"], f"NRR at {r['60-Day Negative Review Rate (NRR)']:.1f}% (>5%)", "Quality"))

    # OOS
    if "Out-of-stock products" in df_current.columns:
        for _, r in df_current[df_current["Out-of-stock products"].notna() & (df_current["Out-of-stock products"] > 10)][["Brand", "Out-of-stock products"]].iterrows():
            alerts.append(("warn", r["Brand"], f"{int(r['Out-of-stock products'])} products OOS", "Inventory"))

    # Low qualified pages (now in percentage form)
    if "Qualified product page %" in df_current.columns:
        for _, r in df_current[df_current["Qualified product page %"].notna() & (df_current["Qualified product page %"] < 50)][["Brand", "Qualified product page %"]].iterrows():
            alerts.append(("warn", r["Brand"], f"Only {r['Qualified product page %']:.0f}% qualified pages", "Content Quality"))

    # Low brand auth
    if "Products with brand authorization %" in df_current.columns:
        for _, r in df_current[df_current["Products with brand authorization %"].notna() & (df_current["Products with brand authorization %"] < 50)][["Brand", "Products with brand authorization %"]].iterrows():
            alerts.append(("warn", r["Brand"], f"Only {r['Products with brand authorization %']:.0f}% brand authorized", "Compliance"))

    # Low listing quality
    if "# of lisitng quality not good" in df_current.columns:
        for _, r in df_current[df_current["# of lisitng quality not good"].notna() & (df_current["# of lisitng quality not good"] > 0)][["Brand", "# of lisitng quality not good"]].iterrows():
            alerts.append(("warn", r["Brand"], f"{int(r['# of lisitng quality not good'])} low-quality listings", "Content Quality"))

    high = [a for a in alerts if a[0] == "high"]
    warn = [a for a in alerts if a[0] == "warn"]

    a1, a2 = st.columns(2)
    with a1:
        st.markdown(f"**🔴 Critical ({len(high)})**")
        if high:
            for sev, brand, msg, area in high[:25]:
                st.markdown(f'<div class="tv-alert"><strong>{brand}</strong> · {area}<br/>{msg}</div>', unsafe_allow_html=True)
        else:
            st.success("No critical issues!")
    with a2:
        st.markdown(f"**🟡 Warnings ({len(warn)})**")
        if warn:
            for sev, brand, msg, area in warn[:25]:
                st.markdown(f'<div class="tv-alert tv-alert-warn"><strong>{brand}</strong> · {area}<br/>{msg}</div>', unsafe_allow_html=True)
        else:
            st.success("No warnings!")

    # Alert summary chart
    if alerts:
        section_header("📊", "Issue Breakdown")
        area_counts = pd.DataFrame(alerts, columns=["sev", "brand", "msg", "area"])
        area_summary = area_counts.groupby(["area", "sev"]).size().reset_index(name="count")
        fig = px.bar(area_summary, x="area", y="count", color="sev",
            color_discrete_map={"high": RED, "warn": AMBER},
            labels={"area": "Category", "count": "Issues", "sev": "Severity"})
        fig.update_layout(title="Issues by Category", barmode="stack")
        st.plotly_chart(chart_layout(fig, 320), use_container_width=True)


# ═══════════════════════════════════════════════════════════════════════════
# TAB 6: DATA EXPORT
# ═══════════════════════════════════════════════════════════════════════════
with tab_data:
    section_header("📥", "Filtered Data", f"{len(df_current)} rows")

    if "Brand" in df_current.columns:
        brand_list = sorted([str(x) for x in df_current["Brand"].dropna().unique().tolist()])
        sel_brand = st.selectbox("Quick brand lookup", ["— All —"] + brand_list)
        if sel_brand != "— All —":
            st.dataframe(df_current[df_current["Brand"] == sel_brand], use_container_width=True, hide_index=True)
        else:
            st.dataframe(df_current, use_container_width=True, height=500, hide_index=True)
    else:
        st.dataframe(df_current, use_container_width=True, height=500, hide_index=True)

    csv = df_current.to_csv(index=False)
    st.download_button("📥 Download CSV", csv,
        file_name=f"broadway_{selected_am.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.csv", mime="text/csv")
