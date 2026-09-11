import io
from pathlib import Path
from datetime import datetime
from urllib.parse import urlparse, parse_qsl, urlencode, urlunparse

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st


# ============================================================
# PAGE / CONFIG
# ============================================================

st.set_page_config(
    page_title="EGAT Fault Dashboard",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

PLOTLY_CONFIG = {
    "displayModeBar": False,
    "responsive": True,
    "scrollZoom": False,
}

EXCEL_SHAREPOINT_URL = (
    "https://egatucc-my.sharepoint.com/:x:/g/personal/599320_egat_co_th/"
    "IQCa3BYjVfj4RqpmHfP8_VEEAWZ5i0luFrsoelLzPAT46Nw?e=yyLIJx"
)
SHEET_NAME = "Logbook"
HEADER_ROW = 1

SNAPSHOT_PATH = Path(__file__).parent / "data" / "fault_snapshot.csv.gz"
SNAPSHOT_META_PATH = Path(__file__).parent / "data" / "fault_snapshot.meta.txt"


BAR_COLORS = [
    "#2478F3",  # blue
    "#FF7A18",  # orange
    "#19C36A",  # green
    "#A855F7",  # purple
    "#FFB81C",  # amber
    "#20C4E8",  # cyan
    "#EF4B55",  # red
    "#8BD450",  # lime
    "#F04BB3",  # pink
    "#7B8FA6",  # slate
]

def apply_bar_colors(fig, colors=None):
    """Give each bar/category a distinct readable color."""
    colors = colors or BAR_COLORS
    for trace in fig.data:
        n = len(trace.x) if trace.orientation != "h" else len(trace.y)
        trace.update(marker_color=[colors[i % len(colors)] for i in range(n)])
    return fig


THAI_MONTHS = [
    "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
    "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."
]


@st.cache_resource
def get_http_session():
    """Reuse HTTP connections while the Streamlit worker is alive."""
    session = requests.Session()
    session.headers.update({"User-Agent": "Mozilla/5.0 EGAT-Fault-Dashboard/8.0"})
    return session


# ============================================================
# THEME
# ============================================================

st.markdown("""
<style>
:root {
  --bg: #06111d;
  --panel: #081a2a;
  --panel2: #0a2033;
  --border: #1d4c6a;
  --text: #f4f7fb;
  --muted: #9ab0c4;
}
html, body, [class*="css"] { font-family: "Arial", "Noto Sans Thai", sans-serif; }
.stApp {
  background:
    radial-gradient(circle at 78% 0%, rgba(0,96,170,.12), transparent 26%),
    linear-gradient(135deg, #06111d 0%, #071523 55%, #04101a 100%);
}
.block-container {
  max-width: 1720px;
  padding-top: 3.4rem;
  padding-bottom: 2rem;
}
[data-testid="stSidebar"] {
  background: linear-gradient(180deg,#071827,#081b2b);
  border-right: 1px solid #1a4058;
}
[data-testid="stSidebar"] .block-container { padding-top: 1.4rem; }
[data-testid="stMetric"] {
  background: linear-gradient(135deg,#0a2238,#081725);
  border: 1px solid #215a7f;
  border-radius: 14px;
  padding: 14px 16px;
  min-height: 118px;
}
[data-testid="stMetricLabel"] { color:#dbe8f3; font-weight:700; }
[data-testid="stMetricValue"] { color:#fff; font-weight:800; }
div[data-testid="stPlotlyChart"] {
  background: rgba(6,22,36,.66);
  border: 1px solid #183f59;
  border-radius: 14px;
  padding: 6px 8px;
  overflow: hidden;
}
header[data-testid="stHeader"] {
  background: rgba(5,14,23,.96);
}
[data-testid="stToolbar"] {
  top: .35rem;
}
.stTabs [data-baseweb="tab-list"] {
  gap: 4px;
  border-bottom: 1px solid #173d57;
}
.stTabs [data-baseweb="tab"] {
  color:#dbe6ef;
  background:transparent;
  border-radius:7px 7px 0 0;
  padding:8px 12px;
}
.stTabs [aria-selected="true"] {
  background:#1677ff !important;
  color:white !important;
}
.hero {
  border:1px solid #1d6697;
  background:linear-gradient(135deg,#09223c,#07182a);
  border-radius:15px;
  padding:15px 18px 13px;
  margin:0 0 10px;
}
.hero-title {font-size:29px;font-weight:900;color:white;margin:0}
.hero-sub {font-size:12px;color:#9fb5c8;margin-top:7px}
.kpi {
  border-radius:14px;
  min-height:114px;
  padding:15px 16px 11px;
  border:1px solid;
  background:linear-gradient(135deg,#0c2948,#0b1a2a);
}
.kpi.blue {border-color:#1687ff;background:linear-gradient(135deg,#0d3764,#0a2038)}
.kpi.gold {border-color:#9d7900;background:linear-gradient(135deg,#3b3211,#17190d)}
.kpi.green {border-color:#119b59;background:linear-gradient(135deg,#06482f,#08251d)}
.kpi.purple {border-color:#9148a8;background:linear-gradient(135deg,#3a1b4b,#171324)}
.kpi.cyan {border-color:#1184a5;background:linear-gradient(135deg,#07364a,#08202d)}
.kpi-label{font-size:14px;font-weight:800;color:#f5f7fa}
.kpi-value{font-size:31px;font-weight:900;color:white;margin-top:9px}
.kpi-foot{font-size:12px;font-weight:700;margin-top:7px}
.blue .kpi-foot{color:#55a9ff}.gold .kpi-foot{color:#ffd438}
.green .kpi-foot{color:#34e58a}.purple .kpi-foot{color:#d77cff}.cyan .kpi-foot{color:#39c9e8}
.summary-card {
  border:1px solid #214864; background:#091d2d; border-radius:13px;
  min-height:105px; padding:14px 16px;
}
.summary-label {font-size:13px;color:#98aec0;font-weight:700}
.summary-value {font-size:20px;color:#fff;font-weight:900;margin-top:8px}
.small-note {font-size:12px;color:#91a8bb}
hr {border-color:#173a52}

[data-testid="stSidebar"] .stSelectbox,
[data-testid="stSidebar"] .stTextInput,
[data-testid="stSidebar"] .stFileUploader {
  margin-bottom: .35rem;
}
[data-testid="stSidebar"] label {
  color:#d8e5ef !important;
  font-weight:700 !important;
}
[data-testid="stSidebar"] button[kind="primary"] {
  background: linear-gradient(90deg,#1677ff,#2c8cff);
  border: 0;
  font-weight: 800;
}
[data-testid="stSidebar"] button[kind="secondary"] {
  border-color:#28506c;
}


@media (max-width: 1200px){
  .kpi-value{font-size:26px}
  .hero-title{font-size:24px}
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# DATA HELPERS
# ============================================================

def make_download_url(url: str) -> str:
    p = urlparse(url)
    q = dict(parse_qsl(p.query, keep_blank_values=True))
    q["download"] = "1"
    return urlunparse(p._replace(query=urlencode(q)))


@st.cache_data(ttl=1800, show_spinner=False)
def download_sharepoint(url: str) -> bytes:
    r = get_http_session().get(
        make_download_url(url),
        timeout=45,
        allow_redirects=True,
    )
    r.raise_for_status()
    b = r.content
    ct = (r.headers.get("content-type") or "").lower()
    head = b[:500].lstrip().lower()
    if "text/html" in ct or head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        raise RuntimeError("SharePoint ต้อง Sign in จึงไม่สามารถดาวน์โหลดไฟล์ผ่าน public sharing link ได้")
    return b


@st.cache_data(ttl=1800, show_spinner=False)
def read_excel_bytes(b: bytes) -> pd.DataFrame:
    return pd.read_excel(
        io.BytesIO(b),
        sheet_name=SHEET_NAME,
        header=HEADER_ROW,
        engine="openpyxl",
    )


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    required = ["วันที่", "Voltage (kV)", "สายส่ง", "Trip type", "เหตุการณ์"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("ไม่พบคอลัมน์: " + ", ".join(missing))
    df["วันที่"] = pd.to_datetime(df["วันที่"], errors="coerce")
    df["year_calc"] = df["วันที่"].dt.year
    df["month_calc"] = df["วันที่"].dt.month
    return df


@st.cache_data(ttl=1800, show_spinner=False)
def load_sharepoint_dataframe(url: str) -> pd.DataFrame:
    """Download, parse and prepare SharePoint Excel once, then reuse it."""
    raw = download_sharepoint(url)
    return prepare(read_excel_bytes(raw))


@st.cache_data(ttl=1800, show_spinner=False)
def load_snapshot(path_str: str) -> pd.DataFrame:
    path = Path(path_str)
    df = pd.read_csv(path, compression="gzip", low_memory=False)
    if "วันที่" in df.columns:
        df["วันที่"] = pd.to_datetime(df["วันที่"], errors="coerce")
    if "year_calc" not in df.columns:
        df["year_calc"] = df["วันที่"].dt.year
    if "month_calc" not in df.columns:
        df["month_calc"] = df["วันที่"].dt.month
    return df


def save_runtime_snapshot(df: pd.DataFrame) -> None:
    try:
        SNAPSHOT_PATH.parent.mkdir(parents=True, exist_ok=True)
        df.to_csv(SNAPSHOT_PATH, index=False, compression="gzip", encoding="utf-8-sig")
        SNAPSHOT_META_PATH.write_text(datetime.now().isoformat(timespec="seconds"), encoding="utf-8")
    except Exception:
        pass


def text_options(df, col):
    if col not in df.columns:
        return []
    s = df[col].dropna().astype(str).str.strip()
    return sorted(s[s.ne("")].unique().tolist())


@st.cache_data(ttl=1800, show_spinner=False)
def annual_counts(df_all: pd.DataFrame) -> pd.DataFrame:
    out = (
        df_all.dropna(subset=["year_calc"])
        .assign(year_calc=lambda x: x["year_calc"].astype(int))
        .groupby("year_calc").size()
        .sort_index()
        .reset_index(name="Fault")
    )
    out["year_calc"] = out["year_calc"].astype(str)
    return out


@st.cache_data(ttl=1800, show_spinner=False)
def cumulative_counts(df_all: pd.DataFrame, selected_year: int) -> pd.DataFrame:
    compare_years = [selected_year, selected_year - 1, selected_year - 2]
    rows = []
    now = datetime.now()
    for y in compare_years:
        yd = df_all[df_all["year_calc"] == y]
        cnt = (
            yd.dropna(subset=["month_calc"])
            .assign(month_calc=lambda x: x["month_calc"].astype(int))
            .groupby("month_calc").size()
            .reindex(range(1, 13), fill_value=0)
        )
        cum = cnt.cumsum()
        max_m = now.month if y == now.year else 12
        for m in range(1, max_m + 1):
            rows.append({
                "เดือน": THAI_MONTHS[m - 1],
                "ปี": str(y),
                "Fault สะสม": int(cum.loc[m]),
            })
    return pd.DataFrame(rows)


def style_fig(fig, height=330, legend_top=True):
    fig.update_layout(
        template="plotly_dark",
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=42, r=22, t=55, b=42),
        font=dict(color="#eef5fb", size=12),
        hoverlabel=dict(bgcolor="#0b2236"),
    )
    if legend_top:
        fig.update_layout(legend=dict(
            orientation="h", yanchor="bottom", y=1.02,
            xanchor="right", x=1
        ))
    fig.update_xaxes(gridcolor="rgba(113,148,174,.22)", zeroline=False)
    fig.update_yaxes(gridcolor="rgba(113,148,174,.22)", zeroline=False)
    return fig


def donut(values, names, title, center_text):
    fig = go.Figure(go.Pie(
        labels=names,
        values=values,
        hole=.58,
        marker=dict(colors=BAR_COLORS),
        textinfo="none",
        hovertemplate="%{label}: %{value} ครั้ง<extra></extra>",
    ))
    fig.add_annotation(
        text=center_text,
        x=.5, y=.5, showarrow=False,
        font=dict(size=22, color="white")
    )
    fig.update_layout(title=title, showlegend=True)
    return style_fig(fig, 340, legend_top=False)


def kpi_card(label, value, foot, cls):
    st.markdown(
        f'<div class="kpi {cls}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'<div class="kpi-foot">{foot}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ============================================================
# LOAD
# ============================================================

with st.sidebar:
    st.markdown("## 🔎 ตัวกรองข้อมูล")

    source_mode = st.radio(
        "แหล่งข้อมูล",
        ["SharePoint link", "Upload Excel"],
        horizontal=False,
        label_visibility="collapsed",
    )

    uploaded = None
    if source_mode == "Upload Excel":
        uploaded = st.file_uploader("เลือก Excel", type=["xlsx", "xlsm"])
    else:
        share_url = EXCEL_SHAREPOINT_URL
        if st.button("↻ Refresh Excel", use_container_width=True):
            st.cache_data.clear()
            st.session_state["force_live_refresh"] = True
            st.rerun()

try:
    if source_mode == "SharePoint link":
        force_live = bool(st.session_state.pop("force_live_refresh", False))
        if SNAPSHOT_PATH.exists() and not force_live:
            df_all = load_snapshot(str(SNAPSHOT_PATH))
            source_text = "Local snapshot (fast)"
        else:
            with st.spinner("กำลังเชื่อมต่อ SharePoint และเตรียม Dashboard..."):
                df_all = load_sharepoint_dataframe(share_url)
            save_runtime_snapshot(df_all)
            source_text = "Excel SharePoint (live)"
    else:
        if uploaded is None:
            st.info("กรุณาเลือกไฟล์ Excel")
            st.stop()
        df_all = pd.read_excel(uploaded, sheet_name=SHEET_NAME, header=HEADER_ROW, engine="openpyxl")
        df_all = prepare(df_all)
        source_text = uploaded.name
except Exception as e:
    st.error("อ่าน Excel ไม่สำเร็จ")
    st.code(str(e))
    st.info("ถ้า SharePoint บังคับ Sign in ให้ใช้ Upload Excel ทดสอบก่อน")
    st.stop()


# ============================================================
# FILTERS
# ============================================================

years = sorted(df_all["year_calc"].dropna().astype(int).unique().tolist(), reverse=True)
default_year = datetime.now().year if datetime.now().year in years else years[0]

with st.sidebar:
    selected_year = st.selectbox("ปี (ค.ศ.)", years, index=years.index(default_year))

    c1, c2 = st.columns(2)
    with c1:
        m_from = st.selectbox("ช่วงเดือน", ["ทั้งหมด"] + list(range(1,13)), index=0)
    with c2:
        m_to = st.selectbox("สิ้นสุด", ["ทั้งหมด"] + list(range(1,13)), index=0)

    voltages = sorted(
        pd.to_numeric(df_all["Voltage (kV)"], errors="coerce")
        .dropna().astype(int).unique().tolist()
    )
    selected_voltage = st.selectbox("ระดับแรงดัน (kV)", ["ทั้งหมด"] + voltages)

    line_search = st.text_input("สายส่ง", placeholder="เช่น CC-PA1#2")

    cause_options = text_options(df_all, "เหตุการณ์")
    selected_cause = st.selectbox(
        "สาเหตุ",
        ["ทั้งหมด"] + cause_options,
        index=0,
    )

    trip_options = text_options(df_all, "Trip type")
    selected_trip = st.selectbox("Trip type", ["ทั้งหมด"] + trip_options)

    st.button("🔎 ค้นหาข้อมูล", use_container_width=True, type="primary")

    if st.button("↻ รีเซ็ตตัวกรอง", use_container_width=True):
        st.session_state.clear()
        st.rerun()

df = df_all[df_all["year_calc"] == selected_year].copy()

if m_from != "ทั้งหมด":
    df = df[df["month_calc"] >= int(m_from)]
if m_to != "ทั้งหมด":
    df = df[df["month_calc"] <= int(m_to)]
if selected_voltage != "ทั้งหมด":
    v = pd.to_numeric(df["Voltage (kV)"], errors="coerce")
    df = df[v == int(selected_voltage)]
if line_search.strip():
    df = df[df["สายส่ง"].astype(str).str.contains(line_search.strip(), case=False, na=False, regex=False)]
if selected_cause != "ทั้งหมด":
    df = df[df["เหตุการณ์"].fillna("").astype(str).str.strip() == selected_cause]
if selected_trip != "ทั้งหมด":
    df = df[df["Trip type"].astype(str) == selected_trip]

with st.sidebar:
    st.divider()
    st.markdown("### สรุปเงื่อนไขที่เลือก")
    st.markdown(
        f"""
<div class="small-note">
ปี: {selected_year}<br>
เดือน: {m_from} → {m_to}<br>
แรงดัน: {selected_voltage}<br>
สายส่ง: {line_search or "ทั้งหมด"}<br>
สาเหตุ: {selected_cause}<br>
Trip type: {selected_trip}
</div>
""",
        unsafe_allow_html=True,
    )


# ============================================================
# HEADER
# ============================================================

st.markdown(
    f"""
<div class="hero">
  <div class="hero-title">⚡ Dashboard สรุปสถิติไฟฟ้าขัดข้อง ในแผนก ทสก2-ส.</div>
  <div class="hero-sub">
    Source of Truth: {source_text} / Sheet {SHEET_NAME}
    • อัปเดตล่าสุด {datetime.now().strftime("%d/%m/%Y %H:%M")}
    • ไม่มี AI • V9 Hybrid Fast Load
  </div>
</div>
""",
    unsafe_allow_html=True,
)

view = st.segmented_control(
    "Dashboard section",
    options=["▣ Overview", "↻ Recurring & Lockout", "▤ Event Log"],
    default="▣ Overview",
    label_visibility="collapsed",
)

if df.empty:
    st.warning("ไม่พบข้อมูลตามเงื่อนไขที่เลือก")
    st.stop()


# ============================================================
# COMMON FACTS
# ============================================================

total = len(df)
trip_norm = df["Trip type"].fillna("").astype(str).str.strip().str.casefold()
lock_df = df[trip_norm.eq("lockout")].copy()
lockout = int(len(lock_df))
reclose = int(trip_norm.eq("reclose").sum())

blackout = 0
if "Black out" in df.columns:
    blackout = int(
        df["Black out"].fillna("").astype(str).str.strip().str.upper()
        .isin(["Y","YES","TRUE","1"]).sum()
    )

affected_lines = int(
    df["สายส่ง"].dropna().astype(str).str.strip()
    .replace("", np.nan).dropna().nunique()
)


# ============================================================
# OVERVIEW
# ============================================================

if view == "▣ Overview":

    cols = st.columns(5)
    vals = [
        ("⚡ Fault ทั้งหมด", f"{total} ครั้ง", "100% ของข้อมูลที่กรอง", "blue"),
        ("🔒 Lockout", f"{lockout} ครั้ง", f"{lockout/total*100:.0f}% ของทั้งหมด", "gold"),
        ("♻ Reclose", f"{reclose} ครั้ง", f"{reclose/total*100:.0f}% ของทั้งหมด", "green"),
        ("⏻ Black out", f"{blackout} ครั้ง", f"{blackout/total*100:.0f}% ของทั้งหมด", "purple"),
        ("🗼 สายส่งที่ได้รับผลกระทบ", f"{affected_lines} Line", "จำนวนสายส่งที่พบเหตุ", "cyan"),
    ]
    for col, item in zip(cols, vals):
        with col:
            kpi_card(*item)

    st.write("")

    # Row 1: cause / monthly / voltage
    c1, c2, c3 = st.columns([1, 1.08, 1])

    cause_counts = (
        df["เหตุการณ์"].fillna("ไม่ทราบสาเหตุ").replace("", "ไม่ทราบสาเหตุ")
        .value_counts()
    )
    with c1:
        st.plotly_chart(
            donut(
                cause_counts.values,
                cause_counts.index,
                "1) สาเหตุไฟฟ้าขัดข้อง",
                f"<b>{total}</b><br><span style='font-size:12px'>ครั้ง</span>",
            ),
            use_container_width=True, config=PLOTLY_CONFIG,
        )

    monthly = (
        df.dropna(subset=["month_calc"])
        .assign(month_calc=lambda x: x["month_calc"].astype(int))
        .groupby("month_calc").size()
        .reindex(range(1,13), fill_value=0)
    )
    display_months = 12
    if selected_year == datetime.now().year:
        display_months = datetime.now().month
    month_df = pd.DataFrame({
        "เดือน": THAI_MONTHS[:display_months],
        "Fault": [int(monthly.get(m,0)) for m in range(1,display_months+1)]
    })
    with c2:
        fig = px.bar(month_df, x="เดือน", y="Fault", text="Fault",
                     title=f"2) แนวโน้มรายเดือน (ปี {selected_year})")
        month_colors = [
            BAR_COLORS[i % len(BAR_COLORS)] for i in range(len(month_df))
        ]
        if selected_year == datetime.now().year and len(month_colors) >= datetime.now().month:
            month_colors[datetime.now().month - 1] = "#FFB81C"
        fig.update_traces(
            textposition="outside",
            marker_color=month_colors,
            cliponaxis=False,
        )
        fig.update_layout(xaxis_title="เดือน", yaxis_title="จำนวนครั้ง")
        st.plotly_chart(style_fig(fig,360), use_container_width=True, config=PLOTLY_CONFIG)

    volt_counts = (
        pd.to_numeric(df["Voltage (kV)"], errors="coerce")
        .dropna().astype(int).value_counts().sort_index()
    )
    with c3:
        names = [f"{x} kV" for x in volt_counts.index]
        st.plotly_chart(
            donut(
                volt_counts.values,
                names,
                "3) แยกตามระดับแรงดัน (kV)",
                f"<b>{total}</b><br><span style='font-size:12px'>ครั้ง</span>",
            ),
            use_container_width=True, config=PLOTLY_CONFIG,
        )

    # Row 2: TMU / annual / top line
    c4, c5, c6 = st.columns([1,1,1.15])

    tmu_counts = (
        df["TMU"].dropna().astype(str).str.strip().replace("", np.nan)
        .dropna().value_counts()
    ) if "TMU" in df.columns else pd.Series(dtype=int)

    with c4:
        if len(tmu_counts):
            tmu_df = tmu_counts.reset_index()
            tmu_df.columns = ["TMU","Fault"]
            fig = px.bar(tmu_df, x="Fault", y="TMU", orientation="h",
                         text="Fault", title="4) แยกตาม TMU")
            fig.update_layout(
                yaxis={"categoryorder":"total ascending"},
                xaxis_title="จำนวนครั้ง",
                yaxis_title="TMU",
            )
            apply_bar_colors(fig)
            fig.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(style_fig(fig,360), use_container_width=True, config=PLOTLY_CONFIG)
        else:
            st.info("ไม่มีข้อมูล TMU")

    annual = annual_counts(df_all)

    with c5:
        fig = px.bar(annual, x="year_calc", y="Fault", text="Fault",
                     title="5) เปรียบเทียบจำนวน Fault รายปี")
        annual_colors = []
        past_palette = ["#2478F3", "#20C4E8", "#19C36A", "#A855F7", "#FFB81C"]
        pi = 0
        for y in annual["year_calc"]:
            if str(y) == str(selected_year):
                annual_colors.append("#FF7A18")
            else:
                annual_colors.append(past_palette[pi % len(past_palette)])
                pi += 1
        fig.update_traces(
            textposition="outside",
            marker_color=annual_colors,
            cliponaxis=False,
        )
        fig.update_layout(xaxis_title=None, yaxis_title=None)
        st.plotly_chart(style_fig(fig,360), use_container_width=True, config=PLOTLY_CONFIG)

    top_lines = (
        df["สายส่ง"].dropna().astype(str).str.strip().replace("",np.nan)
        .dropna().value_counts().head(10)
        .reset_index()
    )
    top_lines.columns = ["Line","Fault"]

    with c6:
        fig = px.bar(top_lines, x="Fault", y="Line", orientation="h",
                     text="Fault", title="6) สายส่งที่พบ Fault (Top 10)")
        rank_palette = [
            "#FF7A18", "#FFB81C", "#2478F3", "#20C4E8", "#19C36A",
            "#A855F7", "#EF4B55", "#8BD450", "#F04BB3", "#7B8FA6"
        ]
        top_colors = [rank_palette[i % len(rank_palette)] for i in range(len(top_lines))]
        fig.update_traces(
            marker_color=top_colors,
            textposition="outside",
            cliponaxis=False
        )
        fig.update_layout(
            yaxis={"categoryorder":"total ascending"},
            xaxis_title="จำนวนครั้ง",
            yaxis_title=None,
            margin=dict(l=70, r=55, t=55, b=42),
        )
        st.plotly_chart(style_fig(fig,360), use_container_width=True, config=PLOTLY_CONFIG)

    # Section 7
    compare_years = [selected_year, selected_year-1, selected_year-2]
    cum_df = cumulative_counts(df_all, selected_year)
    fig = px.line(
        cum_df, x="เดือน", y="Fault สะสม", color="ปี", markers=True,
        category_orders={"เดือน":THAI_MONTHS, "ปี":[str(y) for y in compare_years]},
        title=f"7) Fault สะสมรายเดือน — เปรียบเทียบ {compare_years[0]}, {compare_years[1]}, {compare_years[2]}",
    )

    color_map = {
        str(compare_years[0]): "#FFE600",
        str(compare_years[1]): "#00E5EE",
        str(compare_years[2]): "#FF00D4",
    }
    for trace in fig.data:
        trace.update(
            line=dict(width=3, color=color_map.get(trace.name)),
            marker=dict(size=8, color=color_map.get(trace.name)),
            mode="lines+markers+text",
            text=[str(v) for v in trace.y],
            textposition="top center",
            textfont=dict(size=11),
            cliponaxis=False,
        )
    fig.update_layout(
        xaxis_title="เดือน",
        yaxis_title="Fault สะสม (ครั้ง)",
        hovermode="x unified",
        margin=dict(l=50, r=35, t=55, b=45),
    )
    st.plotly_chart(style_fig(fig,540), use_container_width=True, config=PLOTLY_CONFIG)

    # Summary cards
    latest_lockout = "ไม่มีข้อมูล"
    if not lock_df.empty:
        latest = lock_df.sort_values("วันที่", ascending=False).iloc[0]
        d = latest["วันที่"].strftime("%Y-%m-%d") if not pd.isna(latest["วันที่"]) else "ไม่มีวันที่"
        v = pd.to_numeric(pd.Series([latest.get("Voltage (kV)")]), errors="coerce").iloc[0]
        vtxt = f"{int(v)} kV" if not pd.isna(v) else "ไม่ทราบแรงดัน"
        ltxt = str(latest.get("สายส่ง") or "ไม่มีข้อมูล")
        latest_lockout = f"{d}<br>{vtxt} {ltxt}"

    voltage_levels = int(
        pd.to_numeric(df["Voltage (kV)"], errors="coerce").dropna().nunique()
    )

    cards = st.columns(5)
    card_data = [
        ("📅 ช่วงข้อมูลที่แสดง", str(selected_year)),
        (f"♻ จำนวนเหตุการณ์ (ปี {selected_year})", f"{total} ครั้ง"),
        (f"🗼 สายส่งที่ได้รับผลกระทบ (ปี {selected_year})", f"{affected_lines} Line"),
        (f"▣ ระดับแรงดันที่พบ (ปี {selected_year})", f"{voltage_levels} ระดับ"),
        ("🔒 Lockout ล่าสุด", latest_lockout),
    ]
    for col,(label,value) in zip(cards,card_data):
        with col:
            st.markdown(
                f'<div class="summary-card"><div class="summary-label">{label}</div>'
                f'<div class="summary-value">{value}</div></div>',
                unsafe_allow_html=True
            )


# ============================================================
# RECURRING + LOCKOUT
# ============================================================

if view == "↻ Recurring & Lockout":
    st.subheader("Recurring Fault — สายส่ง + สาเหตุ")

    rr = df.copy()
    rr["cause_clean"] = rr["เหตุการณ์"].fillna("ไม่ทราบสาเหตุ").replace("", "ไม่ทราบสาเหตุ")
    rr["line_clean"] = rr["สายส่ง"].fillna("ไม่มีข้อมูล").replace("", "ไม่มีข้อมูล")
    recur = (
        rr.groupby(["line_clean","cause_clean"]).size()
        .reset_index(name="จำนวนครั้ง")
    )
    recur = recur[recur["จำนวนครั้ง"] >= 2].sort_values("จำนวนครั้ง", ascending=False)

    if recur.empty:
        st.info("ไม่พบเหตุซ้ำตั้งแต่ 2 ครั้งขึ้นไป")
    else:
        recur["Pattern"] = recur["line_clean"] + " — " + recur["cause_clean"]
        fig = px.bar(recur, x="จำนวนครั้ง", y="Pattern", orientation="h",
                     text="จำนวนครั้ง", title="Recurring Fault Pattern")
        fig.update_layout(yaxis={"categoryorder":"total ascending"})
        st.plotly_chart(style_fig(fig,420), use_container_width=True, config=PLOTLY_CONFIG)
        st.dataframe(
            recur[["line_clean","cause_clean","จำนวนครั้ง"]].rename(
                columns={"line_clean":"สายส่ง","cause_clean":"สาเหตุ"}
            ),
            use_container_width=True, hide_index=True
        )

    st.subheader("Lockout Events")
    show_cols = [c for c in [
        "วันที่","Time","TMU","สายส่ง","Voltage (kV)",
        "เหตุการณ์","Phase","Tower Damaged"
    ] if c in lock_df.columns]
    if lock_df.empty:
        st.info("ไม่มี Lockout ตามเงื่อนไขที่เลือก")
    else:
        st.dataframe(
            lock_df[show_cols].sort_values("วันที่", ascending=False),
            use_container_width=True, hide_index=True
        )


# ============================================================
# EVENT LOG
# ============================================================

if view == "▤ Event Log":
    st.subheader("Event Log")

    preferred = [
        "วันที่","Time","TMU","สายส่ง","Voltage (kV)","เหตุการณ์",
        "Phase","Tower Damaged","Trip type","Black out","Remark"
    ]
    cols = [c for c in preferred if c in df.columns]
    events = df[cols].sort_values("วันที่", ascending=False).copy()
    if "วันที่" in events.columns:
        events["วันที่"] = events["วันที่"].dt.strftime("%Y-%m-%d")

    st.dataframe(events, use_container_width=True, hide_index=True, height=560)

    csv = events.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Download filtered data (CSV)",
        data=csv,
        file_name=f"fault_{selected_year}_filtered.csv",
        mime="text/csv",
    )

st.markdown(
    '<div style="text-align:center;color:#6f879a;font-size:11px;padding:20px 0 4px">'
    'EGAT Fault Dashboard • V9 Hybrid Fast Load • Excel → Pandas → Plotly → Streamlit • No AI'
    '</div>',
    unsafe_allow_html=True,
)
