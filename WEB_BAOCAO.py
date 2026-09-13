from datetime import date, datetime
import io
import os
import re
import sqlite3
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ================= 1. KẾT NỐI CSDL TURSO CLOUD / LOCAL DUAL MODE =================
USE_TURSO = "TURSO_DATABASE_URL" in st.secrets

if USE_TURSO:
  import libsql_experimental as libsql


def get_db_connection():
  if USE_TURSO:
    conn = libsql.connect(
        database=st.secrets["TURSO_DATABASE_URL"],
        auth_token=st.secrets["TURSO_AUTH_TOKEN"],
    )
  else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DB_PATH = os.path.join(BASE_DIR, "Report_Database.db")
    conn = sqlite3.connect(DB_PATH, timeout=20.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.row_factory = sqlite3.Row
  return conn


# ================= 2. CẤU HÌNH DASHBOARD & HỆ THỐNG THIẾT KẾ =================
st.set_page_config(
    page_title="EMIC QC Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        header[data-testid="stHeader"] { background: transparent !important; box-shadow: none !important; }

        :root {
            --bg: #F5F6F9; --bg-accent: #EEF1FA; --surface: #FFFFFF; --border: #E6E9F0;
            --border-strong: #D6DAE5; --text: #0F1222; --text-muted: #6B7280;
            --primary: #4F46E5; --primary-soft: #EEF0FF; --success: #0EA968;
            --danger: #E23D4D; --warning: #EA8A0A; --purple: #9333EA; --radius: 14px;
        }

        html, body, [class*="css"], .stApp {
            font-family: 'Inter', -apple-system, sans-serif !important; color: var(--text);
        }
        .stApp {
            background: radial-gradient(1100px 480px at 12% -8%, var(--bg-accent) 0%, var(--bg) 55%) !important;
        }
        .main .block-container {
            padding: 1.2rem 2rem 3rem 2rem !important; max-width: 1720px !important; margin: 0 auto !important;
        }
        .filter-banner {
            background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius);
            padding: 16px 22px; box-shadow: 0 1px 3px rgba(15,18,34,0.05); margin-bottom: 22px;
        }
        .filter-title { font-size: 13px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; margin-bottom: 12px; }
        .page-header { display: flex; align-items: flex-end; justify-content: space-between; padding-bottom: 16px; margin-bottom: 18px; border-bottom: 1px solid var(--border); }
        .ph-title { font-size: 24px; font-weight: 900; color: var(--text); margin: 0; }
        .ph-subtitle { font-size: 12.5px; color: var(--text-muted); margin: 4px 0 0 0; }
        .ph-meta { font-size: 12px; font-weight: 700; color: #3730A3; background: var(--primary-soft); padding: 7px 15px; border-radius: 999px; }
        
        .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 20px; }
        .kpi-card { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); padding: 18px 20px; position: relative; }
        .kpi-card::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: var(--accent, var(--primary)); }
        .kpi-label { font-size: 11.5px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; }
        .kpi-value { font-size: 27px; font-weight: 800; color: var(--text); margin-top: 4px; }
        .chart-card { background-color: var(--surface); border-radius: var(--radius); border: 1px solid var(--border); padding: 18px 22px 10px 22px; margin-bottom: 20px; }
    </style>
""",
    unsafe_allow_html=True,
)


def render_page_header(title, subtitle, meta_text):
  st.markdown(
      f"""<div class="page-header"><div><p class="ph-title">{title}</p><p class="ph-subtitle">{subtitle}</p></div><div class="ph-meta">📅 {meta_text}</div></div>""",
      unsafe_allow_html=True,
  )


def render_kpi_cards(items):
  cards_html = ""
  for it in items:
    color = it.get("color", "#4F46E5")
    sub = it.get("sub", "")
    sub_html = f'<div style="font-size:11.5px; color:#6B7280; margin-top:6px;">{sub}</div>' if sub else ""
    cards_html += f"""<div class="kpi-card" style="--accent: {color};"><div style="display:flex; justify-content:space-between;"><span class="kpi-label">{it['label']}</span><span>{it.get('icon', '📌')}</span></div><div class="kpi-value">{it['value']}</div>{sub_html}</div>"""
  st.markdown(f'<div class="kpi-row">{cards_html}</div>', unsafe_allow_html=True)


COLOR_SUCCESS, COLOR_PRIMARY, COLOR_DANGER, COLOR_WARNING, COLOR_PURPLE = (
    "#0EA968",
    "#4F46E5",
    "#E23D4D",
    "#EA8A0A",
    "#9333EA",
)


@st.cache_data(ttl=15)
def load_data(tu_date, den_date):
  tu_iso = tu_date.strftime("%Y-%m-%d 00:00:00")
  den_iso = den_date.strftime("%Y-%m-%d 23:59:59")
  try:
    conn = get_db_connection()
    df_qa32 = pd.read_sql_query(
        "SELECT * FROM tb_sap_qa32 WHERE ngay_ve_dt >= ? AND ngay_ve_dt <= ?",
        conn,
        params=(tu_iso, den_iso),
    )
    df_coois = pd.read_sql_query(
        "SELECT * FROM tb_sap_coois WHERE ngay_lenh_dt >= ? AND ngay_lenh_dt"
        " <= ?",
        conn,
        params=(tu_iso, den_iso),
    )
    conn.close()
    return df_qa32, df_coois
  except Exception:
    return pd.DataFrame(), pd.DataFrame()


# ================= 3. BỘ LỌC ĐỈNH TRANG =================
st.markdown('<div class="filter-banner">', unsafe_allow_html=True)
st.markdown(
    '<div class="filter-title">📅 BỘ LỌC THỜI GIAN BÁO CÁO TOÀN HỆ THỐNG</div>',
    unsafe_allow_html=True,
)
col_f1, col_f2, col_f3 = st.columns([1.5, 1.5, 1.2])
with col_f1:
  tu_date = st.date_input("Từ ngày:", date(datetime.now().year, 1, 1))
with col_f2:
  den_date = st.date_input("Đến ngày:", date.today())
with col_f3:
  st.markdown("<div style='height: 25px;'></div>", unsafe_allow_html=True)
  if st.button("🔄 CẬP NHẬT BÁO CÁO", use_container_width=True, type="primary"):
    st.cache_data.clear()
    st.rerun()
st.markdown("</div>", unsafe_allow_html=True)

df_qa32, df_coois = load_data(tu_date, den_date)

render_page_header(
    "📊 EMIC QC Dashboard",
    f"Hệ Thống Báo Cáo QC {'(Cloud Turso)' if USE_TURSO else '(Máy Tính Nội Bộ)'}",
    f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
)

tab_vat_tu, tab_co_khi, tab_danh_sach = st.tabs(
    ["📋 Báo Cáo Vật Tư", "⚙️ Báo Cáo Sản Xuất", "🔍 Danh Sách Chi Tiết"]
)

with tab_vat_tu:
  if df_qa32.empty:
    st.info("💡 Chưa có dữ liệu QA32 trong dải ngày đã chọn.")
  else:
    tot_ft = (
        df_qa32["ft_qty"].sum()
        if "ft_qty" in df_qa32.columns
        else len(df_qa32)
    )
    tot_ca = (
        df_qa32["ca_qty"].sum() if "ca_qty" in df_qa32.columns else 0.0
    )
    render_kpi_cards([
        {
            "label": "Tổng Lô Về (FT)",
            "value": f"{len(df_qa32):,}",
            "icon": "📦",
            "color": COLOR_PRIMARY,
        },
        {
            "label": "Tổng Số Lượng (FT)",
            "value": f"{int(tot_ft):,}",
            "icon": "📊",
            "color": COLOR_SUCCESS,
        },
        {
            "label": "SL Bị Block (CA)",
            "value": f"{int(tot_ca):,}",
            "icon": "🚫",
            "color": COLOR_DANGER,
        },
        {
            "label": "Trạng Thái CSDL",
            "value": "TURSO CLOUD" if USE_TURSO else "LOCAL DB",
            "icon": "🌐",
            "color": COLOR_PURPLE,
        },
    ])
    st.dataframe(df_qa32, use_container_width=True, hide_index=True)

with tab_co_khi:
  if df_coois.empty:
    st.info("💡 Chưa có dữ liệu COOIS trong dải ngày đã chọn.")
  else:
    st.dataframe(df_coois, use_container_width=True, hide_index=True)

with tab_danh_sach:
  st.markdown("### 🔍 Danh Sách Báo Cáo QC Đầu Vào Đã Lưu")
  try:
    conn = get_db_connection()
    df_dau_vao = pd.read_sql_query(
        "SELECT * FROM tb_qc_dau_vao ORDER BY id DESC", conn
    )
    conn.close()
    if not df_dau_vao.empty:
      st.dataframe(df_dau_vao, use_container_width=True, hide_index=True)
    else:
      st.warning("⚠️ Chưa có báo cáo nhập liệu nào trong cơ sở dữ liệu.")
  except Exception as e:
    st.info("💡 Chưa khởi tạo bảng báo cáo đầu vào.")