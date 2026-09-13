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

# ================= 1. KẾT NỐI CSDL AN TOÀN (TURSO CLOUD / LOCAL SQLITE) =================
try:
  USE_TURSO = "TURSO_DATABASE_URL" in st.secrets
except Exception:
  USE_TURSO = False

if USE_TURSO:
  import libsql_experimental as libsql


def get_db_connection():
  if USE_TURSO:
    conn = libsql.connect(
        database=st.secrets["TURSO_DATABASE_URL"],
        auth_token=st.secrets.get("TURSO_AUTH_TOKEN", ""),
    )
  else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    DB_PATH = os.path.join(BASE_DIR, "Report_Database.db")
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.row_factory = sqlite3.Row
  return conn


# ================= 2. CẤU HÌNH DASHBOARD & HỆ THỐNG THIẾT KẾ =================
st.set_page_config(
    page_title="EMIC QC Dashboard Tổng Hợp",
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
            --bg: #F5F6F9;
            --bg-accent: #EEF1FA;
            --surface: #FFFFFF;
            --border: #E6E9F0;
            --border-strong: #D6DAE5;
            --text: #0F1222;
            --text-muted: #6B7280;
            --primary: #4F46E5;
            --primary-dark: #3730A3;
            --primary-soft: #EEF0FF;
            --success: #0EA968;
            --success-soft: #ECFDF5;
            --danger: #E23D4D;
            --danger-soft: #FEF2F3;
            --warning: #EA8A0A;
            --warning-soft: #FFF8EB;
            --purple: #9333EA;
            --radius: 14px;
        }

        html, body, [class*="css"], .stApp {
            font-family: 'Inter', -apple-system, sans-serif !important;
            color: var(--text);
        }
        .stApp {
            background: radial-gradient(1100px 480px at 12% -8%, var(--bg-accent) 0%, var(--bg) 55%) !important;
        }

        .main .block-container {
            padding: 1.2rem 2rem 3rem 2rem !important;
            max-width: 1720px !important;
            margin: 0 auto !important;
        }

        .filter-banner {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 16px 22px;
            box-shadow: 0 1px 3px rgba(15,18,34,0.05);
            margin-bottom: 22px;
        }
        .filter-banner .filter-title {
            font-size: 13px; font-weight: 700; color: var(--text-muted);
            text-transform: uppercase; margin-bottom: 12px;
        }

        .page-header {
            display: flex; align-items: flex-end; justify-content: space-between;
            padding-bottom: 16px; margin-bottom: 18px;
            border-bottom: 1px solid var(--border);
        }
        .page-header .ph-title { font-size: 24px; font-weight: 900; color: var(--text); margin: 0; }
        .page-header .ph-subtitle { font-size: 12.5px; color: var(--text-muted); margin: 4px 0 0 0; }
        .page-header .ph-meta {
            font-size: 12px; font-weight: 700; color: var(--primary-dark);
            background: var(--primary-soft); padding: 7px 15px; border-radius: 999px;
        }

        .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 20px; }
        .kpi-card {
            background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius);
            padding: 18px 20px; position: relative; overflow: hidden;
        }
        .kpi-card::before {
            content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
            background: var(--accent, var(--primary));
        }
        .kpi-card .kpi-label { font-size: 11.5px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; }
        .kpi-card .kpi-value { font-size: 27px; font-weight: 800; color: var(--text); margin-top: 4px; }
        .chart-card {
            background-color: var(--surface); border-radius: var(--radius); border: 1px solid var(--border);
            padding: 18px 22px 10px 22px; margin-bottom: 20px;
        }
        .section-heading { font-size: 16.5px; font-weight: 800; color: var(--text); margin: 6px 0 14px 2px; }
    </style>
""",
    unsafe_allow_html=True,
)


def render_page_header(title, subtitle, meta_text):
  st.markdown(
      f"""<div class="page-header"><div><p class="ph-title">{title}</p><p class="ph-subtitle">{subtitle}</p></div><div class="ph-meta">📅 {meta_text}</div></div>""",
      unsafe_allow_html=True,
  )


def render_section_heading(text):
  st.markdown(
      f'<div class="section-heading">{text}</div>', unsafe_allow_html=True
  )


def render_kpi_cards(items):
  cards_html = ""
  for it in items:
    color = it.get("color", "#4F46E5")
    sub = it.get("sub", "")
    sub_html = (
        f'<div style="font-size:11.5px; color:#6B7280;'
        f' margin-top:6px;">{sub}</div>'
        if sub
        else ""
    )
    cards_html += f"""<div class="kpi-card" style="--accent: {color};"><div style="display:flex; justify-content:space-between;"><span class="kpi-label">{it['label']}</span><span>{it.get('icon', '📌')}</span></div><div class="kpi-value">{it['value']}</div>{sub_html}</div>"""
  st.markdown(f'<div class="kpi-row">{cards_html}</div>', unsafe_allow_html=True)


COLOR_SUCCESS, COLOR_PRIMARY, COLOR_DANGER, COLOR_WARNING, COLOR_PURPLE = (
    "#0EA968",
    "#4F46E5",
    "#E23D4D",
    "#EA8A0A",
    "#9333EA",
)
COLOR_TEXT = "#0F1222"
DISTINCT_COLORS = [
    "#4F46E5",
    "#0EA968",
    "#EA8A0A",
    "#E23D4D",
    "#9333EA",
    "#0891B2",
]
PLOTLY_FONT = "Inter, -apple-system, Segoe UI, sans-serif"
PLOTLY_GRID, PLOTLY_AXIS_TEXT = "#F1F2F6", "#6B7280"

plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = [
    "Calibri",
    "Arial",
    "DejaVu Sans",
    "sans-serif",
]
plt.rcParams["font.size"] = 9


def clean_emoji(text):
  return re.sub(r"[^\w\s\(\)\-\/\.\,\:]", "", str(text)).strip()


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
        "SELECT * FROM tb_sap_coois WHERE ngay_lenh_dt >= ? AND ngay_lenh_dt <="
        " ?",
        conn,
        params=(tu_iso, den_iso),
    )
    conn.close()
    return df_qa32, df_coois
  except Exception:
    return pd.DataFrame(), pd.DataFrame()


# ================= 3. BỘ LỌC THỜI GIAN =================
st.markdown('<div class="filter-banner">', unsafe_allow_html=True)
st.markdown(
    '<div class="filter-title">📅 BỘ LỌC THỜI GIAN BÁO CÁO TOÀN HỆ THỐNG</div>',
    unsafe_allow_html=True,
)

today = date.today()
first_day_of_month = date(today.year, today.month, 1)

col_f1, col_f2, col_f3 = st.columns([1.5, 1.5, 1.2])
with col_f1:
  tu_date = st.date_input("Từ ngày:", first_day_of_month)
with col_f2:
  den_date = st.date_input("Đến ngày:", today)
with col_f3:
  st.markdown("<div style='height: 25px;'></div>", unsafe_allow_html=True)
  if st.button("🔄 CẬP NHẬT BÁO CÁO", use_container_width=True, type="primary"):
    st.cache_data.clear()
    st.rerun()
st.markdown("</div>", unsafe_allow_html=True)

df_qa32, df_coois = load_data(tu_date, den_date)

render_page_header(
    "📊 EMIC QC Dashboard Tổng Hợp",
    "Tổng Công Ty Thiết Bị Điện EMIC · Hệ Thống Quản Lý Chất Lượng Tự Động"
    f" {'(Turso Cloud)' if USE_TURSO else '(Máy Tính Nội Bộ)'}",
    f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
)


# ================= 4. NAVIGATION TABS (5 TABS) =================
tab_vat_tu, tab_co_khi, tab_tuti, tab_cong_to, tab_danh_sach = st.tabs([
    "📋 Báo Cáo Vật Tư",
    "⚙️ Báo Cáo Cơ Khí",
    "🔌 Báo Cáo TU/TI",
    "⚡ Báo Cáo Công Tơ",
    "🔍 Danh Sách Chi Tiết & Năng Suất",
])

with tab_vat_tu:
  if df_qa32.empty:
    st.info("💡 Chưa có dữ liệu QA32 trong khoảng thời gian đã chọn.")
  else:
    tot_ft = df_qa32["ft_qty"].sum() if "ft_qty" in df_qa32.columns else 0
    render_kpi_cards([
        {
            "label": "Tổng Lô Về (FT)",
            "value": f"{len(df_qa32):,}",
            "icon": "📦",
            "color": COLOR_PRIMARY,
        },
        {
            "label": "Tổng Số Lượng Về",
            "value": f"{int(tot_ft):,}",
            "icon": "📊",
            "color": COLOR_SUCCESS,
        },
    ])
    st.dataframe(df_qa32, use_container_width=True, hide_index=True)

with tab_co_khi:
  if df_coois.empty:
    st.info("💡 Chưa có dữ liệu sản xuất Cơ khí.")
  else:
    st.dataframe(
        df_coois[df_coois["phan_he"] == "CO_KHI"],
        use_container_width=True,
        hide_index=True,
    )

with tab_tuti:
  if df_coois.empty:
    st.info("💡 Chưa có dữ liệu sản xuất TU/TI.")
  else:
    st.dataframe(
        df_coois[df_coois["phan_he"] == "TU_TI"],
        use_container_width=True,
        hide_index=True,
    )

with tab_cong_to:
  if df_coois.empty:
    st.info("💡 Chưa có dữ liệu sản xuất Công tơ.")
  else:
    st.dataframe(
        df_coois[df_coois["phan_he"] == "CONG_TO"],
        use_container_width=True,
        hide_index=True,
    )

# ================= 5. TAB 5: DANH SÁCH CHI TIẾT & BÁO CÁO NĂNG SUẤT =================
with tab_danh_sach:
  render_section_heading(
      "🔍 QUẢN LÝ DANH SÁCH CHI TIẾT VẬT TƯ, LỆNH SX & BÁO CÁO NĂNG SUẤT"
  )

  tab_sub_vt, tab_sub_lenh, tab_sub_nangsuat = st.tabs([
      "📦 1. Danh Sách Vật Tư (QA32)",
      "⚙️ 2. Danh Sách Lệnh Sản Xuất (COOIS)",
      "👨‍💼 3. Báo Cáo Năng Suất Cá Nhân (Nhật Ký QC)",
  ])

  with tab_sub_vt:
    if not df_qa32.empty:
      st.dataframe(df_qa32, use_container_width=True, hide_index=True)
    else:
      st.info("💡 Chưa có dữ liệu vật tư.")

  with tab_sub_lenh:
    if not df_coois.empty:
      st.dataframe(df_coois, use_container_width=True, hide_index=True)
    else:
      st.info("💡 Chưa có dữ liệu lệnh sản xuất.")

  # SUB-TAB BÁO CÁO NĂNG SUẤT CÁ NHÂN THEO TỪNG NGƯỜI / TỪNG NGÀY
  with tab_sub_nangsuat:
    try:
      conn = get_db_connection()
      df_qc_logs = pd.read_sql_query(
          "SELECT id, loai_qc, so_lot, ma_vt, ten_vt, ncc, tong_sl_ve, sl_kiem,"
          " sl_khong_dat, sl_dat, ket_luan, nguoi_kiem, ngay_kiem, ghi_chu FROM"
          " tb_qc_dau_vao ORDER BY ngay_kiem DESC",
          conn,
      )
      conn.close()

      if not df_qc_logs.empty:
        df_qc_logs["ngay_kiem_dt"] = pd.to_datetime(
            df_qc_logs["ngay_kiem"], errors="coerce"
        )
        df_qc_logs["Ngay_Format"] = df_qc_logs["ngay_kiem_dt"].dt.strftime(
            "%d/%m/%Y"
        )

        col_flt_person, col_flt_type = st.columns([1.5, 1])

        list_inspectors = ["Tất cả nhân sự"] + sorted([
            str(x).strip()
            for x in df_qc_logs["nguoi_kiem"].dropna().unique()
            if str(x).strip()
        ])

        with col_flt_person:
          selected_inspector = st.selectbox(
              "👤 Chọn Nhân sự QC:", list_inspectors, key="ns_inspector"
          )

        with col_flt_type:
          selected_loai_qc = st.selectbox(
              "🎯 Phân hệ kiểm:",
              ["Tất cả", "QC Đầu Vào (DAU_VAO)", "QC Sản Xuất (SAN_XUAT)"],
              key="ns_loai_qc",
          )

        df_filtered = df_qc_logs.copy()

        if selected_inspector != "Tất cả nhân sự":
          df_filtered = df_filtered[
              df_filtered["nguoi_kiem"] == selected_inspector
          ]

        if selected_loai_qc == "QC Đầu Vào (DAU_VAO)":
          df_filtered = df_filtered[df_filtered["loai_qc"] == "DAU_VAO"]
        elif selected_loai_qc == "QC Sản Xuất (SAN_XUAT)":
          df_filtered = df_filtered[df_filtered["loai_qc"] == "SAN_XUAT"]

        tot_luot = len(df_filtered)
        tot_sl_kiem = (
            df_filtered["sl_kiem"].sum() if "sl_kiem" in df_filtered else 0
        )
        tot_sl_loi = (
            df_filtered["sl_khong_dat"].sum()
            if "sl_khong_dat" in df_filtered
            else 0
        )
        ty_le_loi = (
            (tot_sl_loi / tot_sl_kiem * 100) if tot_sl_kiem > 0 else 0.0
        )

        render_kpi_cards([
            {
                "label": "TỔNG LƯỢT KIỂM TRẢ",
                "value": f"{tot_luot:,} lượt",
                "icon": "📝",
                "color": COLOR_PRIMARY,
            },
            {
                "label": "TỔNG SỐ LƯỢNG ĐÃ KIỂM",
                "value": f"{int(tot_sl_kiem):,}",
                "icon": "🔍",
                "color": COLOR_SUCCESS,
            },
            {
                "label": "TỔNG SỐ LƯỢNG LỖI",
                "value": f"{int(tot_sl_loi):,}",
                "icon": "⚠️",
                "color": COLOR_DANGER,
            },
            {
                "label": "TỶ LỆ LỖI PHÁT HIỆN",
                "value": f"{ty_le_loi:.1f}%",
                "icon": "📈",
                "color": COLOR_WARNING,
            },
        ])

        df_display = pd.DataFrame()
        df_display["STT"] = np.arange(1, len(df_filtered) + 1)
        df_display["Ngày kiểm"] = df_filtered["Ngay_Format"].values
        df_display["Thời gian"] = df_filtered["ngay_kiem"].values
        df_display["Người kiểm tra"] = df_filtered["nguoi_kiem"].values
        df_display["Loại QC"] = df_filtered["loai_qc"].map(
            {"DAU_VAO": "QC Đầu Vào", "SAN_XUAT": "QC Sản Xuất"}
        )
        df_display["Lô / Lệnh SX"] = df_filtered["so_lot"].values
        df_display["Mã mặt hàng"] = df_filtered["ma_vt"].values
        df_display["Tên mặt hàng"] = df_filtered["ten_vt"].values
        df_display["Đơn vị / NCC"] = df_filtered["ncc"].values
        df_display["SL Kiểm"] = df_filtered["sl_kiem"].values
        df_display["SL Lỗi"] = df_filtered["sl_khong_dat"].values
        df_display["SL Đạt"] = df_filtered["sl_dat"].values
        df_display["Kết luận"] = df_filtered["ket_luan"].values
        df_display["Ghi chú"] = df_filtered["ghi_chu"].values

        st.markdown(
            f"##### 📋 BẢNG NHẬT KÝ KIỂM TRẢ CHI TIẾT ({len(df_display):,} bản"
            " ghi)"
        )
        st.dataframe(
            df_display,
            column_config={
                "STT": st.column_config.NumberColumn("STT", width="small"),
                "SL Kiểm": st.column_config.NumberColumn(
                    "SL Kiểm", format="%d"
                ),
                "SL Lỗi": st.column_config.NumberColumn("SL Lỗi", format="%d"),
                "SL Đạt": st.column_config.NumberColumn("SL Đạt", format="%d"),
            },
            use_container_width=True,
            hide_index=True,
            height=450,
        )

        buf_ns = io.BytesIO()
        with pd.ExcelWriter(buf_ns, engine="openpyxl") as writer:
          df_display.to_excel(
              writer, sheet_name="NangSuat_QC_ChiTiet", index=False
          )

        st.download_button(
            label="📥 XUẤT BÁO CÁO NĂNG SUẤT QC (EXCEL)",
            data=buf_ns.getvalue(),
            file_name=(
                "BaoCao_NangSuat_QC_"
                f"{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            ),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
      else:
        st.info("💡 Chưa có nhật ký báo cáo QC nào trong Cơ sở dữ liệu.")
    except Exception as e:
      st.error(f"⚠️ Lỗi khi tải nhật ký QC: {e}")