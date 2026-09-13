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

# ================= 1. KẾT NỐI CSDL AN TOÀN & TỰ ĐỘNG CẬP NHẬT CẤU TRÚC =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Report_Database.db")
IMG_DIR = os.path.join(BASE_DIR, "Anh_kiem_tra_dau_vao")

os.makedirs(IMG_DIR, exist_ok=True)


def get_db_connection():
  try:
    USE_TURSO = "TURSO_DATABASE_URL" in st.secrets
  except Exception:
    USE_TURSO = False

  if USE_TURSO:
    import libsql_experimental as libsql

    return libsql.connect(
        database=st.secrets["TURSO_DATABASE_URL"],
        auth_token=st.secrets.get("TURSO_AUTH_TOKEN", ""),
    )
  else:
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
  try:
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_qc_dau_vao (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                loai_qc TEXT DEFAULT 'DAU_VAO',
                so_lot TEXT, ma_vt TEXT, ten_vt TEXT, ncc TEXT, ngay_ve TEXT,
                tong_sl_ve REAL, sl_kiem REAL, sl_khong_dat REAL, sl_dat REAL,
                ket_luan TEXT, nguoi_kiem TEXT, ngay_kiem DATETIME, ghi_chu TEXT,
                kieu_loi TEXT DEFAULT '', cong_viec_con TEXT DEFAULT '', img1 TEXT, img2 TEXT
            )
        """)

    cursor.execute("PRAGMA table_info(tb_qc_dau_vao)")
    cols = [col[1] for col in cursor.fetchall()]
    if "loai_qc" not in cols:
      cursor.execute(
          "ALTER TABLE tb_qc_dau_vao ADD COLUMN loai_qc TEXT DEFAULT 'DAU_VAO'"
      )
    if "kieu_loi" not in cols:
      cursor.execute(
          "ALTER TABLE tb_qc_dau_vao ADD COLUMN kieu_loi TEXT DEFAULT ''"
      )
    if "cong_viec_con" not in cols:
      cursor.execute(
          "ALTER TABLE tb_qc_dau_vao ADD COLUMN cong_viec_con TEXT DEFAULT ''"
      )

    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_dm_loai_loi (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                phan_he TEXT,
                ten_loi TEXT
            )
        """)

    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_dm_cong_viec (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                phan_he TEXT,
                ten_cong_viec TEXT
            )
        """)

    conn.commit()
    conn.close()
  except Exception:
    pass


init_db()

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
            --text-faint: #9CA3AF;
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
            --radius-lg: 18px;
            --radius: 14px;
            --radius-sm: 10px;
            --shadow-xs: 0 1px 2px rgba(15,18,34,0.05);
            --shadow: 0 1px 3px rgba(15,18,34,0.05), 0 6px 16px -6px rgba(15,18,34,0.08);
            --shadow-hover: 0 8px 24px -6px rgba(15,18,34,0.14);
        }

        html, body, [class*="css"], .stApp {
            font-family: 'Inter', -apple-system, 'Segoe UI', sans-serif !important;
            color: var(--text);
        }
        .stApp {
            background: radial-gradient(1100px 480px at 12% -8%, var(--bg-accent) 0%, var(--bg) 55%) !important;
            background-attachment: fixed !important;
        }

        .main .block-container {
            padding: 1.2rem 2rem 3rem 2rem !important;
            max-width: 1720px !important;
            margin: 0 auto !important;
        }

        section[data-testid="stSidebar"] {
            background-color: #FFFFFF !important;
            border-right: 1px solid var(--border) !important;
        }
        
        .sidebar-brand {
            background: linear-gradient(135deg, #1E1B4B 0%, #4F46E5 100%);
            color: white; padding: 16px; border-radius: 14px; text-align: center; margin-bottom: 16px;
            box-shadow: 0 4px 12px rgba(79, 70, 229, 0.25);
        }
        .sidebar-brand .sb-logo { font-size: 10px; font-weight: 800; letter-spacing: 1.5px; color: #93C5FD; text-transform: uppercase; }
        .sidebar-brand .sb-title { font-size: 16px; font-weight: 900; color: #FFFFFF; margin-top: 3px; }

        .tree-node-title {
            font-size: 11px; font-weight: 800; color: var(--primary);
            text-transform: uppercase; letter-spacing: 0.08em; margin: 12px 0 6px 4px;
        }

        .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 20px; }
        .kpi-card {
            background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius);
            padding: 18px 20px; box-shadow: var(--shadow); position: relative; overflow: hidden;
            transition: box-shadow 0.18s ease, transform 0.18s ease, border-color 0.18s ease;
        }
        .kpi-card:hover { box-shadow: var(--shadow-hover); transform: translateY(-2px); border-color: var(--border-strong); }
        .kpi-card::before {
            content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
            background: var(--accent, var(--primary));
        }
        .kpi-card .kpi-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
        .kpi-card .kpi-label { font-size: 11.5px; font-weight: 700; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.05em; }
        .kpi-card .kpi-icon {
            width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center;
            justify-content: center; font-size: 14px; background: var(--accent-soft, var(--primary-soft));
        }
        .kpi-card .kpi-value { font-size: 27px; font-weight: 800; color: var(--text); line-height: 1.15; letter-spacing: -0.02em; }
        .kpi-card .kpi-sub { font-size: 11.5px; color: var(--text-muted); margin-top: 6px; font-weight: 500; }

        .chart-card {
            background-color: var(--surface); border-radius: var(--radius); border: 1px solid var(--border);
            padding: 18px 22px 10px 22px; box-shadow: 0 2px 4px rgba(15,18,34,0.06), 0 12px 28px -8px rgba(15,18,34,0.14);
            margin-bottom: 20px; transition: box-shadow 0.18s ease, transform 0.18s ease;
        }
        .chart-card:hover { box-shadow: 0 4px 8px rgba(15,18,34,0.08), 0 18px 36px -8px rgba(15,18,34,0.18); transform: translateY(-1px); }
        .chart-card-header {
            display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px;
            padding-bottom: 10px; border-bottom: 1px solid var(--border);
        }
        .chart-card-title {
            font-size: 15.5px; font-weight: 800; color: var(--text); letter-spacing: -0.01em;
            display: flex; align-items: center; gap: 8px;
        }
        .chart-card-title::before {
            content: ""; width: 4px; height: 15px; border-radius: 2px;
            background: var(--primary); display: inline-block;
        }
        .chart-card-caption {
            font-size: 11px; color: var(--text-muted); font-weight: 600;
            background: var(--bg); padding: 3px 9px; border-radius: 999px;
        }

        .section-heading { font-size: 16.5px; font-weight: 800; color: var(--text); margin: 6px 0 14px 2px; letter-spacing: -0.01em; }

        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border) !important; border-radius: var(--radius) !important;
            box-shadow: 0 2px 4px rgba(15,18,34,0.05), 0 8px 20px -6px rgba(15,18,34,0.10);
            overflow: hidden;
        }
        div[data-testid="stDataFrame"] [data-testid="stDataFrameResizable"] { font-family: 'Inter', sans-serif !important; }

        [data-testid="stWidgetLabel"] p {
            font-size: 12px !important; font-weight: 700 !important; color: var(--text-muted) !important;
            text-transform: uppercase; letter-spacing: 0.03em;
        }
        div[data-baseweb="select"] > div {
            border-radius: var(--radius-sm) !important; border-color: var(--border) !important;
            font-family: 'Inter', sans-serif !important; background-color: var(--surface) !important;
        }
        .stDateInput input, .stTextInput input {
            border-radius: var(--radius-sm) !important; font-family: 'Inter', sans-serif !important;
            border-color: var(--border) !important; background-color: var(--surface) !important;
            font-weight: 600 !important;
        }

        .stButton button, .stDownloadButton button {
            border-radius: var(--radius-sm) !important; font-weight: 600 !important;
            font-family: 'Inter', sans-serif !important; border: 1px solid var(--border) !important;
            transition: transform 0.12s ease, box-shadow 0.12s ease, background-color 0.12s ease;
        }
        .stButton button:hover, .stDownloadButton button:hover { transform: translateY(-1px); }
        .stButton button[kind="primary"], .stDownloadButton button[kind="primary"] {
            background-color: var(--primary) !important; border: none !important;
            box-shadow: 0 3px 10px rgba(79, 70, 229, 0.32) !important;
        }

        ::-webkit-scrollbar { width: 9px; height: 9px; }
        ::-webkit-scrollbar-thumb { background: #C9CEDA; border-radius: 8px; }
        ::-webkit-scrollbar-track { background: transparent; }
    </style>
""",
    unsafe_allow_html=True,
)


def _html(raw):
  return "".join(line.strip() for line in raw.strip().splitlines())


def render_page_header(title, subtitle, meta_text):
  st.markdown(
      _html(f"""
      <div class="page-header">
          <div>
              <p class="ph-title">{title}</p>
              <p class="ph-subtitle">{subtitle}</p>
          </div>
          <div class="ph-meta">📅 {meta_text}</div>
      </div>
      """),
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
    color_soft = it.get("color_soft", "#EEF0FF")
    sub = it.get("sub", "")
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    cards_html += _html(f"""
        <div class="kpi-card" style="--accent: {color}; --accent-soft: {color_soft};">
            <div class="kpi-head">
                <span class="kpi-label">{it['label']}</span>
                <span class="kpi-icon">{it.get('icon', '📌')}</span>
            </div>
            <div class="kpi-value">{it['value']}</div>
            {sub_html}
        </div>
    """)
  st.markdown(f'<div class="kpi-row">{cards_html}</div>', unsafe_allow_html=True)


def chart_card_open(title, caption=""):
  cap_html = f'<span class="chart-card-caption">{caption}</span>' if caption else ""
  st.markdown(
      _html(f"""<div class="chart-card">
          <div class="chart-card-header">
              <span class="chart-card-title">{title}</span>
              {cap_html}
          </div>"""),
      unsafe_allow_html=True,
  )


def chart_card_close():
  st.markdown("</div>", unsafe_allow_html=True)


COLOR_SUCCESS = "#0EA968"
COLOR_PRIMARY = "#4F46E5"
COLOR_DANGER = "#E23D4D"
COLOR_WARNING = "#EA8A0A"
COLOR_PURPLE = "#9333EA"
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
PLOTLY_GRID = "#F1F2F6"
PLOTLY_AXIS_TEXT = "#6B7280"

plt.rcParams["font.family"] = "sans-serif"
plt.rcParams["font.sans-serif"] = [
    "Calibri",
    "Arial",
    "DejaVu Sans",
    "sans-serif",
]
plt.rcParams["font.size"] = 9
plt.rcParams["axes.unicode_minus"] = False
plt.rcParams["axes.edgecolor"] = "#CBD5E1"
plt.rcParams["axes.linewidth"] = 0.8


def clean_emoji(text):
  return re.sub(r"[^\w\s\(\)\-\/\.\,\:]", "", str(text)).strip()


@st.cache_data(ttl=15)
def load_data(tu_date, den_date):
  if not os.path.exists(DB_PATH):
    return pd.DataFrame(), pd.DataFrame()

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


# ================= 3. HÀM TẠO EXCEL BÁO CÁO =================
def generate_print_ready_excel(
    phan_he_code,
    title_clean,
    tu_date,
    den_date,
    df_monthly,
    df_plan,
    df_family,
    df_year,
    fig1_mpl,
    fig2_mpl,
    fig3_mpl,
    fig4_mpl,
):
  output = io.BytesIO()
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  font_company = Font(name="Calibri", size=10, bold=True, color="1F4E79")
  font_title = Font(name="Calibri", size=14, bold=True, color="000000")
  font_subtitle = Font(name="Calibri", size=9, italic=True, color="595959")
  font_section = Font(name="Calibri", size=11, bold=True, color="1F4E79")
  font_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
  font_data = Font(name="Calibri", size=9)
  fill_header = PatternFill(
      start_color="1F4E79", end_color="1F4E79", fill_type="solid"
  )
  fill_zebra = PatternFill(
      start_color="F9FBFD", end_color="F9FBFD", fill_type="solid"
  )
  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )
  header_border = Border(
      left=Side(style="thin", color="FFFFFF"),
      right=Side(style="thin", color="FFFFFF"),
      top=Side(style="medium", color="1F4E79"),
      bottom=Side(style="medium", color="1F4E79"),
  )
  align_center = Alignment(
      horizontal="center", vertical="center", wrap_text=True
  )
  align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
  align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

  sheets_data = [
      (
          "TienDo_Thang",
          "1. TIẾN ĐỘ SẢN XUẤT THEO THÁNG",
          df_monthly,
          fig1_mpl,
          "I7",
      ),
      (
          "TongQuan_KeHoach",
          "2. TỔNG QUAN CHỈ TIÊU KẾ HOẠCH",
          df_plan,
          fig2_mpl,
          "E7",
      ),
      (
          "Dong_SanPham",
          "3. CHI TIẾT THEO DÒNG SẢN PHẨM",
          df_family,
          fig3_mpl,
          "D7",
      ),
      (
          "SanLuong_CaNam",
          "4. TỔNG SẢN LƯỢNG CẢ NĂM MÃ ĐẦU 5",
          df_year,
          fig4_mpl,
          "D7",
      ),
  ]

  for sheet_name, section_title, df_table, fig_obj, img_pos in sheets_data:
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ws["A1"] = (
        "TỔNG CÔNG TY THIẾT BỊ ĐIỆN EMIC - PHÒNG QUẢN LÝ CHẤT LƯỢNG (QC)"
    )
    ws["A1"].font = font_company
    ws["A2"] = f"BÁO CÁO TỔNG HỢP DỮ LIỆU - {title_clean.upper()}"
    ws["A2"].font = font_title
    ws["A3"] = (
        f"Thời gian: {tu_date.strftime('%d/%m/%Y')} -"
        f" {den_date.strftime('%d/%m/%Y')} | Ngày xuất:"
        f" {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )
    ws["A3"].font = font_subtitle
    ws["A5"] = section_title
    ws["A5"].font = font_section

    start_row = 7
    for col_idx, col_name in enumerate(df_table.columns, start=1):
      cell = ws.cell(row=start_row, column=col_idx, value=col_name)
      cell.font = font_header
      cell.fill = fill_header
      cell.alignment = align_center
      cell.border = header_border

    for row_idx, row_data in enumerate(df_table.values, start=start_row + 1):
      row_fill = (
          fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
      )
      for col_idx, val in enumerate(row_data, start=1):
        cell = ws.cell(row=row_idx, column=col_idx)
        col_name_lower = str(df_table.columns[col_idx - 1]).lower()
        if isinstance(val, (int, float, np.number)):
          if "%" in col_name_lower or "tỷ lệ" in col_name_lower:
            cell.value = float(val) / 100.0 if val > 1.0 else float(val)
            cell.number_format = "0.0%"
          else:
            cell.value = float(val)
            cell.number_format = (
                "#,##0" if float(val).is_integer() else "#,##0.00"
            )
          cell.alignment = align_right
        else:
          cell.value = str(val)
          cell.alignment = (
              align_center if col_idx == 1 or len(str(val)) < 10 else align_left
          )
        cell.font = font_data
        if row_fill.fill_type:
          cell.fill = row_fill
        cell.border = thin_border

    for col in ws.columns:
      max_len = 0
      col_letter = get_column_letter(col[0].column)
      for cell in col:
        if cell.row >= start_row and cell.value is not None:
          max_len = max(max_len, len(str(cell.value)))
      ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    if fig_obj is not None:
      buf = io.BytesIO()
      fig_obj.savefig(
          buf, format="png", dpi=200, bbox_inches="tight", facecolor="#FFFFFF"
      )
      buf.seek(0)
      img = OpenpyxlImage(buf)
      ws.add_image(img, img_pos)

  wb.save(output)
  return output.getvalue()


# ================= 4. SIDEBAR MENU CÂY 2 CẤP & BỘ LỌC THỜI GIAN =================
with st.sidebar:
  st.markdown(
      """
        <div class='sidebar-brand'>
            <div class='sb-logo'>🏢 TỔNG CÔNG TY EMIC</div>
            <div class='sb-title'>HỆ THỐNG BÁO CÁO QC</div>
        </div>
    """,
      unsafe_allow_html=True,
  )

  st.markdown(
      "<div class='tree-node-title'>📅 BỘ LỌC THỜI GIAN</div>",
      unsafe_allow_html=True,
  )
  today = date.today()
  first_day_of_month = date(today.year, today.month, 1)

  tu_date = st.date_input("Từ ngày:", first_day_of_month, key="sb_tu_date")
  den_date = st.date_input("Đến ngày:", today, key="sb_den_date")

  if st.button("🔄 CẬP NHẬT DỮ LIỆU", type="primary", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

  st.markdown("<hr style='margin:14px 0;'>", unsafe_allow_html=True)
  st.markdown(
      "<div class='tree-node-title'>🌳 PHÂN CẤP MENU QUẢN LÝ</div>",
      unsafe_allow_html=True,
  )

  tree_node_category = st.selectbox(
      "📁 CHỌN NHÓM BÁO CÁO:",
      [
          "1. 📊 Báo Cáo Tiến Độ & Tổng Quan (Charts)",
          "2. 🔍 Tra Cứu Danh Sách Chi Tiết (SAP Data)",
          "3. 👨‍💼 Quản Lý Năng Suất & Thiết Lập",
      ],
      index=0,
  )

  if "1." in tree_node_category:
    selected_sub_leaf = st.radio(
        "📄 Chọn biểu đồ báo cáo:",
        [
            "📋 Báo Cáo Vật Tư (QA32)",
            "⚙️ Báo Cáo Cơ Khí (3012)",
            "🔌 Báo Cáo TU/TI (3011)",
            "⚡ Báo Cáo Công Tơ (3013, 3016)",
        ],
        index=0,
    )
  elif "2." in tree_node_category:
    selected_sub_leaf = st.radio(
        "📄 Chọn bảng chi tiết:",
        [
            "📦 Chi Tiết Vật Tư (QA32)",
            "⚙️ Chi Tiết Lệnh SX (COOIS)",
        ],
        index=0,
    )
  else:
    selected_sub_leaf = st.radio(
        "📄 Chọn quản lý & thiết lập:",
        [
            "👨‍💼 Báo Cáo Năng Suất Cá Nhân QC",
            "⚙️ Quản Lý Công Việc Con Theo Xưởng",
            "🚨 Báo Cáo Sai Hỏng Chi Tiết & DM Lỗi",
        ],
        index=0,
    )

df_qa32, df_coois = load_data(tu_date, den_date)


# ================= 5. HÀM COOIS TÍCH HỢP TỰ ĐỘNG SỐ LIỆU SAI HỎNG =================
def render_coois_tab_layout(phan_he_code, title_text):
  render_page_header(
      title_text,
      "Tiến độ sản xuất, tỷ lệ giao hàng hoàn thành, sản lượng và phân tích"
      " sai hỏng theo từng dòng sản phẩm",
      f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
  )

  df_sub = (
      df_coois[df_coois["phan_he"] == phan_he_code]
      if not df_coois.empty
      else pd.DataFrame()
  )
  if df_sub.empty:
    st.info(f"💡 Chưa có dữ liệu sản xuất cho phân hệ {title_text}.")
    return

  # Truy vấn số liệu QC Sai Hỏng từ CSDL Mobile
  try:
    conn = get_db_connection()
    df_qc_sub = pd.read_sql_query(
        "SELECT so_lot, ma_vt, ten_vt, sl_kiem, sl_khong_dat, kieu_loi,"
        " cong_viec_con, nguoi_kiem, ngay_kiem FROM tb_qc_dau_vao WHERE loai_qc"
        " = 'SAN_XUAT' AND (sl_khong_dat > 0 OR (kieu_loi IS NOT NULL AND"
        " kieu_loi != '')) ORDER BY ngay_kiem DESC",
        conn,
    )
    conn.close()
  except Exception:
    df_qc_sub = pd.DataFrame()

  # Ghép nối Lô/Lệnh COOIS với Dòng sản phẩm (mat_prefix)
  if not df_qc_sub.empty and not df_sub.empty:
    order_map = dict(
        zip(df_sub["so_lenh"].astype(str), df_sub["mat_prefix"].astype(str))
    )
    code_map = dict(
        zip(df_sub["ma_tp"].astype(str), df_sub["mat_prefix"].astype(str))
    )

    df_qc_sub["mat_prefix"] = df_qc_sub["so_lot"].astype(str).map(order_map)
    df_qc_sub["mat_prefix"] = (
        df_qc_sub["mat_prefix"]
        .fillna(df_qc_sub["ma_vt"].astype(str).map(code_map))
        .fillna("Khác")
    )

    # Lọc các ca lỗi thuộc phân hệ xưởng này
    allowed_orders = set(df_sub["so_lenh"].astype(str).unique())
    allowed_codes = set(df_sub["ma_tp"].astype(str).unique())
    df_qc_sub = df_qc_sub[
        df_qc_sub["so_lot"].astype(str).isin(allowed_orders)
        | df_qc_sub["ma_vt"].astype(str).isin(allowed_codes)
    ].copy()
  else:
    df_qc_sub = pd.DataFrame()

  months_labels = [f"T{i}" for i in range(1, 13)]
  m_comp_qty, m_uncomp_qty = [0.0] * 12, [0.0] * 12
  m_tot_orders, m_uncomp_orders = [0] * 12, [0] * 12
  tot_qty_all, deliv_qty_all = 0.0, 0.0

  for _, r in df_sub.iterrows():
    try:
      m_idx = (
          datetime.strptime(
              str(r["ngay_lenh_dt"]).split()[0], "%Y-%m-%d"
          ).month
          - 1
      )
    except Exception:
      m_idx = 0
    if not (0 <= m_idx < 12):
      m_idx = 0
    sl_t, sl_h = float(r["sl_tong"]), float(r["sl_ht"])
    uncomp_q = max(0.0, sl_t - sl_h)
    tot_qty_all += sl_t
    deliv_qty_all += sl_h
    m_comp_qty[m_idx] += sl_h
    m_uncomp_qty[m_idx] += uncomp_q
    m_tot_orders[m_idx] += 1
    if sl_h < sl_t:
      m_uncomp_orders[m_idx] += 1

  m_comp_orders = [m_tot_orders[i] - m_uncomp_orders[i] for i in range(12)]
  title_clean = clean_emoji(title_text)

  rem_qty_kpi = max(0.0, tot_qty_all - deliv_qty_all)
  pct_deliv_kpi = (
      (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0.0
  )
  total_orders_kpi = sum(m_tot_orders)
  tot_defect_qty = (
      df_qc_sub["sl_khong_dat"].sum() if not df_qc_sub.empty else 0.0
  )
  pct_defect_kpi = (
      (tot_defect_qty / deliv_qty_all * 100.0) if deliv_qty_all > 0 else 0.0
  )

  render_kpi_cards([
      {
          "label": "Tổng Kế Hoạch",
          "value": f"{int(tot_qty_all):,}",
          "icon": "🎯",
          "color": COLOR_PRIMARY,
          "sub": f"{int(total_orders_kpi):,} lệnh sản xuất",
      },
      {
          "label": "Đã Hoàn Thành",
          "value": f"{int(deliv_qty_all):,}",
          "icon": "✅",
          "color": COLOR_SUCCESS,
          "sub": f"Tỷ lệ SL: {pct_deliv_kpi:.1f}%",
      },
      {
          "label": "TỔNG SỐ LƯỢNG LỖI",
          "value": f"{int(tot_defect_qty):,}",
          "icon": "⚠️",
          "color": COLOR_DANGER,
          "sub": f"Tỷ lệ sai hỏng: {pct_defect_kpi:.2f}%",
      },
      {
          "label": "Còn Lại Chưa Xong",
          "value": f"{int(rem_qty_kpi):,}",
          "icon": "⏳",
          "color": COLOR_WARNING,
          "sub": f"{int(sum(m_uncomp_orders)):,} lệnh chưa xong",
      },
  ])

  PLOT_HEIGHT = 380

  col1, col2 = st.columns([2.1, 1.0])
  with col1:
    chart_card_open(
        f"Sản Lượng & Tỷ Lệ Hoàn Thành — {title_clean}", "Theo tháng"
    )
    fig1 = make_subplots(specs=[[{"secondary_y": True}]])

    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_comp_qty,
            name="SL Hoàn Thành",
            marker_color=COLOR_SUCCESS,
            text=[f"{int(v):,}" if v > 0 else "" for v in m_comp_qty],
            textposition="inside",
            textfont=dict(size=9.5, family=PLOTLY_FONT, color="#FFFFFF"),
        ),
        secondary_y=False,
    )
    fig1.add_trace(
        go.Bar(
            x=months_labels,
            y=m_uncomp_qty,
            name="SL Chưa Xong",
            marker_color=COLOR_WARNING,
            base=m_comp_qty,
            text=[f"{int(v):,}" if v > 0 else "" for v in m_uncomp_qty],
            textposition="inside",
            textfont=dict(size=9.5, family=PLOTLY_FONT, color="#FFFFFF"),
        ),
        secondary_y=False,
    )
    pct_hoanthanh_m = [
        (m_comp_qty[i] / (m_comp_qty[i] + m_uncomp_qty[i]) * 100.0)
        if (m_comp_qty[i] + m_uncomp_qty[i]) > 0
        else None
        for i in range(12)
    ]
    fig1.add_trace(
        go.Scatter(
            x=months_labels,
            y=pct_hoanthanh_m,
            name="% Hoàn Thành",
            mode="lines+markers+text",
            line=dict(color=COLOR_PRIMARY, width=2.5),
            marker=dict(size=6, color=COLOR_PRIMARY),
            text=[f"{v:.0f}%" if v is not None else "" for v in pct_hoanthanh_m],
            textposition="top center",
            textfont=dict(size=10, family=PLOTLY_FONT, color=COLOR_PRIMARY),
            connectgaps=False,
        ),
        secondary_y=True,
    )

    fig1.update_layout(
        barmode="stack",
        margin=dict(l=30, r=20, t=8, b=55),
        height=PLOT_HEIGHT,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.22,
            xanchor="center",
            x=0.5,
            font=dict(size=11, family=PLOTLY_FONT, color="#6B7280"),
        ),
    )
    fig1.update_xaxes(
        showgrid=False, tickfont=dict(size=11, family=PLOTLY_FONT, color="#6B7280")
    )
    fig1.update_yaxes(
        title_text="← Sản Lượng",
        title_font=dict(size=12, color=COLOR_SUCCESS),
        tickformat="~s",
        secondary_y=False,
        showgrid=True,
        gridcolor=PLOTLY_GRID,
    )
    fig1.update_yaxes(
        title_text="% Hoàn Thành →",
        title_font=dict(size=12, color=COLOR_PRIMARY),
        tickformat=".0f",
        ticksuffix="%",
        range=[0, 105],
        secondary_y=True,
        showgrid=False,
    )

    st.plotly_chart(
        fig1,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig1_{phan_he_code}",
    )
    chart_card_close()

  with col2:
    chart_card_open("Tỷ Lệ Hoàn Thành Tổng Quan")
    pct_deliv = (deliv_qty_all / tot_qty_all * 100) if tot_qty_all > 0 else 0
    pct_rem = 100.0 - pct_deliv if tot_qty_all > 0 else 0.0

    fig2 = go.Figure(
        data=[
            go.Pie(
                labels=["Hoàn thành", "Chưa xong"],
                values=[deliv_qty_all, rem_qty_kpi],
                hole=0.6,
                marker=dict(
                    colors=[COLOR_SUCCESS, COLOR_WARNING],
                    line=dict(color="#FFFFFF", width=3),
                ),
                text=[
                    f"<b>Hoàn thành</b><br>{pct_deliv:.1f}%<br>({int(deliv_qty_all):,})",
                    f"<b>Chưa xong</b><br>{pct_rem:.1f}%<br>({int(rem_qty_kpi):,})",
                ],
                textinfo="text",
                textposition="outside",
                direction="clockwise",
                sort=False,
            )
        ]
    )
    fig2.update_layout(
        margin=dict(l=95, r=95, t=30, b=30),
        height=PLOT_HEIGHT,
        paper_bgcolor="#FFFFFF",
        showlegend=False,
    )
    st.plotly_chart(
        fig2,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig2_{phan_he_code}",
    )
    chart_card_close()

  # HÀNG 2: BỘ LỌC DÒNG SP VÀ PHÂN TÍCH SAI HỎNG TƯƠNG ỨNG
  sub_5 = (
      df_sub[
          df_sub["ma_tp"]
          .astype(str)
          .str.split(".")
          .str[0]
          .str.lstrip("0")
          .str.startswith("5")
      ].copy()
      if not df_sub.empty
      else pd.DataFrame()
  )
  raw_fams = (
      [
          str(x).strip()
          for x in sub_5["mat_prefix"].unique()
          if pd.notna(x)
          and str(x).strip()
          and str(x).strip().lower() not in ["none", "nan"]
      ]
      if not sub_5.empty
      else []
  )
  available_fams = ["Tất cả dòng sản phẩm"] + sorted(list(set(raw_fams)))
  sel_fam = st.selectbox(
      "🎯 Chọn Dòng SP (Đầu 5):", available_fams, key=f"cb_{phan_he_code}"
  )

  col3, col4 = st.columns([1, 1])

  with col3:
    chart_card_open(f"Sản Lượng — Dòng: {clean_emoji(sel_fam)}")
    m3_qty = [0.0] * 12
    if not sub_5.empty:
      sub_5_df = sub_5.copy()
      sub_5_df["month"] = pd.to_datetime(
          sub_5_df["ngay_lenh_dt"], errors="coerce"
      ).dt.month
      sub_filtered = (
          sub_5_df[sub_5_df["mat_prefix"] == sel_fam]
          if sel_fam != "Tất cả dòng sản phẩm"
          else sub_5_df
      )
      for _, r in sub_filtered.iterrows():
        m_val = r["month"]
        if pd.notna(m_val) and 1 <= int(m_val) <= 12:
          m3_qty[int(m_val) - 1] += float(r["sl_ht"])

    fig3 = go.Figure()
    fig3.add_trace(
        go.Bar(
            x=months_labels,
            y=m3_qty,
            name="SL Sản Xuất",
            marker=dict(color=COLOR_PRIMARY, cornerradius=6),
            text=[f"{int(v):,}" if v > 0 else "" for v in m3_qty],
            textposition="outside",
            textfont=dict(color=COLOR_PRIMARY, size=11, family=PLOTLY_FONT),
        )
    )

    fig3.update_layout(
        margin=dict(l=30, r=20, t=8, b=36),
        height=PLOT_HEIGHT - 40,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        showlegend=False,
    )
    st.plotly_chart(
        fig3,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig3_{phan_he_code}",
    )
    chart_card_close()

  with col4:
    chart_card_open("Tổng Sản Lượng Cả Năm Các Mã Đầu 5")
    if not sub_5.empty:
      summary_fams = (
          sub_5.groupby("mat_prefix")[["sl_ht"]].sum().reset_index()
      )
      fams_x = [
          str(val)
          for val in summary_fams["mat_prefix"].tolist()
          if pd.notna(val)
          and str(val).strip()
          and str(val).strip().lower() not in ["none", "nan"]
      ]
      if not fams_x:
        fams_x, deliv_fams = ["Trống"], [0.0]
      else:
        summary_fams = summary_fams[summary_fams["mat_prefix"].isin(fams_x)]
        fams_x, deliv_fams = (
            summary_fams["mat_prefix"].tolist(),
            summary_fams["sl_ht"].values,
        )
    else:
      fams_x, deliv_fams = ["Không có SP"], [0.0]

    if len(fams_x) > 1:
      pairs = sorted(zip(fams_x, deliv_fams), key=lambda p: p[1], reverse=True)
      fams_x = [p[0] for p in pairs]
      deliv_fams = [p[1] for p in pairs]

    bar_colors = [
        DISTINCT_COLORS[i % len(DISTINCT_COLORS)] for i in range(len(fams_x))
    ]

    fig4 = go.Figure()
    fig4.add_trace(
        go.Bar(
            x=fams_x,
            y=deliv_fams,
            marker=dict(color=bar_colors, cornerradius=6),
            text=[f"{int(v):,}" if v > 0 else "" for v in deliv_fams],
            textposition="outside",
            textfont=dict(color=bar_colors, size=11, family=PLOTLY_FONT),
        )
    )

    fig4.update_layout(
        margin=dict(l=30, r=20, t=8, b=36),
        height=PLOT_HEIGHT - 40,
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        showlegend=False,
    )
    fig4.update_yaxes(type="log", dtick=1)
    st.plotly_chart(
        fig4,
        use_container_width=True,
        config={"displayModeBar": False},
        key=f"coois_fig4_{phan_he_code}",
    )
    chart_card_close()

  # HÀNG 3: BỔ SUNG BIỂU ĐỒ VÀ BẢNG SAI HỎNG ỨNG VỚI DÒNG SẢN PHẨM ĐÃ CHỌN
  st.markdown(
      f"#### 🚨 PHÂN TÍCH SAI HỎNG DÒNG SẢN PHẨM: {clean_emoji(sel_fam)}"
  )
  col_sh_fam1, col_sh_fam2 = st.columns([1.1, 1.5])

  df_qc_fam = df_qc_sub.copy()
  if sel_fam != "Tất cả dòng sản phẩm" and not df_qc_fam.empty:
    df_qc_fam = df_qc_fam[df_qc_fam["mat_prefix"] == sel_fam]

  with col_sh_fam1:
    chart_card_open(f"Top Kiểu Lỗi — Dòng: {clean_emoji(sel_fam)}")
    if not df_qc_fam.empty and "kieu_loi" in df_qc_fam.columns:
      df_err_agg = (
          df_qc_fam.groupby("kieu_loi")["sl_khong_dat"].sum().reset_index()
      )
      df_err_agg = df_err_agg[df_err_agg["kieu_loi"] != ""].sort_values(
          by="sl_khong_dat", ascending=True
      )

      if not df_err_agg.empty:
        fig_err_fam = go.Figure(
            go.Bar(
                x=df_err_agg["sl_khong_dat"],
                y=df_err_agg["kieu_loi"],
                orientation="h",
                marker=dict(color=COLOR_DANGER),
                text=[f"{v:,.0f}" for v in df_err_agg["sl_khong_dat"]],
                textposition="outside",
            )
        )
        fig_err_fam.update_layout(
            margin=dict(l=10, r=40, t=10, b=10),
            height=260,
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
        )
        st.plotly_chart(
            fig_err_fam,
            use_container_width=True,
            key=f"err_fam_chart_{phan_he_code}",
        )
      else:
        st.success("🎉 Không ghi nhận kiểu sai hỏng nào!")
    else:
      st.success("🎉 Không có ca báo lỗi nào thuộc dòng SP này!")
    chart_card_close()

  with col_sh_fam2:
    chart_card_open(f"Nhật Ký Sai Hỏng Chi Tiết — Dòng: {clean_emoji(sel_fam)}")
    if not df_qc_fam.empty:
      df_qc_fam_disp = df_qc_fam[[
          "ngay_kiem",
          "nguoi_kiem",
          "so_lot",
          "ma_vt",
          "ten_vt",
          "cong_viec_con",
          "kieu_loi",
          "sl_khong_dat",
      ]].copy()
      df_qc_fam_disp.columns = [
          "Thời Gian",
          "Người Kiểm",
          "Lệnh SX",
          "Mã SP",
          "Tên SP",
          "Công Đoạn",
          "Kiểu Sai Hỏng",
          "SL Lỗi",
      ]
      st.dataframe(
          df_qc_fam_disp, use_container_width=True, hide_index=True, height=230
      )
    else:
      st.info("💡 Chưa có nhật ký báo lỗi cho dòng sản phẩm này.")
    chart_card_close()

  # ================= XUẤT ẢNH MATPLOTLIB ẨN CHO EXCEL =================
  fig1_mpl, ax_m1 = plt.subplots(figsize=(8, 3.2), dpi=200)
  ax_m1_twin = ax_m1.twinx()
  ax_m1.bar(
      np.arange(12) - 0.18,
      m_comp_orders,
      width=0.35,
      color=COLOR_PRIMARY,
      label="Lệnh HT",
  )
  ax_m1.bar(
      np.arange(12) - 0.18,
      m_uncomp_orders,
      width=0.35,
      bottom=m_comp_orders,
      color=COLOR_DANGER,
      label="Lệnh Chưa Xong",
  )
  ax_m1_twin.bar(
      np.arange(12) + 0.18,
      m_comp_qty,
      width=0.35,
      color=COLOR_SUCCESS,
      label="SL HT",
  )
  ax_m1_twin.bar(
      np.arange(12) + 0.18,
      m_uncomp_qty,
      width=0.35,
      bottom=m_comp_qty,
      color=COLOR_WARNING,
      label="SL Chưa Xong",
  )
  ax_m1.set_title(f"TIẾN ĐỘ SẢN XUẤT - {title_clean}", fontweight="bold")
  ax_m1.set_xticks(np.arange(12))
  ax_m1.set_xticklabels(months_labels)
  ax_m1_twin.set_yscale("symlog", linthresh=100)
  plt.close(fig1_mpl)

  fig2_mpl, ax_m2 = plt.subplots(figsize=(4, 3.2), dpi=200)
  ax_m2.set_aspect("equal")
  if tot_qty_all > 0:
    ax_m2.pie(
        [deliv_qty_all, rem_qty_kpi],
        labels=["Hoàn thành", "Chưa xong"],
        colors=[COLOR_SUCCESS, COLOR_WARNING],
        autopct="%1.1f%%",
        startangle=140,
        wedgeprops=dict(width=0.35, edgecolor="white"),
    )
  ax_m2.set_title("TỶ LỆ HOÀN THÀNH TỔNG QUAN", fontweight="bold")
  plt.close(fig2_mpl)

  fig3_mpl, ax_m3 = plt.subplots(figsize=(6, 3.0), dpi=200)
  ax_m3.bar(np.arange(12), m3_qty, width=0.45, color=COLOR_PRIMARY)
  ax_m3.set_xticks(np.arange(12))
  ax_m3.set_xticklabels(months_labels)
  ax_m3.set_yscale("symlog", linthresh=100)
  ax_m3.set_title(f"SẢN LƯỢNG - DÒNG: {clean_emoji(sel_fam)}", fontweight="bold")
  plt.close(fig3_mpl)

  fig4_mpl, ax_m4 = plt.subplots(figsize=(6, 3.0), dpi=200)
  ax_m4.bar(np.arange(len(fams_x)), deliv_fams, width=0.45, color=bar_colors)
  ax_m4.set_xticks(np.arange(len(fams_x)))
  ax_m4.set_xticklabels(fams_x)
  ax_m4.set_yscale("symlog", linthresh=100)
  ax_m4.set_title("TỔNG SẢN LƯỢNG CẢ NĂM CÁC MÃ ĐẦU 5", fontweight="bold")
  plt.close(fig4_mpl)

  # ================= BẢNG SỐ LIỆU TỔNG HỢP & NÚT EXCEL =================
  pct_orders_m = [
      ((m_comp_orders[i] / m_tot_orders[i]) * 100.0)
      if m_tot_orders[i] > 0
      else 0.0
      for i in range(12)
  ]
  pct_qty_m = [
      ((m_comp_qty[i] / (m_comp_qty[i] + m_uncomp_qty[i])) * 100.0)
      if (m_comp_qty[i] + m_uncomp_qty[i]) > 0
      else 0.0
      for i in range(12)
  ]
  total_m3 = sum(m3_qty)
  pct_fam_contrib = [
      ((m3_qty[i] / total_m3) * 100.0) if total_m3 > 0 else 0.0
      for i in range(12)
  ]
  total_yr = sum(deliv_fams)
  pct_yr_share = [
      ((v / total_yr) * 100.0) if total_yr > 0 else 0.0 for v in deliv_fams
  ]

  df_monthly_summary = pd.DataFrame({
      "Tháng": months_labels,
      "Lệnh Hoàn Thành": m_comp_orders,
      "Lệnh Chưa Xong": m_uncomp_orders,
      "Tỷ Lệ HT Lệnh (%)": pct_orders_m,
      "SL Hoàn Thành": [int(v) for v in m_comp_qty],
      "SL Chưa Xong": [int(v) for v in m_uncomp_qty],
      "Tỷ Lệ HT SL (%)": pct_qty_m,
  })
  df_plan_summary = pd.DataFrame({
      "Chỉ Tiêu": [
          "Tổng Kế Hoạch",
          "Đã Giao Hoàn Thành",
          "Còn Lại Chưa Xong",
      ],
      "Số Lượng": [
          int(tot_qty_all),
          int(deliv_qty_all),
          int(rem_qty_kpi),
      ],
      "Tỷ Lệ Cơ Cấu (%)": [
          100.0,
          (deliv_qty_all / tot_qty_all * 100.0) if tot_qty_all > 0 else 0.0,
          (rem_qty_kpi / tot_qty_all * 100.0) if tot_qty_all > 0 else 0.0,
      ],
  })
  df_family_summary = pd.DataFrame({
      "Tháng": months_labels,
      f"SL Sản Xuất ({sel_fam})": [int(v) for v in m3_qty],
      "Tỷ Lệ Đóng Góp Tháng (%)": pct_fam_contrib,
  })
  df_year_summary = pd.DataFrame({
      "Mã / Dòng SP": fams_x,
      "Tổng SL Hoàn Thành Cả Năm": [int(v) for v in deliv_fams],
      "Tỷ Lệ Cơ Cấu (%)": pct_yr_share,
  })

  st.markdown(
      "<hr style='margin: 8px 0 18px 0; border-color: #E4E8F0;'>",
      unsafe_allow_html=True,
  )
  col_hdr, col_btn = st.columns([2.5, 1.2])
  with col_hdr:
    render_section_heading(
        "📑 Bảng Số Liệu Tổng Hợp & Tỷ Lệ Hiệu Chỉnh —"
        f" {title_clean.upper()}"
    )
  excel_bytes = generate_print_ready_excel(
      phan_he_code,
      title_clean,
      tu_date,
      den_date,
      df_monthly_summary,
      df_plan_summary,
      df_family_summary,
      df_year_summary,
      fig1_mpl,
      fig2_mpl,
      fig3_mpl,
      fig4_mpl,
  )
  with col_btn:
    st.download_button(
        label="📥 XUẤT BÁO CÁO EXCEL CHUYÊN NGHIỆP",
        data=excel_bytes,
        file_name=(
            f"BaoCao_EMIC_{phan_he_code}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        ),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

  t_col1, t_col2 = st.columns([1.6, 1.0])
  with t_col1:
    st.markdown("##### 1. Tiến Độ Sản Xuất Theo Tháng")
    styled_monthly = (
        df_monthly_summary.style.background_gradient(
            subset=["Lệnh Hoàn Thành"], cmap="Blues"
        )
        .background_gradient(subset=["Lệnh Chưa Xong"], cmap="Reds")
        .background_gradient(subset=["SL Hoàn Thành"], cmap="Greens")
        .background_gradient(subset=["SL Chưa Xong"], cmap="Oranges")
    )
    st.dataframe(
        styled_monthly,
        column_config={
            "Tỷ Lệ HT Lệnh (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ HT Lệnh", format="%.1f%%", min_value=0, max_value=100
            ),
            "Tỷ Lệ HT SL (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ HT SL", format="%.1f%%", min_value=0, max_value=100
            ),
            "SL Hoàn Thành": st.column_config.NumberColumn(
                "SL Hoàn Thành", format="%d"
            ),
            "SL Chưa Xong": st.column_config.NumberColumn(
                "SL Chưa Xong", format="%d"
            ),
        },
        use_container_width=True,
        hide_index=True,
    )
  with t_col2:
    st.markdown("##### 2. Tổng Quan Chỉ Tiêu Kế Hoạch")
    styled_plan = df_plan_summary.style.background_gradient(
        subset=["Số Lượng"], cmap="Blues"
    )
    st.dataframe(
        styled_plan,
        column_config={
            "Tỷ Lệ Cơ Cấu (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ Cơ Cấu", format="%.1f%%", min_value=0, max_value=100
            ),
            "Số Lượng": st.column_config.NumberColumn("Số Lượng", format="%d"),
        },
        use_container_width=True,
        hide_index=True,
    )

  t_col3, t_col4 = st.columns([1.0, 1.0])
  with t_col3:
    st.markdown(f"##### 3. Chi Tiết Sản Lượng Dòng SP ({sel_fam})")
    styled_family = df_family_summary.style.background_gradient(
        subset=[f"SL Sản Xuất ({sel_fam})"], cmap="Purples"
    )
    st.dataframe(
        styled_family,
        column_config={
            "Tỷ Lệ Đóng Góp Tháng (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ Đóng Góp Tháng", format="%.1f%%", min_value=0, max_value=100
            ),
            f"SL Sản Xuất ({sel_fam})": st.column_config.NumberColumn(
                f"SL Sản Xuất ({sel_fam})", format="%d"
            ),
        },
        use_container_width=True,
        hide_index=True,
    )
  with t_col4:
    st.markdown("##### 4. Tổng Sản Lượng Cả Năm Các Mã Đầu 5")
    styled_year = df_year_summary.style.background_gradient(
        subset=["Tổng SL Hoàn Thành Cả Năm"], cmap="BuGn"
    )
    st.dataframe(
        styled_year,
        column_config={
            "Tỷ Lệ Cơ Cấu (%)": st.column_config.ProgressColumn(
                "Tỷ Lệ Cơ Cấu", format="%.1f%%", min_value=0, max_value=100
            ),
            "Tổng SL Hoàn Thành Cả Năm": st.column_config.NumberColumn(
                "Tổng SL Cả Năm", format="%d"
            ),
        },
        use_container_width=True,
        hide_index=True,
    )


# ================= 6. HIỂN THỊ DỮ LIỆU THEO LÁ CÂY ĐÃ CHỌN =================

if selected_sub_leaf == "📋 Báo Cáo Vật Tư (QA32)":
  render_page_header(
      "📋 Báo Cáo Chất Lượng Vật Tư Nhập Mua (QA32)",
      "Tổng hợp tiến độ kiểm tra, tỷ lệ đạt/đặc nhượng và danh sách vật tư bị"
      " block",
      f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
  )

  if df_qa32.empty:
    st.info("💡 Chưa có dữ liệu QA32 trong khoảng thời gian đã chọn.")
  else:
    months_labels = [f"T{i}" for i in range(1, 13)]
    ud01_m, ud02_m, ud03_m, uninspected_m = (
        [0] * 12,
        [0] * 12,
        [0] * 12,
        [0] * 12,
    )
    ft_qty_m, by_inspected_m, by_uninspected_m = (
        [0.0] * 12,
        [0.0] * 12,
        [0.0] * 12,
    )
    total_ca_block, total_ft_all = 0.0, 0.0
    top_block_dict = {}

    for _, r in df_qa32.iterrows():
      try:
        m_idx = (
            datetime.strptime(
                str(r["ngay_ve_dt"]).split()[0], "%Y-%m-%d"
            ).month
            - 1
        )
      except Exception:
        m_idx = 0
      if not (0 <= m_idx < 12):
        m_idx = 0

      st_clean = (
          str(r["xac_nhan_sap"]).strip().upper().replace(" ", "")
          if "xac_nhan_sap" in r and pd.notna(r["xac_nhan_sap"])
          else ""
      )
      ft_val = (
          float(r["ft_qty"])
          if ("ft_qty" in r and pd.notna(r["ft_qty"]))
          else 0.0
      )
      by_val = (
          float(r["by_sample"])
          if ("by_sample" in r and pd.notna(r["by_sample"]))
          else 0.0
      )
      ca_val = (
          float(r["ca_qty"])
          if ("ca_qty" in r and pd.notna(r["ca_qty"]))
          else 0.0
      )

      ft_qty_m[m_idx] += ft_val
      total_ft_all += ft_val
      total_ca_block += ca_val

      is_uninspected = (
          "CHƯA" in st_clean or not st_clean or st_clean in ["NAN", "NONE", "❌CHƯAXN"]
      )
      is_ud02 = any(
          k in st_clean for k in ["02", "UD2", "ĐẶCNHƯỢNG", "DACNHUONG"]
      )
      is_ud03 = any(
          k in st_clean
          for k in ["03", "UD3", "TRẢLẠI", "TRALAI", "TỪCHỐI", "TUCHOI", "KHÔNG", "KHONG"]
      )
      is_ud01 = any(k in st_clean for k in ["01", "UD1", "ĐẠT", "DAT"]) and not (
          is_ud02 or is_ud03
      )

      if is_uninspected:
        uninspected_m[m_idx] += 1
        by_uninspected_m[m_idx] += by_val
      else:
        by_inspected_m[m_idx] += by_val
        if is_ud02:
          ud02_m[m_idx] += 1
        elif is_ud03:
          ud03_m[m_idx] += 1
        else:
          ud01_m[m_idx] += 1

      ma_vt_str = (
          str(r["ma_vt"]).strip()
          if "ma_vt" in r and pd.notna(r["ma_vt"])
          else ""
      )
      ten_vt_str = (
          str(r["ten_vt"]).strip()
          if "ten_vt" in r and pd.notna(r["ten_vt"])
          else ""
      )
      ncc_str = (
          str(r["ncc"]).strip() if "ncc" in r and pd.notna(r["ncc"]) else ""
      )

      if (
          is_ud02
          or is_ud03
          or ca_val > 0
          or (st_clean and not is_ud01 and not is_uninspected)
      ):
        if "VIHA" in ncc_str.upper():
          key = ("Mặt số công tơ", ncc_str if ncc_str else "Cty TNHH CN VIHA")
          ma_display, ten_display = "Mặt số công tơ", "Mặt số công tơ"
        else:
          key = (ma_vt_str, ncc_str)
          ma_display, ten_display = ma_vt_str, ten_vt_str

        if key not in top_block_dict:
          top_block_dict[key] = {
              "ma_vt": ma_display,
              "ten_vt": ten_display,
              "ncc": key[1],
              "ud02": 0,
              "ud03": 0,
              "ca_block": 0.0,
              "ft_total": 0.0,
          }
        if is_ud02:
          top_block_dict[key]["ud02"] += 1
        if is_ud03:
          top_block_dict[key]["ud03"] += 1
        top_block_dict[key]["ca_block"] += ca_val
        top_block_dict[key]["ft_total"] += ft_val

    total_ud01 = sum(ud01_m)
    total_ud02 = sum(ud02_m)
    total_ud03 = sum(ud03_m)
    total_uninspected = sum(uninspected_m)
    total_lots = total_ud01 + total_ud02 + total_ud03 + total_uninspected
    pct_ud01 = (total_ud01 / total_lots * 100) if total_lots > 0 else 0.0
    total_by_all = sum(by_inspected_m) + sum(by_uninspected_m)

    render_kpi_cards([
        {
            "label": "Tổng Vật Tư Về (FT)",
            "value": f"{int(total_ft_all):,}",
            "icon": "📦",
            "color": COLOR_PRIMARY,
            "sub": f"{int(total_by_all):,} mẫu đã kiểm (BY)",
        },
        {
            "label": "Tỷ Lệ Đạt (UD 01)",
            "value": f"{pct_ud01:.1f}%",
            "icon": "✅",
            "color": COLOR_SUCCESS,
            "sub": f"{int(total_ud01):,} / {int(total_lots):,} lô",
        },
        {
            "label": "Đặc Nhượng / Trả Lại",
            "value": f"{int(total_ud02 + total_ud03):,}",
            "icon": "⚠️",
            "color": COLOR_WARNING,
            "sub": f"UD02: {int(total_ud02):,} · UD03: {int(total_ud03):,}",
        },
        {
            "label": "SL Bị Block (CA)",
            "value": f"{int(total_ca_block):,}",
            "icon": "🚫",
            "color": COLOR_DANGER,
            "sub": f"{len(top_block_dict):,} mã vật tư liên quan",
        },
    ])

    col1, col2 = st.columns([2.1, 1.0])
    with col1:
      chart_card_open(
          "Số Lượng Lệnh Kiểm & Tổng Vật Tư Về / Số Mẫu Kiểm", "Theo tháng"
      )
      fig1 = make_subplots(specs=[[{"secondary_y": True}]])
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud01_m,
              name="UD 01 (Đạt)",
              marker_color=COLOR_SUCCESS,
          ),
          secondary_y=False,
      )
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud02_m,
              name="UD 02 (Đặc nhượng)",
              marker_color=COLOR_WARNING,
          ),
          secondary_y=False,
      )
      fig1.add_trace(
          go.Bar(
              x=months_labels,
              y=ud03_m,
              name="UD 03 (Trả lại)",
              marker_color=COLOR_DANGER,
          ),
          secondary_y=False,
      )

      total_by = [by_inspected_m[i] + by_uninspected_m[i] for i in range(12)]
      fig1.add_trace(
          go.Scatter(
              x=months_labels,
              y=total_by,
              name="Số mẫu phải kiểm (BY)",
              mode="lines+markers+text",
              line=dict(color=COLOR_PURPLE, width=2, dash="dash"),
              text=[f"{int(v):,}" if v > 0 else "" for v in total_by],
              textposition="top center",
          ),
          secondary_y=True,
      )
      fig1.add_trace(
          go.Scatter(
              x=months_labels,
              y=ft_qty_m,
              name="Tổng số hàng về (FT)",
              mode="lines+markers+text",
              line=dict(color=COLOR_PRIMARY, width=2),
              text=[f"{int(v):,}" if v > 0 else "" for v in ft_qty_m],
              textposition="bottom center",
          ),
          secondary_y=True,
      )

      fig1.update_layout(
          barmode="stack",
          margin=dict(l=30, r=20, t=8, b=55),
          height=380,
          paper_bgcolor="#FFFFFF",
          plot_bgcolor="#FFFFFF",
      )
      fig1.update_yaxes(
          title_text="← Số Lượng Lệnh", secondary_y=False, showgrid=True
      )
      fig1.update_yaxes(
          title_text="Vật Tư / Mẫu (Log) →",
          type="log",
          dtick=1,
          secondary_y=True,
          showgrid=False,
      )
      st.plotly_chart(
          fig1,
          use_container_width=True,
          config={"displayModeBar": False},
          key="leaf_vt_fig1",
      )
      chart_card_close()

    with col2:
      chart_card_open("Tỷ Lệ Vật Tư Đạt vs Bị Block Lỗi")
      ok_cnt = max(0.0, total_ft_all - total_ca_block)
      pct_ok = (ok_cnt / total_ft_all * 100) if total_ft_all > 0 else 0
      pct_block = (
          (total_ca_block / total_ft_all * 100) if total_ft_all > 0 else 0
      )

      fig2 = go.Figure(
          data=[
              go.Pie(
                  labels=["Vật tư Đạt", "Bị Block (Lỗi)"],
                  values=[ok_cnt, total_ca_block],
                  hole=0.6,
                  marker=dict(
                      colors=[COLOR_SUCCESS, COLOR_DANGER],
                      line=dict(color="#FFFFFF", width=3),
                  ),
                  text=[
                      f"<b>Vật tư Đạt</b><br>{pct_ok:.1f}%<br>({ok_cnt:,.0f})",
                      f"<b>Bị Block (Lỗi)</b><br>{pct_block:.1f}%<br>({total_ca_block:,.0f})",
                  ],
                  textinfo="text",
                  textposition="outside",
                  direction="clockwise",
                  sort=False,
              )
          ]
      )
      fig2.update_layout(
          margin=dict(l=90, r=90, t=30, b=30),
          height=380,
          paper_bgcolor="#FFFFFF",
          showlegend=False,
      )
      st.plotly_chart(
          fig2,
          use_container_width=True,
          config={"displayModeBar": False},
          key="leaf_vt_fig2",
      )
      chart_card_close()

    sorted_blocks = sorted(
        top_block_dict.values(),
        key=lambda x: (x["ud03"] + x["ud02"], x["ca_block"], x["ft_total"]),
        reverse=True,
    )
    if sorted_blocks:
      col_rank, col_table = st.columns([1.0, 1.6])
      supplier_agg = {}
      for item in sorted_blocks:
        ncc_name = item["ncc"].strip() if item["ncc"] else "Không rõ NCC"
        supplier_agg[ncc_name] = (
            supplier_agg.get(ncc_name, 0.0) + item["ca_block"]
        )
      top_suppliers = sorted(
          supplier_agg.items(), key=lambda x: x[1], reverse=True
      )[:8]

      with col_rank:
        if top_suppliers:
          sup_names = [clean_emoji(s[0])[:28] for s in top_suppliers][::-1]
          sup_vals = [s[1] for s in top_suppliers][::-1]
          chart_card_open("Top Nhà Cung Cấp Bị Block Nhiều Nhất")
          fig_sup = go.Figure(
              go.Bar(
                  x=sup_vals,
                  y=sup_names,
                  orientation="h",
                  marker=dict(color=COLOR_DANGER),
                  text=[f"{v:,.0f}" for v in sup_vals],
                  textposition="outside",
              )
          )
          fig_sup.update_layout(
              margin=dict(l=10, r=55, t=10, b=10),
              height=max(230, 32 * len(sup_names)),
              paper_bgcolor="#FFFFFF",
              plot_bgcolor="#FFFFFF",
              showlegend=False,
          )
          st.plotly_chart(
              fig_sup,
              use_container_width=True,
              config={"displayModeBar": False},
              key="leaf_vt_fig_sup",
          )
          chart_card_close()

      with col_table:
        chart_card_open(
            "🚨 Danh Sách Vật Tư Bị Block & UD02, UD03",
            f"{len(sorted_blocks)} mã vật tư",
        )
        df_block = pd.DataFrame(sorted_blocks)
        df_block["Tổng SL Block (CA) / SL Về"] = df_block.apply(
            lambda r: f"{r['ca_block']:,.0f} / {r['ft_total']:,.0f}", axis=1
        )
        df_block = df_block[[
            "ma_vt",
            "ten_vt",
            "ncc",
            "ud02",
            "ud03",
            "Tổng SL Block (CA) / SL Về",
        ]]
        df_block.columns = [
            "Mã Vật Tư",
            "Tên Vật Tư",
            "Nhà Cung Cấp",
            "Số Lượt UD 02",
            "Số Lượt UD 03",
            "Tổng SL Block (CA) / SL Về",
        ]
        st.dataframe(df_block, use_container_width=True, hide_index=True)
        chart_card_close()

elif selected_sub_leaf == "⚙️ Báo Cáo Cơ Khí (3012)":
  render_coois_tab_layout("CO_KHI", "⚙️ BÁO CÁO CƠ KHÍ (LỆNH 3012)")

elif selected_sub_leaf == "🔌 Báo Cáo TU/TI (3011)":
  render_coois_tab_layout("TU_TI", "🔌 BÁO CÁO TUTI (LỆNH 3011)")

elif selected_sub_leaf == "⚡ Báo Cáo Công Tơ (3013, 3016)":
  render_coois_tab_layout("CONG_TO", "⚡ BÁO CÁO CÔNG TƠ (LỆNH 3013, 3016)")

elif selected_sub_leaf == "📦 Chi Tiết Vật Tư (QA32)":
  render_page_header(
      "📦 Danh Sách Chi Tiết Vật Tư QA32",
      "Tra cứu chi tiết tình trạng lô hàng vật tư nhập mua từ SAP QA32",
      f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
  )
  if not df_qa32.empty:
    st.dataframe(df_qa32, use_container_width=True, hide_index=True)
  else:
    st.info("💡 Chưa có dữ liệu vật tư.")

elif selected_sub_leaf == "⚙️ Chi Tiết Lệnh SX (COOIS)":
  render_page_header(
      "⚙️ Danh Sách Chi Tiết Lệnh Sản Xuất COOIS",
      "Tra cứu chi tiết tình trạng hoàn thành các lệnh sản xuất từ SAP COOIS",
      f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
  )
  if not df_coois.empty:
    st.dataframe(df_coois, use_container_width=True, hide_index=True)
  else:
    st.info("💡 Chưa có dữ liệu lệnh sản xuất.")

elif selected_sub_leaf == "👨‍💼 Báo Cáo Năng Suất Cá Nhân QC":
  render_page_header(
      "👨‍💼 Báo Cáo Năng Suất Làm Việc Cá Nhân QC",
      "Theo dõi tổng lượt kiểm, sản lượng đạt/lỗi và nhật ký kiểm tra theo"
      " từng nhân sự",
      f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
  )
  try:
    conn = get_db_connection()
    df_qc_logs = pd.read_sql_query(
        "SELECT id, loai_qc, so_lot, ma_vt, ten_vt, ncc, sl_kiem, sl_khong_dat,"
        " sl_dat, ket_luan, nguoi_kiem, ngay_kiem, ghi_chu, kieu_loi,"
        " cong_viec_con FROM tb_qc_dau_vao ORDER BY ngay_kiem DESC",
        conn,
    )
    conn.close()

    if not df_qc_logs.empty:
      df_qc_logs["Ngay_Format"] = pd.to_datetime(
          df_qc_logs["ngay_kiem"], errors="coerce"
      ).dt.strftime("%d/%m/%Y")

      col_flt_person, col_flt_type = st.columns([1.5, 1])
      list_inspectors = ["Tất cả nhân sự"] + sorted([
          str(x).strip()
          for x in df_qc_logs["nguoi_kiem"].dropna().unique()
          if str(x).strip()
      ])

      with col_flt_person:
        selected_inspector = st.selectbox(
            "👤 Chọn Nhân sự QC:", list_inspectors, key="leaf_ns_inspector"
        )
      with col_flt_type:
        selected_loai_qc = st.selectbox(
            "🎯 Phân hệ kiểm:",
            ["Tất cả", "QC Đầu Vào (DAU_VAO)", "QC Sản Xuất (SAN_XUAT)"],
            key="leaf_ns_loai_qc",
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
      ty_le_loi = (tot_sl_loi / tot_sl_kiem * 100) if tot_sl_kiem > 0 else 0.0

      render_kpi_cards([
          {
              "label": "TỔNG LƯỢT KIỂM",
              "value": f"{tot_luot:,} lượt",
              "icon": "📝",
              "color": COLOR_PRIMARY,
          },
          {
              "label": "TỔNG SL ĐÃ KIỂM",
              "value": f"{int(tot_sl_kiem):,}",
              "icon": "🔍",
              "color": COLOR_SUCCESS,
          },
          {
              "label": "TỔNG SL LỖI",
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

      st.dataframe(df_filtered, use_container_width=True, hide_index=True)
    else:
      st.info("💡 Chưa có nhật ký báo cáo QC nào trong Cơ sở dữ liệu.")
  except Exception as e:
    st.error(f"⚠️ Lỗi khi tải nhật ký QC: {e}")

elif selected_sub_leaf == "⚙️ Quản Lý Công Việc Con Theo Xưởng":
  render_page_header(
      "⚙️ Thiết Lập & Quản Lý Danh Mục Công Việc Con",
      "Khai báo các công đoạn/công việc con tương ứng với từng xưởng sản xuất"
      " để nhân viên chọn trên App Mobile",
      f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
  )

  col_cv1, col_cv2, col_cv3 = st.columns([1, 1.5, 1])
  with col_cv1:
    add_ph_cv = st.selectbox(
        "Chọn Xưởng / Phân hệ:",
        ["DAU_VAO", "CO_KHI", "TU_TI", "CONG_TO"],
        key="leaf_add_ph_cv",
    )
  with col_cv2:
    add_ten_cv = st.text_input(
        "Tên công việc con mới:", placeholder="Gõ tên công đoạn..."
    )
  with col_cv3:
    st.markdown("<div style='height:25px;'></div>", unsafe_allow_html=True)
    if st.button("➕ Thêm Công Việc", type="primary", use_container_width=True):
      if add_ten_cv.strip():
        try:
          conn = get_db_connection()
          cursor = conn.cursor()
          cursor.execute(
              "INSERT INTO tb_dm_cong_viec (phan_he, ten_cong_viec) VALUES (?,"
              " ?)",
              (add_ph_cv, add_ten_cv.strip()),
          )
          conn.commit()
          conn.close()
          st.success(f"✅ Đã thêm công việc: {add_ten_cv.strip()}")
          st.rerun()
        except Exception as ex:
          st.error(f"Lỗi thêm công việc con: {ex}")
      else:
        st.warning("⚠️ Vui lòng nhập tên công việc con!")

  try:
    conn = get_db_connection()
    df_dm_cv = pd.read_sql_query(
        "SELECT id, phan_he, ten_cong_viec FROM tb_dm_cong_viec ORDER BY phan_he"
        " ASC, ten_cong_viec ASC",
        conn,
    )
    conn.close()

    if not df_dm_cv.empty:
      st.markdown("##### 📜 Danh Mục Các Công Việc Con Đang Được Áp Dụng")
      df_dm_cv.columns = ["ID", "Phân Hệ / Xưởng", "Tên Công Việc Con"]
      st.dataframe(df_dm_cv, use_container_width=True, hide_index=True)
  except Exception as ex:
    st.error(f"Lỗi nạp danh mục công việc: {ex}")

elif selected_sub_leaf == "🚨 Báo Cáo Sai Hỏng Chi Tiết & DM Lỗi":
  render_page_header(
      "🚨 Báo Cáo Phân Tích Sai Hỏng & Quản Lý Danh Mục Lỗi",
      "Thống kê các kiểu lỗi phát sinh nhiều nhất và quản lý thiết lập danh"
      " mục loại lỗi cho xưởng",
      f"{tu_date.strftime('%d/%m/%Y')} → {den_date.strftime('%d/%m/%Y')}",
  )

  sub_sh1, sub_sh2 = st.tabs(
      ["📊 1. Thống Kê & Phân Tích Sai Hỏng", "⚙️ 2. Quản Lý Danh Mục Loại Lỗi"]
  )

  with sub_sh1:
    try:
      conn = get_db_connection()
      df_defects = pd.read_sql_query(
          "SELECT * FROM tb_qc_dau_vao WHERE sl_khong_dat > 0 OR (kieu_loi IS"
          " NOT NULL AND kieu_loi != '') ORDER BY ngay_kiem DESC",
          conn,
      )
      conn.close()

      if not df_defects.empty:
        col_sh_f1, col_sh_f2 = st.columns(2)
        with col_sh_f1:
          sh_filter_loai = st.selectbox(
              "Lọc Phân Hệ:",
              ["Tất cả", "DAU_VAO", "CO_KHI", "TU_TI", "CONG_TO"],
              key="leaf_sh_flt_ph",
          )

        df_sh_view = df_defects.copy()
        if sh_filter_loai != "Tất cả":
          df_sh_view = df_sh_view[
              df_sh_view["loai_qc"].str.contains(sh_filter_loai, na=False)
              | df_sh_view["ncc"].str.contains(sh_filter_loai, na=False)
          ]

        if not df_sh_view.empty and "kieu_loi" in df_sh_view.columns:
          defect_counts = (
              df_sh_view.groupby("kieu_loi")["sl_khong_dat"].sum().reset_index()
          )
          defect_counts = defect_counts[defect_counts["kieu_loi"] != ""]
          defect_counts = defect_counts.sort_values(
              by="sl_khong_dat", ascending=False
          )

          if not defect_counts.empty:
            fig_err = go.Figure(
                go.Bar(
                    x=defect_counts["sl_khong_dat"],
                    y=defect_counts["kieu_loi"],
                    orientation="h",
                    marker=dict(color=COLOR_DANGER),
                    text=[f"{v:,.0f}" for v in defect_counts["sl_khong_dat"]],
                    textposition="outside",
                )
            )
            fig_err.update_layout(
                title="<b>TOP CÁC KIỂU SAI HỎNG PHÁT HIỆN NHIỀU NHẤT</b>",
                margin=dict(l=10, r=40, t=40, b=10),
                height=320,
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
            )
            st.plotly_chart(
                fig_err,
                use_container_width=True,
                key="leaf_sh_chart_fig_err",
            )

        st.markdown("##### 📋 Danh Sách Ca Báo Lỗi Chi Tiết")
        st.dataframe(df_sh_view, use_container_width=True, hide_index=True)
      else:
        st.success("🎉 Chưa ghi nhận ca phát sinh sai hỏng nào!")
    except Exception as e:
      st.error(f"Lỗi tải báo cáo sai hỏng: {e}")

  with sub_sh2:
    st.markdown("##### ⚙️ THÊM MỚI KIỂU SAI HỎNG CHO CÁC XƯỞNG")
    col_add1, col_add2, col_add3 = st.columns([1, 1.5, 1])

    with col_add1:
      add_phan_he = st.selectbox(
          "Chọn Xưởng / Phân hệ:",
          ["DAU_VAO", "CO_KHI", "TU_TI", "CONG_TO"],
          key="leaf_add_ph_loi",
      )
    with col_add2:
      add_ten_loi = st.text_input(
          "Tên kiểu sai hỏng mới:", placeholder="Gõ tên loại lỗi..."
      )
    with col_add3:
      st.markdown("<div style='height:25px;'></div>", unsafe_allow_html=True)
      if st.button("➕ Thêm Loại Lỗi", type="primary", use_container_width=True):
        if add_ten_loi.strip():
          try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO tb_dm_loai_loi (phan_he, ten_loi) VALUES (?, ?)",
                (add_phan_he, add_ten_loi.strip()),
            )
            conn.commit()
            conn.close()
            st.success(f"✅ Đã thêm loại lỗi: {add_ten_loi.strip()}")
            st.rerun()
          except Exception as ex:
            st.error(f"Lỗi thêm loại lỗi: {ex}")
        else:
          st.warning("⚠️ Vui lòng nhập tên loại lỗi!")

    try:
      conn = get_db_connection()
      df_dm_loi = pd.read_sql_query(
          "SELECT id, phan_he, ten_loi FROM tb_dm_loai_loi ORDER BY phan_he"
          " ASC, ten_loi ASC",
          conn,
      )
      conn.close()

      if not df_dm_loi.empty:
        st.markdown("##### 📜 Danh Mục Các Loại Lỗi Đang Được Áp Dụng")
        df_dm_loi.columns = ["ID", "Phân Hệ / Xưởng", "Tên Kiểu Sai Hỏng"]
        st.dataframe(df_dm_loi, use_container_width=True, hide_index=True)
    except Exception as ex:
      st.error(f"Lỗi nạp danh mục loại lỗi: {ex}")