from datetime import date, datetime
import os
import re
import sqlite3
import unicodedata
import pandas as pd
import streamlit as st

# ================= 1. KẾT NỐI CSDL LOCAL SQLITE =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Report_Database.db")
IMG_DIR = os.path.join(BASE_DIR, "Anh_kiem_tra_dau_vao")

os.makedirs(IMG_DIR, exist_ok=True)


def get_db_connection():
  conn = sqlite3.connect(DB_PATH, timeout=30.0)
  conn.execute("PRAGMA journal_mode=WAL;")
  conn.row_factory = sqlite3.Row
  return conn


def to_ddmmyyyy(iso_date_str):
  """Chuyển 'YYYY-MM-DD' (lấy từ CSDL) sang 'DD/MM/YYYY' để hiển thị."""
  try:
    return datetime.strptime(str(iso_date_str).strip(), "%Y-%m-%d").strftime(
        "%d/%m/%Y"
    )
  except Exception:
    return str(iso_date_str)


# ---------- Lưu / đọc cấu hình (tên người kiểm gần nhất, bộ lọc gần nhất) ----------
def get_setting(key, default=""):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT val FROM tb_settings WHERE key = ?", (key,))
    row = cursor.fetchone()
    conn.close()
    if row and row["val"] is not None:
      return row["val"]
  except Exception:
    pass
  return default


def set_setting(key, val):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO tb_settings (key, val) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET val = excluded.val",
        (key, str(val)),
    )
    conn.commit()
    conn.close()
  except Exception:
    pass


def option_index(options, saved_value, default_idx=0):
  return options.index(saved_value) if saved_value in options else default_idx


# ---------- Thư mục lưu ảnh theo từng xưởng / phân hệ ----------
IMG_SUBDIRS = {
    "DAU_VAO": "Dau_Vao",
    "CO_KHI": "Co_Khi",
    "TU_TI": "TUTI",
    "CONG_TO": "Cong_To",
}


def get_img_subdir(phan_he_code):
  sub_name = IMG_SUBDIRS.get(phan_he_code, "Khac")
  path = os.path.join(IMG_DIR, sub_name)
  os.makedirs(path, exist_ok=True)
  return path


def remove_accents(text):
  text = unicodedata.normalize("NFD", str(text))
  text = "".join(c for c in text if unicodedata.category(c) != "Mn")
  return text.replace("đ", "d").replace("Đ", "D")


def safe_filename_part(text, max_len=30):
  text = remove_accents(text)
  text = re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")
  return text[:max_len] if text else "NA"


# ================= 2. CẤU HÌNH GIAO DIỆN MOBILE =================
st.set_page_config(
    page_title="EMIC QC Mobile Pro",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');
    * { font-family: 'Plus Jakarta Sans', sans-serif !important; }
    
    .stApp { 
        background: linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 50%, #F3E8FF 100%) !important; 
    }
    header, #MainMenu, footer { display: none !important; }
    .block-container { padding: 0.8rem 0.6rem 3rem 0.6rem !important; max-width: 100% !important; }
    
    .app-header {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 50%, #1E3A8A 100%);
        color: white; 
        padding: 18px 15px; 
        border-radius: 18px; 
        text-align: center;
        box-shadow: 0 10px 25px -5px rgba(37, 99, 235, 0.45); 
        margin-bottom: 14px;
        border: 1px solid rgba(255, 255, 255, 0.3);
    }
    .emic-logo { font-size: 11px; font-weight: 900; letter-spacing: 2px; color: #FDE047; text-transform: uppercase; margin-bottom: 3px; }
    .app-title { font-size: 19px; font-weight: 900; color: #FFFFFF; text-shadow: 0 2px 4px rgba(0,0,0,0.2); }

    .stTextInput input, .stSelectbox select, .stNumberInput input {
        font-size: 13.5px !important; height: 44px !important; border-radius: 12px !important;
        border: 2px solid #CBD5E1 !important; font-weight: 700 !important; background-color: #FFFFFF !important; color: #0F172A !important;
    }
    label { font-size: 11.5px !important; color: #475569 !important; font-weight: 800 !important; text-transform: uppercase; }

    /* Sửa chữ trong ô tải ảnh bị bé/che do CSS label chung ở trên */
    .stFileUploader label, [data-testid="stFileUploader"] label {
        font-size: 12.5px !important; text-transform: none !important; white-space: normal !important;
    }
    [data-testid="stFileUploaderDropzone"] {
        min-height: 100px !important; padding: 14px 10px !important; overflow: visible !important; height: auto !important;
    }
    [data-testid="stFileUploaderDropzone"] div,
    [data-testid="stFileUploaderDropzone"] span,
    [data-testid="stFileUploaderDropzone"] small,
    [data-testid="stFileUploaderDropzone"] p {
        font-size: 11.5px !important; text-transform: none !important; white-space: normal !important;
        overflow: visible !important; line-height: 1.35 !important;
    }
    [data-testid="stFileUploader"] button {
        height: auto !important; min-height: 34px !important; white-space: nowrap !important;
        padding: 6px 14px !important; font-size: 12px !important; text-transform: none !important;
    }

    /* Sửa nút "BỘ LỌC TÌM KIẾM" bị che chữ */
    [data-testid="stExpander"] {
        border-radius: 14px !important; border: 2px solid #CBD5E1 !important;
        background-color: #FFFFFF !important; overflow: hidden !important; margin-bottom: 10px !important;
    }
    [data-testid="stExpander"] summary {
        min-height: 46px !important; padding: 10px 14px !important; display: flex !important;
        align-items: center !important; gap: 8px !important;
    }
    [data-testid="stExpander"] summary p, .streamlit-expanderHeader p {
        font-size: 13px !important; font-weight: 800 !important; color: #1E293B !important;
        margin: 0 !important; white-space: normal !important; overflow: visible !important; text-transform: none !important;
    }

    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #FF416C 0%, #FF4B2B 100%) !important;
        color: white !important; font-size: 16px !important; min-height: 52px !important;
        border-radius: 14px !important; border: none !important; font-weight: 900 !important;
        box-shadow: 0 8px 20px rgba(255, 75, 43, 0.4) !important;
    }

    .mobile-card {
        background-color: #FFFFFF; padding: 14px; border-radius: 16px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.05); border: 1px solid #E2E8F0; margin-bottom: 12px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class='app-header'>
        <div class='emic-logo'>⚡ EMIC - PHÒNG QUẢN LÝ CHẤT LƯỢNG</div>
        <div class='app-title'>📱 PHẦN MỀM BÁO CÁO KIỂM TRA</div>
    </div>
""",
    unsafe_allow_html=True,
)


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

    cursor.execute("""
            CREATE TABLE IF NOT EXISTS tb_settings (
                key TEXT PRIMARY KEY,
                val TEXT
            )
        """)

    # Khởi tạo công việc con mặc định nếu chưa có
    cursor.execute("SELECT COUNT(*) FROM tb_dm_cong_viec")
    if cursor.fetchone()[0] == 0:
      default_tasks = [
          ("DAU_VAO", "Kiểm tra ngoại quan & kích thước"),
          ("DAU_VAO", "Kiểm tra thông số kỹ thuật"),
          ("DAU_VAO", "Kiểm tra đóng gói & nhãn mác"),
          ("CO_KHI", "Gia công đột dập cơ khí"),
          ("CO_KHI", "Gia công phay / tiện / hàn"),
          ("CO_KHI", "Sơn / Mạ bề mặt"),
          ("CO_KHI", "Lắp ráp cụm vỏ cơ khí"),
          ("TU_TI", "Quấn dây sơ cấp / thứ cấp"),
          ("TU_TI", "Đổ keo đúc Epoxy"),
          ("TU_TI", "Thử nghiệm cao áp & cách điện"),
          ("TU_TI", "Kiểm tra tỷ số biến & sai số"),
          ("CONG_TO", "Lắp ráp bo mạch điện tử"),
          ("CONG_TO", "Hiệu chuẩn & Bật điểm số"),
          ("CONG_TO", "Thử nghiệm gá đặt nhiệt độ"),
          ("CONG_TO", "Kiểm tra dán tem & Đóng gói"),
      ]
      cursor.executemany(
          "INSERT INTO tb_dm_cong_viec (phan_he, ten_cong_viec) VALUES (?, ?)",
          default_tasks,
      )

    conn.commit()
    conn.close()
  except Exception:
    pass


def get_last_inspector_name():
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT nguoi_kiem FROM tb_qc_dau_vao WHERE nguoi_kiem IS NOT NULL AND"
        " nguoi_kiem != '' ORDER BY id DESC LIMIT 1"
    )
    row = cursor.fetchone()
    conn.close()
    if row and row["nguoi_kiem"]:
      return row["nguoi_kiem"]
  except Exception:
    pass
  return ""


def get_defect_types(phan_he_code):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT ten_loi FROM tb_dm_loai_loi WHERE phan_he = ? ORDER BY"
        " ten_loi ASC",
        (phan_he_code,),
    )
    rows = cursor.fetchall()
    conn.close()
    return [r["ten_loi"] for r in rows] if rows else ["Chưa xác định", "Khác"]
  except Exception:
    return ["Chưa xác định", "Khác"]


def get_sub_tasks(phan_he_code):
  try:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "SELECT ten_cong_viec FROM tb_dm_cong_viec WHERE phan_he = ? ORDER BY"
        " ten_cong_viec ASC",
        (phan_he_code,),
    )
    rows = cursor.fetchall()
    conn.close()
    return (
        [r["ten_cong_viec"] for r in rows]
        if rows
        else ["Kiểm tra chung", "Khác"]
    )
  except Exception:
    return ["Kiểm tra chung", "Khác"]


init_db()

if "saved_inspector_name" not in st.session_state:
  st.session_state["saved_inspector_name"] = get_last_inspector_name()

# ================= 3. BỘ LỌC VÀ THÔNG TIN BÁO CÁO =================
st.markdown("<div class='mobile-card'>", unsafe_allow_html=True)

nguoi_kiem_input = st.text_input(
    "👤 HỌ VÀ TÊN NGƯỜI KIỂM TRA:",
    value=st.session_state["saved_inspector_name"],
    placeholder="Gõ họ tên người kiểm...",
)
if nguoi_kiem_input != st.session_state["saved_inspector_name"]:
  st.session_state["saved_inspector_name"] = nguoi_kiem_input

loai_qc_options_list = ["📦 QC Đầu Vào (Vật tư)", "⚙️ QC Sản Xuất (Lệnh SX)"]
loai_qc_option = st.radio(
    "🎯 CHỌN PHÂN HỆ PHẠM VI KIỂM TRA:",
    loai_qc_options_list,
    index=option_index(loai_qc_options_list, get_setting("last_phan_he", loai_qc_options_list[0])),
    horizontal=True,
)
set_setting("last_phan_he", loai_qc_option)
is_qc_dau_vao = "Vật tư" in loai_qc_option

today = date.today()
first_day_of_month = date(today.year, today.month, 1)

with st.expander("🔍 BỘ LỌC TÌM KIẾM", expanded=False):
  col_d1, col_d2 = st.columns(2)
  with col_d1:
    tu_date = st.date_input(
        "Từ ngày:", first_day_of_month, format="DD/MM/YYYY"
    )
  with col_d2:
    den_date = st.date_input("Đến ngày:", today, format="DD/MM/YYYY")

  if is_qc_dau_vao:
    status_opts = ["Chưa kiểm", "Đã kiểm", "Tất cả"]
    status_filter = st.selectbox(
        "Trạng thái kiểm:",
        status_opts,
        index=option_index(status_opts, get_setting("status_filter_dv", "Chưa kiểm")),
    )
    set_setting("status_filter_dv", status_filter)
  else:
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
      xuong_opts = ["Tất cả", "Cơ khí (3012)", "TU/TI (3011)", "Công tơ (3013, 3016)"]
      filter_xuong = st.selectbox(
          "Xưởng / Phân hệ:",
          xuong_opts,
          index=option_index(xuong_opts, get_setting("filter_xuong", "Tất cả")),
      )
      set_setting("filter_xuong", filter_xuong)
    with col_f2:
      loaisp_opts = ["Tất cả", "Bán thành phẩm (Đầu 4)", "Sản phẩm (Đầu 5)"]
      filter_loai_sp = st.selectbox(
          "Loại sản phẩm:",
          loaisp_opts,
          index=option_index(loaisp_opts, get_setting("filter_loai_sp", "Tất cả")),
      )
      set_setting("filter_loai_sp", filter_loai_sp)
    with col_f3:
      trangthai_opts = ["Tất cả", "Hoàn thành", "Đang SX", "Chưa SX"]
      filter_trang_thai_sx = st.selectbox(
          "Trạng thái SX:",
          trangthai_opts,
          index=option_index(trangthai_opts, get_setting("filter_trang_thai_sx", "Tất cả")),
      )
      set_setting("filter_trang_thai_sx", filter_trang_thai_sx)

search_kw = st.text_input(
    "🔎 Tìm kiếm nhanh:",
    "",
    placeholder="Mã/Tên/Lot/NCC..." if is_qc_dau_vao else "Số lệnh/Mã/Tên SP...",
)

st.markdown("</div>", unsafe_allow_html=True)


# ================= 4. NẠP DỮ LIỆU TỪ CSDL =================
@st.cache_data(ttl=5)
def load_qc_data(is_dau_vao, tu_d, den_d):
  tu_iso = tu_d.strftime("%Y-%m-%d 00:00:00")
  den_iso = den_d.strftime("%Y-%m-%d 23:59:59")
  try:
    conn = get_db_connection()
    if is_dau_vao:
      df_raw = pd.read_sql_query(
          "SELECT * FROM tb_sap_qa32 WHERE ngay_ve_dt >= ? AND ngay_ve_dt <="
          " ?",
          conn,
          params=(tu_iso, den_iso),
      )
    else:
      df_raw = pd.read_sql_query(
          "SELECT * FROM tb_sap_coois WHERE ngay_lenh_dt >= ? AND ngay_lenh_dt"
          " <= ?",
          conn,
          params=(tu_iso, den_iso),
      )

    try:
      target_type = "DAU_VAO" if is_dau_vao else "SAN_XUAT"
      df_rep = pd.read_sql_query(
          "SELECT so_lot FROM tb_qc_dau_vao WHERE loai_qc = ?",
          conn,
          params=(target_type,),
      )
      checked_set = set(df_rep["so_lot"].dropna().unique())
    except Exception:
      checked_set = set()

    conn.close()
    return df_raw, checked_set
  except Exception:
    return pd.DataFrame(), set()


df_raw, checked_set = load_qc_data(is_qc_dau_vao, tu_date, den_date)

filtered_items = []

if not df_raw.empty:
  if is_qc_dau_vao:
    for _, r in df_raw.iterrows():
      so_lot = str(
          r.get(
              "lot",
              r.get(
                  "so_lot", f"{r.get('ma_vt', '')}_{r.get('ngay_ve_dt', '')}"
              ),
          )
      ).strip()
      ma_vt = str(r.get("ma_vt", "")).strip()
      ten_vt = str(r.get("ten_vt", "")).strip()
      ncc = str(r.get("ncc", "")).strip()
      is_checked = so_lot in checked_set

      if status_filter == "Chưa kiểm" and is_checked:
        continue
      if status_filter == "Đã kiểm" and not is_checked:
        continue

      if search_kw.strip():
        kw = search_kw.strip().lower()
        if not (
            kw in ma_vt.lower()
            or kw in ten_vt.lower()
            or kw in so_lot.lower()
            or kw in ncc.lower()
        ):
          continue

      filtered_items.append({
          "so_lot": so_lot,
          "ma_vt": ma_vt,
          "ten_vt": ten_vt,
          "ncc": ncc,
          "ngay_ve": str(r.get("ngay_ve_dt", "")).split()[0],
          "ngay_ve_display": to_ddmmyyyy(str(r.get("ngay_ve_dt", "")).split()[0]),
          "tong_sl": float(r.get("ft_qty", 0.0) or 0.0),
          "co_mau": float(r.get("by_sample", 5.0) or 5.0),
          "phan_he": "DAU_VAO",
          "is_checked": is_checked,
      })
  else:
    for _, r in df_raw.iterrows():
      so_lenh = str(r.get("lenh_sx", r.get("so_lenh", ""))).strip()
      ma_tp = str(r.get("ma_tp", "")).strip()
      ten_tp = str(r.get("ten_tp", "")).strip()
      phan_he = str(r.get("phan_he", "")).strip()
      trang_thai_sx = str(r.get("trang_thai", "")).strip()
      is_checked = so_lenh in checked_set

      if filter_xuong != "Tất cả":
        if "Cơ khí" in filter_xuong and phan_he != "CO_KHI":
          continue
        if "TU/TI" in filter_xuong and phan_he != "TU_TI":
          continue
        if "Công tơ" in filter_xuong and phan_he != "CONG_TO":
          continue

      if (
          filter_trang_thai_sx != "Tất cả"
          and trang_thai_sx != filter_trang_thai_sx
      ):
        continue

      clean_code = ma_tp.lstrip("0")
      if (
          filter_loai_sp == "Bán thành phẩm (Đầu 4)"
          and not clean_code.startswith("4")
      ):
        continue
      if filter_loai_sp == "Sản phẩm (Đầu 5)" and not clean_code.startswith(
          "5"
      ):
        continue

      if search_kw.strip():
        kw = search_kw.strip().lower()
        if not (
            kw in so_lenh.lower() or kw in ma_tp.lower() or kw in ten_tp.lower()
        ):
          continue

      filtered_items.append({
          "so_lot": so_lenh,
          "ma_vt": ma_tp,
          "ten_vt": ten_tp,
          "ncc": f"Xưởng {phan_he}" if phan_he else "Xưởng Sản Xuất",
          "ngay_ve": str(r.get("ngay_lenh_dt", r.get("ngay_lenh", ""))).split()[
              0
          ],
          "ngay_ve_display": to_ddmmyyyy(
              str(r.get("ngay_lenh_dt", r.get("ngay_lenh", ""))).split()[0]
          ),
          "tong_sl": float(r.get("sl_tong", 0.0) or 0.0),
          "co_mau": float(r.get("sl_tong", 10.0) or 10.0),
          "phan_he": phan_he if phan_he else "CO_KHI",
          "is_checked": is_checked,
          "trang_thai_sx": trang_thai_sx,
      })

options_list = ["-- Chạm để chọn đối tượng nhập báo cáo --"] + [
    f"{'✅' if item['is_checked'] else '⏳'} {item['so_lot']} | {item['ma_vt']} -"
    f" {item['ten_vt']}"
    for item in filtered_items
]

selected_opt = st.selectbox(
    "📋 DANH SÁCH ĐỐI TƯỢNG CẦN KIỂM TRẢ:", options_list, index=0
)

# ================= 5. FORM NHẬP BÁO CÁO KẾT QUẢ =================
if selected_opt != options_list[0]:
  idx = options_list.index(selected_opt) - 1
  selected_item = filtered_items[idx]

  st.markdown(
      f"""
    <div style="background-color:#FFFFFF; padding:16px; border-radius:18px; border:2px solid #C7D2FE; box-shadow:0 8px 20px rgba(99, 102, 241, 0.08); margin-bottom:14px;">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:2px dashed #E2E8F0; padding-bottom:10px; margin-bottom:12px;">
            <span style="font-size:15px; font-weight:900; color:#4F46E5;">📌 {'LÔ VẬT TƯ' if is_qc_dau_vao else 'LỆNH SẢN XUẤT'}: {selected_item['so_lot']}</span>
        </div>
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; font-size:13.5px;">
            <div><span style="color:#64748B; font-size:11px; font-weight:800; display:block;">MÃ HÀNG</span><b style="color:#0F172A; font-size:14.5px;">{selected_item['ma_vt']}</b></div>
            <div><span style="color:#64748B; font-size:11px; font-weight:800; display:block;">NGÀY VỀ / LỆNH</span><b style="color:#0F172A; font-size:14.5px;">{selected_item['ngay_ve_display']}</b></div>
            <div style="grid-column: span 2;"><span style="color:#64748B; font-size:11px; font-weight:800; display:block;">TÊN MẶT HÀNG</span><b style="color:#334155;">{selected_item['ten_vt']}</b></div>
            <div style="grid-column: span 2;"><span style="color:#64748B; font-size:11px; font-weight:800; display:block;">ĐƠN VỊ / NCC / XƯỞNG</span><b style="color:#334155;">{selected_item['ncc']}</b></div>
            <div style="background:linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%); padding:6px 10px; border-radius:8px; border:1px solid #BFDBFE;"><span style="color:#1E40AF; font-size:11px; font-weight:800; display:block;">TỔNG SỐ LƯỢNG</span><b style="color:#1E3A8A; font-size:15.5px;">{selected_item['tong_sl']:,.0f}</b></div>
            <div style="background:linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%); padding:6px 10px; border-radius:8px; border:1px solid #FECACA;"><span style="color:#991B1B; font-size:11px; font-weight:800; display:block;">SỐ MẪU KIỂM YÊU CẦU</span><b style="color:#7F1D1D; font-size:15.5px;">{selected_item['co_mau']:,.0f}</b></div>
        </div>
    </div>
    """,
      unsafe_allow_html=True,
  )

  # BỔ SUNG Ô CHỌN CÔNG VIỆC CON / CÔNG ĐOẠN
  sub_task_list = get_sub_tasks(selected_item["phan_he"]) + [
      "Công việc khác / Nhập tay"
  ]
  col_st1, col_st2 = st.columns([1.5, 1])
  with col_st1:
    sub_task_opt = st.selectbox(
        "⚙️ CÔNG VIỆC CON / CÔNG ĐOẠN KIỂM:", sub_task_list
    )
    if sub_task_opt == "Công việc khác / Nhập tay":
      cong_viec_con_selected = st.text_input(
          "Nhập tên công việc cụ thể:", placeholder="Gõ tên công việc..."
      )
    else:
      cong_viec_con_selected = sub_task_opt

  col_q1, col_q2, col_q3 = st.columns(3)
  with col_q1:
    sl_kiem = st.number_input(
        "SL KIỂM:", min_value=1, value=int(selected_item["co_mau"]), step=1
    )
  with col_q2:
    sl_khong_dat = st.number_input("SL LỖI:", min_value=0, value=0, step=1)
  with col_q3:
    sl_dat = max(0, sl_kiem - sl_khong_dat)
    st.number_input("SL ĐẠT:", value=int(sl_dat), disabled=True)

  ket_luan = (
      "Đạt tiêu chuẩn (UD 01)"
      if sl_khong_dat == 0
      else "Không đạt - Trả lại (UD 03)"
  )

  defect_list = get_defect_types(selected_item["phan_he"]) + [
      "Lỗi khác / Nhập tay"
  ]
  kieu_loi_selected = ""

  if sl_khong_dat > 0:
    col_err1, col_err2 = st.columns(2)
    with col_err1:
      loai_loi_opt = st.selectbox(
          "🚨 KIỂU SAI HỎNG / PHÂN LOẠI LỖI:", defect_list
      )
      if loai_loi_opt == "Lỗi khác / Nhập tay":
        kieu_loi_selected = st.text_input(
            "Tên lỗi cụ thể:", placeholder="Gõ mô tả lỗi ngắn..."
        )
      else:
        kieu_loi_selected = loai_loi_opt
    with col_err2:
      ghi_chu = st.text_input(
          "📝 GHI CHÚ BỔ SUNG:", placeholder="Mô tả chi tiết lỗi..."
      )
  else:
    ghi_chu = st.text_input(
        "📝 GHI CHÚ BỔ SUNG:", placeholder="Mô tả chi tiết lỗi..."
    )

  st.markdown("<hr style='border-color:#CBD5E1;'>", unsafe_allow_html=True)
  img_col1, img_col2 = st.columns(2)
  with img_col1:
    img1_file = st.file_uploader("📷 Chụp / Tải Ảnh 1", type=["png", "jpg", "jpeg"])
  with img_col2:
    img2_file = st.file_uploader("📷 Chụp / Tải Ảnh 2", type=["png", "jpg", "jpeg"])

  if st.button("🚀 GỬI BÁO CÁO VỀ HỆ THỐNG", type="primary", use_container_width=True):
    if not nguoi_kiem_input.strip():
      st.error("❌ Vui lòng nhập Họ và tên Người kiểm tra!")
    else:
      img1_path, img2_path = "", ""
      time_str = datetime.now().strftime("%d%m%Y_%H%M%S")
      img_folder = get_img_subdir(selected_item["phan_he"])
      name_lenh = safe_filename_part(selected_item["so_lot"])
      name_sp = safe_filename_part(selected_item["ten_vt"])
      name_nguoi = safe_filename_part(nguoi_kiem_input)
      base_name = f"{time_str}_{name_lenh}_{name_sp}_{name_nguoi}"

      if img1_file:
        img1_path = os.path.join(img_folder, f"{base_name}_anh1.png")
        with open(img1_path, "wb") as f:
          f.write(img1_file.getbuffer())
      if img2_file:
        img2_path = os.path.join(img_folder, f"{base_name}_anh2.png")
        with open(img2_path, "wb") as f:
          f.write(img2_file.getbuffer())

      try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            """
                    INSERT INTO tb_qc_dau_vao 
                    (loai_qc, so_lot, ma_vt, ten_vt, ncc, ngay_ve, tong_sl_ve, sl_kiem, sl_khong_dat, sl_dat, ket_luan, nguoi_kiem, ngay_kiem, ghi_chu, kieu_loi, cong_viec_con, img1, img2)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
            (
                "DAU_VAO" if is_qc_dau_vao else "SAN_XUAT",
                selected_item["so_lot"],
                selected_item["ma_vt"],
                selected_item["ten_vt"],
                selected_item["ncc"],
                selected_item["ngay_ve"],
                selected_item["tong_sl"],
                sl_kiem,
                sl_khong_dat,
                sl_dat,
                ket_luan,
                nguoi_kiem_input.strip(),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ghi_chu,
                kieu_loi_selected,
                cong_viec_con_selected,
                img1_path,
                img2_path,
            ),
        )
        conn.commit()
        conn.close()

        st.session_state["saved_inspector_name"] = nguoi_kiem_input.strip()

        st.success(
            f"🎉 Đã gửi báo cáo QC thành công cho Lô/Lệnh {selected_item['so_lot']}!"
        )
        st.cache_data.clear()
      except Exception as e:
        st.error(f"Lỗi khi lưu dữ liệu: {e}")