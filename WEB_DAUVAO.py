from datetime import date, datetime
import os
import re
import sqlite3
import pandas as pd
from PIL import Image
import streamlit as st

# ================= 1. KẾT NỐI CSDL CHUNG TRÊN THƯ MỤC GOOGLE DRIVE =================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "Report_Database.db")
IMG_DIR = os.path.join(BASE_DIR, "Anh_kiem_tra_dau_vao")

os.makedirs(IMG_DIR, exist_ok=True)


def get_db_connection():
  # Timeout 30s + WAL mode chống khóa file khi nhiều thiết bị mở chung
  conn = sqlite3.connect(DB_PATH, timeout=30.0)
  conn.execute("PRAGMA journal_mode=WAL;")
  conn.row_factory = sqlite3.Row
  return conn


# ================= 2. CẤU HÌNH GIAO DIỆN STREAMLIT MOBILE =================
st.set_page_config(
    page_title="EMIC QC Đầu Vào Mobile",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');
    * { font-family: 'Plus Jakarta Sans', sans-serif !important; }
    .stApp { background-color: #F0F4F8; }
    header, #MainMenu, footer { display: none !important; }
    .block-container { padding: 1rem 0.6rem 3rem 0.6rem !important; max-width: 100% !important; }
    
    .app-header {
        background: linear-gradient(135deg, #1E1B4B 0%, #3B82F6 100%);
        color: white; padding: 18px 15px; border-radius: 16px; text-align: center;
        box-shadow: 0 8px 20px rgba(59, 130, 246, 0.25); margin-bottom: 16px;
    }
    .emic-logo { font-size: 11px; font-weight: 700; letter-spacing: 1.5px; color: #93C5FD; text-transform: uppercase; }
    .app-title { font-size: 18px; font-weight: 900; color: #FFFFFF; }

    .stTextInput input, .stSelectbox select, .stNumberInput input {
        font-size: 13px !important; height: 44px !important; border-radius: 12px !important;
        border: 1.5px solid #E2E8F0 !important; font-weight: 600 !important;
    }
    label { font-size: 11px !important; color: #64748B !important; font-weight: 800 !important; text-transform: uppercase; }
    
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #FF512F 0%, #DD2476 100%) !important;
        color: white !important; font-size: 15px !important; min-height: 52px !important;
        border-radius: 12px !important; border: none !important; font-weight: 800 !important;
    }
    .mobile-card {
        background-color: #FFFFFF; padding: 14px; border-radius: 16px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04); border: 1px solid #E2E8F0; margin-bottom: 14px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class='app-header'>
        <div class='emic-logo'>🏢 EMIC - PHÒNG QUẢN LÝ CHẤT LƯỢNG</div>
        <div class='app-title'>BÁO CÁO QC ĐẦU VÀO MOBILE</div>
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
                so_lot TEXT, ma_vt TEXT, ten_vt TEXT, ncc TEXT, ngay_ve TEXT,
                tong_sl_ve REAL, sl_kiem REAL, sl_khong_dat REAL, sl_dat REAL,
                ket_luan TEXT, nguoi_kiem TEXT, ngay_kiem DATETIME, ghi_chu TEXT,
                img1 TEXT, img2 TEXT
            )
        """)
    conn.commit()
    conn.close()
  except Exception:
    pass


init_db()

# ================= 3. BỘ LỌC NGƯỜI KIỂM VÀ TRẠNG THÁI =================
st.markdown("<div class='mobile-card'>", unsafe_allow_html=True)
col_top1, col_top2 = st.columns([1.2, 1])

with col_top1:
  inspector_list = [
      "Nguyễn Thị Lý",
      "Nguyễn Văn A",
      "Trần Thị B",
      "Nguyễn Mạnh Long",
      "Khác",
  ]
  nguoi_kiem_selected = st.selectbox("👤 NGƯỜI KIỂM TRA:", inspector_list, index=0)
  if nguoi_kiem_selected == "Khác":
    nguoi_kiem_final = st.text_input(
        "Họ tên người kiểm:", value="", placeholder="Gõ tên..."
    )
  else:
    nguoi_kiem_final = nguoi_kiem_selected

with col_top2:
  status_filter = st.selectbox(
      "📌 TRẠNG THÁI:", ["Chưa kiểm", "Đã kiểm", "Tất cả"]
  )

col_d1, col_d2 = st.columns(2)
with col_d1:
  tu_date = st.date_input("Từ ngày:", date(datetime.now().year, 1, 1))
with col_d2:
  den_date = st.date_input("Đến ngày:", date.today())

search_kw = st.text_input("🔍 TÌM KIẾM NHANH (Mã/Tên/Lot/NCC):", "")
st.markdown("</div>", unsafe_allow_html=True)


# ================= 4. TẢI DỮ LIỆU QA32 TỪ CSDL =================
@st.cache_data(ttl=10)
def load_data_mobile(tu_d, den_d):
  tu_iso = tu_d.strftime("%Y-%m-%d 00:00:00")
  den_iso = den_d.strftime("%Y-%m-%d 23:59:59")
  try:
    conn = get_db_connection()
    df_qa32 = pd.read_sql_query(
        "SELECT * FROM tb_sap_qa32 WHERE ngay_ve_dt >= ? AND ngay_ve_dt <= ?",
        conn,
        params=(tu_iso, den_iso),
    )
    try:
      df_rep = pd.read_sql_query("SELECT so_lot FROM tb_qc_dau_vao", conn)
      checked_set = set(df_rep["so_lot"].dropna().unique())
    except Exception:
      checked_set = set()
    conn.close()
    return df_qa32, checked_set
  except Exception:
    return pd.DataFrame(), set()


df_qa32_raw, checked_lots = load_data_mobile(tu_date, den_date)

filtered_items = []
if not df_qa32_raw.empty:
  for _, r in df_qa32_raw.iterrows():
    so_lot = str(
        r.get(
            "lot",
            r.get("so_lot", f"{r.get('ma_vt', '')}_{r.get('ngay_ve_dt', '')}"),
        )
    ).strip()
    ma_vt = str(r.get("ma_vt", "")).strip()
    ten_vt = str(r.get("ten_vt", "")).strip()
    ncc = str(r.get("ncc", "")).strip()
    is_checked = so_lot in checked_lots

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
        "tong_sl": float(r.get("ft_qty", 0.0) or 0.0),
        "co_mau": float(r.get("by_sample", 5.0) or 5.0),
        "ky_hieu_bv": str(r.get("ky_hieu_bv", "-")).strip(),
        "is_checked": is_checked,
    })

options_list = ["-- Chạm để chọn lô hàng cần báo cáo --"] + [
    f"{'✅' if item['is_checked'] else '⏳'} {item['so_lot']} | {item['ma_vt']} -"
    f" {item['ten_vt']}"
    for item in filtered_items
]

selected_opt = st.selectbox("📋 DANH SÁCH LÔ HÀNG CẦN KIỂM:", options_list, index=0)

# ================= 5. FORM NHẬP BÁO CÁO =================
if selected_opt != options_list[0]:
  idx = options_list.index(selected_opt) - 1
  selected_lot = filtered_items[idx]

  st.markdown(
      f"""
    <div style="background-color:#FFFFFF; padding:14px; border-radius:16px; border:1px solid #E2E8F0; margin-bottom:14px;">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1.5px dashed #E2E8F0; padding-bottom:8px; margin-bottom:10px;">
            <span style="font-size:13px; font-weight:900; color:#1E293B;">📦 LÔ HÀNG: {selected_lot['so_lot']}</span>
        </div>
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:8px; font-size:12px;">
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">MÃ VT</span><b>{selected_lot['ma_vt']}</b></div>
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">BV</span><b>{selected_lot['ky_hieu_bv']}</b></div>
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">TÊN VT</span><b>{selected_lot['ten_vt'][:25]}</b></div>
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">NCC</span><b>{selected_lot['ncc'][:20]}</b></div>
            <div style="background:#F1F5F9; padding:4px; border-radius:6px;">TỔNG VỀ: <b>{selected_lot['tong_sl']:,.0f}</b></div>
            <div style="background:#FFF1F2; padding:4px; border-radius:6px;">MẪU BY: <b>{selected_lot['co_mau']:,.0f}</b></div>
        </div>
    </div>
    """,
      unsafe_allow_html=True,
  )

  col_q1, col_q2, col_q3 = st.columns(3)
  with col_q1:
    sl_kiem = st.number_input(
        "SL KIỂM:", min_value=1, value=int(selected_lot["co_mau"]), step=1
    )
  with col_q2:
    sl_khong_dat = st.number_input("SL LỖI:", min_value=0, value=0, step=1)
  with col_q3:
    sl_dat = max(0, sl_kiem - sl_khong_dat)
    st.number_input("SL ĐẠT:", value=int(sl_dat), disabled=True)

  col_res1, col_res2 = st.columns(2)
  with col_res1:
    ket_luan = st.selectbox(
        "📊 KẾT LUẬN:",
        ["Lô hàng đạt (UD 01)", "Lô hàng không đạt (UD 03)", "Đặc nhượng (UD 02)"],
        index=0 if sl_khong_dat == 0 else 1,
    )
  with col_res2:
    ghi_chu = st.text_input("📝 GHI CHÚ:", placeholder="Nhập mô tả lỗi...")

  st.markdown("<hr>", unsafe_allow_html=True)
  img_col1, img_col2 = st.columns(2)
  with img_col1:
    img1_file = st.file_uploader("📷 Up Ảnh 1", type=["png", "jpg", "jpeg"])
  with img_col2:
    img2_file = st.file_uploader("📷 Up Ảnh 2", type=["png", "jpg", "jpeg"])

  if st.button("🚀 GỬI BÁO CÁO VỀ HỆ THỐNG", type="primary", use_container_width=True):
    if not nguoi_kiem_final.strip():
      st.error("❌ Vui lòng điền tên Người kiểm tra!")
    else:
      img1_path, img2_path = "", ""
      time_str = datetime.now().strftime("%Y%m%d_%H%M%S")

      if img1_file:
        img1_path = os.path.join(IMG_DIR, f"{time_str}_img1.png")
        with open(img1_path, "wb") as f:
          f.write(img1_file.getbuffer())
      if img2_file:
        img2_path = os.path.join(IMG_DIR, f"{time_str}_img2.png")
        with open(img2_path, "wb") as f:
          f.write(img2_file.getbuffer())

      try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            """
                    INSERT INTO tb_qc_dau_vao 
                    (so_lot, ma_vt, ten_vt, ncc, ngay_ve, tong_sl_ve, sl_kiem, sl_khong_dat, sl_dat, ket_luan, nguoi_kiem, ngay_kiem, ghi_chu, img1, img2)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
            (
                selected_lot["so_lot"],
                selected_lot["ma_vt"],
                selected_lot["ten_vt"],
                selected_lot["ncc"],
                selected_lot["ngay_ve"],
                selected_lot["tong_sl"],
                sl_kiem,
                sl_khong_dat,
                sl_dat,
                ket_luan,
                nguoi_kiem_final,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                ghi_chu,
                img1_path,
                img2_path,
            ),
        )
        conn.commit()
        conn.close()
        st.success(f"🎉 Gửi báo cáo thành công cho Lô {selected_lot['so_lot']}!")
        st.cache_data.clear()
      except Exception as e:
        st.error(f"Lỗi khi lưu dữ liệu: {e}")