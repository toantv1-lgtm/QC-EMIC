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


# ================= 2. CẤU HÌNH GIAO DIỆN STREAMLIT MOBILE NỔI BẬT & SẶC SỠ =================
st.set_page_config(
    page_title="EMIC QC Mobile Pro",
    page_icon="📱",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Style CSS tùy biến giao diện sặc sỡ, chuyên nghiệp
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
    
    /* Vibrant Header Gradient */
    .app-header {
        background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 50%, #EC4899 100%);
        color: white; 
        padding: 20px 16px; 
        border-radius: 20px; 
        text-align: center;
        box-shadow: 0 10px 25px -5px rgba(124, 58, 237, 0.4); 
        margin-bottom: 16px;
        border: 1px solid rgba(255, 255, 255, 0.3);
    }
    .emic-logo { 
        font-size: 11px; 
        font-weight: 900; 
        letter-spacing: 2px; 
        color: #FDE047; 
        text-transform: uppercase; 
        margin-bottom: 4px;
    }
    .app-title { 
        font-size: 20px; 
        font-weight: 900; 
        color: #FFFFFF; 
        text-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }

    /* Input Controls Custom */
    .stTextInput input, .stSelectbox select, .stNumberInput input {
        font-size: 13.5px !important; 
        height: 46px !important; 
        border-radius: 12px !important;
        border: 2px solid #CBD5E1 !important; 
        font-weight: 700 !important;
        background-color: #FFFFFF !important;
        color: #0F172A !important;
    }
    .stTextInput input:focus, .stSelectbox select:focus {
        border-color: #7C3AED !important;
        box-shadow: 0 0 0 3px rgba(124, 58, 237, 0.2) !important;
    }
    label { 
        font-size: 11.5px !important; 
        color: #475569 !important; 
        font-weight: 800 !important; 
        text-transform: uppercase; 
        letter-spacing: 0.5px;
    }
    
    /* Vibrant Action Button */
    div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #FF416C 0%, #FF4B2B 100%) !important;
        color: white !important; 
        font-size: 16px !important; 
        min-height: 54px !important;
        border-radius: 14px !important; 
        border: none !important; 
        font-weight: 900 !important;
        letter-spacing: 0.5px;
        box-shadow: 0 8px 20px rgba(255, 75, 43, 0.4) !important;
        transition: all 0.2s ease !important;
    }
    div.stButton > button[kind="primary"]:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 12px 25px rgba(255, 75, 43, 0.5) !important;
    }

    /* Cards */
    .mobile-card {
        background-color: #FFFFFF; 
        padding: 16px; 
        border-radius: 18px;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.05); 
        border: 1px solid #E2E8F0; 
        margin-bottom: 14px;
    }
    
    /* Custom Badges */
    .badge-label {
        background: linear-gradient(135deg, #8B5CF6 0%, #6D28D9 100%);
        color: white; padding: 3px 8px; border-radius: 6px; font-weight: 800; font-size: 11px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class='app-header'>
        <div class='emic-logo'>⚡ EMIC - QUALITY CONTROL SYSTEM</div>
        <div class='app-title'>📱 BÁO CÁO QC ĐẦU VÀO MOBILE PRO</div>
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


def get_last_inspector():
  """Tự động truy vấn lấy tên người kiểm tra từ lượt nhập gần đây nhất"""
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
  return "Nguyễn Thị Lý"


init_db()

# Lấy tên người kiểm trước đó
last_inspector = get_last_inspector()

# ================= 3. BỘ LỌC NGƯỜI KIỂM VÀ TRẠNG THÁI =================
st.markdown("<div class='mobile-card'>", unsafe_allow_html=True)
col_top1, col_top2 = st.columns([1.3, 1])

inspector_list = [
    "Nguyễn Thị Lý",
    "Nguyễn Văn A",
    "Trần Thị B",
    "Nguyễn Mạnh Long",
    "Khác",
]
default_idx = (
    inspector_list.index(last_inspector)
    if last_inspector in inspector_list
    else (len(inspector_list) - 1)
)

with col_top1:
  nguoi_kiem_selected = st.selectbox(
      "👤 NGƯỜI KIỂM (TỰ NHỚ LẦN TRƯỚC):",
      inspector_list,
      index=default_idx,
  )
  if nguoi_kiem_selected == "Khác":
    custom_default = (
        last_inspector if last_inspector not in inspector_list else ""
    )
    nguoi_kiem_final = st.text_input(
        "Họ tên người kiểm:",
        value=custom_default,
        placeholder="Nhập tên người kiểm...",
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
    <div style="background-color:#FFFFFF; padding:16px; border-radius:18px; border:2px solid #C7D2FE; box-shadow:0 8px 20px rgba(99, 102, 241, 0.08); margin-bottom:14px;">
        <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:2px dashed #E2E8F0; padding-bottom:10px; margin-bottom:12px;">
            <span style="font-size:14px; font-weight:900; color:#4F46E5;">📦 LÔ HÀNG: {selected_lot['so_lot']}</span>
        </div>
        <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; font-size:12px;">
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">MÃ VT</span><b style="color:#0F172A; font-size:13px;">{selected_lot['ma_vt']}</b></div>
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">BẢN VẼ</span><b style="color:#0F172A; font-size:13px;">{selected_lot['ky_hieu_bv']}</b></div>
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">TÊN VT</span><b style="color:#334155;">{selected_lot['ten_vt'][:25]}</b></div>
            <div><span style="color:#64748B; font-size:10px; font-weight:800; display:block;">NHÀ CUNG CẤP</span><b style="color:#334155;">{selected_lot['ncc'][:20]}</b></div>
            <div style="background:linear-gradient(135deg, #EFF6FF 0%, #DBEAFE 100%); padding:6px 10px; border-radius:8px; border:1px solid #BFDBFE;"><span style="color:#1E40AF; font-size:10px; font-weight:800; display:block;">TỔNG VỀ</span><b style="color:#1E3A8A; font-size:14px;">{selected_lot['tong_sl']:,.0f}</b></div>
            <div style="background:linear-gradient(135deg, #FEF2F2 0%, #FEE2E2 100%); padding:6px 10px; border-radius:8px; border:1px solid #FECACA;"><span style="color:#991B1B; font-size:10px; font-weight:800; display:block;">MẪU KIỂM (BY)</span><b style="color:#7F1D1D; font-size:14px;">{selected_lot['co_mau']:,.0f}</b></div>
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

  st.markdown("<hr style='border-color:#CBD5E1;'>", unsafe_allow_html=True)
  img_col1, img_col2 = st.columns(2)
  with img_col1:
    img1_file = st.file_uploader("📷 Chụp/Up Ảnh 1", type=["png", "jpg", "jpeg"])
  with img_col2:
    img2_file = st.file_uploader("📷 Chụp/Up Ảnh 2", type=["png", "jpg", "jpeg"])

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