"""
╔══════════════════════════════════════════════════════════════════════════════╗
║  HỆ THỐNG AI QUẢN TRỊ & CHUẨN HÓA VẬT TƯ KỸ THUẬT — v6.0 INDUSTRIAL HUD  ║
║  PHÒNG KỸ THUẬT - CƠ ĐIỆN | CÔNG TY CP PHÂN BÓN & HÓA CHẤT CẦN THƠ           ║
║  Trợ Lý Kỹ Thuật Chuyên Sâu (Dữ Liệu Kho + Mạng) | Bảng & Thẻ Đẹp Mắt         ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import os
import io
import json
import re
import time
import datetime
import base64
import hashlib
import unicodedata
import html
import sqlite3
import shutil
import zoneinfo
import threading
from pathlib import Path

try:
    from streamlit.runtime.scriptrunner import add_script_run_ctx
except ImportError:
    try:
        from streamlit.runtime.scriptrunner_utils.script_run_context import add_script_run_ctx
    except ImportError:
        add_script_run_ctx = None

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from rapidfuzz import fuzz, process as rfprocess
from PIL import Image

# ── Google Generative AI ──
import google.generativeai as genai
import agent2_manager
from sku_matching_engine import UniversalTechnicalSKUMatcher

# ── PDF Parser ──
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError:
    HAS_FITZ = False

def render_pdf_pages_as_images(pdf_bytes: bytes, max_pages: int = 5, dpi: int = 150) -> list[bytes]:
    """Chuyển đổi các trang PDF sang danh sách ảnh PNG bytes để phân tích OCR/Vision."""
    if not HAS_FITZ or not pdf_bytes:
        return []
    images = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for p_no in range(min(len(doc), max_pages)):
            page = doc[p_no]
            pix = page.get_pixmap(dpi=dpi)
            images.append(pix.tobytes("png"))
        doc.close()
    except Exception as e:
        print(f"Lỗi render PDF sang ảnh: {e}")
    return images

def render_pdf_first_page_as_image(pdf_bytes: bytes, dpi: int = 150) -> bytes | None:
    """Chuyển đổi trang đầu tiên của file PDF sang ảnh PNG bytes."""
    imgs = render_pdf_pages_as_images(pdf_bytes, max_pages=1, dpi=dpi)
    return imgs[0] if imgs else None

# ═══════════════════════════════════════════════════════════════════════════
# CẤU HÌNH TRANG STREAMLIT (BẮT BUỘC ĐẦU TIÊN)
# ═══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="QUẢN TRỊ VẬT TƯ CƠ - ĐIỆN | AI HUD",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ═══════════════════════════════════════════════════════════════════════════
# CẤU HÌNH ĐƯỜNG DẪN HỆ THỐNG
# ═══════════════════════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).resolve().parent
MASTER_DIR = BASE_DIR / "Master_Inventory"
DAILY_DIR = BASE_DIR / "Daily_Input"
REPORT_DIR = BASE_DIR / "Processed_Reports"
MASTER_FILE = MASTER_DIR / "Master_Inventory.xlsx"
PASTE_BOX_DIR = BASE_DIR / "paste_box"
PUBLIC_URL_FILE = BASE_DIR / "public_url.txt"

for d in [MASTER_DIR, DAILY_DIR, REPORT_DIR, PASTE_BOX_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ── Đọc link Public Internet (Cloudflare Tunnel hoặc Cloud) ──
def get_public_url() -> str:
    if PUBLIC_URL_FILE.exists():
        try:
            u = PUBLIC_URL_FILE.read_text(encoding="utf-8").strip()
            if u.startswith("http"):
                return u
        except Exception:
            pass
    return "https://quanlyvattucfc.streamlit.app"

PUBLIC_URL = get_public_url()

# ── Xử lý múi giờ Việt Nam (GMT+7) & Đóng gói Cây Báo Cáo ──
def get_vn_now() -> datetime.datetime:
    """Lấy thời gian hiện tại chuẩn múi giờ Việt Nam (GMT+7 / Asia/Ho_Chi_Minh)."""
    try:
        tz = zoneinfo.ZoneInfo("Asia/Ho_Chi_Minh")
        return datetime.datetime.now(tz)
    except Exception:
        return datetime.datetime.utcnow() + datetime.timedelta(hours=7)

def format_vn_time(dt_or_timestamp) -> str:
    """Định dạng ngày giờ chuẩn Việt Nam: DD/MM/YYYY HH:mm:ss."""
    try:
        tz = zoneinfo.ZoneInfo("Asia/Ho_Chi_Minh")
        if isinstance(dt_or_timestamp, (int, float)):
            dt = datetime.datetime.fromtimestamp(dt_or_timestamp, tz=tz)
        elif isinstance(dt_or_timestamp, datetime.datetime):
            dt = dt_or_timestamp.astimezone(tz) if dt_or_timestamp.tzinfo else dt_or_timestamp
        else:
            return str(dt_or_timestamp)
        return dt.strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return str(dt_or_timestamp)

def get_report_file_date(f: Path) -> tuple[int, int, int]:
    """Trích xuất Năm, Tháng, Ngày của file báo cáo chuẩn múi giờ Việt Nam."""
    m = re.search(r'(\d{4})(\d{2})(\d{2})_\d{6}', f.name)
    if m:
        return int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        tz = zoneinfo.ZoneInfo("Asia/Ho_Chi_Minh")
        dt = datetime.datetime.fromtimestamp(f.stat().st_mtime, tz=tz)
        return dt.year, dt.month, dt.day
    except Exception:
        dt = datetime.datetime.now()
        return dt.year, dt.month, dt.day

def organize_reports_into_tree():
    """Tự động đóng gói tất cả các file báo cáo vào cây thư mục Năm / Tháng / Ngày (chuẩn giờ VN)."""
    if not REPORT_DIR.exists():
        return
    for f in list(REPORT_DIR.glob("*.xlsx")):
        if f.is_file():
            y, m, d = get_report_file_date(f)
            target_dir = REPORT_DIR / str(y) / f"Thang_{m:02d}" / f"Ngay_{d:02d}"
            target_dir.mkdir(parents=True, exist_ok=True)
            try:
                shutil.move(str(f), str(target_dir / f.name))
            except Exception:
                pass

# ═══════════════════════════════════════════════════════════════════════════
# CƠ SỞ DỮ LIỆU BỀN VỮNG & BỘ NHỚ ÁNH XẠ (SQLITE materials_memory.db)
# ═══════════════════════════════════════════════════════════════════════════
DB_PATH = BASE_DIR / "materials_memory.db"

_DB_INITIALIZED = False

def init_db():
    global _DB_INITIALIZED
    if _DB_INITIALIZED:
        return
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS approved_materials (
        code TEXT PRIMARY KEY,
        name_normalized TEXT,
        spec TEXT,
        unit TEXT,
        group_code TEXT,
        stock REAL DEFAULT 0,
        source_doc TEXT,
        notes TEXT,
        approved_at TEXT
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS alias_mapping (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        raw_name TEXT UNIQUE,
        target_code TEXT,
        target_name TEXT,
        spec TEXT,
        unit TEXT,
        created_at TEXT
    )
    """)
    conn.commit()
    conn.close()
    _DB_INITIALIZED = True

init_db()

@st.cache_data(ttl=60)
def get_all_approved_materials() -> list[dict]:
    init_db()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM approved_materials ORDER BY approved_at DESC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

@st.cache_data(ttl=60)
def get_all_aliases() -> list[dict]:
    init_db()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM alias_mapping ORDER BY id DESC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

def match_from_alias_memory(raw_query: str, threshold: float = 85.0) -> dict | None:
    """Kiểm tra ưu tiên: Quét Từ điển Ánh xạ đã học trước khi gọi AI."""
    if not raw_query or not str(raw_query).strip():
        return None
    q_norm = unicodedata.normalize("NFC", str(raw_query)).strip().lower()
    aliases = get_all_aliases()
    if not aliases:
        return None

    # 1. Khớp chính xác 100%
    for al in aliases:
        rn = unicodedata.normalize("NFC", str(al.get("raw_name", ""))).strip().lower()
        if rn == q_norm:
            return al

    # 2. Khớp tương đồng cao: Phải cùng danh từ chính và có token_sort_ratio >= 85
    q_words = [w for w in q_norm.split() if len(w) > 1]
    q_head = q_words[0] if q_words else ""

    for al in aliases:
        rn = unicodedata.normalize("NFC", str(al.get("raw_name", ""))).strip().lower()
        rn_words = [w for w in rn.split() if len(w) > 1]
        rn_head = rn_words[0] if rn_words else ""

        # Tuyệt đối không cho 'kim' khớp nhầm với 'máy'
        if q_head and rn_head and q_head != rn_head:
            continue

        sort_score = fuzz.token_sort_ratio(q_norm, rn)
        if sort_score >= threshold:
            return al

    return None

def append_material_to_master_excel(code: str, name_normalized: str, unit: str = "Cái", spec: str = ""):
    """Ghi trực tiếp mã mới vào Master_Inventory.xlsx bảo đảm đầy đủ 12 cột và text có số 0 đầu."""
    if not MASTER_FILE.exists():
        return
    try:
        df = pd.read_excel(MASTER_FILE, dtype={"Mã hàng": str, "Tên kho": str, "Tên hàng": str, "ĐVT": str})
        df = df.loc[:, ~df.columns.duplicated()]
        existing_codes = set(df["Mã hàng"].dropna().astype(str).str.strip()) if "Mã hàng" in df.columns else set()
        if code.strip() in existing_codes:
            return
        full_name = f"{name_normalized} {spec}".strip() if spec else name_normalized
        new_row_dict = {
            "Tên kho": "Kho vật tư phòng Kỹ thuật 01",
            "Mã hàng": str(code).strip(),
            "Tên hàng": full_name,
            "ĐVT": str(unit).strip() if unit else "Cái",
            "Đầu kỳ SL": 0,
            "Đầu kỳ GT": 0,
            "Nhập kho SL": 0,
            "Nhập kho GT": 0,
            "Xuất kho SL": 0,
            "Xuất kho GT": 0,
            "Cuối kỳ SL": 0,
            "Cuối kỳ GT": 0
        }
        df = pd.concat([df, pd.DataFrame([new_row_dict])], ignore_index=True)
        for sc in ["Tên kho", "Mã hàng", "Tên hàng", "ĐVT"]:
            df[sc] = df[sc].astype(str).replace("nan", "").fillna("").str.strip()
        num_cols = ["Đầu kỳ SL", "Đầu kỳ GT", "Nhập kho SL", "Nhập kho GT", "Xuất kho SL", "Xuất kho GT", "Cuối kỳ SL", "Cuối kỳ GT"]
        for nc in num_cols:
            df[nc] = pd.to_numeric(df[nc], errors="coerce").fillna(0)
        df.to_excel(MASTER_FILE, index=False)
    except Exception as e:
        print(f"Error appending to Master_Inventory.xlsx: {e}")

def save_approved_material(
    code: str,
    name_normalized: str,
    spec: str = "",
    unit: str = "Cái",
    group_code: str = "03",
    raw_name: str = "",
    source_doc: str = "",
    notes: str = ""
):
    """Lưu vĩnh viễn: ghi vào SQLite approved_materials, alias_mapping và Master_Inventory.xlsx."""
    init_db()
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
    INSERT INTO approved_materials (code, name_normalized, spec, unit, group_code, stock, source_doc, notes, approved_at)
    VALUES (?, ?, ?, ?, ?, 0, ?, ?, ?)
    ON CONFLICT(code) DO UPDATE SET
        name_normalized=excluded.name_normalized,
        spec=excluded.spec,
        unit=excluded.unit,
        notes=excluded.notes,
        approved_at=excluded.approved_at
    """, (code, name_normalized, spec, unit, group_code, source_doc, notes, now_str))
    
    if raw_name:
        cur.execute("""
        INSERT INTO alias_mapping (raw_name, target_code, target_name, spec, unit, created_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(raw_name) DO UPDATE SET
            target_code=excluded.target_code,
            target_name=excluded.target_name,
            spec=excluded.spec,
            unit=excluded.unit,
            created_at=excluded.created_at
        """, (raw_name.strip(), code, name_normalized, spec, unit, now_str))
        
    conn.commit()
    conn.close()
    
    append_material_to_master_excel(code, name_normalized, unit, spec)
    try:
        load_master_df.clear()
        get_all_approved_materials.clear()
        get_all_aliases.clear()
    except Exception:
        pass

def parse_inventory_upload(file_bytes: bytes, filename: str) -> tuple[pd.DataFrame | None, str]:
    """Quét thông minh file kho Excel/CSV từ phần mềm kế toán/ERP (MISA, FAST, Bravo, Excel tự lập)."""
    try:
        if filename.lower().endswith(".csv"):
            try:
                raw_df = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str, encoding="utf-8")
            except Exception:
                raw_df = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str, encoding="latin1")
        else:
            raw_df = pd.read_excel(io.BytesIO(file_bytes), header=None, dtype=str)
    except Exception as e:
        return None, f"Không thể đọc file: {e}"

    if raw_df.empty:
        return None, "File tải lên không có dữ liệu."

    # Quét tìm dòng header (chứa Mã hàng và Tên hàng) trong 25 dòng đầu tiên
    header_idx = None
    for i in range(min(25, len(raw_df))):
        row_vals = [str(x).strip().lower() for x in raw_df.iloc[i].dropna()]
        has_code = any(any(k in x for k in ["mã hàng", "mã vật tư", "mã vt", "mã sp", "mã hàng hóa", "item code"]) for x in row_vals)
        has_name = any(any(k in x for k in ["tên hàng", "tên vật tư", "tên vt", "tên sp", "tên hàng hóa", "item name"]) for x in row_vals)
        if has_code and has_name:
            header_idx = i
            break

    if header_idx is None:
        row0 = [str(x).strip().lower() for x in raw_df.iloc[0].dropna()]
        if any("mã" in x for x in row0) and any("tên" in x for x in row0):
            header_idx = 0
        else:
            return None, "File không chứa dòng tiêu đề có cột 'Mã hàng' và 'Tên hàng'."

    h_row = raw_df.iloc[header_idx].tolist()
    has_sub = False
    if header_idx + 1 < len(raw_df):
        next_row_vals = [str(x).strip().lower() for x in raw_df.iloc[header_idx + 1].dropna()]
        if any("số lượng" in x or "sl" in x for x in next_row_vals) and any("giá trị" in x or "gt" in x for x in next_row_vals):
            has_sub = True
            sub_row = raw_df.iloc[header_idx + 1].tolist()

    new_cols = []
    last_main = ""
    for col_i in range(len(h_row)):
        main_val = str(h_row[col_i]).strip() if pd.notna(h_row[col_i]) else ""
        if main_val and main_val.lower() != "nan":
            last_main = main_val
        sub_val = ""
        if has_sub and col_i < len(sub_row):
            s = str(sub_row[col_i]).strip() if pd.notna(sub_row[col_i]) else ""
            if s and s.lower() != "nan":
                sub_val = s
        if sub_val:
            col_name = f"{last_main} {sub_val}" if last_main else sub_val
        else:
            col_name = main_val if main_val else f"Col_{col_i}"
        new_cols.append(col_name)

    start_data_idx = header_idx + 2 if has_sub else header_idx + 1
    data_df = raw_df.iloc[start_data_idx:].copy()
    data_df.columns = new_cols

    def _map_col(c: str) -> str:
        c_low = c.lower().strip()
        # Ưu tiên nhận diện các cột số lượng và giá trị 2 cấp
        if "nhập kho" in c_low and any(k in c_low for k in ["số lượng", "sl"]):
            return "Nhập kho SL"
        if "nhập kho" in c_low and any(k in c_low for k in ["giá trị", "gt", "tiền"]):
            return "Nhập kho GT"
        if "xuất kho" in c_low and any(k in c_low for k in ["số lượng", "sl"]):
            return "Xuất kho SL"
        if "xuất kho" in c_low and any(k in c_low for k in ["giá trị", "gt", "tiền"]):
            return "Xuất kho GT"
        if "cuối kỳ" in c_low and any(k in c_low for k in ["số lượng", "sl"]):
            return "Cuối kỳ SL"
        if "cuối kỳ" in c_low and any(k in c_low for k in ["giá trị", "gt", "tiền"]):
            return "Cuối kỳ GT"
        if "đầu kỳ" in c_low and any(k in c_low for k in ["số lượng", "sl"]):
            return "Đầu kỳ SL"
        if "đầu kỳ" in c_low and any(k in c_low for k in ["giá trị", "gt", "tiền"]):
            return "Đầu kỳ GT"
        # Các cột định danh vật tư
        if any(k in c_low for k in ["mã hàng", "mã vật tư", "mã vt", "mã sp", "item code"]):
            return "Mã hàng"
        if any(k in c_low for k in ["tên hàng", "tên vật tư", "tên vt", "tên sp", "item name"]):
            return "Tên hàng"
        if any(k in c_low for k in ["đvt", "đơn vị tính", "đơn vị", "uom", "unit"]):
            return "ĐVT"
        if "tên kho" in c_low or c_low == "kho" or "kho hàng" in c_low:
            return "Tên kho"
        return c

    col_mapping = {c: _map_col(c) for c in data_df.columns}
    data_df = data_df.rename(columns=col_mapping)
    # Khử trùng lặp tên cột nếu có
    data_df = data_df.loc[:, ~data_df.columns.duplicated()]

    if "Mã hàng" not in data_df.columns or "Tên hàng" not in data_df.columns:
        return None, "File thiếu cột bắt buộc 'Mã hàng' hoặc 'Tên hàng'."

    # Lọc bỏ các dòng tiêu đề nhóm hoặc dòng tổng cộng không phải mã vật tư
    data_df = data_df[data_df["Mã hàng"].notna()]
    data_df["Mã hàng"] = data_df["Mã hàng"].astype(str).str.strip()
    data_df = data_df[~data_df["Mã hàng"].str.lower().isin(["", "nan", "none", "null"])]
    data_df = data_df[~data_df["Mã hàng"].str.contains("tên kho|số dòng|cộng|tổng cộng|ngày|từ ngày", case=False, na=False)]

    data_df = data_df[data_df["Tên hàng"].notna()]
    data_df["Tên hàng"] = data_df["Tên hàng"].astype(str).str.strip()
    data_df = data_df[~data_df["Tên hàng"].str.lower().isin(["", "nan", "none", "null"])]

    # Chuyển đổi các cột số lượng và giá trị
    num_cols = ["Cuối kỳ SL", "Cuối kỳ GT", "Đầu kỳ SL", "Đầu kỳ GT", "Nhập kho SL", "Nhập kho GT", "Xuất kho SL", "Xuất kho GT"]
    for num_col in num_cols:
        if num_col in data_df.columns:
            data_df[num_col] = pd.to_numeric(data_df[num_col].astype(str).str.replace(",", "").str.strip(), errors="coerce").fillna(0)
        else:
            data_df[num_col] = 0.0

    if "Tên kho" not in data_df.columns:
        data_df["Tên kho"] = "Kho vật tư phòng Kỹ thuật 01"
    if "ĐVT" not in data_df.columns:
        data_df["ĐVT"] = "Cái"

    # Tổng hợp nếu mã hàng xuất hiện ở nhiều kho hoặc nhiều dòng trong bảng
    agg_dict = {
        "Tên kho": "first",
        "Tên hàng": "first",
        "ĐVT": "first"
    }
    for nc in num_cols:
        agg_dict[nc] = "sum"

    data_df = data_df.groupby("Mã hàng", as_index=False).agg(agg_dict)

    standard_cols = ["Tên kho", "Mã hàng", "Tên hàng", "ĐVT", "Đầu kỳ SL", "Đầu kỳ GT", "Nhập kho SL", "Nhập kho GT", "Xuất kho SL", "Xuất kho GT", "Cuối kỳ SL", "Cuối kỳ GT"]
    final_df = data_df[standard_cols].copy()
    final_df = final_df.reset_index(drop=True)
    return final_df, f"Đã nhận diện {len(final_df)} vật tư."


def update_master_inventory(file_bytes: bytes, filename: str, mode: str = "merge") -> tuple[bool, str]:
    """Cập nhật Master_Inventory: tự động bóc tách định dạng ERP/Kế toán, luôn cố định chế độ Hợp nhất (Merge/Upsert)."""
    try:
        new_df, parse_msg = parse_inventory_upload(file_bytes, filename)
        if new_df is None:
            return False, parse_msg

        # Luôn thực hiện chế độ Hợp nhất (Merge / Upsert)
        if MASTER_FILE.exists():
            cur_df = pd.read_excel(MASTER_FILE, dtype={"Mã hàng": str})
            cur_df["Mã hàng"] = cur_df["Mã hàng"].astype(str).str.strip()
            cur_df = cur_df.loc[:, ~cur_df.columns.duplicated()]
            new_df = new_df.reset_index(drop=True)
            cur_df = cur_df.reset_index(drop=True)
            # Cập nhật thông tin mới từ new_df, giữ các mã cũ không có trong new_df
            new_codes = set(new_df["Mã hàng"])
            remaining_cur = cur_df[~cur_df["Mã hàng"].isin(new_codes)].reset_index(drop=True)
            merged_df = pd.concat([new_df, remaining_cur], ignore_index=True)
        else:
            merged_df = new_df

        merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]
        merged_df = merged_df.reset_index(drop=True)

        # Bảo toàn tất cả các mã đã được kỹ sư duyệt trong SQLite materials_memory.db
        appr = get_all_approved_materials()
        if appr:
            appr_codes = set(merged_df["Mã hàng"])
            rows_to_add = []
            for a in appr:
                c = a["code"]
                if c not in appr_codes:
                    full_name = f"{a['name_normalized']} {a.get('spec', '')}".strip()
                    rows_to_add.append({
                        "Tên kho": "Kho vật tư phòng Kỹ thuật 01",
                        "Mã hàng": c,
                        "Tên hàng": full_name,
                        "ĐVT": a.get("unit", "Cái"),
                        "Đầu kỳ SL": 0,
                        "Đầu kỳ GT": 0,
                        "Nhập kho SL": 0,
                        "Nhập kho GT": 0,
                        "Xuất kho SL": 0,
                        "Xuất kho GT": 0,
                        "Cuối kỳ SL": a.get("stock", 0),
                        "Cuối kỳ GT": 0
                    })
            if rows_to_add:
                merged_df = pd.concat([merged_df, pd.DataFrame(rows_to_add)], ignore_index=True)

        merged_df = merged_df.sort_values(by="Mã hàng").reset_index(drop=True)
        merged_df.to_excel(MASTER_FILE, index=False)
        load_master_df.clear()
        build_search_corpus.clear()
        try:
            get_global_sku_matcher.clear()
        except Exception:
            pass
        return True, f"Hợp nhất thành công kho vật tư: Tổng {len(merged_df):,} mã hàng (Đã cập nhật từ file '{filename}')."
    except Exception as e:
        return False, f"Lỗi cập nhật kho: {e}"

# ── Khởi tạo Custom Component nhận Ctrl + V ──
_paste_component = components.declare_component("ctrl_v_paste_box", path=str(PASTE_BOX_DIR))

def render_ctrl_v_box(key="main_paste_box"):
    return _paste_component(key=key, default=None)

# ═══════════════════════════════════════════════════════════════════════════
# SYSTEM PROMPT 1: DÀNH CHO TAB ĐỐI CHIẾU CHỨNG TỪ (BẮT BUỘC JSON)
# ═══════════════════════════════════════════════════════════════════════════
SYSTEM_PROMPT_RECONCILE = """
Bạn là **Kỹ sư Trưởng phòng Cơ – Điện** kiêm Giám đốc Kỹ thuật nhà máy công nghiệp nặng.
Nhiệm vụ: Phân tích, chuẩn hóa và đối chiếu vật tư kỹ thuật từ chứng từ vào kho gốc nhà máy.

KIẾN THỨC BẮT BUỘC:
• Cơ khí & Truyền động: Vòng bi (bi đạn, tang trống, gối đỡ UCP/UCF), Bánh răng (module m, số răng Z, SCM440/S45C), Hộp giảm tốc (R/K/F/S series), Băng tải (EP200-EP400, gân V, cao su chịu lực), Sàng rung (lưới thép, PU, SUP9), Dây curoa (A/B/C/3V/5V/8V/răng).
• Thiết bị Điện: Động cơ 3 pha (kW ↔ HP), Aptomat MCCB/MCB, Contactor, Rơ le nhiệt, Biến tần (Inverter), Cáp, Đầu cos.
• Tiêu hao & Phụ kiện: Bu lông Gr 8.8/10.9/Inox, Phớt dầu NBR/Viton/FKM, Que hàn E6013/E7016/Inox, Dầu mỡ công nghiệp (EP2, VG68, Gear oil 140).
• Tiếng lóng thợ thuyền: "bạc đạn" = vòng bi, "curoa" = dây đai, "ru lô" = con lăn, "nhông" = bánh răng, "mô tơ" = động cơ, "hộp số" = hộp giảm tốc, "áp tô mát" = aptomat, "khởi động từ" = contactor, "mỡ bò" = mỡ bôi trơn, "que hàn xanh" = E7016, "10 ngựa" = 7.5kW.

NGUYÊN TẮC ĐỐI CHIẾU VÀ MÃ ĐỀ XUẤT (THUẬT TOÁN ĐỐI CHIẾU THÔNG MINH ĐA THUỘC TÍNH):
1. Trích xuất chính xác từng dòng: tên vật tư, quy cách kỹ thuật, số lượng, đơn vị tính.
2. Chuẩn hóa tên kỹ thuật, xóa bỏ cách gọi dân dã/tiếng lóng.
3. QUY ĐỔI ĐỒNG NHẤT QUY CÁCH ÁP DỤNG CHO TẤT CẢ VẬT TƯ TOÀN NHÀ MÁY:
   - Cơ khí / Kim khí / Làm kín: "1,5ly" = "1.5mm" (ly = li = mm); "1m0" = "1m"; "Lục giác M8x25" = "Bu lông M8x25"; "Tắc kê sắt 12x80" = "Bu lông nở 12x80"; "Phốt dầu" = "Phớt"; "Ron/Sin cao su" = "Gioăng".
   - Đường ống, Van & Phụ kiện: "1 inch" = "phi 34" = "DN25"; "2 inch" = "phi 60" = "DN50"; "Co ren" = "Cút"; "Măng sông" = "Nối ren"; "Cà rá" = "Lơ thu"; "Bầu thu" = "Côn thu"; "Rắc co" = "Union".
   - Điện & Thiết bị đóng cắt: "Khởi động từ MC-100a/150a" = "Contactor 100A/150A"; "Cầu dao tép/MCB" = "Aptomat"; "Đèn led bán nguyệt 1.2m/1m2" = "Đèn led 1,2m"; "Đèn led âm trần 9W" = "Bóng đèn led 9w".
   - Truyền động, Dầu mỡ & Tiêu hao: "Mỡ bôi trơn EP2" = "Mỡ bò"; "Dây curoa Bando B65" = "Dây curoa B65"; "Đá cắt 355" = "Đá Cắt 350"; "Bạc đạn" = "Vòng bi"; "Dây rút" = "Lạt nhựa".
4. NGUYÊN TẮC ĐỐI CHIẾU THÔNG MINH (QUY CÁCH + CÔNG DỤNG + TỪ ĐỒNG NGHĨA):
   - TUYỆT ĐỐI KHÔNG cứng nhắc tra đúng 100% chữ cái. Trong nhà máy công nghiệp, vật tư thường có nhiều tên thương mại hoặc hãng sản xuất khác nhau nhưng CÙNG CÔNG DỤNG & CÙNG THÔNG SỐ CỐT LÕI (Ví dụ: Co ren 1" và Co 34 kẽm; Bu lông M8x25 và Lục giác M8*25; Van bi DN50 và Van bi phi 60; Mỡ bôi trơn EP2 và Mỡ bò; Đá cắt 355 và Đá cắt 350; Contactor 100A và MC-100a; Đèn led 1m2 và Đèn led 1,2m).
   - Hãy căn cứ vào danh sách ỨNG VIÊN KHO được cung cấp: Nếu ứng viên có cùng nhóm chủng loại/công dụng và thông số kỹ thuật cốt lõi tương đương, BẮT BUỘC ưu tiên KHỚP về mã kho đó (trang_thai: 'KHỚP', ma_vat_tu_cu: mã kho tương ứng, ma_de_xuat: null). TUYỆT ĐỐI TRÁNH tự ý sinh mã mới gây trùng lặp danh mục kho.
   - KHỚP: Cùng chủng loại, công dụng và thông số -> Trả mã hàng cũ chuẩn (ví dụ 03.1221, 02.1186, 03.0853, 03.1141, 03.0586, 02.0083...), tên trong kho, tồn kho hiện tại.
   - MỚI: Chỉ khi kho hoàn toàn không có vật tư cùng chủng loại/thông số tương đương -> Quy cách mã BẮT BUỘC dạng số 0x.xxxx (ví dụ 03.9347, 03.9348, 02.2841). TUYỆT ĐỐI KHÔNG dùng mã chữ cái như VB.INOX hay TB.KN.
   - CẦN_XÁC_NHẬN: Thiếu thông số quan trọng để phân biệt.

ĐỊNH DẠNG ĐẦU RA BẮT BUỘC:
Trả về khối JSON thuần túy trong block ```json ... ```:
[
  {
    "stt": 1,
    "ten_trich_xuat": "tên trên chứng từ",
    "ten_chuan_hoa": "tên kỹ thuật chuẩn",
    "quy_cach": "thông số quy cách",
    "so_luong": 10,
    "don_vi": "Cái",
    "trang_thai": "KHỚP" | "MỚI" | "CẦN_XÁC_NHẬN",
    "ma_vat_tu_cu": "02.0083 hoặc null",
    "ten_trong_kho": "tên gốc trong kho hoặc null",
    "ton_kho": 120,
    "ma_de_xuat": "03.xxxx hoặc null",
    "ghi_chu": "nhận định ngắn gọn"
  }
]
"""

# ═══════════════════════════════════════════════════════════════════════════
# SYSTEM PROMPT 2: DÀNH CHO TRỢ LÝ CƠ - ĐIỆN (TRẢ LỜI BẢNG & THẺ ĐẸP MẮT)
# ═══════════════════════════════════════════════════════════════════════════
SYSTEM_PROMPT_ASSISTANT = """
Bạn là **Kỹ sư Trưởng phòng Cơ - Điện kiêm Chuyên gia Quản trị Kỹ thuật Nhà máy** của Công ty CP Phân bón & Hóa chất Cần Thơ.
Nhiệm vụ của bạn là giải đáp mọi câu hỏi của kỹ sư, ban giám đốc về:
1. Thông tin tồn kho thực tế của nhà máy (từ dữ liệu 1.318 mã kho được cung cấp).
2. Toàn bộ kiến thức kỹ thuật công nghiệp sâu rộng trên mạng & thực tế: quy cách chi tiết, kích thước danh nghĩa (DN/OD/ID), cấp độ dày (SCH10, SCH20, SCH40, SCH80...), tiêu chuẩn chế tạo (ASTM, JIS, DIN, BS, TCVN, ISO, IEC), barem trọng lượng, vật liệu, hãng sản xuất uy tín và kinh nghiệm lắp đặt bảo trì.

QUY TẮC BẮT BUỘC VỀ ĐỊNH DẠNG:
- TUYỆT ĐỐI KHÔNG xuất mã JSON thô. Không được trả về code block JSON.
- Phải dùng ĐỊNH DẠNG MARKDOWN TRỰC QUAN: Bảng Markdown có kẻ viền rõ ràng, Thẻ thông tin in đậm (Blockquote / Callout), Icon sinh động, danh sách mạch lạc.

BỐ CỤC CÂU TRẢ LỜI LUÔN THEO 3 PHẦN KHOA HỌC:

### 1. 📦 TÌNH TRẠNG TỒN KHO THỰC TẾ TẠI NHÀ MÁY
(Căn cứ vào dữ liệu 1.318 mã kho nội bộ bên dưới)
Trình bày dạng Bảng Markdown:
| Mã hàng | Tên vật tư trong kho | ĐVT | Tồn cuối kỳ | Đánh giá kho |
- Đánh giá kho: "🟢 Còn hàng (Sẵn sàng xuất)", "🟡 Tồn ít (Cần lưu ý)", hoặc "🔴 Hết hàng (Tồn = 0)".
- Nếu trong kho có sẵn các mã liên quan hoặc tương đương: Liệt kê đầy đủ để kỹ sư nắm rõ.
- Nếu kho KHÔNG CÓ: Nêu rõ "Kho hiện chưa có mã đích danh này" và gợi ý các mã gần giống nhất đang có trong kho.

### 2. 🛠️ THÔNG SỐ KỸ THUẬT & TIÊU CHUẨN CÔNG NGHIỆP (KIẾN THỨC MỞ RỘNG)
Trình bày dạng Thẻ thông tin & Bảng barem:
- Đường kính danh nghĩa DN, Đường kính ngoài thực tế (OD mm/inch), Độ dày thành ống (SCH), áp lực làm việc (PN/Bar/PSI).
- Tiêu chuẩn sản xuất quốc tế & Việt Nam: ASTM (A53, A106, A312...), JIS, DIN, BS 1387, TCVN... (đối với ống/thép); SKF/NSK (đối với vòng bi); IEC/Schneider (đối với đồ điện).
- Barem trọng lượng lý thuyết (kg/m hoặc kg/cái) để thuận tiện tính toán vận chuyển và lắp dựng.

### 3. 💡 TƯ VẤN KỸ THUẬT & LƯU Ý BẢO TRÌ NHÀ MÁY
- Phương án thay thế tương đương nếu cần xử lý sự cố gấp.
- Phụ kiện đồng bộ cần xuất kèm (mặt bích, co hàn, tê hàn, gioăng chịu nhiệt/chịu hóa chất, bulon...).
- Lưu ý môi trường phân bón & hóa chất: độ ăn mòn, khuyến nghị vật liệu (mạ kẽm nhúng nóng, Inox 304/316, nhựa composite/HDPE/PPR).
"""

# ═══════════════════════════════════════════════════════════════════════════
# CẤU HÌNH GIAO DIỆN LIGHT TECHNICAL SCADA / INDUSTRIAL HUD
# ═══════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    code, pre, .mono-font {
        font-family: 'JetBrains Mono', monospace !important;
    }
    
    /* Light Modern Background */
    .stApp {
        background-color: #F8FAFC !important;
        background-image: none !important;
        color: #0F172A !important;
    }
    
    /* ═══════════════════════════════════════════════════════════════
       QUY CHUẨN SINGLE-VIEWPORT (ZERO-SCROLL): GỌN TRONG 1 MÀN HÌNH
       ═══════════════════════════════════════════════════════════════ */
    header[data-testid="stHeader"] {
        background: transparent !important;
        height: 1.5rem !important;
    }
    .block-container {
        padding-top: 1.4rem !important;
        padding-bottom: 0.35rem !important;
        padding-left: 1.2rem !important;
        padding-right: 1.2rem !important;
        max-width: 100% !important;
    }
    
    /* Top Bar HUD / Clean White Navbar */
    .hud-header {
        width: 100% !important;
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05), 0 1px 2px rgba(0, 0, 0, 0.03) !important;
        border-radius: 14px !important;
        padding: 0.75rem 1.4rem;
        margin-bottom: 0.65rem;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .hud-title {
        font-size: 1.18rem;
        font-weight: 700;
        letter-spacing: -0.2px;
        color: #0F172A;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    .hud-title span {
        color: #2563EB !important;
        background: none !important;
        -webkit-text-fill-color: initial !important;
    }
    .hud-subtitle {
        font-size: 0.76rem;
        color: #64748B;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        font-weight: 600;
        margin-top: 2px;
    }
    .hud-chips {
        display: flex;
        gap: 8px;
        align-items: center;
    }
    .hud-badge {
        padding: 5px 12px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-family: 'JetBrains Mono', monospace;
        font-weight: 600;
        display: inline-flex;
        align-items: center;
        gap: 6px;
        border: 1px solid #E2E8F0;
        background: #F1F5F9;
        color: #334155;
        text-decoration: none;
        transition: all 0.2s ease;
    }
    .badge-cyan   { border-color: #BFDBFE !important; color: #2563EB !important; background: #EFF6FF !important; }
    .badge-green  { border-color: #A7F3D0 !important; color: #059669 !important; background: #ECFDF5 !important; }
    .badge-public {
        border-color: #FDE68A !important;
        color: #D97706 !important;
        background: #FEF3C7 !important;
        box-shadow: none !important;
    }
    
    /* CSS Pulse Animation for Status Dot */
    @keyframes pulse-ring {
        0% { transform: scale(0.9); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); }
        70% { transform: scale(1.05); box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); }
        100% { transform: scale(0.9); box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }
    }
    .status-dot-pulse {
        width: 7px;
        height: 7px;
        background-color: #10B981;
        border-radius: 50%;
        display: inline-block;
        animation: pulse-ring 1.8s infinite ease-in-out;
    }
    
    /* Sub-tabs modern styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: #FFFFFF !important;
        padding: 4px 8px !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        margin-bottom: 0.65rem;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03) !important;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 6px 16px !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
        color: #64748B !important;
        border-radius: 8px !important;
        border: none !important;
        transition: all 0.2s ease !important;
    }
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: #EFF6FF !important;
        color: #2563EB !important;
        font-weight: 700 !important;
    }
    .stTabs [data-baseweb="tab-highlight"] {
        background-color: #2563EB !important;
    }

    /* Container Card styling */
    div[data-testid="stVerticalBlockBorderWrapper"] {
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 14px !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04) !important;
        padding: 14px !important;
    }
    div[data-testid="stExpander"] {
        background: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 12px !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.03) !important;
        margin-bottom: 8px !important;
    }

    /* Modern Buttons */
    button[kind="primary"] {
        background: #2563EB !important;
        color: #FFFFFF !important;
        border: none !important;
        border-radius: 9999px !important;
        font-weight: 600 !important;
        padding: 0.45rem 1.2rem !important;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2) !important;
        transition: all 0.2s ease !important;
    }
    button[kind="primary"]:hover {
        background: #1D4ED8 !important;
        box-shadow: 0 4px 8px rgba(37, 99, 235, 0.3) !important;
        transform: translateY(-1px);
    }
    button[kind="secondary"] {
        background: #FFFFFF !important;
        color: #1E293B !important;
        border: 1px solid #E2E8F0 !important;
        border-radius: 9999px !important;
        font-weight: 500 !important;
        padding: 0.45rem 1.2rem !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04) !important;
        transition: all 0.2s ease !important;
    }
    button[kind="secondary"]:hover {
        background: #F8FAFC !important;
        border-color: #CBD5E1 !important;
        color: #0F172A !important;
    }

    /* Segmented Control Pill styling */
    [data-testid="stSegmentedControl"] {
        background: #F1F5F9 !important;
        border-radius: 9999px !important;
        padding: 3px !important;
        border: 1px solid #E2E8F0 !important;
    }
    [data-testid="stSegmentedControl"] button {
        border-radius: 9999px !important;
        border: none !important;
        color: #64748B !important;
        font-weight: 600 !important;
    }
    [data-testid="stSegmentedControl"] button[aria-selected="true"] {
        background: #FFFFFF !important;
        color: #2563EB !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08) !important;
    }

    /* ═══════════════════════════════════════════════════════════════
       BẢNG KỸ THUẬT GỌN GÀNG — SÁNG SỦA, TINH TẾ, CHUẨN THIẾT KẾ
       ═══════════════════════════════════════════════════════════════ */
    .scada-compact-table-wrap {
        width: 100% !important;
        max-width: 100% !important;
        border-radius: 12px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
        overflow: hidden;
        margin: 8px 0 12px 0;
        background: #FFFFFF;
    }
    .scada-compact-table {
        width: 100% !important;
        border-collapse: collapse;
        font-family: 'Inter', -apple-system, sans-serif;
        font-size: 0.84rem;
        color: #1E293B;
        table-layout: fixed;
        background: #FFFFFF;
    }
    .scada-compact-table th {
        position: sticky;
        top: 0;
        z-index: 10;
        background: #F8FAFC !important;
        color: #475569 !important;
        font-weight: 700;
        padding: 10px 8px;
        border-bottom: 2px solid #E2E8F0 !important;
        font-size: 0.76rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .scada-compact-table td {
        padding: 10px 8px;
        border-bottom: 1px solid #F1F5F9;
        vertical-align: top;
        white-space: normal !important;
        word-break: break-word !important;
        line-height: 1.4;
        color: #1E293B;
    }
    .scada-compact-table tr:hover {
        background: #F8FAFC !important;
    }

    /* Custom 4px Slim Scrollbar (Zero-Scroll Standard) */
    ::-webkit-scrollbar {
        width: 5px;
        height: 5px;
    }
    ::-webkit-scrollbar-track {
        background: #F1F5F9;
    }
    ::-webkit-scrollbar-thumb {
        background: #CBD5E1;
        border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
        background: #94A3B8;
    }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# QUẢN LÝ DỮ LIỆU KHO MASTER (1.318 MÃ)
# ═══════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def load_master_df() -> pd.DataFrame:
    if not MASTER_FILE.exists():
        return pd.DataFrame()
    df = pd.read_excel(MASTER_FILE, dtype={"Mã hàng": str, "Tên kho": str, "Tên hàng": str, "ĐVT": str})
    df = df.loc[:, ~df.columns.duplicated()]
    for col in ["Tên kho", "Mã hàng", "Tên hàng", "ĐVT"]:
        if col in df.columns:
            df[col] = df[col].astype(str).replace("nan", "").fillna("").str.strip()
    return df


@st.cache_resource(show_spinner=False)
def get_global_sku_matcher() -> UniversalTechnicalSKUMatcher:
    """Khởi tạo hoặc lấy matcher đối chiếu vật tư 5 tầng (Singleton cached in memory)."""
    df = load_master_df()
    return UniversalTechnicalSKUMatcher(df)


def format_vnd_currency(val) -> str:
    """Định dạng số tiền sang chuẩn kế toán Việt Nam: 1.000.000, 10.000.000 (dùng dấu chấm ngăn cách hàng nghìn)."""
    if pd.isna(val) or val is None or str(val).strip() == "" or str(val).strip().lower() in ("nan", "none", "-"):
        return "0"
    try:
        n = float(val)
        if n == 0:
            return "0"
        return f"{int(round(n)):,}".replace(",", ".")
    except Exception:
        return str(val)


def count_files_processed_today() -> int:
    today_vn = get_vn_now().date()
    cnt = 0
    if REPORT_DIR.exists():
        for f in REPORT_DIR.rglob("*.xlsx"):
            if f.is_file():
                try:
                    y, m, d = get_report_file_date(f)
                    if datetime.date(y, m, d) == today_vn:
                        cnt += 1
                except Exception:
                    pass
    return cnt


@st.cache_data
def build_search_corpus(df: pd.DataFrame):
    """Xây dựng chỉ mục chuẩn hóa Unicode NFC cho tiếng Việt."""
    corpus = []
    for idx, row in df.iterrows():
        m = str(row.get("Mã hàng", "")).strip() if pd.notna(row.get("Mã hàng")) else ""
        t = unicodedata.normalize("NFC", str(row.get("Tên hàng", ""))).strip() if pd.notna(row.get("Tên hàng")) else ""
        d = str(row.get("ĐVT", "")).strip() if pd.notna(row.get("ĐVT")) else ""
        ton = row.get("Cuối kỳ SL", 0) if pd.notna(row.get("Cuối kỳ SL")) else 0
        combined = f"{m} | {t} | {d}"
        corpus.append((combined, idx, m, t, d, ton))
    return corpus


# ═══════════════════════════════════════════════════════════════════════════
# THUẬT TOÁN ĐỐI CHIẾU THÔNG MINH 3 CẤP ĐỘ (3-TIER TECHNICAL MATCHER)
# 1. Trùng khớp Tên / Chủng loại cốt lõi
# 2. Trùng khớp Quy cách hàng hóa (Kích thước, Cỡ phi, Độ dày, Mắt lưới, Mã chuẩn)
# 3. Trùng khớp Đặc tính kỹ thuật sử dụng (Đúc vs Hàn vs Kẽm vs Inox vs Áp lực)
# ═══════════════════════════════════════════════════════════════════════════

CAT_TAXONOMY = {
    "thep_ong": {
        "name": "Thép ống các loại",
        "keywords": ["thép ống", "ống thép", "ống đúc", "ống kẽm", "ống áp lực", "uốn ống", "ống tròn"],
        "exclude": ["ống nhựa", "ống gen", "ống mềm", "ống tuýp", "ống cao su", "ống hút", "ống xả", "kính quan sát", "ke đỡ", "kẹp", "băng lót"]
    },
    "ong_inox": {
        "name": "Ống Inox",
        "keywords": ["ống inox", "inox ống", "ống đúc inox"],
        "exclude": ["ống mềm", "ống nối mềm"]
    },
    "ong_nhua_mem": {
        "name": "Ống nhựa & Ống mềm",
        "keywords": ["ống nhựa", "ống pvc", "ống ppr", "ống hdpe", "ống bm", "ống mềm", "ống cao su", "ống khí nén", "ống gen", "ống dẻo"],
        "exclude": []
    },
    "thep_tam": {
        "name": "Thép tấm & Lá thép",
        "keywords": ["thép tấm", "tấm thép", "lá thép", "bản mã"],
        "exclude": ["tấm nhựa", "tấm xốp", "tấm lợp"]
    },
    "thep_hinh": {
        "name": "Thép hình (V/U/I/H)",
        "keywords": ["thép v", "thép u", "thép i", "thép h", "thép góc", "thép hộp", "inox v", "v63", "v50", "v40", "u100", "i150"],
        "exclude": ["ống"]
    },
    "vong_bi": {
        "name": "Vòng bi & Gối đỡ",
        "keywords": ["vòng bi", "bạc đạn", "gối đỡ", "ổ bi", "ucp", "ucf", "ucfl", "ucpa"],
        "exclude": []
    },
    "mai_cat": {
        "name": "Đá mài & Đá cắt",
        "keywords": ["đá mài", "đá cắt", "đĩa mài", "đĩa cắt", "lưỡi cắt", "lưỡi mài", "đá chà"],
        "exclude": []
    },
    "nham": {
        "name": "Nhám đánh bóng",
        "keywords": ["nhám", "giấy nhám", "vải nhám", "bánh nhám", "đĩa nhám"],
        "exclude": []
    },
    "cuoc": {
        "name": "Cước đánh rỉ",
        "keywords": ["cước", "chén cước", "bát cước", "chổi cước"],
        "exclude": []
    },
    "luoi": {
        "name": "Lưới các loại",
        "keywords": ["lưới", "luoi"],
        "exclude": ["ống lưới"]
    },
    "nhua_kt": {
        "name": "Nhựa kỹ thuật",
        "keywords": ["tấm nhựa", "cây nhựa", "nhựa pe", "nhựa pom", "nhựa teflon", "nhựa ptfe", "nhựa mc", "nhựa rrp"],
        "exclude": ["ống nhựa"]
    },
    "bulon": {
        "name": "Bu lông & Ốc vít",
        "keywords": ["bu lông", "bulon", "bulông", "ốc vít", "tán", "đai ốc", "long đền", "vòng đệm"],
        "exclude": []
    },
    "lam_kin": {
        "name": "Phớt & Gioăng làm kín",
        "keywords": ["phớt", "phốt", "gioăng", "sin cao su", "vòng sin", "seal", "o-ring", "vòng đệm chì"],
        "exclude": []
    },
    "curoa": {
        "name": "Dây curoa",
        "keywords": ["curoa", "dây curoa", "dây đai"],
        "exclude": []
    },
    "van": {
        "name": "Van công nghiệp",
        "keywords": ["van bi", "van bướm", "van một chiều", "van cổng", "van an toàn", "van cầu", "van điện từ", "van"],
        "exclude": ["vỏ van", "tay van"]
    },
    "bich": {
        "name": "Mặt bích",
        "keywords": ["mặt bích", "bích"],
        "exclude": []
    },
    "que_han": {
        "name": "Que hàn & Dây hàn",
        "keywords": ["que hàn", "dây hàn", "thuốc hàn"],
        "exclude": []
    },
    "dong_co": {
        "name": "Động cơ & Hộp số",
        "keywords": ["động cơ", "mô tơ", "motor", "hộp số", "hộp giảm tốc"],
        "exclude": []
    },
    "contactor": {
        "name": "Khởi động từ (Contactor)",
        "keywords": ["khởi động từ", "contactor", "mc-", "m-c", "kđt"],
        "exclude": []
    },
    "den_led": {
        "name": "Đèn LED & Chiếu sáng",
        "keywords": ["đèn led", "bóng đèn led", "bóng đèn", "máng đèn", "đuôi đèn", "đèn pha", "đèn âm trần", "đèn bán nguyệt", "đèn tuýp"],
        "exclude": []
    },
    "aptomat": {
        "name": "Aptomat & Thiết bị đóng cắt",
        "keywords": ["aptomat", "áp tô mát", "mcb", "mccb", "cầu dao", "cb"],
        "exclude": []
    },
    "bien_tan": {
        "name": "Biến tần",
        "keywords": ["biến tần", "inverter"],
        "exclude": []
    },
    "ro_le": {
        "name": "Rơ le",
        "keywords": ["rơ le", "relay", "rơle"],
        "exclude": []
    },
    "son": {
        "name": "Sơn & Hóa chất",
        "keywords": ["sơn", "sơn dầu", "sơn chống rỉ", "dung môi", "xăng nhật"],
        "exclude": []
    },
    "phu_kien_ong": {
        "name": "Phụ kiện đường ống & Khớp nối",
        "keywords": ["co", "tê", "măng sông", "cà rá", "bầu thu", "côn thu", "rắc co", "khớp nối mềm", "nắp bịt", "kép ren", "chén hàn", "nút ren"],
        "exclude": []
    },
    "dau_mo": {
        "name": "Dầu mỡ bôi trơn",
        "keywords": ["mỡ bò", "mỡ bôi trơn", "dầu nhớt", "dầu thủy lực", "dầu hộp số", "dầu cầu", "dầu bôi trơn", "dầu động cơ", "dầu máy", "mỡ chịu nhiệt", "vg68", "vg46", "gear oil"],
        "exclude": []
    },
    "dung_cu": {
        "name": "Dụng cụ & Thiết bị thi công",
        "keywords": ["kìm", "cờ lê", "mỏ lết", "cần điếu", "lục giác", "mũi khoan", "taro", "búa", "cảo", "đầu tuýp"],
        "exclude": []
    },
    "bao_ho": {
        "name": "Bảo hộ lao động",
        "keywords": ["găng tay", "khẩu trang", "nón bảo hộ", "mũ bảo hộ", "kính hàn", "kính bảo hộ", "ủng", "dây an toàn"],
        "exclude": []
    },
    "do_kiem": {
        "name": "Thiết bị đo kiểm & Tự động hóa",
        "keywords": ["đồng hồ áp", "áp kế", "nhiệt kế", "cảm biến", "sensor", "công tắc hành trình", "timer", "bộ điều khiển nhiệt", "flow meter"],
        "exclude": []
    },
    "xich_tai": {
        "name": "Xích tải & Nhông truyền động",
        "keywords": ["xích", "nhông xích", "đĩa xích", "khóa xích", "xích tải", "xích đôi"],
        "exclude": []
    }
}

def normalize_factory_terms(text: str) -> str:
    """Chuẩn hóa thuật ngữ, tiếng lóng thợ kỹ thuật, quy đổi đơn vị đo lường cho tất cả vật tư."""
    t = unicodedata.normalize("NFC", str(text)).lower()
    t = re.sub(r'(\d+),(\d+)', r'\1.\2', t)

    # 1. Điện & Thiết bị đóng cắt
    t = re.sub(r'\bkhởi động từ\b|\bkđt\b|\bkdt\b', 'contactor', t)
    t = re.sub(r'\báp tô mát\b|\báp-tô-mát\b|\bap to mat\b', 'aptomat', t)
    t = re.sub(r'\bcầu dao tự động\b|\bcầu dao tép\b', 'aptomat', t)
    t = re.sub(r'\bmcb\b|\bmccb\b', 'aptomat', t)
    t = re.sub(r'\bbóng đèn led\b|\bđèn tuýp led\b|\bđèn led bán nguyệt\b|\bđèn led âm trần\b|\bđèn downlight\b', 'đèn led', t)
    t = re.sub(r'\brơle nhiệt\b|\brơ le nhiệt\b|\brelay nhiệt\b', 'rơle nhiệt', t)
    t = re.sub(r'\bbiến tần\b|\binverter\b', 'biến tần', t)
    t = re.sub(r'\b3 pha\b|\b3p\b|\bba pha\b', '3pha', t)
    t = re.sub(r'\b1 pha\b|\b1p\b|\bmột pha\b', '1pha', t)

    # 2. Cơ khí & Truyền động
    t = re.sub(r'\bbạc đạn\b|\bổ bi\b', 'vòng bi', t)
    t = re.sub(r'\bmô tơ\b|\bmôtơ\b|\bmotor\b', 'động cơ', t)
    t = re.sub(r'\bhộp số giảm tốc\b|\bhộp số\b|\bhộp nhông\b', 'hộp giảm tốc', t)
    t = re.sub(r'\bru lô\b|\brulo\b', 'con lăn', t)
    t = re.sub(r'\bnhông xích\b|\bđĩa xích\b', 'nhông', t)
    t = re.sub(r'\bdây đai\b|\bcuroa\b', 'dây curoa', t)

    # 3. Phụ kiện đường ống & Van
    t = re.sub(r'\bcút ren\b|\bcút hàn\b|\bcút 90\b|\bco lơi\b|\bco 90\b', 'co', t)
    t = re.sub(r'\bchia 3\b|\bnối tê\b', 'tê', t)
    t = re.sub(r'\bống nối ren\b|\bkhớp nối ren\b|\bmăng-sông\b|\bmang song\b', 'măng sông', t)
    t = re.sub(r'\blơ thu\b|\blơ thu ren\b|\bnối thu ren\b|\bca ra\b|\bcà-rá\b', 'cà rá', t)
    t = re.sub(r'\bcôn thu\b|\bống thu\b|\bbầu giảm\b', 'bầu thu', t)
    t = re.sub(r'\brắc-co\b|\bracco\b|\braco\b', 'rắc co', t)
    t = re.sub(r'\bkhớp nối chống rung\b|\bống nối mềm\b', 'khớp nối mềm', t)
    t = re.sub(r'\bnắp bịt\b|\bnút ren\b|\bchén hàn\b', 'nắp bịt', t)

    # 4. Bu lông, ốc vít & Làm kín
    t = re.sub(r'\bbulon\b|\bbulông\b|\blục giác chìm\b|\blục giác thép\b|\blục giác\b', 'bu lông lục giác', t)
    t = re.sub(r'\bđai ốc\b|\becu\b|\bê cu\b|\bê-cu\b', 'tán', t)
    t = re.sub(r'\bvòng đệm\b|\blông đền\b|\blong đen\b', 'long đền', t)
    t = re.sub(r'\bbu lông nở\b|\bnở sắt\b|\bnở rút\b', 'tắc kê sắt', t)
    t = re.sub(r'\bnở inox\b', 'tắc kê inox', t)
    t = re.sub(r'\bphốt dầu\b|\boil seal\b|\bphốt\b', 'phớt', t)
    t = re.sub(r'\bsin cao su\b|\bo-ring\b|\bvòng sin\b', 'gioăng', t)
    t = re.sub(r'\bron\b|\bgasket\b', 'gioăng', t)

    # 5. Dầu mỡ & Hóa chất
    t = re.sub(r'\bmỡ bôi trơn\b|\bmỡ chịu nhiệt\b', 'mỡ bò', t)
    t = re.sub(r'\bbăng tan\b|\bbăng cao su non\b', 'băng keo non', t)
    t = re.sub(r'\blạt nhựa\b', 'dây rút', t)

    # 6. Quy đổi kích thước thông dụng
    t = re.sub(r'\b1m2\b', '1.2m', t)
    t = re.sub(r'\b0m6\b', '0.6m', t)
    t = re.sub(r'(\d+(?:\.\d+)?)\s*(?:ly|li)\b', r'\1mm', t)
    t = re.sub(r'\b([Mm]\d+)\s*[\*xX]\s*(\d+)\b', r'\1x\2', t)
    t = re.sub(r'\b(\d+)\s*[\*]\s*(\d+)\b', r'\1x\2', t)

    # Đá cắt / đá mài: 355mm chính là đá 350
    t = re.sub(r'355\s*(?:x|\*|mm|\b)', '350 ', t)

    # 7. Quy đổi hệ inch ống sang phi / DN phổ biến
    t = re.sub(r'\b1/2\s*(?:inch|\")\b|\bdn15\b', 'phi 21', t)
    t = re.sub(r'\b3/4\s*(?:inch|\")\b|\bdn20\b', 'phi 27', t)
    t = re.sub(r'\b1\s*(?:inch|\")\b|\bdn25\b', 'phi 34', t)
    t = re.sub(r'\b1\s*[-/.]\s*1/4\s*(?:inch|\")\b|\bdn32\b', 'phi 42', t)
    t = re.sub(r'\b1\s*[-/.]\s*1/2\s*(?:inch|\")\b|\bdn40\b', 'phi 49', t)
    t = re.sub(r'\b2\s*(?:inch|\")\b|\bdn50\b', 'phi 60', t)
    t = re.sub(r'\b2\s*[-/.]\s*1/2\s*(?:inch|\")\b|\bdn65\b', 'phi 76', t)
    t = re.sub(r'\b3\s*(?:inch|\")\b|\bdn80\b', 'phi 90', t)
    t = re.sub(r'\b4\s*(?:inch|\")\b|\bdn100\b', 'phi 114', t)

    # 8. Quy đổi HP -> kW
    t = re.sub(r'\b1\s*(?:hp|ngựa)\b', '0.75kw', t)
    t = re.sub(r'\b2\s*(?:hp|ngựa)\b', '1.5kw', t)
    t = re.sub(r'\b3\s*(?:hp|ngựa)\b', '2.2kw', t)
    t = re.sub(r'\b5\s*(?:hp|ngựa)\b', '3.7kw', t)
    t = re.sub(r'\b7\.5\s*(?:hp|ngựa)\b', '5.5kw', t)
    t = re.sub(r'\b10\s*(?:hp|ngựa)\b', '7.5kw', t)
    t = re.sub(r'\b15\s*(?:hp|ngựa)\b', '11kw', t)
    t = re.sub(r'\b20\s*(?:hp|ngựa)\b', '15kw', t)
    t = re.sub(r'\b30\s*(?:hp|ngựa)\b', '22kw', t)

    return t.strip()

def detect_category(text: str) -> str:
    t = unicodedata.normalize("NFC", str(text)).lower()
    for cat_key, info in CAT_TAXONOMY.items():
        if any(ex in t for ex in info["exclude"]):
            continue
        if any(kw in t for kw in info["keywords"]):
            return cat_key
    return "khac"

def extract_tech_specs(text: str) -> dict:
    t = normalize_factory_terms(text)

    spec = {
        "cat": detect_category(t),
        "dia": None,
        "thick": None,
        "length": None,
        "bearing": None,
        "mesh": None,
        "amps": [],
        "watts": [],
        "bolt": None,
        "is_duc": "đúc" in t,
        "is_kem": "kẽm" in t,
        "is_ap_luc": "áp lực" in t,
        "is_inox": "inox" in t or "304" in t or "316" in t
    }

    # 1. Bóc tách dòng định mức Ampe (100A, 150A, 85A, 40A, 22A, 12A, 9A, mc-100a, mc-150a...)
    m_amps = re.findall(r'(?:mc\s*-\s*|\b)(\d+)\s*a\b', t)
    if m_amps:
        spec["amps"] = [int(x) for x in m_amps if 5 <= int(x) <= 2500]

    # 2. Bóc tách công suất Watt & kW (9W, 18W, 25W, 36W, 40W, 150W, 1.5kW, 7.5kW...)
    m_watts = re.findall(r'\b(\d+(?:\.\d+)?)\s*w\b', t)
    if m_watts:
        spec["watts"] = [float(x) for x in m_watts if float(x) <= 50000]
    m_kw = re.findall(r'\b(\d+(?:\.\d+)?)\s*kw\b', t)
    if m_kw:
        spec["watts"].extend([float(x) * 1000 for x in m_kw if float(x) <= 1000])

    # 3. Kích thước mắt lưới (ô 1.5mm, ô 1,5ly, ô 10ly...) - Bóc tách trước chiều dài
    m_mesh = re.search(r'ô\s*(?:vuông\s*)?(\d+(?:\.\d+)?)\s*(?:mm)?', t)
    if m_mesh:
        spec["mesh"] = float(m_mesh.group(1))
        spec["cat"] = "luoi"

    # 4. Bóc tách chiều dài m (1.2m, 1,2m, 1m2, 6m/cây...)
    m_len_dim = re.search(r'\b(\d+)m(\d+)\b', t)
    if m_len_dim:
        spec["length"] = float(f"{m_len_dim.group(1)}.{m_len_dim.group(2)}")
    else:
        m_len = re.search(r'(?:x|\*|\s|^)\s*(\d+(?:\.\d+)?)\s*(?:m/cây|mét|m\b)', t)
        if m_len:
            spec["length"] = float(m_len.group(1))

    # 5. Mã vòng bi chuẩn kỹ thuật (loại trừ trường hợp đèn chiếu sáng có nhiệt độ màu Kelvin như 6500K, 4000K, 3000K)
    t_no_kelvin = re.sub(r'\b\d{3,4}\s*k\b', '', t)
    is_lighting = spec["cat"] == "den_led" or any(k in t for k in ["đèn", "bóng", "chiếu sáng", "led", "âm trần", "bán nguyệt"])
    if not is_lighting:
        m_brg = re.search(r'\b(6\d{2,3}|7\d{3}|2\d{4}|3\d{4}|uc[pf]?\s*\d{3}|n[jup]\s*\d{3,4})(?:[a-z]{1,4})?\b', t_no_kelvin)
        if m_brg:
            code_clean = re.sub(r'\s+', '', m_brg.group(1)).upper()
            spec["bearing"] = code_clean
            spec["cat"] = "vong_bi"

    # 6. Đường kính phi (phi 60, fi 60, d60, ø60, dn50, đá 350...)
    m_dia = re.search(r'(?:phi|fi|d|ø|dn)\s*(\d+(?:\.\d+)?)', t)
    if m_dia:
        spec["dia"] = float(m_dia.group(1))
        if spec["dia"] == 10 and any(k in t for k in ["đá", "mài", "cắt"]):
            spec["dia"] = 100.0  # phi 10 đá mài thợ gọi = 100mm
    elif spec["cat"] in ("mai_cat", "nham"):
        m_disc = re.search(r'\b(350|355|100|125|150|180)\b', t)
        if m_disc:
            spec["dia"] = 350.0 if m_disc.group(1) in ("350", "355") else float(m_disc.group(1))

    # 7. Cặp kích thước AxB (ví dụ: 60x7mm, 27x5mm...)
    m_dim = re.search(r'(?:phi|fi|d|ø)?\s*(\d+(?:\.\d+)?)\s*(?:x|\*)\s*(\d+(?:\.\d+)?)\s*(?:mm)?', t)
    if m_dim:
        v1 = float(m_dim.group(1))
        v2 = float(m_dim.group(2))
        if v1 > v2 and v2 <= 35:
            if spec["dia"] is None:
                spec["dia"] = v1
            if spec["thick"] is None:
                spec["thick"] = v2

    # 8. Độ dày thành ống / tấm (6ly, 6mm, dày 5mm...)
    if spec["thick"] is None:
        m_th = re.search(r'(?:dày\s*)?(\d+(?:\.\d+)?)\s*mm\b', t)
        if m_th:
            v_th = float(m_th.group(1))
            if v_th != spec["dia"]:
                spec["thick"] = v_th

    return spec

def compute_technical_match_pct(q_text: str, cand_text: str, cand_ton: float = 0) -> tuple[int, str]:
    """Tính toán % độ trùng khớp theo 3 tầng kỹ thuật: Tên chủng loại (35%) -> Quy cách (45%) -> Đặc tính sử dụng (20%)."""
    q_norm = normalize_factory_terms(q_text)
    c_norm = normalize_factory_terms(cand_text)
    q_spec = extract_tech_specs(q_norm)
    c_spec = extract_tech_specs(c_norm)

    # ── TẦNG 1: TRÙNG KHỚP TÊN / CHỦNG LOẠI CỐT LÕI (35%) ──
    cat_score = 100
    if q_spec["cat"] != "khac" and c_spec["cat"] != "khac":
        if q_spec["cat"] != c_spec["cat"]:
            COMPATIBLE_PAIRS = [
                {"thep_ong", "ong_inox"},
                {"thep_ong", "phu_kien_ong"},
                {"ong_inox", "phu_kien_ong"},
                {"thep_ong", "ong_nhua_mem"},
                {"aptomat", "contactor"},
                {"ro_le", "contactor"},
                {"bulon", "thep_hinh"}
            ]
            if {q_spec["cat"], c_spec["cat"]} in COMPATIBLE_PAIRS:
                cat_score = 70
            else:
                base_set_r = fuzz.token_set_ratio(q_norm, c_norm)
                if base_set_r >= 80:
                    cat_score = base_set_r
                else:
                    return 0, "Khác chủng loại vật tư"
    else:
        base_fuzz = fuzz.token_set_ratio(q_norm, c_norm)
        if base_fuzz < 35:
            return 0, "Tên gọi không tương đồng"
        cat_score = base_fuzz

    # ── TẦNG 2: TRÙNG KHỚP QUY CÁCH HÀNG HÓA (45%) ──
    spec_score = 100
    spec_notes = []

    # 2.1 Khởi động từ / Contactor: BẮT BUỘC KHỚP DÒNG ĐỊNH MỨC AMPE (100A, 150A...)
    if q_spec["cat"] == "contactor" and c_spec["cat"] == "contactor":
        q_a = set(q_spec.get("amps", []))
        c_a = set(c_spec.get("amps", []))
        if q_a and c_a:
            if q_a & c_a:
                matched_amp = list(q_a & c_a)[0]
                spec_score = 100
                spec_notes.append(f"Cùng dòng định mức {matched_amp}A")
                return 100, f"Trùng khớp Contactor {matched_amp}A trong kho"
            else:
                return 0, f"Khác dòng định mức ({list(q_a)[0]}A vs {list(c_a)[0]}A)"
        else:
            spec_score = max(fuzz.token_set_ratio(q_text, cand_text), fuzz.token_sort_ratio(q_text, cand_text))

    # 2.2 Đèn LED & Thiết bị chiếu sáng: Khớp chiều dài (1.2m = 1,2m = 1m2) hoặc công suất Watt (9W, 36W...)
    elif q_spec["cat"] == "den_led" and c_spec["cat"] == "den_led":
        score_parts = []
        if q_spec.get("length") is not None and c_spec.get("length") is not None:
            if abs(q_spec["length"] - c_spec["length"]) < 0.05:
                score_parts.append(100)
                spec_notes.append(f"Cùng chiều dài {q_spec['length']}m")
            else:
                score_parts.append(40)
        if q_spec.get("watts") and c_spec.get("watts"):
            q_w = set(q_spec["watts"])
            c_w = set(c_spec["watts"])
            if q_w & c_w:
                score_parts.append(100)
                spec_notes.append(f"Cùng công suất {int(list(q_w & c_w)[0])}W")
            else:
                score_parts.append(60)
        if score_parts:
            spec_score = max(score_parts)
            if spec_score == 100:
                all_notes = spec_notes
                return 95, ", ".join(all_notes)
        else:
            spec_score = max(fuzz.token_set_ratio(q_text, cand_text), fuzz.token_sort_ratio(q_text, cand_text))

    # 2.3 Vòng bi: Bắt buộc đúng mã bi
    elif q_spec["bearing"]:
        if c_spec["bearing"] == q_spec["bearing"]:
            spec_score = 100
            spec_notes.append(f"Cùng mã bi {q_spec['bearing']}")
        else:
            return 0, "Khác mã vòng bi"

    # 2.4 Mắt lưới
    elif q_spec["mesh"] is not None and c_spec["mesh"] is not None:
        if abs(q_spec["mesh"] - c_spec["mesh"]) < 0.05:
            spec_score = 100
            spec_notes.append(f"Cùng ô {q_spec['mesh']}mm")
        else:
            diff = abs(q_spec["mesh"] - c_spec["mesh"])
            spec_score = max(30, int(100 - diff * 20))
            spec_notes.append(f"Ô {c_spec['mesh']}mm")

    # 2.5 Đường kính phi & Độ dày
    elif q_spec["dia"] is not None and c_spec["dia"] is not None:
        qd = q_spec["dia"]
        cd = c_spec["dia"]
        if abs(qd - cd) < 0.5:
            dia_score = 100
            spec_notes.append(f"Cùng phi {int(qd)}")
        else:
            ratio = max(qd, cd) / min(qd, cd)
            if ratio > 1.4:
                return 0, "Khác cỡ phi quá lớn"
            dia_score = max(40, int(100 - (ratio - 1.0) * 150))
            spec_notes.append(f"Cỡ phi {int(cd)}")

        thick_score = 100
        if q_spec["thick"] is not None and c_spec["thick"] is not None:
            qt = q_spec["thick"]
            ct = c_spec["thick"]
            if abs(qt - ct) < 0.2:
                thick_score = 100
                spec_notes.append(f"dày {qt}mm")
            else:
                thick_diff = abs(qt - ct)
                thick_score = max(50, int(100 - thick_diff * 15))
                spec_notes.append(f"dày {ct}mm vs {qt}mm")
        elif q_spec["thick"] is not None or c_spec["thick"] is not None:
            thick_score = 80

        spec_score = int(dia_score * 0.7 + thick_score * 0.3)
    else:
        # Dùng max của token_set_ratio và token_sort_ratio sau khi đã chuẩn hóa tiếng lóng kỹ thuật
        set_r = fuzz.token_set_ratio(q_norm, c_norm)
        sort_r = fuzz.token_sort_ratio(q_norm, c_norm)
        spec_score = max(set_r, sort_r)
        if set_r >= 85:
            spec_score = max(spec_score, 90)
            spec_notes.append("Trùng khớp cụm từ khóa chính")

    # ── TẦNG 3: TRÙNG KHỚP ĐẶC TÍNH KỸ THUẬT SỬ DỤNG (20%) ──
    feat_score = 100
    feat_notes = []
    if q_spec["is_duc"]:
        if c_spec["is_duc"]:
            feat_score = 100
            feat_notes.append("thép đúc chịu áp")
        elif c_spec["is_ap_luc"]:
            feat_score = 90
            feat_notes.append("ống chịu áp lực")
        elif c_spec["is_kem"]:
            feat_score = 65
            feat_notes.append("thép kẽm hàn")
        elif c_spec["is_inox"]:
            feat_score = 75
            feat_notes.append("vật liệu Inox")
        else:
            feat_score = 70

    total_pct = int(cat_score * 0.35 + spec_score * 0.45 + feat_score * 0.20)
    total_pct = min(100, max(0, total_pct))

    all_notes = spec_notes + feat_notes
    reason_str = ", ".join(all_notes) if all_notes else "Tương đương kỹ thuật"
    return total_pct, reason_str


def get_candidate_matches(query: str, corpus: list = None, top_k: int = 3, input_price: float = 0.0) -> list[dict]:
    """
    Thuật toán đối chiếu thông minh 5 tầng (Universal Technical SKU Matching Engine):
    - Tầng 1: Tiền xử lý & chuẩn hóa kỹ thuật (thập phân, tấc/phân, phi, dải thông số).
    - Tầng 2: Từ điển đồng nghĩa 10 nhóm ngành kỹ thuật CFC.
    - Tầng 3: Bóc tách cấu trúc vật tư (chủng loại, thông số định danh, model, quy cách).
    - Tầng 4: Chấm điểm trọng số thông minh (Tên 35đ, Thông số 45đ, Model/Quy cách 20đ).
    - Tầng 5: Phân cấp 3 nhóm (EXACT_MATCH >= 75đ, GENERIC_MATCH 60-74đ, SIMILAR 35-59đ).
    """
    if not query or not str(query).strip():
        return []

    try:
        matcher = get_global_sku_matcher()
        match_res = matcher.match_item(str(query), input_price=input_price, top_k=top_k)
        top_cands = match_res.get("top_candidates", [])
        if top_cands:
            results = []
            for c in top_cands:
                score = c["score"]
                reasons_str = ", ".join(c["reasons"]) if c.get("reasons") else "Tương đương kỹ thuật"
                results.append({
                    "score": score,
                    "pct_str": f"[{int(score)}%]",
                    "ma": str(c["ma_hang"]),
                    "ten": str(c["ten_hang"]),
                    "dvt": str(c["dvt"]),
                    "ton": c["cuoi_ky_sl"],
                    "same_cat": True,
                    "ly_do": reasons_str,
                    "tier": c.get("tier", "SIMILAR_MISMATCH"),
                    "don_gia_bq": c.get("don_gia_bq", 0.0),
                    "price_diff_pct": c.get("price_diff_pct"),
                    "data_alerts": c.get("data_alerts", []),
                    "reasons": c.get("reasons", [])
                })
            return results
    except Exception as e:
        pass

    # Fallback dự phòng nếu có corpus
    if not corpus:
        return []
    results = []
    for item in corpus:
        text_val, real_idx, ma, ten, dvt, ton = item
        pct, reason = compute_technical_match_pct(str(query), str(ten), ton)
        if pct >= 45:
            results.append({
                "score": pct,
                "pct_str": f"[{pct}%]",
                "ma": str(ma),
                "ten": str(ten),
                "dvt": str(dvt),
                "ton": ton,
                "same_cat": True,
                "ly_do": reason,
                "tier": "EXACT_MATCH" if pct >= 75 else ("GENERIC_MATCH" if pct >= 60 else "SIMILAR_MISMATCH"),
                "don_gia_bq": 0.0,
                "price_diff_pct": None,
                "data_alerts": [],
                "reasons": [reason]
            })

    results.sort(key=lambda x: x["score"], reverse=True)
    return results[:top_k]


def retrieve_inventory_data_for_chat(query: str, df: pd.DataFrame, max_rows=15) -> str:
    """Tra cứu kho toàn diện phục vụ trả lời chatbot (kết hợp từ khóa + fuzzy)."""
    if df.empty or not query:
        return "Kho hiện chưa có dữ liệu."

    q_norm = unicodedata.normalize("NFC", query).lower().strip()
    stop_words = {"anh", "em", "muốn", "hỏi", "về", "có", "không", "mấy", "cái", "bao", "nhiêu", "trong", "kho", "cho", "xin", "thông", "tin", "xem", "tra", "cứu", "được", "gì"}
    raw_words = [w for w in re.findall(r'\w+', q_norm) if w not in stop_words and len(w) > 1]

    matched_set = set()

    # 1. Tìm theo từ khóa (keywords)
    for idx, row in df.iterrows():
        line = unicodedata.normalize("NFC", f"{row.get('Mã hàng', '')} {row.get('Tên hàng', '')} {row.get('ĐVT', '')}").lower()
        if raw_words and all(w in line for w in raw_words):
            matched_set.add(idx)
        elif len(raw_words) >= 2 and sum(1 for w in raw_words if w in line) >= 2:
            matched_set.add(idx)

    # 2. Bổ sung thêm bằng Fuzzy Matching
    corpus = build_search_corpus(df)
    cands = get_candidate_matches(query, corpus, top_k=10)
    for c in cands:
        if c["score"] >= 40:
            for item in corpus:
                if item[2] == c["ma"]:
                    matched_set.add(item[1])
                    break

    if not matched_set:
        return "THÔNG BÁO TỪ KHO: Không tìm thấy vật tư nào khớp với từ khóa trong 1.318 mã kho."

    matched_df = df.loc[list(matched_set)[:max_rows]]
    table_lines = ["DỮ LIỆU TỒN KHO THỰC TẾ NHÀ MÁY (1.318 MÃ GỐC):"]
    table_lines.append("| Mã hàng | Tên vật tư trong kho | ĐVT | Tồn cuối kỳ | Nhập kỳ | Xuất kỳ | Giá trị tồn kho |")
    table_lines.append("|---|---|---|---|---|---|---|")
    for _, r in matched_df.iterrows():
        ma = r.get("Mã hàng", "")
        ten = r.get("Tên hàng", "")
        dvt = r.get("ĐVT", "")
        ton = r.get("Cuối kỳ SL", 0)
        nhap = r.get("Nhập kho SL", 0)
        xuat = r.get("Xuất kho SL", 0)
        gt = r.get("Cuối kỳ GT", 0)
        gt_str = f"{gt:,.0f} đ" if pd.notna(gt) else "0 đ"
        table_lines.append(f"| {ma} | {ten} | {dvt} | {ton} | {nhap} | {xuat} | {gt_str} |")

    return "\n".join(table_lines)


def allocate_next_codes(df_master: pd.DataFrame, num_needed: int, group: str = "03") -> list[str]:
    """Cấp phát danh sách mã số kế tiếp chuẩn hóa 0x.xxxx (DẠNG SỐ THUẦN TÚY), đồng bộ cả Master & SQLite."""
    group_prefix = f"{group}."
    existing_nums = []
    if not df_master.empty and "Mã hàng" in df_master.columns:
        for code in df_master["Mã hàng"].dropna().astype(str):
            if code.startswith(group_prefix):
                try:
                    suffix = int(code.split(".", 1)[1])
                    existing_nums.append(suffix)
                except ValueError:
                    pass
    try:
        appr_list = get_all_approved_materials()
        for item in appr_list:
            c = str(item.get("code", "")).strip()
            if c.startswith(group_prefix):
                try:
                    suffix = int(c.split(".", 1)[1])
                    existing_nums.append(suffix)
                except ValueError:
                    pass
    except Exception:
        pass
    max_num = max(existing_nums) if existing_nums else 0
    return [f"{group}.{max_num + i:04d}" for i in range(1, num_needed + 1)]


def find_candidates(query: str, df: pd.DataFrame, corpus: list, top_k: int = 3) -> str:
    matches = get_candidate_matches(query, corpus, top_k=top_k)
    lines = []
    for c in matches:
        ly_do_str = f" [{c['ly_do']}]" if c.get("ly_do") else ""
        lines.append(f"  • [{c['score']}%] Mã: {c['ma']} | Tên: {c['ten']} | ĐVT: {c['dvt']} | Tồn: {c['ton']}{ly_do_str}")
    return "\n".join(lines)


def build_candidates_for_items(items_text: str, df: pd.DataFrame) -> str:
    corpus = build_search_corpus(df)
    lines = [l.strip() for l in items_text.split("\n") if l.strip()]
    seen = set()
    result = []
    for line in lines[:40]:
        clean = line[:70]
        if clean not in seen:
            seen.add(clean)
            cands = find_candidates(clean, df, corpus, top_k=3)
            if cands:
                result.append(f"• Item '{clean}':\n{cands}")
    return "\n\n".join(result)


# ═══════════════════════════════════════════════════════════════════════════
# HÀM GỌI GEMINI AN TOÀN — HỖ TRỢ CHỌN SYSTEM PROMPT THEO TỪNG TAB
# ═══════════════════════════════════════════════════════════════════════════
def generate_content_safe(prompt_or_parts, preferred_model="gemini-4-argon", system_instruction=None):
    instruction = system_instruction if system_instruction is not None else SYSTEM_PROMPT_RECONCILE
    candidate_models = [
        preferred_model,
        "gemini-2.0-flash",
        "gemini-1.5-flash",
        "gemini-1.5-pro",
        "gemini-4-argon",
        "gemini-3.8-flash",
        "gemini-3.5-flash",
        "gemini-3.5-flash-lite"
    ]
    models_to_try = []
    for m in candidate_models:
        if m not in models_to_try:
            models_to_try.append(m)

    last_err = None
    for m in models_to_try:
        try:
            mdl = genai.GenerativeModel(model_name=m, system_instruction=instruction)
            return mdl.generate_content(prompt_or_parts)
        except Exception as e:
            last_err = e
            err_msg = str(e).lower()
            if "429" in str(e) or "404" in str(e) or "quota" in err_msg:
                continue
            raise e
    raise last_err


def parse_json_from_response(text: str) -> list | None:
    if not text:
        return None
    # 1. Thử trích xuất từ khối markdown ```json ... ``` hoặc ``` ... ```
    m = re.search(r"```(?:json)?\s*([\s\S]*?)```", text)
    if m:
        try:
            cleaned = m.group(1).strip()
            cleaned = re.sub(r",\s*([\]\}])", r"\1", cleaned)
            return json.loads(cleaned)
        except Exception:
            pass
    # 2. Thử trích xuất mảng JSON [...] trực tiếp từ chuỗi
    m_arr = re.search(r"\[\s*\{[\s\S]*\}\s*\]", text)
    if m_arr:
        try:
            cleaned = m_arr.group(0).strip()
            cleaned = re.sub(r",\s*([\]\}])", r"\1", cleaned)
            return json.loads(cleaned)
        except Exception:
            pass
    # 3. Thử parse trực tiếp
    try:
        cleaned = text.strip()
        cleaned = re.sub(r",\s*([\]\}])", r"\1", cleaned)
        return json.loads(cleaned)
    except Exception:
        return None


def save_batch_report_excel(results: list, batch_name: str = "TongHop") -> Path:
    now_vn = get_vn_now()
    timestamp = now_vn.strftime("%Y%m%d_%H%M%S")
    safe_name = "".join(c if c.isalnum() or c in "._-" else "_" for c in batch_name)
    report_name = f"BaoCao_TongHop_{safe_name}_{timestamp}.xlsx"

    # Đóng gói trực tiếp vào cây thư mục Năm / Tháng / Ngày chuẩn giờ VN
    target_dir = REPORT_DIR / str(now_vn.year) / f"Thang_{now_vn.month:02d}" / f"Ngay_{now_vn.day:02d}"
    target_dir.mkdir(parents=True, exist_ok=True)
    report_path = target_dir / report_name

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tong_Hop_Doi_Chieu"

    headers = [
        "STT", "Chứng từ gốc", "Tên trích xuất", "Tên chuẩn hóa", "Quy cách",
        "Số lượng", "ĐVT", "Phân loại", "Mã kho",
        "Tồn kho", "Gợi ý tương đương / Gần giống", "Mã đề xuất mới (0x.xxxx)", "Ghi chú"
    ]
    h_font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    h_fill = PatternFill(start_color="1A2634", end_color="1A2634", fill_type="solid")
    f_khop = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
    f_moi = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    f_xn = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = h_font
        c.fill = h_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    field_map = [
        "stt", "chung_tu_goc", "ten_trich_xuat", "ten_chuan_hoa", "quy_cach",
        "so_luong", "don_vi", "trang_thai", "ma_vat_tu_cu", "ton_kho",
        "goi_y_tuong_duong", "ma_de_xuat", "ghi_chu"
    ]
    for r, item in enumerate(results, 2):
        for c, key in enumerate(field_map, 1):
            val = item.get(key, "")
            cell = ws.cell(row=r, column=c, value="" if val is None else val)
            cell.border = border
            cell.alignment = Alignment(vertical="center")
        stt = str(item.get("trang_thai", ""))
        fill = f_khop if stt == "KHỚP" else (f_moi if stt == "MỚI" else f_xn)
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).fill = fill

    ws.freeze_panes = "A2"
    wb.save(report_path)
    return report_path


# ═══════════════════════════════════════════════════════════════════════════
# BỘ HIỂN THỊ NỘI DUNG PHIẾU BÁO CÁO GỌN GÀNG (100% TRONG KHUNG, KHÔNG TRƯỢT NGANG)
# ═══════════════════════════════════════════════════════════════════════════
def render_report_detail_html(df: pd.DataFrame) -> str:
    """Hiển thị nội dung phiếu chi tiết gọn gàng 100% trong khung hình, tự động ngắt dòng, tuyệt đối không trượt ngang."""
    if df.empty:
        return "<div style='color:#94a3b8; font-size:0.85rem; padding:8px;'>Phiếu không có dữ liệu.</div>"

    rows = []
    for _, r in df.iterrows():
        stt = r.get("STT", "")
        doc = html.escape(str(r.get("Chứng từ gốc", "") or ""))
        t_raw = html.escape(str(r.get("Tên trích xuất", "") or ""))
        t_norm = html.escape(str(r.get("Tên chuẩn hóa", "") or t_raw))
        spec = html.escape(str(r.get("Quy cách", "") or "-"))
        qty = r.get("Số lượng", "-")
        unit = html.escape(str(r.get("ĐVT", "") or ""))
        cat = str(r.get("Phân loại", "")).upper()

        old_code = r.get("Mã kho", "")
        old_code_str = "" if pd.isna(old_code) or str(old_code).lower() in ["none", "nan", "-"] else str(old_code).strip()

        new_code = r.get("Mã đề xuất mới (0x.xxxx)", "")
        new_code_str = "" if pd.isna(new_code) or str(new_code).lower() in ["none", "nan", "-"] else str(new_code).strip()

        ton = r.get("Tồn kho", 0)
        note = html.escape(str(r.get("Ghi chú", "") or "-"))

        # Phân loại Badge
        if "KHỚP" in cat or old_code_str:
            badge_html = (
                f'<div style="text-align:center;">'
                f'<span style="display:inline-block; padding:3px 8px; border-radius:9999px; font-size:0.72rem; font-weight:700; background:#ECFDF5; color:#059669; border:1px solid #A7F3D0; margin-bottom:3px;">🟢 KHỚP KHO</span>'
                f'<div style="font-family:JetBrains Mono, monospace; font-weight:700; color:#059669; font-size:0.88rem;">{old_code_str}</div>'
                f'<div style="font-size:0.72rem; color:#64748B;">Tồn: <strong style="color:#0F172A;">{ton} {unit}</strong></div>'
                f'</div>'
            )
        else:
            badge_html = (
                f'<div style="text-align:center;">'
                f'<span style="display:inline-block; padding:3px 8px; border-radius:9999px; font-size:0.72rem; font-weight:700; background:#FEF3C7; color:#D97706; border:1px solid #FDE68A; margin-bottom:3px;">🟡 CẤP MÃ MỚI</span>'
                f'<div style="font-family:JetBrains Mono, monospace; font-weight:700; color:#D97706; font-size:0.88rem;">{new_code_str}</div>'
                f'</div>'
            )

        doc_tag = f'<div style="color:#94a3b8; font-size:0.7rem; font-family:JetBrains Mono, monospace; margin-top:2px;">📄 {doc}</div>' if doc and doc != '-' else ''

        row_html = (
            f'<tr style="border-bottom: 1px solid #F1F5F9;">'
            f'<td style="text-align:center; font-family:JetBrains Mono, monospace; font-weight:700; color:#2563EB; font-size:0.82rem; padding:8px 4px;">#{stt}</td>'
            f'<td style="padding:8px 8px; word-break:break-word;">'
            f'<div style="font-weight:700; color:#0F172A; font-size:0.86rem; line-height:1.35; margin-bottom:2px;">{t_norm}</div>'
            f'<div style="color:#64748B; font-size:0.75rem; line-height:1.3;">Gốc: {t_raw}</div>'
            f'{doc_tag}'
            f'</td>'
            f'<td style="padding:8px 8px; color:#2563EB; font-size:0.8rem; line-height:1.35; word-break:break-word; font-weight:500;">{spec}</td>'
            f'<td style="text-align:center; padding:8px 4px; white-space:nowrap;">'
            f'<div style="font-size:0.95rem; font-weight:700; color:#0F172A;">{qty}</div>'
            f'<div style="font-size:0.72rem; color:#64748B; font-weight:600;">{unit}</div>'
            f'</td>'
            f'<td style="padding:8px 4px;">{badge_html}</td>'
            f'<td style="padding:8px 8px; color:#475569; font-size:0.78rem; line-height:1.35; word-break:break-word;">{note}</td>'
            f'</tr>'
        )
        rows.append(row_html)

    table_html = (
        f'<div style="width:100%; border:1px solid #E2E8F0; border-radius:12px; overflow:hidden; background:#FFFFFF; margin-top:6px; box-shadow: 0 1px 3px rgba(0,0,0,0.03);">'
        f'<table style="width:100%; border-collapse:collapse; table-layout:fixed; font-size:0.82rem; text-align:left;">'
        f'<thead>'
        f'<tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-weight:700; font-size:0.75rem; text-transform:uppercase;">'
        f'<th style="width:5%; text-align:center; padding:8px 4px;">STT</th>'
        f'<th style="width:28%; padding:8px 8px;">Vật tư (Chuẩn & Gốc)</th>'
        f'<th style="width:25%; padding:8px 8px;">Quy cách kỹ thuật</th>'
        f'<th style="width:8%; text-align:center; padding:8px 4px;">SL</th>'
        f'<th style="width:14%; text-align:center; padding:8px 4px;">Phân loại & Mã</th>'
        f'<th style="width:20%; padding:8px 8px;">Nhận định / Ghi chú</th>'
        f'</tr>'
        f'</thead>'
        f'<tbody>'
        f'{"".join(rows)}'
        f'</tbody>'
        f'</table>'
        f'</div>'
    )
    return table_html


# ═══════════════════════════════════════════════════════════════════════════
# BỘ HIỂN THỊ BẢNG GỌN GÀNG (100% TRONG 1 MÀN HÌNH, TỰ ĐỘNG XUỐNG DÒNG)
# ═══════════════════════════════════════════════════════════════════════════
def render_compact_table_html(items: list) -> str:
    rows_html = []
    for item in items:
        stt = item.get("stt", 1)
        chung_tu_goc = html.escape(str(item.get("chung_tu_goc", "") or "-"))
        ten_trich_xuat = html.escape(str(item.get("ten_trich_xuat", "") or "-"))
        ten_chuan_hoa = html.escape(str(item.get("ten_chuan_hoa", "") or item.get("ten_trich_xuat", "") or "-"))
        quy_cach = html.escape(str(item.get("quy_cach", "") or "-"))
        so_luong = item.get("so_luong", "-")
        don_vi = html.escape(str(item.get("don_vi", "") or ""))
        trang_thai = str(item.get("trang_thai", "MỚI")).upper()
        ma_vat_tu_cu = html.escape(str(item.get("ma_vat_tu_cu", "") or "-"))
        ton_kho = item.get("ton_kho", 0)
        ma_de_xuat = html.escape(str(item.get("ma_de_xuat", "") or "-"))
        ghi_chu = html.escape(str(item.get("ghi_chu", "") or "-"))

        raw_goi_y = str(item.get("goi_y_tuong_duong", "") or "-")
        if raw_goi_y and raw_goi_y != "-":
            raw_bullets = [b.strip() for b in re.split(r'[|\n]', raw_goi_y) if b.strip()]
            formatted_bullets = []
            for b in raw_bullets[:3]:
                b_clean = b.lstrip("•").strip()
                formatted_bullets.append(f"<div style='margin-bottom:3px; line-height:1.35; color:#334155; font-size:0.8rem;'>• {html.escape(b_clean)}</div>")
            goi_y_html = "".join(formatted_bullets) if formatted_bullets else "<span style='color:#94a3b8;'>-</span>"
        else:
            goi_y_html = "<span style='color:#94a3b8;'>-</span>"

        is_generic = "danh mục chung" in ghi_chu.lower() or item.get("tier") == "GENERIC_MATCH"
        price_alert_html = f"<div style='font-size:0.7rem; color:#D97706; font-weight:600; margin-top:2px;'>{html.escape(str(item['canh_bao_gia']))}</div>" if item.get("canh_bao_gia") else ""
        alert_tags = item.get("canh_bao_kho") or []
        stock_alert_html = "".join([f"<span style='display:inline-block; margin-top:2px; font-size:0.68rem; padding:1px 5px; border-radius:4px; background:#FEE2E2; color:#DC2626; font-weight:600;'>{html.escape(a)}</span>" for a in alert_tags])

        if trang_thai == "KHỚP":
            badge_bg = "#EEF2FF" if is_generic else "#ECFDF5"
            badge_color = "#4338CA" if is_generic else "#059669"
            badge_border = "#C7D2FE" if is_generic else "#A7F3D0"
            badge_text = "🔵 KHỚP MÃ CHUNG" if is_generic else "🟢 KHỚP KHO"
            status_col = (
                f"<div style='text-align:center;'>"
                f"<span style='display:inline-block; padding:3px 9px; border-radius:9999px; font-size:0.72rem; font-weight:700; background:{badge_bg}; color:{badge_color}; border:1px solid {badge_border}; margin-bottom:4px;'>{badge_text}</span>"
                f"<div style='font-family:JetBrains Mono, monospace; font-weight:700; color:{badge_color}; font-size:0.95rem;'>{ma_vat_tu_cu}</div>"
                f"<div style='font-size:0.75rem; color:#64748B;'>Tồn: <strong style='color:#0F172A;'>{ton_kho} {don_vi}</strong></div>"
                f"{stock_alert_html}"
                f"{price_alert_html}"
                f"</div>"
            )
        elif trang_thai == "MỚI":
            status_col = (
                f"<div style='text-align:center;'>"
                f"<span style='display:inline-block; padding:3px 9px; border-radius:9999px; font-size:0.72rem; font-weight:700; background:#FEF3C7; color:#D97706; border:1px solid #FDE68A; margin-bottom:4px;'>🟡 CẤP MÃ MỚI</span>"
                f"<div style='font-family:JetBrains Mono, monospace; font-weight:700; color:#D97706; font-size:0.95rem;'>{ma_de_xuat}</div>"
                f"<div style='font-size:0.72rem; color:#B45309;'>(Mã số mới 0x.xxxx)</div>"
                f"</div>"
            )
        else:
            status_col = (
                f"<div style='text-align:center;'>"
                f"<span style='display:inline-block; padding:3px 9px; border-radius:9999px; font-size:0.72rem; font-weight:700; background:#FEF2F2; color:#DC2626; border:1px solid #FECACA; margin-bottom:4px;'>🔴 CẦN XÁC NHẬN</span>"
                f"<div style='font-size:0.75rem; color:#DC2626;'>Thiếu quy cách</div>"
                f"</div>"
            )

        row = (
            f"<tr>"
            f"<td style='text-align:center; font-family:JetBrains Mono, monospace; font-weight:700; color:#2563EB; font-size:0.88rem;'>#{stt}</td>"
            f"<td>"
            f"<div style='font-weight:700; color:#0F172A; font-size:0.92rem; line-height:1.35; margin-bottom:3px;'>{ten_chuan_hoa}</div>"
            f"<div style='color:#64748B; font-size:0.78rem; line-height:1.35;'><strong>Gốc:</strong> {ten_trich_xuat}</div>"
            f"<div style='color:#94a3b8; font-size:0.72rem; font-family:JetBrains Mono, monospace; margin-top:2px;'>📄 {chung_tu_goc}</div>"
            f"</td>"
            f"<td>"
            f"<div style='color:#2563EB; font-size:0.84rem; font-weight:500; line-height:1.4;'>{quy_cach}</div>"
            f"</td>"
            f"<td style='text-align:center;'>"
            f"<div style='font-size:1.05rem; font-weight:700; color:#0F172A;'>{so_luong}</div>"
            f"<div style='font-size:0.75rem; color:#64748B; font-weight:600;'>{don_vi}</div>"
            f"</td>"
            f"<td>{status_col}</td>"
            f"<td>{goi_y_html}</td>"
            f"<td>"
            f"<div style='font-size:0.8rem; color:#475569; line-height:1.4;'>💡 {ghi_chu}</div>"
            f"</td>"
            f"</tr>"
        )
        rows_html.append(row)

    table_html = (
        "<div class='scada-compact-table-wrap'>"
        "<table class='scada-compact-table'>"
        "<thead>"
        "<tr>"
        "<th style='width: 4%; text-align: center;'>STT</th>"
        "<th style='width: 25%;'>VẬT TƯ (CHUẨN HÓA & GỐC)</th>"
        "<th style='width: 20%;'>QUY CÁCH KỸ THUẬT</th>"
        "<th style='width: 6%; text-align: center;'>SL</th>"
        "<th style='width: 14%; text-align: center;'>PHÂN LOẠI & MÃ</th>"
        "<th style='width: 16%;'>GỢI Ý TƯƠNG ĐƯƠNG</th>"
        "<th style='width: 15%;'>NHẬN ĐỊNH KỸ THUẬT</th>"
        "</tr>"
        "</thead>"
        "<tbody>"
        + "".join(rows_html) +
        "</tbody>"
        "</table>"
        "</div>"
    )
    return table_html


# ═══════════════════════════════════════════════════════════════════════════
# MINI CHAT ASSISTANT (TRỢ LÝ CƠ - ĐIỆN DẠNG POPOVER NỔI CỐ ĐỊNH)
# ═══════════════════════════════════════════════════════════════════════════
@st.fragment
def render_mini_chat_assistant(df_master: pd.DataFrame, api_key: str, model_name: str):
    st.markdown("""
    <div style="font-weight: 700; color: #0F172A; font-size: 0.95rem; margin-bottom: 2px;">
        👨‍🔧 TRỢ LÝ KỸ SƯ TRƯỞNG CƠ - ĐIỆN
    </div>
    <div style="font-size: 0.76rem; color: #64748B; margin-bottom: 8px;">
        Tra cứu kho 1.318 mã + tra barem, tiêu chuẩn công nghiệp (ASTM/JIS/DIN/IEC) và tương đương realtime.
    </div>
    """, unsafe_allow_html=True)

    c_q1, c_q2 = st.columns(2)
    quick_q = None
    with c_q1:
        if st.button("🔩 Thép ống 60?", key="pop_q1", use_container_width=True):
            quick_q = "Kiểm tra thép ống phi 60 trong kho còn bao nhiêu, quy cách tiêu chuẩn thế nào?"
        if st.button("⚡ Động cơ 15kW?", key="pop_q3", use_container_width=True):
            quick_q = "Thông tin động cơ điện 15kW 4P trong kho và thông số dòng định mức, tốc độ?"
    with c_q2:
        if st.button("⚙️ Vòng bi 6205?", key="pop_q2", use_container_width=True):
            quick_q = "Vòng bi 6205 trong kho còn mấy cái, thông số kỹ thuật và hãng tương đương?"
        if st.button("⚠️ Tồn kho = 0?", key="pop_q4", use_container_width=True):
            quick_q = "Liệt kê các mã vật tư đang hết hàng (tồn bằng 0) cần chú ý bổ sung gấp?"

    if "ag1_chat_history" not in st.session_state:
        st.session_state.ag1_chat_history = []

    chat_container = st.container(height=340)
    with chat_container:
        if not st.session_state.ag1_chat_history:
            st.caption("Chưa có tin nhắn. Nhập câu hỏi bên dưới hoặc bấm gợi ý nhanh.")
        for msg in st.session_state.ag1_chat_history:
            with st.chat_message(msg["role"], avatar="👨‍🔧" if msg["role"] == "user" else "🤖"):
                st.markdown(msg["content"])

    input_text = st.chat_input("Hỏi quy cách, tồn kho, tiêu chuẩn...", key="pop_chat_input")
    q = quick_q or input_text

    if q:
        if not api_key:
            st.error("Chưa nhập Google AI Key tại sidebar!")
        else:
            st.session_state.ag1_chat_history.append({"role": "user", "content": q})
            with chat_container:
                with st.chat_message("user", avatar="👨‍🔧"):
                    st.markdown(q)
                with st.chat_message("assistant", avatar="🤖"):
                    with st.spinner("Kỹ sư Trưởng AI đang tra cứu kho 1.318 mã & barem kỹ thuật..."):
                        inv_info = retrieve_inventory_data_for_chat(q, df_master)
                        prompt = f"""
{inv_info}

CÂU HỎI CỦA KỸ SƯ / NGƯỜI DÙNG:
"{q}"

Hãy trả lời thật chi tiết, trực quan, chuyên nghiệp theo đúng 3 phần:
1. 📦 TÌNH TRẠNG TỒN KHO THỰC TẾ TẠI NHÀ MÁY (lập Bảng Markdown rõ ràng gồm Mã hàng, Tên hàng, ĐVT, Tồn cuối kỳ, Đánh giá)
2. 🛠️ THÔNG SỐ KỸ THUẬT & TIÊU CHUẨN CÔNG NGHIỆP (kiến thức mở rộng: quy cách DN, OD, ID, cấp độ dày SCH, barem trọng lượng, tiêu chuẩn ASTM/JIS/DIN/TCVN/IEC...)
3. 💡 TƯ VẤN KỸ THUẬT & LƯU Ý BẢO TRÌ NHÀ MÁY (phụ kiện đi kèm khi lắp đặt, vật tư thay thế tương đương, lưu ý ăn mòn phân bón/hóa chất)

TRÌNH BÀY DẠNG BẢNG VÀ THẺ THÔNG TIN MARKDOWN ĐẸP MẮT. TUYỆT ĐỐI KHÔNG XUẤT JSON.
"""
                        try:
                            ans = generate_content_safe(
                                prompt,
                                preferred_model=model_name,
                                system_instruction=SYSTEM_PROMPT_ASSISTANT
                            ).text
                        except Exception as err:
                            ans = f"Lỗi: {err}"
                        st.markdown(ans)
                        st.session_state.ag1_chat_history.append({"role": "assistant", "content": ans})
            st.rerun()

    if st.session_state.ag1_chat_history:
        if st.button("🗑️ XÓA LỊCH SỬ CHAT", key="pop_clear_chat", use_container_width=True):
            st.session_state.ag1_chat_history = []
            st.rerun()

# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR — CẤU HÌNH & THIẾT LẬP
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### 🤖 PHÂN HỆ ĐANG CHỌN")
    active_name = st.session_state.get("main_agent_segmented_ui") or st.session_state.get("current_main_agent", "🤖 AGENT 1")
    is_ag2 = "AGENT 2" in str(active_name)
    if is_ag2:
        status_color = "#10B981"
        status_label = "⚡ Agent 2: Theo dõi kho xưởng Live"
    else:
        status_color = "#2563EB"
        status_label = "🤖 Agent 1: Đối chiếu & Cấp mã"

    st.markdown(f"""
    <div style="background:#F8FAFC; border:1px solid #CBD5E1; border-radius:8px; padding:8px 10px; margin-bottom:8px;">
        <div style="font-size:0.75rem; color:#64748B; font-weight:600;">ĐANG ĐIỀU HÀNH:</div>
        <div style="font-size:0.86rem; font-weight:700; color:{status_color}; margin-top:2px;">
            {status_label}
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("💡 Chuyển đổi qua lại giữa Agent 1, Agent 2 hoặc bật Chế độ Song song ngay trên đỉnh màn hình chính.")
    st.markdown("---")
    st.markdown("---")
    st.markdown("### ⚙️ THIẾT LẬP HỆ THỐNG")
    
    default_key = os.environ.get("GOOGLE_API_KEY", "")
    if not default_key:
        try:
            default_key = st.secrets.get("GOOGLE_API_KEY", "")
        except Exception:
            pass

    api_key = st.text_input(
        "🔑 GOOGLE AI KEY (Agent 1)",
        type="password",
        value=default_key,
    )
    if api_key:
        os.environ["GOOGLE_API_KEY"] = api_key
        genai.configure(api_key=api_key)

    default_nv_key = os.environ.get("NVIDIA_API_KEY", "")
    if not default_nv_key:
        try:
            default_nv_key = st.secrets.get("NVIDIA_API_KEY", "")
        except Exception:
            pass
    if not default_nv_key:
        default_nv_key = "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"

    nvidia_api_key = st.text_input(
        "🟢 NVIDIA NIM KEY (Agent 2)",
        type="password",
        value=default_nv_key,
        help="Khóa API của NVIDIA giúp Agent 2 chạy độc lập 100% tài nguyên, không nghẽn với Agent 1"
    )
    if nvidia_api_key:
        os.environ["NVIDIA_API_KEY"] = nvidia_api_key

    model_name = st.selectbox(
        "⚡ ENGINE",
        ["gemini-4-argon", "gemini-3.8-flash", "gemini-3.5-flash", "gemini-3.5-flash-lite"],
        index=0,
    )

    df_master = load_master_df()
    current_public_url = get_public_url()
    st.markdown("---")
    st.markdown(f"**📦 KHO GỐC**: `{len(df_master):,}` mã")
    st.markdown(f"**🏠 MẠNG NỘI BỘ LAN**: `http://192.168.1.189:8501`")
    st.markdown(f"**🌐 TOÀN CẦU (INTERNET)**:\n[{current_public_url}]({current_public_url})")
    st.caption("✓ Bất kỳ ai đều truy cập được trên điện thoại / 4G / máy tính ngoài nhà máy, không giới hạn mạng.")
    st.markdown(f"**📁 THƯ MỤC BÁO CÁO**: `Processed_Reports/`")

# ═══════════════════════════════════════════════════════════════════════════
# HEADER CONSOLE — CHÍNH QUY & TOÀN KHUNG HÌNH (FULL-WIDTH HUD)
# ═══════════════════════════════════════════════════════════════════════════
current_public_url = get_public_url()
st.markdown(f"""
<div class="hud-header">
    <div>
        <div class="hud-title">
            ⚙️ <span>HỆ THỐNG AI QUẢN TRỊ & CHUẨN HÓA VẬT TƯ KỸ THUẬT</span>
        </div>
        <div class="hud-subtitle">
            🏢 PHÒNG KỸ THUẬT - CƠ ĐIỆN | CÔNG TY CP PHÂN BÓN & HÓA CHẤT CẦN THƠ
        </div>
    </div>
    <div class="hud-chips">
        <span class="hud-badge badge-green">
            <span class="status-dot-pulse"></span> READY
        </span>
        <span class="hud-badge badge-cyan">KHO: {len(df_master):,} MÃ</span>
        <span class="hud-badge">LAN: 192.168.1.189:8501</span>
        <a class="hud-badge badge-public" href="{current_public_url}" target="_blank" title="Mở link toàn cầu cho internet truy cập">
            🌐 CLOUD ↗
        </a>
    </div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# THANH ĐIỀU HÀNH TRUNG TÂM: 2 PHÂN HỆ ĐỘC LẬP & XỬ LÝ NỀN (BACKGROUND THREADING)
# ═══════════════════════════════════════════════════════════════════════════
MODE_AGENT1 = "🤖 AGENT 1: ĐỐI CHIẾU KHO GỐC"
MODE_AGENT2 = "⚡ AGENT 2: THEO DÕI KHO XƯỞNG (LIVE)"
AGENT_OPTIONS = [MODE_AGENT1, MODE_AGENT2]

# Khởi tạo trạng thái tiến trình ngầm cho Agent 1
if "ag1_is_running" not in st.session_state:
    st.session_state.ag1_is_running = False
if "ag1_task_completed" not in st.session_state:
    st.session_state.ag1_task_completed = False
if "ag1_results" not in st.session_state:
    st.session_state.ag1_results = None
if "ag1_progress_val" not in st.session_state:
    st.session_state.ag1_progress_val = 0.0
if "ag1_status_msg" not in st.session_state:
    st.session_state.ag1_status_msg = ""
if "ag1_task_state" not in st.session_state:
    st.session_state["ag1_task_state"] = {
        "is_running": False,
        "task_completed": False,
        "results": None,
        "progress_val": 0.0,
        "status_msg": "",
        "total_docs": 0,
        "file_path": "",
        "file_name": "",
        "error": None
    }

# Đồng bộ trạng thái từ worker thread vào session state
ts_now = st.session_state["ag1_task_state"]
st.session_state.ag1_is_running = ts_now.get("is_running", False)
st.session_state.ag1_task_completed = ts_now.get("task_completed", False)
st.session_state.ag1_results = ts_now.get("results")
st.session_state.ag1_progress_val = ts_now.get("progress_val", 0.0)
st.session_state.ag1_status_msg = ts_now.get("status_msg", "")

if st.session_state.ag1_task_completed and st.session_state.ag1_results and not st.session_state.get("ag1_reconciled_results"):
    st.session_state["ag1_reconciled_results"] = st.session_state.ag1_results
    st.session_state["ag1_reconciled_total_docs"] = ts_now.get("total_docs", 1)
    st.session_state["ag1_reconciled_file_path"] = ts_now.get("file_path", "")
    st.session_state["ag1_reconciled_file_name"] = ts_now.get("file_name", "")

# Hỗ trợ URL Query Parameters (?tab=agent1, ?tab=agent2)
qp_tab = (st.query_params.get("tab") or st.query_params.get("agent") or "").strip().lower()
if "current_main_agent" not in st.session_state:
    if qp_tab in ["agent2", "2"]:
        st.session_state["current_main_agent"] = MODE_AGENT2
    else:
        st.session_state["current_main_agent"] = MODE_AGENT1

current_agent_active = st.session_state.get("main_agent_segmented_ui") or st.session_state.get("current_main_agent", MODE_AGENT1)
if current_agent_active not in AGENT_OPTIONS:
    current_agent_active = MODE_AGENT1

is_agent2_only = (current_agent_active == MODE_AGENT2)

base_url = get_public_url()
st.markdown(f"""
<div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; margin: 2px 0 6px 0;">
    <span style="font-size:0.73rem; color:#64748B;">Mở tab mới:</span>
    <a href="{base_url}/?tab=agent1" target="_blank" title="Mở Agent 1 trên tab trình duyệt riêng biệt" style="text-decoration:none; font-size:0.72rem; background:#EFF6FF; color:#0071E3; border:1px solid rgba(0,113,227,0.25); padding:2px 8px; border-radius:980px; font-weight:600;">
        🤖 Agent 1 ↗
    </a>
    <a href="{base_url}/?tab=agent2" target="_blank" title="Mở Agent 2 trên tab trình duyệt riêng biệt" style="text-decoration:none; font-size:0.72rem; background:#ECFDF5; color:#10B981; border:1px solid rgba(16,185,129,0.25); padding:2px 8px; border-radius:980px; font-weight:600;">
        ⚡ Agent 2 ↗
    </a>
</div>
""", unsafe_allow_html=True)

# Trợ lý Cơ-Điện popover (chỉ hiển thị tại Agent 1)
if not is_agent2_only:
    with st.sidebar:
        with st.popover("💬 Trợ lý Cơ-Điện (Master)", use_container_width=True):
            render_mini_chat_assistant(df_master, api_key, model_name)

st.markdown("<div style='margin-bottom: 6px;'></div>", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# AGENT 1: BACKGROUND WORKER ENGINE (NON-BLOCKING THREADED PROCESSING)
# ═══════════════════════════════════════════════════════════════════════════
def run_agent1_background_worker(all_queue, df_master, corpus, model_name, task_state):
    """
    Tiến trình chạy ngầm độc lập cho Agent 1.
    Chạy trên background thread, không phụ thuộc UI Streamlit để tránh bị ngắt quãng khi người dùng chuyển sang Agent 2.
    """
    try:
        task_state["is_running"] = True
        task_state["task_completed"] = False
        task_state["results"] = None
        task_state["error"] = None

        total_items = len(all_queue)
        task_state["total_docs"] = total_items
        raw_collected_results = []

        for idx, item in enumerate(all_queue):
            pct = idx / max(total_items, 1)
            task_state["progress_val"] = round(pct * 0.85, 2)
            task_state["status_msg"] = f"🤖 [{idx + 1}/{total_items}] Đang phân tích chứng từ: {item['name']}..."

            save_dest = DAILY_DIR / item["name"]
            try:
                with open(save_dest, "wb") as f:
                    f.write(item["bytes"])
            except Exception:
                pass

            try:
                # ── EXCEL ──
                if item["type"] == "excel":
                    df_in = pd.read_excel(io.BytesIO(item["bytes"]))
                    in_text = df_in.to_string(index=False)
                    cands = build_candidates_for_items(in_text, df_master)
                    prompt = (
                        f"DANH SÁCH VẬT TƯ CHỨNG TỪ '{item['name']}':\n{in_text}\n\n"
                        f"ỨNG VIÊN KHO:\n{cands}\n\n"
                        f"Trả JSON đối chiếu theo quy định."
                    )
                    response = generate_content_safe(prompt, preferred_model=model_name, system_instruction=SYSTEM_PROMPT_RECONCILE)

                # ── PDF ──
                elif item["type"] == "pdf":
                    pdf_txt = ""
                    if HAS_FITZ:
                        try:
                            doc = fitz.open(stream=item["bytes"], filetype="pdf")
                            pdf_txt = "\n".join(page.get_text() for page in doc)
                            doc.close()
                        except Exception as pdf_read_err:
                            print(f"Lỗi đọc text PDF: {pdf_read_err}")
                            pdf_txt = ""

                    if pdf_txt.strip():
                        cands = build_candidates_for_items(pdf_txt, df_master)
                        prompt = (
                            f"NỘI DUNG VĂN BẢN PDF '{item['name']}':\n{pdf_txt}\n\n"
                            f"ỨNG VIÊN KHO:\n{cands}\n\n"
                            f"Bóc tách vật tư và đối chiếu kho. Trả JSON đối chiếu."
                        )
                        response = generate_content_safe(prompt, preferred_model=model_name, system_instruction=SYSTEM_PROMPT_RECONCILE)
                    else:
                        # Scanned PDF hoặc PDF dạng ảnh scan
                        pdf_imgs = render_pdf_pages_as_images(item["bytes"], max_pages=3)
                        if pdf_imgs:
                            cands = build_candidates_for_items(item["name"], df_master)
                            parts = [
                                f"CHỨNG TỪ QUÉT DẠNG ẢNH TỪ PDF '{item['name']}'.\n"
                                f"ỨNG VIÊN KHO GẦN NHẤT:\n{cands}\n\n"
                                f"Bóc tách danh mục vật tư, tên gọi, quy cách, ĐVT, số lượng và trả JSON đối chiếu."
                            ]
                            for p_img in pdf_imgs:
                                parts.append({"mime_type": "image/png", "data": p_img})
                            first_res = generate_content_safe(parts, preferred_model=model_name, system_instruction=SYSTEM_PROMPT_RECONCILE)
                            ai_txt = first_res.text if hasattr(first_res, "text") else ""

                            prelim_data = parse_json_from_response(ai_txt)
                            cand_lines = []
                            if prelim_data and isinstance(prelim_data, list):
                                for p_row in prelim_data:
                                    p_name = p_row.get("ten_trich_xuat", "") or ""
                                    p_spec = p_row.get("quy_cach", "") or ""
                                    p_full = f"{p_name} {p_spec}".strip()
                                    cands_str = find_candidates(p_full, df_master, corpus, top_k=4)
                                    if cands_str:
                                        cand_lines.append(f"• Vật tư '{p_name}' (Quy cách: '{p_spec}'):\n{cands_str}")

                            if cand_lines:
                                followup_prompt = (
                                    f"DANH SÁCH ỨNG VIÊN KHO GỐC (ĐÃ ĐỐI CHIẾU THÔNG MINH THEO CHỦNG LOẠI & QUY CÁCH KỸ THUẬT):\n"
                                    + "\n\n".join(cand_lines)
                                    + "\n\nQUY TẮC BẮT BUỘC KHI ĐỐI CHIẾU:\n"
                                    "1. Quy đổi tương đương chuẩn công nghiệp áp dụng cho TẤT CẢ VẬT TƯ.\n"
                                    "2. Khi danh sách ứng viên có mã kho cùng chủng loại và quy cách/thông số tương đương: "
                                    "BẮT BUỘC KHỚP MÃ CŨ. TUYỆT ĐỐI KHÔNG sinh mã mới khi kho đã có sẵn!\n"
                                    "3. Chỉ cấp mã MỚI (0x.xxxx) khi kho hoàn toàn không có vật tư cùng chủng loại và quy cách.\n"
                                    "Hãy hoàn thiện bảng JSON đối chiếu chính xác nhất."
                                )
                                response = generate_content_safe(f"{ai_txt}\n\n{followup_prompt}", preferred_model=model_name, system_instruction=SYSTEM_PROMPT_RECONCILE)
                            else:
                                response = first_res
                        else:
                            raise ValueError(f"Không thể đọc nội dung văn bản hoặc trích xuất ảnh từ file PDF '{item['name']}'.")

                # ── HÌNH ẢNH ──
                else:
                    mime = item.get("mime", "image/png")
                    cands = build_candidates_for_items(item["name"], df_master)
                    parts = [
                        f"CHỨNG TỪ VẬT TƯ DẠNG ẢNH '{item['name']}'.\n"
                        f"ỨNG VIÊN KHO GẦN NHẤT:\n{cands}\n\n"
                        f"Bóc tách danh sách vật tư dạng bảng gồm tên, quy cách, ĐVT, số lượng. "
                        f"Chuẩn hóa tên gọi và quy cách kỹ thuật. "
                        f"Đối chiếu với kho 1.318 mã để xác định xem là KHỚP mã cũ hay đề xuất MỚI dạng số 0x.xxxx. "
                        f"Trả kết quả trong khối JSON thuần túy.",
                        {"mime_type": mime, "data": item["bytes"]}
                    ]
                    first_res = generate_content_safe(parts, preferred_model=model_name, system_instruction=SYSTEM_PROMPT_RECONCILE)
                    ai_txt = first_res.text

                    prelim_data = parse_json_from_response(ai_txt)
                    cand_lines = []
                    if prelim_data and isinstance(prelim_data, list):
                        for p_row in prelim_data:
                            p_name = p_row.get("ten_trich_xuat", "") or ""
                            p_spec = p_row.get("quy_cach", "") or ""
                            p_full = f"{p_name} {p_spec}".strip()
                            cands_str = find_candidates(p_full, df_master, corpus, top_k=4)
                            if cands_str:
                                cand_lines.append(f"• Vật tư '{p_name}' (Quy cách: '{p_spec}'):\n{cands_str}")
                    else:
                        extracted_names = re.findall(r'"ten_trich_xuat"\s*:\s*"([^"]+)"', ai_txt)
                        for name in extracted_names:
                            cands_str = find_candidates(name, df_master, corpus, top_k=4)
                            if cands_str:
                                cand_lines.append(f"• '{name}':\n{cands_str}")

                    if cand_lines:
                        followup_prompt = (
                            f"DANH SÁCH ỨNG VIÊN KHO GỐC (ĐÃ ĐỐI CHIẾU THÔNG MINH THEO CHỦNG LOẠI & QUY CÁCH KỸ THUẬT):\n"
                            + "\n\n".join(cand_lines)
                            + "\n\nQUY TẮC BẮT BUỘC KHI ĐỐI CHIẾU:\n"
                            "1. Quy đổi tương đương chuẩn công nghiệp áp dụng cho TẤT CẢ VẬT TƯ.\n"
                            "2. Khi danh sách ứng viên có mã kho cùng chủng loại và quy cách/thông số tương đương: "
                            "BẮT BUỘC KHỚP MÃ CŨ. TUYỆT ĐỐI KHÔNG sinh mã mới khi kho đã có sẵn!\n"
                            "3. Chỉ cấp mã MỚI (0x.xxxx) khi kho hoàn toàn không có vật tư cùng chủng loại và quy cách.\n"
                            "Hãy hoàn thiện bảng JSON đối chiếu chính xác nhất."
                        )
                        response = generate_content_safe(f"{ai_txt}\n\n{followup_prompt}", preferred_model=model_name, system_instruction=SYSTEM_PROMPT_RECONCILE)
                    else:
                        response = first_res

                resp_text = response.text if hasattr(response, "text") else str(response)
                doc_results = parse_json_from_response(resp_text)
                if doc_results and isinstance(doc_results, list):
                    for row in doc_results:
                        row["chung_tu_goc"] = item["name"]
                        raw_collected_results.append(row)
                else:
                    task_state["error"] = f"Không trích xuất được bảng danh mục JSON từ file '{item['name']}'."

            except Exception as item_err:
                print(f"Lỗi xử lý file {item.get('name')}: {item_err}")
                task_state["error"] = f"Lỗi phân tích file '{item.get('name', '')}': {str(item_err)}"


            time.sleep(0.5)

        task_state["progress_val"] = 0.90
        task_state["status_msg"] = "🔄 Đang đối chiếu kho master 1.318 mã & cấp mã đề xuất..."

        if raw_collected_results:
            consolidated_results = []
            needed_count = len(raw_collected_results) + 10
            next_03_codes = allocate_next_codes(df_master, needed_count, group="03")
            next_02_codes = allocate_next_codes(df_master, needed_count, group="02")
            idx_03 = 0
            idx_02 = 0

            for i, r in enumerate(raw_collected_results, 1):
                r["stt"] = i
                t_goc = r.get("ten_trich_xuat", "") or ""
                t_chuan = r.get("ten_chuan_hoa", "") or ""
                q_cach = r.get("quy_cach", "") or ""
                search_str = f"{t_goc} {q_cach}".strip()

                # 🔒 [FROZEN CORE START] GHIM CHẾT CODE ĐỐI CHIẾU AGENT 1 - TUYỆT ĐỐI KHÔNG SỬA ĐỔI
                # ĐÃ LOẠI BỎ TOÀN BỘ TRÍ NHỚ ĐÃ HỌC (ALIAS_MEMORY) THEO BƯỚC 2 STRICT CATEGORY GATING
                # Toàn bộ đối chiếu chạy trực tiếp qua 5-Stage Industrial SKU Resolution Pipeline

                try:
                    in_price = float(r.get("don_gia", 0) or 0)
                except Exception:
                    in_price = 0.0

                candidates = get_candidate_matches(search_str, corpus, top_k=5, input_price=in_price)
                r["candidates"] = candidates
                best = candidates[0] if candidates else None

                sim_items = []
                for c in candidates:
                    ly_do = f" [{c['ly_do']}]" if c.get("ly_do") else ""
                    alert_txt = f" {' '.join(c.get('data_alerts', []))}" if c.get("data_alerts") else ""
                    price_txt = f" | BQ: {c['don_gia_bq']:,.0f}đ" if c.get("don_gia_bq", 0) > 0 else ""
                    sim_items.append(f"• [{c['score']}%] {c['ma']}: {c['ten']} (Tồn: {c['ton']}{price_txt}{alert_txt}{ly_do})")
                r["goi_y_tuong_duong"] = "\n".join(sim_items) if sim_items else "Không có mã tương tự"

                # CHỈ TỰ ĐỘNG KHỚP KHI ĐẠT EXACT_MATCH (ĐÚNG CHỦNG LOẠI & QUY CÁCH >= 75%)
                if best and best.get("tier") == "EXACT_MATCH" and best.get("score", 0) >= 85:
                    r["trang_thai"] = "KHỚP"
                    r["ma_vat_tu_cu"] = best["ma"]
                    r["ten_trong_kho"] = best["ten"]
                    r["ton_kho"] = best["ton"]
                    r["ma_de_xuat"] = "-"
                    alert_tag = f" {' '.join(best.get('data_alerts', []))}" if best.get("data_alerts") else ""
                    r["ghi_chu"] = f"✓ Khớp mã kho {best['ma']} ({best.get('ly_do', 'Trùng khớp quy cách kỹ thuật')}){alert_tag}"
                    if in_price > 0 and best.get("don_gia_bq", 0) > 0:
                        diff = round(((in_price - best["don_gia_bq"]) / best["don_gia_bq"]) * 100, 1)
                        if abs(diff) >= 20:
                            sign = "+" if diff > 0 else ""
                            r["canh_bao_gia"] = f"⚠️ Giá lệch {sign}{diff}% (Xuất: {in_price:,.0f}đ vs BQ: {best['don_gia_bq']:,.0f}đ)"
                            r["ghi_chu"] += f" | {r['canh_bao_gia']}"
                elif r.get("trang_thai") == "KHỚP" and r.get("ma_vat_tu_cu") and str(r.get("ma_vat_tu_cu")).strip() not in ["null", "-", "None", ""]:
                    r["ma_de_xuat"] = "-"
                    old_code = str(r.get("ma_vat_tu_cu")).strip()
                    if not r.get("ghi_chu"):
                        r["ghi_chu"] = f"✓ Khớp mã kho {old_code}"
                else:
                    # KHÔNG CÓ TRONG KHO HOẶC CHỈ LÀ GỢI Ý DANH MỤC CHUNG -> BẮT BUỘC ĐỀ XUẤT MÃ MỚI
                    r["trang_thai"] = "MỚI"
                    r["ma_vat_tu_cu"] = None
                    r["ten_trong_kho"] = "-"
                    r["ton_kho"] = 0
                    full_kw = (t_goc + " " + t_chuan + " " + q_cach).lower()
                    if any(w in full_kw for w in ["xi măng", "cát", "đá 1*", "đá 4*", "gạch", "vôi bột", "bê tông tươi"]):
                        r["ma_de_xuat"] = next_02_codes[idx_02]
                        idx_02 += 1
                    else:
                        r["ma_de_xuat"] = next_03_codes[idx_03]
                        idx_03 += 1

                    if best and best.get("tier") == "GENERIC_MATCH":
                        r["ghi_chu"] = f"Chưa có quy cách chính xác trong kho. Đề xuất mã mới {r['ma_de_xuat']} (Gợi ý tương tự: {best['ma']} - {best['ten']})"
                    elif not r.get("ghi_chu"):
                        r["ghi_chu"] = f"Chưa có trong kho. Đề xuất cấp mã mới {r['ma_de_xuat']}"

                consolidated_results.append(r)

            batch_file = save_batch_report_excel(consolidated_results, f"{total_items}_ChungTu")
            task_state["file_path"] = str(batch_file)
            task_state["file_name"] = batch_file.name
            task_state["results"] = consolidated_results
            task_state["progress_val"] = 1.0
            task_state["status_msg"] = f"✓ Hoàn tất đối chiếu toàn bộ {total_items} chứng từ!"
            task_state["task_completed"] = True
            task_state["error"] = None
        else:
            if not task_state.get("error"):
                task_state["error"] = "Không tìm thấy hoặc không bóc tách được danh mục vật tư nào từ chứng từ đã nạp. Vui lòng kiểm tra lại chất lượng file hoặc hình ảnh!"
            task_state["status_msg"] = "⚠️ Không trích xuất được vật tư."
            task_state["task_completed"] = False

    except Exception as e:
        task_state["error"] = str(e)
        task_state["status_msg"] = f"Lỗi xử lý: {str(e)}"
    finally:
        task_state["is_running"] = False


# ═══════════════════════════════════════════════════════════════════════════
# HÀM HIỂN THỊ PHÂN HỆ: AGENT 1 (ĐỐI CHIẾU CHỨNG TỪ & KHO MASTER)
# ═══════════════════════════════════════════════════════════════════════════
@st.fragment
def render_agent1_ui(df_master: pd.DataFrame, api_key: str = "", model_name: str = ""):
    # ═══════════════════════════════════════════════════════════════════════════
    # PHÂN HỆ: AGENT 1 — 3 SUB-TABS ĐIỀU HƯỚNG CHỨC NĂNG CON
    # ═══════════════════════════════════════════════════════════════════════════
    tab_scan, tab_master, tab_reports = st.tabs([
        "📑 ĐỐI CHIẾU CHỨNG TỪ",
        "📦 KHO VẬT TƯ & CẬP NHẬT DANH MỤC",
        "📋 BÁO CÁO / PHIẾU ĐÃ XỬ LÝ"
    ])

    # ═══════════════════════════════════════════════════════════════════════════
    # SUB-TAB 1: ĐỐI CHIẾU CHỨNG TỪ (CORE / MẶC ĐỊNH)
    # ═══════════════════════════════════════════════════════════════════════════
    with tab_scan:
        corpus = build_search_corpus(df_master)

        # Đồng bộ kết quả từ tiến trình ngầm vào Session State của Agent 1 (Bắt buộc cho @st.fragment)
        ts_live = st.session_state.get("ag1_task_state", {})
        if ts_live.get("task_completed") and ts_live.get("results") and not st.session_state.get("ag1_reconciled_results"):
            st.session_state["ag1_reconciled_results"] = ts_live.get("results")
            st.session_state["ag1_reconciled_total_docs"] = ts_live.get("total_docs", 1)
            st.session_state["ag1_reconciled_file_path"] = ts_live.get("file_path", "")
            st.session_state["ag1_reconciled_file_name"] = ts_live.get("file_name", "")

        # Hiển thị thông báo lỗi từ tiến trình ngầm nếu có
        if ts_live.get("error") and not ts_live.get("is_running"):
            st.error(f"⛔ {ts_live['error']}")
            if st.button("🔄 Đóng thông báo & thử lại", key="btn_clear_ag1_err"):
                ts_live["error"] = None
                st.rerun(scope="fragment")

        # Khởi tạo danh sách ảnh dán clipboard trong Session State
        if "ag1_pasted_gallery" not in st.session_state:
            st.session_state["ag1_pasted_gallery"] = []
        if "ag1_last_pasted_hash" not in st.session_state:
            st.session_state["ag1_last_pasted_hash"] = None

        # ── KHU VỰC NHẬP LIỆU: GOM DÁN ẢNH & TẢI FILE CÙNG 1 CHỖ (GỌN GÀNG, CHUYÊN NGHIỆP) ──
        has_results = bool(st.session_state.get("ag1_reconciled_results"))

        def _render_intake_hub():
            st.markdown("""
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                <div style="font-size: 1.4rem; color: #2563EB; font-weight: 800; line-height: 1;">➕</div>
                <div>
                    <div style="font-weight: 700; color: #0F172A; font-size: 0.95rem;">
                        CỔNG NẠP CHỨNG TỪ ĐA NĂNG (HỖ TRỢ QUÉT & PHÂN TÍCH SONG SONG)
                    </div>
                    <div style="font-size: 0.75rem; color: #64748B;">
                        Chọn hoặc kéo thả đồng thời nhiều file cùng lúc: <strong>Ảnh chụp</strong> (PNG, JPG, WEBP) & <strong>Tài liệu</strong> (PDF, XLSX, XLS, CSV)
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            uploaded_files = st.file_uploader(
                "➕ Bấm để chọn nhiều file hoặc kéo thả vào đây (PDF, Excel, Ảnh, CSV):",
                type=["pdf", "xlsx", "xls", "csv", "png", "jpg", "jpeg", "webp", "bmp"],
                accept_multiple_files=True,
                key="ag1_unified_file_uploader",
                help="Chọn đồng thời tất cả các hóa đơn, phiếu xuất kho cần đối chiếu một lần"
            )

            # Khung hỗ trợ dán ảnh nhanh tự động (Ctrl + V)
            raw_pasted_data = render_ctrl_v_box(key="ctrl_v_catcher")
            if raw_pasted_data and isinstance(raw_pasted_data, str) and ";base64," in raw_pasted_data:
                try:
                    _, b64_part = raw_pasted_data.split(";base64,")
                    pasted_bytes = base64.b64decode(b64_part)
                    img_hash = hashlib.md5(pasted_bytes).hexdigest()

                    if img_hash != st.session_state["ag1_last_pasted_hash"]:
                        st.session_state["ag1_last_pasted_hash"] = img_hash
                        st.session_state["ag1_pasted_gallery"].append({
                            "id": img_hash,
                            "name": f"Clip_{datetime.datetime.now().strftime('%H%M%S')}_{len(st.session_state['ag1_pasted_gallery'])+1}.png",
                            "bytes": pasted_bytes,
                        })
                        st.rerun(scope="fragment")
                except Exception as e:
                    st.error(f"Lỗi đọc clipboard: {e}")

            # Hiển thị ảnh dán từ Clipboard nếu có
            n_pasted = len(st.session_state["ag1_pasted_gallery"])
            if n_pasted > 0:
                col_title, col_clear = st.columns([4, 1])
                with col_title:
                    st.caption(f"📸 Đã dán từ Clipboard: **{n_pasted} ảnh** *(Sẽ được quét song song cùng các file trên)*")
                with col_clear:
                    if st.button("🗑️ Xóa ảnh dán", use_container_width=True, key="btn_clear_pastes_dock"):
                        st.session_state["ag1_pasted_gallery"] = []
                        st.session_state["ag1_last_pasted_hash"] = None
                        st.rerun(scope="fragment")

                cols_per_row = 6
                rows = [st.session_state["ag1_pasted_gallery"][i:i + cols_per_row] for i in range(0, n_pasted, cols_per_row)]
                for r_idx, row_items in enumerate(rows):
                    thumb_cols = st.columns(cols_per_row)
                    for c_idx, item in enumerate(row_items):
                        with thumb_cols[c_idx]:
                            st.image(item["bytes"], caption=item["name"], use_container_width=True)
                            if st.button("❌", key=f"del_paste_{item['id']}"):
                                st.session_state["ag1_pasted_gallery"] = [
                                    x for x in st.session_state["ag1_pasted_gallery"] if x["id"] != item["id"]
                                ]
                                st.rerun(scope="fragment")

            return uploaded_files

        if has_results:
            with st.expander("➕ NẠP THÊM CHỨNG TỪ / TIẾP TỤC ĐỐI CHIẾU (QUÉT SONG SONG)", expanded=False):
                uploaded_files = _render_intake_hub()
        else:
            with st.container(border=True):
                uploaded_files = _render_intake_hub()

        # ── TỔNG HỢP TẤT CẢ CHỨNG TỪ TRONG HÀNG ĐỢI (QUÉT SONG SONG) ──
        all_queue = []
        active_uploads = uploaded_files if uploaded_files is not None else st.session_state.get("ag1_unified_file_uploader")
        if active_uploads:
            for uf in active_uploads:
                ext = uf.name.rsplit(".", 1)[-1].lower() if "." in uf.name else ""
                if hasattr(uf, "getvalue"):
                    file_bytes = uf.getvalue()
                elif hasattr(uf, "read"):
                    if hasattr(uf, "seek"):
                        uf.seek(0)
                    file_bytes = uf.read()
                else:
                    file_bytes = bytes(uf)

                ftype = "excel" if ext in ("xlsx", "xls", "csv") else ("pdf" if ext == "pdf" else "image")
                if len(file_bytes) > 0:
                    all_queue.append({
                        "name": uf.name,
                        "bytes": file_bytes,
                        "ext": ext,
                        "type": ftype
                    })

        for p in st.session_state.get("ag1_pasted_gallery", []):
            all_queue.append({
                "name": p["name"],
                "bytes": p["bytes"],
                "ext": "png",
                "type": "image"
            })

        n_pasted = len(st.session_state["ag1_pasted_gallery"])

        # ── KHI CÓ CHỨNG TỪ: NÚT BẮT ĐẦU ĐỐI CHIẾU HÀNG LOẠT ──
        if all_queue:
            st.markdown(f"""
            <div style="background: #EFF6FF; border: 1px solid #BFDBFE; border-radius: 12px; padding: 12px 18px; margin: 12px 0; display: flex; justify-content: space-between; align-items: center; box-shadow: 0 1px 3px rgba(37, 99, 235, 0.05);">
                <div style="font-weight: 700; color: #1E40AF; font-size: 0.98rem;">
                    ⚡ ĐÃ SẴN SÀNG: <span style="color: #2563EB;">{len(all_queue)} CHỨNG TỪ</span> ({n_pasted} ảnh dán + {len(uploaded_files or [])} file đính kèm)
                </div>
                <div style="color: #64748B; font-size: 0.82rem;">Sẽ tổng hợp thành 01 Báo cáo thống nhất</div>
            </div>
            """, unsafe_allow_html=True)

            ts_live = st.session_state["ag1_task_state"]
            if ts_live.get("is_running", False):
                st.markdown("""
                <div style="background:#EFF6FF; border:1px solid #93C5FD; border-radius:10px; padding:10px 14px; margin-bottom:10px;">
                    <div style="font-size:0.85rem; font-weight:700; color:#1E40AF;">⏳ TIẾN TRÌNH ĐỐI CHIẾU ĐANG CHẠY NGẦM...</div>
                    <div style="font-size:0.75rem; color:#3B82F6; margin-top:2px;">
                        Tiến trình được bảo vệ độc lập trong background worker. Bạn có thể tự do chuyển sang Agent 2 hoặc các chức năng khác mà không lo bị gián đoạn!
                    </div>
                </div>
                """, unsafe_allow_html=True)
                st.progress(ts_live.get("progress_val", 0.05))
                st.caption(f"Trạng thái: **{ts_live.get('status_msg', 'Đang xử lý...')}**")
                time.sleep(1)
                st.rerun(scope="fragment")

            elif st.button(f"🚀 BẮT ĐẦU ĐỐI CHIẾU SONG SONG ({len(all_queue)} CHỨNG TỪ ĐÃ NẠP)", type="primary", use_container_width=True):
                if not api_key:
                    st.error("⛔ Vui lòng nhập Google AI Key tại sidebar bên trái!")
                else:
                    ts_live["is_running"] = True
                    ts_live["task_completed"] = False
                    ts_live["results"] = None
                    ts_live["error"] = None
                    ts_live["progress_val"] = 0.05
                    ts_live["status_msg"] = f"Bắt đầu khởi động background worker cho {len(all_queue)} chứng từ..."
                    st.session_state["ag1_reconciled_results"] = None

                    queue_copy = list(all_queue)
                    corpus = build_search_corpus(df_master)
                    th = threading.Thread(
                        target=run_agent1_background_worker,
                        args=(queue_copy, df_master, corpus, model_name, ts_live),
                        daemon=True
                    )
                    if add_script_run_ctx:
                        try:
                            add_script_run_ctx(th)
                        except Exception:
                            pass
                    th.start()
                    st.rerun(scope="fragment")

        # ═══════════════════════════════════════════════════════════════════════════
        # KHU VỰC HIỂN THỊ KẾT QUẢ ĐỐI CHIẾU TRỰC QUAN (TỐI ƯU CỤC BỘ @st.fragment)
        # ═══════════════════════════════════════════════════════════════════════════
        @st.fragment
        def render_reconciled_results_panel(corpus):
            if not st.session_state.get("ag1_reconciled_results"):
                return
            res_list = st.session_state["ag1_reconciled_results"]
            total_docs = st.session_state.get("ag1_reconciled_total_docs", 1)
            file_path = st.session_state.get("ag1_reconciled_file_path")
            file_name = st.session_state.get("ag1_reconciled_file_name", "BaoCao_DoiChieu.xlsx")

            n_khop = sum(1 for r in res_list if r.get("trang_thai") == "KHỚP")
            n_moi = sum(1 for r in res_list if r.get("trang_thai") == "MỚI")
            n_xn = sum(1 for r in res_list if r.get("trang_thai") == "CẦN_XÁC_NHẬN")

            # ── THANH CÔNG CỤ ĐIỀU KHIỂN DUY NHẤT (SINGLE-ROW HUD TOOLBAR - ZERO-SCROLL) ──
            c_tb1, c_tb2, c_tb3, c_tb4 = st.columns([3.4, 2.2, 2.4, 2.0], gap="small")
            with c_tb1:
                status_filter = st.segmented_control(
                    "Lọc trạng thái:",
                    options=[
                        f"Tất cả ({len(res_list)})",
                        f"🟢 Khớp ({n_khop})",
                        f"🟡 Mới ({n_moi})",
                        f"🔴 Cần xem ({n_xn})"
                    ],
                    default=f"Tất cả ({len(res_list)})",
                    key="res_filter_segmented",
                    label_visibility="collapsed"
                )
            with c_tb2:
                search_kw = st.text_input(
                    "Tìm nhanh:",
                    placeholder="🔍 Tìm tên, mã, quy cách...",
                    key="filter_search_kw",
                    label_visibility="collapsed"
                )
            with c_tb3:
                if n_moi > 0:
                    if st.button(f"✅ DUYỆT {n_moi} MÃ MỚI", type="primary", use_container_width=True, key="btn_appr_all_top"):
                        approved_count = 0
                        for r in res_list:
                            if r.get("trang_thai") == "MỚI":
                                code = r.get("ma_de_xuat")
                                if code and code != "-":
                                    name_norm = r.get("ten_chuan_hoa") or r.get("ten_trich_xuat")
                                    save_approved_material(
                                        code=code,
                                        name_normalized=name_norm,
                                        spec=r.get("quy_cach", ""),
                                        unit=r.get("don_vi", "Cái"),
                                        group_code="02" if str(code).startswith("02.") else "03",
                                        raw_name=r.get("ten_trich_xuat", ""),
                                        source_doc=r.get("chung_tu_goc", ""),
                                        notes=r.get("ghi_chu", "")
                                    )
                                    r["trang_thai"] = "KHỚP"
                                    r["ma_vat_tu_cu"] = code
                                    r["ten_trong_kho"] = name_norm
                                    r["ton_kho"] = 0
                                    r["ma_de_xuat"] = "-"
                                    r["goi_y_tuong_duong"] = f"{code} (Vừa duyệt vào kho)"
                                    r["ghi_chu"] = f"✓ Đã duyệt & học vào bộ nhớ kho (Mã: {code})"
                                    approved_count += 1
                        if approved_count > 0:
                            st.session_state["ag1_reconciled_results"] = res_list
                            batch_file = save_batch_report_excel(res_list, f"{total_docs}_ChungTu")
                            st.session_state["ag1_reconciled_file_path"] = str(batch_file)
                            st.toast(f"✓ Đã phê duyệt và lưu {approved_count} mã mới vào kho vĩnh viễn!", icon="✅")
                            time.sleep(0.3)
                            st.rerun(scope="fragment")
                else:
                    st.markdown("<div style='text-align:center; padding: 7px 0; font-weight:700; color:#059669; font-size:0.82rem; border:1px solid #A7F3D0; border-radius:9999px; background:#ECFDF5;'>✓ ĐÃ CHUẨN HÓA 100%</div>", unsafe_allow_html=True)
            with c_tb4:
                c_dl_btn, c_del_btn = st.columns([1.2, 1], gap="small")
                with c_dl_btn:
                    if file_path and Path(file_path).exists():
                        with open(file_path, "rb") as rf:
                            st.download_button(
                                label="💾 Excel",
                                data=rf.read(),
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True,
                                key="btn_dl_top"
                            )
                with c_del_btn:
                    if st.button("🗑️ Xóa", use_container_width=True, key="btn_clear_top", help="Xóa kết quả này"):
                        st.session_state["ag1_reconciled_results"] = None
                        st.session_state["ag1_reconciled_file_path"] = None
                        st.rerun(scope="fragment")

            # Lọc danh sách
            filtered_items = res_list
            if status_filter and "🟢 Khớp" in status_filter:
                filtered_items = [r for r in filtered_items if r.get("trang_thai") == "KHỚP"]
            elif status_filter and "🟡 Mới" in status_filter:
                filtered_items = [r for r in filtered_items if r.get("trang_thai") == "MỚI"]
            elif status_filter and "🔴 Cần" in status_filter:
                filtered_items = [r for r in filtered_items if r.get("trang_thai") == "CẦN_XÁC_NHẬN"]

            if search_kw:
                kw = search_kw.strip().lower()
                filtered_items = [
                    r for r in filtered_items
                    if kw in str(r.get("ten_trich_xuat", "")).lower()
                    or kw in str(r.get("ten_chuan_hoa", "")).lower()
                    or kw in str(r.get("quy_cach", "")).lower()
                    or kw in str(r.get("ma_vat_tu_cu", "")).lower()
                    or kw in str(r.get("ma_de_xuat", "")).lower()
                    or kw in str(r.get("ghi_chu", "")).lower()
                ]

            # ── 2 TABS NỘI BỘ VỪA VẶN TRONG KHUNG HÌNH (ZERO-SCROLL) ──
            tab_res_table, tab_res_override = st.tabs([
                f"📊 BẢNG KẾT QUẢ ĐỐI CHIẾU ({len(filtered_items)} DÒNG VẬT TƯ)",
                f"🛠️ ĐỔI MÃ & CHỌN TỪ GỢI Ý KHO ({len(filtered_items)})"
            ])

            with tab_res_table:
                if not filtered_items:
                    st.warning("⚠️ Không tìm thấy vật tư nào phù hợp với bộ lọc hiện tại.")
                else:
                    st.markdown(
                        f"<div style='max-height: 54vh; overflow-y: auto; overflow-x: hidden; border: 1px solid #E2E8F0; border-radius: 12px; margin-top: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.04);'>"
                        f"{render_compact_table_html(filtered_items)}"
                        f"</div>",
                        unsafe_allow_html=True
                    )

            with tab_res_override:
                with st.container(height=390):
                    c_top_d, c_top_r = st.columns([3.5, 1.5])
                    with c_top_d:
                        st.markdown("""
                        <div style="font-size:0.8rem; color:#64748B; margin-bottom:8px;">
                            💡 <b>Dò tìm & Gợi ý tương đương chuẩn kỹ thuật (3 tầng):</b> Hệ thống xếp theo % trùng khớp giảm dần (Tối đa 3 món gần nhất). Chọn mã và nhấn <b>Áp dụng</b> để gán và ghi nhớ vĩnh viễn.
                        </div>
                        """, unsafe_allow_html=True)
                    with c_top_r:
                        if st.button("🔄 Quét lại kho với thuật toán mới", key="btn_rescan_smart_top", use_container_width=True):
                            for it in res_list:
                                t_r = it.get("ten_trich_xuat", "")
                                q_s = it.get("quy_cach", "")
                                try:
                                    in_p = float(it.get("don_gia", 0) or 0)
                                except Exception:
                                    in_p = 0.0
                                cands_new = get_candidate_matches(f"{t_r} {q_s}", corpus, top_k=5, input_price=in_p)
                                it["candidates"] = cands_new
                                best_new = cands_new[0] if cands_new else None
                                sim_items = []
                                for c in cands_new:
                                    ly_do = f" [{c['ly_do']}]" if c.get("ly_do") else ""
                                    alert_txt = f" {' '.join(c.get('data_alerts', []))}" if c.get("data_alerts") else ""
                                    price_txt = f" | BQ: {c['don_gia_bq']:,.0f}đ" if c.get("don_gia_bq", 0) > 0 else ""
                                    sim_items.append(f"• [{c['score']}%] {c['ma']}: {c['ten']} (Tồn: {c['ton']}{price_txt}{alert_txt}{ly_do})")
                                it["goi_y_tuong_duong"] = "\n".join(sim_items) if sim_items else "Không có mã tương tự"
                                if best_new and (best_new.get("tier") == "EXACT_MATCH" or best_new["score"] >= 75):
                                    it["trang_thai"] = "KHỚP"
                                    it["ma_vat_tu_cu"] = best_new["ma"]
                                    it["ten_trong_kho"] = best_new["ten"]
                                    it["ton_kho"] = best_new["ton"]
                                    it["ma_de_xuat"] = "-"
                                    alert_tag = f" {' '.join(best_new.get('data_alerts', []))}" if best_new.get("data_alerts") else ""
                                    it["ghi_chu"] = f"✓ Khớp mã kho {best_new['ma']} ({best_new['ly_do']}){alert_tag}"
                                    if in_p > 0 and best_new.get("don_gia_bq", 0) > 0:
                                        diff = round(((in_p - best_new["don_gia_bq"]) / best_new["don_gia_bq"]) * 100, 1)
                                        if abs(diff) >= 20:
                                            sign = "+" if diff > 0 else ""
                                            it["canh_bao_gia"] = f"⚠️ Giá lệch {sign}{diff}% (Xuất: {in_p:,.0f}đ vs BQ: {best_new['don_gia_bq']:,.0f}đ)"
                                            it["ghi_chu"] += f" | {it['canh_bao_gia']}"
                                elif best_new and (best_new.get("tier") == "GENERIC_MATCH" or best_new["score"] >= 60):
                                    it["trang_thai"] = "KHỚP"
                                    it["ma_vat_tu_cu"] = best_new["ma"]
                                    it["ten_trong_kho"] = best_new["ten"]
                                    it["ton_kho"] = best_new["ton"]
                                    it["ma_de_xuat"] = "-"
                                    it["ghi_chu"] = f"✓ Khớp danh mục chung mã {best_new['ma']} ({best_new['ly_do']})"
                            st.session_state["ag1_reconciled_results"] = res_list
                            batch_file = save_batch_report_excel(res_list, f"{total_docs}_ChungTu")
                            st.session_state["ag1_reconciled_file_path"] = str(batch_file)
                            st.toast("✓ Đã quét lại toàn bộ danh sách bằng thuật toán mới!", icon="✅")
                            time.sleep(0.3)
                            st.rerun(scope="fragment")

                    for item in filtered_items:
                        stt_num = item.get("stt", 1)
                        t_raw = item.get("ten_trich_xuat", "")
                        q_spec = item.get("quy_cach", "")
                        curr_code = item.get("ma_vat_tu_cu") if item.get("trang_thai") == "KHỚP" else item.get("ma_de_xuat")
                        status_txt = "🟢 Khớp kho" if item.get("trang_thai") == "KHỚP" else "🟡 Cấp mã mới"

                        item_cands = item.get("candidates") or []
                        if not item_cands or any("pct_str" not in c for c in item_cands):
                            item_cands = get_candidate_matches(f"{t_raw} {q_spec}", corpus, top_k=5)
                            item["candidates"] = item_cands
                            sim_items = []
                            for c in item_cands:
                                ly_do = f" [{c['ly_do']}]" if c.get("ly_do") else ""
                                alert_txt = f" {' '.join(c.get('data_alerts', []))}" if c.get("data_alerts") else ""
                                price_txt = f" | BQ: {c['don_gia_bq']:,.0f}đ" if c.get("don_gia_bq", 0) > 0 else ""
                                sim_items.append(f"• [{c['score']}%] {c['ma']}: {c['ten']} (Tồn: {c['ton']}{price_txt}{alert_txt}{ly_do})")
                            item["goi_y_tuong_duong"] = "\n".join(sim_items) if sim_items else "Không có mã tương tự"

                        options_map = {}
                        curr_label = f"Đang gán: {curr_code} ({item.get('ten_chuan_hoa', '')}) [{status_txt}]"
                        options_map[curr_label] = {"type": "current"}

                        for c in item_cands:
                            ly_do = f" — {c['ly_do']}" if c.get("ly_do") else ""
                            alert_txt = f" {' '.join(c.get('data_alerts', []))}" if c.get("data_alerts") else ""
                            price_txt = f" | BQ: {c['don_gia_bq']:,.0f}đ" if c.get("don_gia_bq", 0) > 0 else ""
                            opt_lbl = f"👉 [{c['score']}%] Chọn mã {c['ma']} | {c['ten']} (Tồn: {c['ton']}{price_txt}{alert_txt}{ly_do})"
                            options_map[opt_lbl] = {
                                "type": "candidate",
                                "code": c["ma"],
                                "name": c["ten"],
                                "ton": c["ton"],
                                "ly_do": c.get("ly_do", "")
                            }

                        c_info, c_sel, c_act = st.columns([3, 4.5, 1.3])
                        with c_info:
                            st.markdown(f"**#{stt_num}. {t_raw}**")
                            st.caption(f"Quy cách: `{q_spec}` | SL: {item.get('so_luong')} {item.get('don_vi')}")
                        with c_sel:
                            chosen_lbl = st.selectbox(
                                f"Gợi ý cho #{stt_num}",
                                options=list(options_map.keys()),
                                key=f"sel_cand_{stt_num}_{item.get('chung_tu_goc')}",
                                label_visibility="collapsed"
                            )
                        with c_act:
                            chosen_val = options_map.get(chosen_lbl, {})
                            if chosen_val.get("type") == "candidate":
                                if st.button("Áp dụng 🔁", key=f"btn_apply_{stt_num}_{item.get('chung_tu_goc')}", type="secondary", use_container_width=True):
                                    new_code = chosen_val["code"]
                                    new_name = chosen_val["name"]
                                    new_ton = chosen_val["ton"]
                                    new_ly_do = chosen_val.get("ly_do", "")

                                    item["trang_thai"] = "KHỚP"
                                    item["ma_vat_tu_cu"] = new_code
                                    item["ten_trong_kho"] = new_name
                                    item["ton_kho"] = new_ton
                                    item["ma_de_xuat"] = "-"
                                    item["ghi_chu"] = f"✓ Kỹ sư chọn mã kho {new_code} ({new_ly_do})"

                                    save_approved_material(
                                        code=new_code,
                                        name_normalized=new_name,
                                        spec=q_spec,
                                        unit=item.get("don_vi", "Cái"),
                                        group_code="02" if str(new_code).startswith("02.") else "03",
                                        raw_name=t_raw,
                                        source_doc=item.get("chung_tu_goc", ""),
                                        notes=f"Kỹ sư chọn từ gợi ý tương đương: {new_ly_do}"
                                    )

                                    st.session_state["ag1_reconciled_results"] = res_list
                                    batch_file = save_batch_report_excel(res_list, f"{total_docs}_ChungTu")
                                    st.session_state["ag1_reconciled_file_path"] = str(batch_file)
                                    st.toast(f"✓ Đã gán mã {new_code} ({new_name}) và ghi nhớ vĩnh viễn!", icon="✅")
                                    time.sleep(0.3)
                                    st.rerun(scope="fragment")

        if st.session_state.get("ag1_reconciled_results"):
            render_reconciled_results_panel(corpus)

    # ═══════════════════════════════════════════════════════════════════════════
    # SUB-TAB 2: KHO VẬT TƯ & CẬP NHẬT DANH MỤC (TỐI ƯU CỤC BỘ @st.fragment)
    # ═══════════════════════════════════════════════════════════════════════════
    @st.fragment
    def render_tab_master(df_master: pd.DataFrame):
        # ── HÀNG ĐIỀU KHIỂN & BỘ LỌC ──
        col_search, col_group, col_zero, col_dl = st.columns([3, 2, 1.3, 2.5])
        with col_search:
            kw = st.text_input("🔍 Tìm kiếm danh mục kho:", placeholder="Nhập tên, mã (02.xxx, 03.xxx), quy cách...", key="master_search_input")
        with col_group:
            group_filter = st.selectbox(
                "Phân nhóm:",
                ["Tất cả danh mục", "Nhóm 02 (VLXD / Tiêu hao)", "Nhóm 03 (Cơ - Điện)"],
                index=0,
                key="master_group_select"
            )
        with col_zero:
            zero_filter = st.checkbox("Chỉ tồn = 0", value=False, key="master_zero_check")
        with col_dl:
            if MASTER_FILE.exists():
                with open(MASTER_FILE, "rb") as mf:
                    st.download_button(
                        label="💾 TẢI FILE KHO HIỆN TẠI (.XLSX)",
                        data=mf.read(),
                        file_name=f"Master_Inventory_CFC_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="btn_dl_master_current"
                    )

        df_display = df_master.copy()
        if group_filter == "Nhóm 02 (VLXD / Tiêu hao)":
            df_display = df_display[df_display["Mã hàng"].astype(str).str.startswith("02.")]
        elif group_filter == "Nhóm 03 (Cơ - Điện)":
            df_display = df_display[df_display["Mã hàng"].astype(str).str.startswith("03.")]

        if kw:
            mask = df_display.apply(lambda r: r.astype(str).str.contains(kw, case=False, na=False).any(), axis=1)
            df_display = df_display[mask]
        if zero_filter and "Cuối kỳ SL" in df_display.columns:
            df_display = df_display[df_display["Cuối kỳ SL"] == 0]

        # Định dạng các cột giá trị tiền (GT) theo chuẩn 1.000.000, 10.000.000
        df_view = df_display.copy()
        gt_cols = [c for c in df_view.columns if "GT" in c or "giá" in c.lower() or "tiền" in c.lower()]
        for col in gt_cols:
            df_view[col] = df_view[col].apply(format_vnd_currency)

        column_config = {
            col: st.column_config.TextColumn(label=col, alignment="right")
            for col in gt_cols
        }

        st.dataframe(df_view, use_container_width=True, height=320, column_config=column_config)
        st.caption(f"Đang hiển thị {len(df_display):,} / {len(df_master):,} mã hàng trong kho.")

        # ── KHU VỰC CẬP NHẬT FILE KHO MỚI & TỪ ĐIỂN ÁNH XẠ ──
        col_up_box, col_alias_box = st.columns([1, 1], gap="medium")

        with col_up_box:
            with st.expander("📤 CẬP NHẬT FILE KHO VẬT TƯ MỚI (HÀNG TUẦN / THÁNG)", expanded=False):
                st.markdown("""
                <div style="font-size: 0.8rem; color: #94a3b8; margin-bottom: 8px;">
                    Tải lên bảng tổng hợp kho mới (.xlsx, .xls, .csv). Hệ thống sẽ tự động đối chiếu và cập nhật cơ sở dữ liệu.
                </div>
                """, unsafe_allow_html=True)

                up_inv_file = st.file_uploader(
                    "Chọn file kho mới:",
                    type=["xlsx", "xls", "csv"],
                    key="uploader_master_inventory"
                )
                st.markdown("""
                <div style="background: #ECFDF5; border: 1px solid #A7F3D0; border-radius: 10px; padding: 10px 14px; margin: 10px 0; font-size: 0.82rem; color: #065F46;">
                    🔄 <b>Chế độ cập nhật cố định: Hợp nhất (Merge / Upsert)</b><br>
                    Tự động nhận diện định dạng ERP/Kế toán, cập nhật số lượng tồn mới và bổ sung các mã hàng mới mà <b>không làm mất</b> danh mục và các mã đã học.
                </div>
                """, unsafe_allow_html=True)

                if st.button("🚀 XÁC NHẬN CẬP NHẬT KHO VẬT TƯ", type="primary", use_container_width=True, key="btn_confirm_up_master"):
                    if not up_inv_file:
                        st.error("Vui lòng chọn file Excel hoặc CSV trước khi bấm cập nhật!")
                    else:
                        success, msg = update_master_inventory(up_inv_file.read(), up_inv_file.name, mode="merge")
                        if success:
                            st.success(f"✓ {msg}")
                            st.toast(f"✓ {msg}", icon="✅")
                            time.sleep(0.5)
                            st.rerun(scope="fragment")
                        else:
                            st.error(f"Thất bại: {msg}")

        with col_alias_box:
            aliases = get_all_aliases()
            with st.expander(f"🧠 TỪ ĐIỂN ÁNH XẠ ĐÃ HỌC ({len(aliases)} MỤC ĐÃ NHỚ)", expanded=False):
                st.markdown("""
                <div style="font-size: 0.8rem; color: #94a3b8; margin-bottom: 8px;">
                    Danh sách từ vựng/tên gọi thợ thuyền đã học. Các lần quét sau gặp lại tên này sẽ tự động khớp 100% về mã kho đã lưu.
                </div>
                """, unsafe_allow_html=True)
                if aliases:
                    df_alias = pd.DataFrame(aliases)[["raw_name", "target_code", "target_name", "created_at"]]
                    df_alias.columns = ["Tên chứng từ gốc", "Mã chuẩn kho", "Tên kỹ thuật chuẩn", "Thời gian học"]
                    st.dataframe(df_alias, use_container_width=True, height=220)
                else:
                    st.info("Chưa có từ khóa học nào trong bộ nhớ. Khi duyệt mã mới ở Tab 1, hệ thống sẽ tự động ghi nhớ vào đây.")

    with tab_master:
        render_tab_master(df_master)

    # ═══════════════════════════════════════════════════════════════════════════
    # SUB-TAB 3: BÁO CÁO / PHIẾU ĐÃ XỬ LÝ (TỐI ƯU CỤC BỘ @st.fragment)
    # ═══════════════════════════════════════════════════════════════════════════
    @st.fragment
    def render_tab_reports():
        # Tự động quét và đóng gói các file báo cáo vào cây thư mục Năm / Tháng / Ngày
        organize_reports_into_tree()
        all_reports = sorted(REPORT_DIR.rglob("*.xlsx"), key=os.path.getmtime, reverse=True)

        if all_reports:
            col_rep_head, col_rep_del_all = st.columns([4, 1.5])
            with col_rep_head:
                st.markdown(f"<div style='font-weight: 700; color: #0F172A; font-size: 0.95rem; margin-bottom: 8px;'>📋 KHO LƯU TRỮ BÁO CÁO DẠNG CÂY ({len(all_reports)} PHIẾU ĐÃ ĐÓNG GÓI THEO NĂM / THÁNG / NGÀY):</div>", unsafe_allow_html=True)
            with col_rep_del_all:
                with st.popover("🗑️ Xóa toàn bộ lịch sử", use_container_width=True):
                    st.markdown("**Xác nhận xóa tất cả báo cáo?**")
                    st.caption("Thao tác này sẽ xóa vĩnh viễn tất cả các file báo cáo trong cây lưu trữ.")
                    if st.button("Xác nhận xóa hết", type="primary", key="btn_confirm_del_all", use_container_width=True):
                        del_cnt = 0
                        for r_item in all_reports:
                            try:
                                r_item.unlink()
                                del_cnt += 1
                            except Exception:
                                pass
                        st.toast(f"✓ Đã xóa {del_cnt} báo cáo!", icon="🗑️")
                        time.sleep(0.3)
                        st.rerun(scope="fragment")

            # Nhóm theo Cây phân cấp: Năm -> Tháng -> Ngày
            tree_reports = {}
            for rf in all_reports:
                y, m, d = get_report_file_date(rf)
                if y not in tree_reports:
                    tree_reports[y] = {}
                if m not in tree_reports[y]:
                    tree_reports[y][m] = {}
                if d not in tree_reports[y][m]:
                    tree_reports[y][m][d] = []
                tree_reports[y][m][d].append(rf)

            current_year = get_vn_now().year
            current_month = get_vn_now().month

            for y in sorted(tree_reports.keys(), reverse=True):
                total_y = sum(len(tree_reports[y][m][d]) for m in tree_reports[y] for d in tree_reports[y][m])
                with st.expander(f"📁 **NĂM {y}** — Tổng {total_y} phiếu đã lưu trữ", expanded=(y == current_year)):
                    for m in sorted(tree_reports[y].keys(), reverse=True):
                        total_m = sum(len(tree_reports[y][m][d]) for d in tree_reports[y][m])
                        with st.expander(f"📂 **Tháng {m:02d}/{y}** — {total_m} phiếu", expanded=(y == current_year and m == current_month)):
                            for d in sorted(tree_reports[y][m].keys(), reverse=True):
                                day_files = tree_reports[y][m][d]
                                st.markdown(f"<div style='color:#1E40AF; font-weight:700; font-size:0.86rem; margin:10px 0 6px 0; border-left:3px solid #2563EB; padding-left:8px;'>📅 Ngày {d:02d}/{m:02d}/{y} ({len(day_files)} phiếu):</div>", unsafe_allow_html=True)

                                for idx, rf in enumerate(day_files):
                                    file_key = f"{rf.stem}_{idx}_{y}_{m}_{d}"
                                    time_str = format_vn_time(rf.stat().st_mtime)

                                    with st.expander(f"📄 {rf.name}  |  🕒 {time_str}", expanded=(idx == 0 and y == current_year and m == current_month)):
                                        c_dl, c_del, c_space = st.columns([2, 1.5, 4.5])
                                        with c_dl:
                                            with open(rf, "rb") as f:
                                                st.download_button(
                                                    label="💾 Tải file Excel (.xlsx)",
                                                    data=f.read(),
                                                    file_name=rf.name,
                                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                    key=f"dl_btn_{file_key}",
                                                    use_container_width=True
                                                )
                                        with c_del:
                                            if st.button("🗑️ Xóa phiếu này", key=f"del_btn_{file_key}", type="secondary", use_container_width=True):
                                                try:
                                                    rf.unlink()
                                                    st.toast(f"✓ Đã xóa thành công phiếu '{rf.name}'!", icon="🗑️")
                                                    time.sleep(0.3)
                                                    st.rerun(scope="fragment")
                                                except Exception as e:
                                                    st.error(f"Lỗi khi xóa: {e}")

                                        # Xem trực tiếp nội dung phiếu dạng bảng compact 100% không trượt ngang
                                        try:
                                            df_view = pd.read_excel(rf)
                                            n_rows = len(df_view)
                                            st.markdown(f"<div style='font-size:0.83rem; color:#94a3b8; margin: 8px 0 4px 0;'>📊 <b>Nội dung phiếu chi tiết ({n_rows} dòng vật tư — Gói gọn trong khung hiển thị):</b></div>", unsafe_allow_html=True)
                                            st.markdown(render_report_detail_html(df_view), unsafe_allow_html=True)
                                        except Exception as e:
                                            st.warning(f"Không thể đọc xem trước nội dung: {e}")
        else:
            st.info("Chưa có báo cáo nào được tạo. Khi đối chiếu xong ở Tab 1, các phiếu sẽ tự động đóng gói theo cây Năm / Tháng / Ngày tại đây.")

    with tab_reports:
        render_tab_reports()


# ═══════════════════════════════════════════════════════════════════════════
# ĐIỀU PHỐI THI HÀNH: ĐA PHÂN HỆ ĐỘC LẬP VỚI @st.fragment (ZERO-BLOCKING)
# ═══════════════════════════════════════════════════════════════════════════
default_tab = MODE_AGENT2 if is_agent2_only else MODE_AGENT1
tab_main_ag1, tab_main_ag2 = st.tabs([MODE_AGENT1, MODE_AGENT2], default=default_tab, key="main_hub_parallel_tabs")

with tab_main_ag1:
    render_agent1_ui(df_master, api_key, model_name)

with tab_main_ag2:
    agent2_manager.render_agent2_ui(api_key=api_key, model_name=model_name, is_split=False, nvidia_api_key=nvidia_api_key)
