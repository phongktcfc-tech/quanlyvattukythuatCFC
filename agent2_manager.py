"""
AGENT 2: THEO DÕI VÀ QUẢN TRỊ KHO VẬT TƯ THỰC TẾ TẠI XƯỞNG CƠ - ĐIỆN
Công ty Cổ Phần Phân Bón & Hóa Chất Cần Thơ (CFC)

Tính năng chính:
1. Đồng bộ dữ liệu Live thời gian thực từ Google Sheet của xưởng cơ điện.
2. Theo dõi liên tục Tồn kho thực tế, Định mức tồn tối thiểu, Cảnh báo thiếu hụt.
3. Trợ lý AI tra cứu ngữ nghĩa đa chiều (vật tư, thiết bị máy móc, người nhận).
4. Tự động hóa phát cảnh báo và gửi Email định kỳ (16:00 hàng ngày, Báo cáo tuần, Báo cáo tháng, Báo cáo năm thanh lý vật tư chết).
"""

import os
import sys
import re
import time
import datetime
import smtplib
import unicodedata
import urllib.request
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
import openpyxl
import streamlit as st

# ── THIẾT LẬP THƯ MỤC & ĐƯỜNG DẪN (TỰ ĐỘNG THÍCH ỨNG LOCAL & CLOUD) ──
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "Agent2_Data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

LOCAL_SHEET_FILE = DATA_DIR / "Kho_CoDien_Live.xlsx"
GOOGLE_SHEET_ID = "13cATeOG2G62QXP9F6luZ0V5OHij-b0kPj_r0-Wb97JE"
DEFAULT_RECIPIENT = "linhb1702638@gmail.com"
TARGET_EMAIL = DEFAULT_RECIPIENT

# Nạp thông tin cấu hình Gmail từ file .env dùng chung hoặc st.secrets / env var
ENV_FILE = Path("D:/Agent EmailAssistant/.env")

# ── MÚI GIỜ VIỆT NAM (UTC+7) & ĐỊNH DẠNG THỜI GIAN ──
VN_TZ = datetime.timezone(datetime.timedelta(hours=7))


def get_vn_now() -> datetime.datetime:
    """Trả về thời gian hiện tại chuẩn Việt Nam (UTC+7)."""
    return datetime.datetime.now(VN_TZ)


def format_vn_time(dt: datetime.datetime | None = None) -> str:
    """Định dạng thời gian chuẩn DD/MM/YYYY HH:mm:ss theo giờ Việt Nam."""
    if dt is None:
        dt = get_vn_now()
    elif dt.tzinfo is None:
        dt = dt.replace(tzinfo=datetime.timezone.utc).astimezone(VN_TZ)
    else:
        dt = dt.astimezone(VN_TZ)
    return dt.strftime("%d/%m/%Y %H:%M:%S")


def compute_item_status(row):
    """
    Phân loại trạng thái vật tư theo quy tắc kỹ thuật nghiêm ngặt:
    1. OUT_OF_STOCK ('🔴 HẾT HÀNG'): Tồn Kho Thực Tế <= 0 (hoặc trống/missing coi như 0)
    2. WARNING ('🟡 SẮP HẾT'): Tồn Kho Thực Tế > 0 AND Định Mức Tồn Tối Thiểu > 0 AND Tồn Kho Thực Tế <= Định Mức Tồn Tối Thiểu
    3. SAFE ('🟢 AN TOÀN'): Tồn Kho Thực Tế > Định Mức Tồn Tối Thiểu (hoặc Định Mức Tồn Tối Thiểu trống/0 và Tồn Kho Thực Tế > 0)
    """
    if isinstance(row, (dict, pd.Series)):
        val_str = str(row.get('Tồn Kho Thực Tế', row.get('ton_thuc_te', ''))).replace(',', '').strip()
        min_str = str(row.get('Định Mức Tồn Tối Thiểu', row.get('dinh_muc_min', ''))).replace(',', '').strip()
    else:
        val_str = str(row).replace(',', '').strip()
        min_str = ''

    try:
        stock = float(val_str) if val_str and val_str.lower() not in ['none', 'nan', ''] else 0.0
    except (ValueError, TypeError):
        stock = 0.0

    try:
        min_val = float(min_str) if min_str and min_str.lower() not in ['none', 'nan', ''] else None
    except (ValueError, TypeError):
        min_val = None

    if stock <= 0:
        return 'OUT_OF_STOCK', '🔴 HẾT HÀNG'
    elif min_val is not None and min_val > 0 and stock <= min_val:
        return 'WARNING', '🟡 SẮP HẾT'
    else:
        return 'SAFE', '🟢 AN TOÀN'


def process_inventory_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Chuẩn hóa và tính toán mã trạng thái nội bộ _status_code từ số thực tuyệt đối.
    Ngăn chặn hoàn toàn lỗi lọc chuỗi (e.g. 'Bình thường' vs 'An toàn') gây rỗng bảng.
    """
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    df = df_raw.copy()

    def to_float(val):
        try:
            s = str(val).replace(',', '').strip()
            return float(s) if s and s.lower() not in ['none', 'nan', ''] else 0.0
        except (ValueError, TypeError):
            return 0.0

    def to_float_or_none(val):
        try:
            s = str(val).replace(',', '').strip()
            return float(s) if s and s.lower() not in ['none', 'nan', ''] else None
        except (ValueError, TypeError):
            return None

    ton_col = 'Tồn Kho Thực Tế' if 'Tồn Kho Thực Tế' in df.columns else 'ton_thuc_te'
    min_col = 'Định Mức Tồn Tối Thiểu' if 'Định Mức Tồn Tối Thiểu' in df.columns else 'dinh_muc_min'

    df['ton_num'] = df[ton_col].apply(to_float)
    df['min_num'] = df[min_col].apply(to_float_or_none) if min_col in df.columns else None

    # Strict Business Logic Classification
    def assign_status(row):
        ton = row['ton_num']
        min_ton = row.get('min_num')
        if ton <= 0:
            return 'OUT_OF_STOCK', '🔴 HẾT HÀNG'
        elif min_ton is not None and min_ton > 0 and ton <= min_ton:
            return 'WARNING', '🟡 SẮP HẾT'
        else:
            return 'SAFE', '🟢 AN TOÀN'

    status_data = df.apply(assign_status, axis=1)
    df['_status_code'] = [s[0] for s in status_data]
    df['status_code'] = df['_status_code']
    df['trang_thai'] = [s[1] for s in status_data]
    df['Trạng Thái Kho'] = df['trang_thai']

    # Đồng bộ các trường số chuẩn cho bảng hiển thị
    df['ton_thuc_te'] = df['ton_num']
    df['dinh_muc_min'] = df['min_num']
    return df


def format_status_badge(val: str) -> str:
    """Định dạng nhãn trạng thái trực quan chuẩn xác (Ưu tiên kiểm tra SẮP HẾT trước HẾT HÀNG)."""
    if not val:
        return "🟢 AN TOÀN"
    s = str(val).strip().upper()
    if "SẮP" in s or "SAP" in s or "WARNING" in s or "THIẾU" in s:
        return "🟡 SẮP HẾT"
    if "HẾT" in s or "HET" in s or "CẠN" in s or "OUT_OF_STOCK" in s:
        return "🔴 HẾT HÀNG"
    if "AN TOÀN" in s or "SAFE" in s or "BÌNH THƯỜNG" in s:
        return "🟢 AN TOÀN"
    return "🟢 AN TOÀN"


def get_smtp_credentials():
    """Lấy thông tin tài khoản SMTP an toàn từ st.secrets, .env hoặc cấu hình chuẩn."""
    creds = {
        "user": "linhb1702638@gmail.com",
        "password": "tgnr ejny ffuy bwmx",
        "server": "smtp.gmail.com",
        "port": 587,
        "recipient": TARGET_EMAIL
    }
    # Thử nạp từ st.secrets nếu đang chạy trên Streamlit Cloud
    try:
        import streamlit as st
        if "EMAIL_APP_PASSWORD" in st.secrets:
            creds["password"] = st.secrets["EMAIL_APP_PASSWORD"]
        if "EMAIL_ADDRESS" in st.secrets:
            creds["user"] = st.secrets["EMAIL_ADDRESS"]
    except Exception:
        pass

    if ENV_FILE.exists():
        try:
            with open(ENV_FILE, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        k = k.strip()
                        v = v.strip().strip("\"'")
                        if k == "EMAIL_ADDRESS":
                            creds["user"] = v
                        elif k == "EMAIL_APP_PASSWORD":
                            creds["password"] = v
                        elif k == "SMTP_SERVER":
                            creds["server"] = v
                        elif k == "SMTP_PORT":
                            creds["port"] = int(v)
                        elif k == "RECIPIENT_EMAIL":
                            creds["recipient"] = v
        except Exception as e:
            print(f"[Agent 2] Lỗi đọc .env: {e}")
    return creds


# ── ĐỒNG BỘ GOOGLE SHEET LIVE ──
def sync_live_google_sheet(sheet_id: str = GOOGLE_SHEET_ID, target_path: Path = LOCAL_SHEET_FILE) -> tuple[bool, str]:
    """Tải trực tiếp file Excel từ link Google Sheet thời gian thực."""
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    try:
        req = urllib.request.Request(
            export_url,
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
        )
        with urllib.request.urlopen(req, timeout=30) as response, open(target_path, "wb") as out_file:
            out_file.write(response.read())
        now_str = format_vn_time(get_vn_now())
        return True, f"Đồng bộ thành công lúc {now_str}"
    except Exception as e:
        return False, f"Lỗi đồng bộ Google Sheet: {e}"


def load_data_from_sheets(force_sync: bool = False) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, str]:
    """
    Nạp dữ liệu từ Google Sheet có tự động cache 60 phút (@st.cache_data(ttl=3600)).
    Trả về: (df_tonghop, df_nhap, df_xuat, sync_time_str)
    """
    try:
        import streamlit as st

        @st.cache_data(ttl=3600, show_spinner=False)
        def _cached_sheet_fetcher():
            sync_live_google_sheet()
            df_th, df_nh, df_xt = load_agent2_data(force_sync=False)
            ts = format_vn_time(get_vn_now())
            return df_th, df_nh, df_xt, ts

        if force_sync:
            try:
                st.cache_data.clear()
            except Exception:
                pass
            sync_live_google_sheet()
            df_th, df_nh, df_xt = load_agent2_data(force_sync=True)
            ts = format_vn_time(get_vn_now())
            return df_th, df_nh, df_xt, ts

        return _cached_sheet_fetcher()
    except Exception:
        if force_sync:
            sync_live_google_sheet()
        df_th, df_nh, df_xt = load_agent2_data(force_sync=False)
        ts = format_vn_time(get_vn_now())
        return df_th, df_nh, df_xt, ts


# ── ĐỌC VÀ CHUẨN HÓA DỮ LIỆU AGENT 2 ──
def load_agent2_data(force_sync: bool = False) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Đọc và phân tách dữ liệu xưởng cơ điện:
    - df_tonghop: Danh mục tồn kho thực tế, định mức, cảnh báo
    - df_nhap: Lịch sử nhập kho từ các sheet tháng
    - df_xuat: Lịch sử xuất kho kèm thiết bị, người nhận từ các sheet tháng
    """
    if force_sync or not LOCAL_SHEET_FILE.exists():
        sync_live_google_sheet()

    if not LOCAL_SHEET_FILE.exists():
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    try:
        wb = openpyxl.load_workbook(LOCAL_SHEET_FILE, data_only=True)
    except Exception as e:
        print(f"[Agent 2] Lỗi mở file Excel: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # 1. Đọc sheet 'Tổng Hợp Vật Tư'
    tonghop_data = []
    if "Tổng Hợp Vật Tư" in wb.sheetnames:
        ws = wb["Tổng Hợp Vật Tư"]
        current_group = "Tổng hợp"
        for r in range(2, ws.max_row + 1):
            col1 = ws.cell(r, 1).value
            col2 = ws.cell(r, 2).value

            # Phát hiện tiêu đề nhóm (ví dụ: Điện, Cơ khí...)
            if col1 and not col2 and isinstance(col1, str) and len(col1.strip()) > 1:
                current_group = col1.strip()
                continue

            if col2 and str(col2).strip():
                name_clean = str(col2).strip()
                dvt = str(ws.cell(r, 4).value or "").strip()
                ton_dau = ws.cell(r, 5).value or 0
                nhap = ws.cell(r, 6).value or 0
                xuat = ws.cell(r, 7).value or 0
                ton_thuc_te = ws.cell(r, 8).value
                if ton_thuc_te is None:
                    try:
                        ton_thuc_te = float(ton_dau) + float(nhap) - float(xuat)
                    except:
                        ton_thuc_te = 0
                else:
                    try:
                        ton_thuc_te = float(ton_thuc_te)
                    except:
                        ton_thuc_te = 0

                dinh_muc = ws.cell(r, 9).value
                try:
                    dinh_muc_val = float(dinh_muc) if dinh_muc is not None and str(dinh_muc).strip() != "" else None
                except:
                    dinh_muc_val = None

                # Đánh giá cảnh báo chuẩn xác qua compute_item_status
                status_code, status_badge = compute_item_status({
                    'Tồn Kho Thực Tế': ton_thuc_te,
                    'Định Mức Tồn Tối Thiểu': dinh_muc_val
                })

                tonghop_data.append({
                    "stt": len(tonghop_data) + 1,
                    "nhom": current_group,
                    "ten_vat_tu": name_clean,
                    "quy_cach": str(ws.cell(r, 3).value or "").strip(),
                    "dvt": dvt,
                    "ton_dau_ky": ton_dau,
                    "tong_nhap": nhap,
                    "tong_xuat": xuat,
                    "ton_thuc_te": ton_thuc_te,
                    "dinh_muc_min": dinh_muc_val,
                    "status_code": status_code,
                    "trang_thai": status_badge,
                    "ghi_chu": str(ws.cell(r, 11).value or "").strip()
                })

    df_tonghop = process_inventory_dataframe(pd.DataFrame(tonghop_data))

    # 2. Đọc các sheet tháng phát sinh
    nhap_data = []
    xuat_data = []
    for sheetname in wb.sheetnames:
        if "tháng" in sheetname.lower():
            ws = wb[sheetname]
            for r in range(2, ws.max_row + 1):
                # Nhập kho (Cột 1 đến 6)
                d_nhap = ws.cell(r, 1).value
                i_nhap = ws.cell(r, 2).value
                if i_nhap and str(i_nhap).strip():
                    try:
                        qty_nhap = float(ws.cell(r, 4).value or 0)
                    except:
                        qty_nhap = 0
                    nhap_data.append({
                        "thang": sheetname.strip(),
                        "ngay": d_nhap,
                        "ten_vat_tu": str(i_nhap).strip(),
                        "dvt": str(ws.cell(r, 3).value or "").strip(),
                        "so_luong": qty_nhap,
                        "nguon_nhap": str(ws.cell(r, 5).value or "").strip(),
                        "ghi_chu": str(ws.cell(r, 6).value or "").strip()
                    })

                # Xuất kho (Cột 8 đến 13)
                d_xuat = ws.cell(r, 8).value
                i_xuat = ws.cell(r, 9).value
                if i_xuat and str(i_xuat).strip():
                    try:
                        qty_xuat = float(ws.cell(r, 11).value or 0)
                    except:
                        qty_xuat = 0
                    xuat_data.append({
                        "thang": sheetname.strip(),
                        "ngay": d_xuat,
                        "ten_vat_tu": str(i_xuat).strip(),
                        "dvt": str(ws.cell(r, 10).value or "").strip(),
                        "so_luong": qty_xuat,
                        "thiet_bi_khu_vuc": str(ws.cell(r, 12).value or "").strip(),
                        "nguoi_nhan": str(ws.cell(r, 13).value or "").strip()
                    })

    df_nhap = pd.DataFrame(nhap_data)
    df_xuat = pd.DataFrame(xuat_data)

    return df_tonghop, df_nhap, df_xuat


# ── PHÂN TÍCH THỐNG KÊ KHO XƯỞNG ──
def get_inventory_kpis(df_tonghop: pd.DataFrame, df_xuat: pd.DataFrame):
    """Tính toán các chỉ số KPI cho giao diện và báo cáo."""
    if df_tonghop.empty:
        return {
            "total_items": 0,
            "safe_items": 0,
            "warning_items": 0,
            "out_of_stock_items": 0,
            "total_issued_qty": 0
        }

    if "_status_code" not in df_tonghop.columns:
        df_tonghop = process_inventory_dataframe(df_tonghop)

    total_items = len(df_tonghop)
    out_of_stock = len(df_tonghop[df_tonghop["_status_code"] == "OUT_OF_STOCK"])
    warning = len(df_tonghop[df_tonghop["_status_code"] == "WARNING"])
    safe = len(df_tonghop[df_tonghop["_status_code"] == "SAFE"])
    total_issues = len(df_xuat) if not df_xuat.empty else 0
    total_issued = df_xuat["so_luong"].sum() if not df_xuat.empty else 0

    return {
        "total_items": total_items,
        "safe": safe,
        "safe_items": safe,
        "warning": warning,
        "warning_items": warning,
        "out_of_stock": out_of_stock,
        "out_of_stock_items": out_of_stock,
        "total_issues": total_issues,
        "total_issued_qty": total_issued
    }


def get_top_consumed_items(df_xuat: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """Lấy danh sách vật tư tiêu hao nhiều nhất phục vụ kế hoạch mua sắm."""
    if df_xuat.empty:
        return pd.DataFrame()
    top_df = df_xuat.groupby(["ten_vat_tu", "dvt"])["so_luong"].sum().reset_index()
    top_df = top_df.sort_values(by="so_luong", ascending=False).head(top_n)
    return top_df


def get_top_equipment_consumers(df_xuat: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """Lấy danh sách thiết bị / khu vực tiêu hao vật tư nhiều nhất."""
    if df_xuat.empty:
        return pd.DataFrame()
    top_eq = df_xuat[df_xuat["thiet_bi_khu_vuc"] != ""].groupby("thiet_bi_khu_vuc")["so_luong"].sum().reset_index()
    top_eq = top_eq.sort_values(by="so_luong", ascending=False).head(top_n)
    return top_eq


def get_dead_stock_items(df_tonghop: pd.DataFrame, df_xuat: pd.DataFrame) -> pd.DataFrame:
    """Truy vết các vật tư tồn kho không phát sinh xuất dùng (đề xuất thanh lý)."""
    if df_tonghop.empty:
        return pd.DataFrame()
    issued_names = set(df_xuat["ten_vat_tu"].dropna().unique()) if not df_xuat.empty else set()
    dead_df = df_tonghop[
        (df_tonghop["ton_thuc_te"] > 0) &
        (df_tonghop["tong_xuat"] == 0) &
        (~df_tonghop["ten_vat_tu"].isin(issued_names))
    ].copy()
    return dead_df.sort_values(by="ton_thuc_te", ascending=False)


# ── TRỢ LÝ AI TÌM KIẾM ĐA CHIỀU ──
def search_agent2_ai(query: str, df_tonghop: pd.DataFrame, df_nhap: pd.DataFrame, df_xuat: pd.DataFrame) -> dict:
    """Tìm kiếm thông minh ngữ nghĩa vật tư, thiết bị, người nhận."""
    if not query or df_tonghop.empty:
        return {"type": "empty", "message": "Vui lòng nhập từ khóa tra cứu."}

    q = unicodedata.normalize("NFC", str(query)).lower().strip()

    # 1. Tìm theo cảnh báo / hết hàng
    if any(k in q for k in ["sắp hết", "hết hàng", "thiếu", "cần mua", "báo động", "định mức"]):
        if "status_code" in df_tonghop.columns:
            alerts = df_tonghop[df_tonghop["status_code"].isin(["OUT_OF_STOCK", "WARNING"])].copy()
        else:
            alerts = df_tonghop[df_tonghop["trang_thai"].str.contains("Hết|Cạn|Sắp", case=False, na=False)].copy()
        return {
            "type": "alerts",
            "title": f"Danh sách {len(alerts)} vật tư đang chạm ngưỡng cảnh báo / hết hàng",
            "data": alerts
        }

    # 2. Tìm theo thiết bị / khu vực máy (ví dụ Bơm DD F4, Sàng F3, Vis PG F5...)
    matches_tb = df_xuat[df_xuat["thiet_bi_khu_vuc"].str.lower().str.contains(q, na=False)]
    if not matches_tb.empty:
        return {
            "type": "equipment",
            "title": f"Lịch sử xuất vật tư cho thiết bị/khu vực: '{query}' ({len(matches_tb)} lượt)",
            "data": matches_tb
        }

    # 3. Tìm theo người nhận (ví dụ anh Hiệp, anh Năm, anh Thành...)
    matches_person = df_xuat[df_xuat["nguoi_nhan"].str.lower().str.contains(q, na=False)]
    if not matches_person.empty:
        return {
            "type": "person",
            "title": f"Lịch sử nhận vật tư của: '{query}' ({len(matches_person)} lượt)",
            "data": matches_person
        }

    # 4. Tìm theo tên vật tư trong Tổng hợp
    matches_item = df_tonghop[df_tonghop["ten_vat_tu"].str.lower().str.contains(q, na=False)]
    if not matches_item.empty:
        item_names = matches_item["ten_vat_tu"].tolist()
        recent_xuat = df_xuat[df_xuat["ten_vat_tu"].isin(item_names)]
        return {
            "type": "items",
            "title": f"Kết quả tra cứu vật tư: '{query}' ({len(matches_item)} mã)",
            "data": matches_item,
            "recent_xuat": recent_xuat
        }

    return {
        "type": "not_found",
        "message": f"Không tìm thấy dữ liệu phù hợp với từ khóa '{query}'."
    }


# ── SOẠN THẢO VÀ GỬI EMAIL BÁO CÁO TỰ ĐỘNG ──
def generate_email_html(report_type: str, df_tonghop: pd.DataFrame, df_nhap: pd.DataFrame, df_xuat: pd.DataFrame) -> tuple[str, str]:
    """Tạo mẫu email HTML chuyên nghiệp theo chuẩn Technical Dashboard."""
    now_str = datetime.datetime.now().strftime("%H:%M ngày %d/%m/%Y")
    kpis = get_inventory_kpis(df_tonghop, df_xuat)

    type_titles = {
        "daily": "BÁO CÁO TỔNG HỢP CUỐI NGÀY & CẢNH BÁO TỒN KHO XƯỞNG CƠ - ĐIỆN",
        "weekly": "BÁO CÁO TUẦN: PHÂN TÍCH TIÊU HAO VẬT TƯ & KẾ HOẠCH MUA SẮM",
        "monthly": "BÁO CÁO THÁNG: CÂN ĐỐI NHẬP - XUẤT - TỒN & DỰ BÁO NHU CẦU",
        "yearly": "BÁO CÁO NĂM: TRUY VẾT VẬT TƯ TỒN KHO LÂU NGÀY / ĐỀ XUẤT THANH LÝ"
    }
    subject = f"[CFC - XƯỞNG CƠ ĐIỆN] {type_titles.get(report_type, 'BÁO CÁO KHO THỰC TẾ')} ({datetime.datetime.now().strftime('%d/%m/%Y')})"

    # Lấy danh sách cần cảnh báo
    if "status_code" in df_tonghop.columns:
        alert_df = df_tonghop[df_tonghop["status_code"].isin(["OUT_OF_STOCK", "WARNING"])].head(25)
    else:
        alert_df = df_tonghop[df_tonghop["trang_thai"].str.contains("Hết|Cạn|Sắp|WARNING|OUT_OF_STOCK", case=False, na=False)].head(25)
    top_consumed = get_top_consumed_items(df_xuat, top_n=10)
    dead_stock = get_dead_stock_items(df_tonghop, df_xuat).head(15)

    alert_rows_html = ""
    if not alert_df.empty:
        for _, r in alert_df.iterrows():
            badge_color = "#EF4444" if (r.get("status_code") == "OUT_OF_STOCK" or "HẾT" in str(r['trang_thai']).upper()) else "#F59E0B"
            alert_rows_html += f"""
            <tr style="border-bottom: 1px solid #E2E8F0;">
                <td style="padding: 10px; font-weight: 600; color: #1E293B;">{r['ten_vat_tu']}</td>
                <td style="padding: 10px; color: #64748B;">{r['quy_cach'] or '-'}</td>
                <td style="padding: 10px; text-align: center;">{r['dvt']}</td>
                <td style="padding: 10px; text-align: right; font-weight: 700; color: #0F172A;">{r['ton_thuc_te']:.1f}</td>
                <td style="padding: 10px; text-align: right; color: #64748B;">{r['dinh_muc_min'] if r['dinh_muc_min'] is not None else '-'}</td>
                <td style="padding: 10px; text-align: center;">
                    <span style="background-color: {badge_color}; color: #FFFFFF; padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: 700;">
                        {r['trang_thai']}
                    </span>
                </td>
            </tr>
            """
    else:
        alert_rows_html = "<tr><td colspan='6' style='padding: 15px; text-align: center; color: #10B981; font-weight: 600;'>✓ Tất cả vật tư đều ở ngưỡng tồn kho an toàn!</td></tr>"

    top_consumed_rows_html = ""
    if not top_consumed.empty:
        for idx, r in top_consumed.iterrows():
            top_consumed_rows_html += f"""
            <tr style="border-bottom: 1px solid #E2E8F0;">
                <td style="padding: 8px 10px; color: #1E293B; font-weight: 600;">#{idx+1}. {r['ten_vat_tu']}</td>
                <td style="padding: 8px 10px; text-align: center;">{r['dvt']}</td>
                <td style="padding: 8px 10px; text-align: right; font-weight: 700; color: #2563EB;">{r['so_luong']:.1f}</td>
            </tr>
            """

    dead_stock_rows_html = ""
    if not dead_stock.empty:
        for _, r in dead_stock.iterrows():
            dead_stock_rows_html += f"""
            <tr style="border-bottom: 1px solid #E2E8F0;">
                <td style="padding: 8px 10px; color: #1E293B; font-weight: 600;">{r['ten_vat_tu']}</td>
                <td style="padding: 8px 10px; text-align: center;">{r['dvt']}</td>
                <td style="padding: 8px 10px; text-align: right; font-weight: 700; color: #475569;">{r['ton_thuc_te']:.1f}</td>
                <td style="padding: 8px 10px; color: #64748B; font-size: 12px;">Chưa phát sinh xuất dùng</td>
            </tr>
            """

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; background-color: #F8FAFC; margin: 0; padding: 20px; color: #1E293B; }}
            .container {{ max-width: 800px; margin: 0 auto; background: #FFFFFF; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); border: 1px solid #E2E8F0; }}
            .header {{ background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%); color: #FFFFFF; padding: 24px; text-align: left; }}
            .header h1 {{ margin: 0 0 6px 0; font-size: 20px; font-weight: 800; letter-spacing: 0.5px; }}
            .header p {{ margin: 0; font-size: 13px; opacity: 0.9; }}
            .kpi-row {{ display: table; width: 100%; table-layout: fixed; background: #F1F5F9; border-bottom: 1px solid #E2E8F0; }}
            .kpi-cell {{ display: table-cell; padding: 16px 12px; text-align: center; border-right: 1px solid #E2E8F0; }}
            .kpi-val {{ font-size: 22px; font-weight: 800; }}
            .kpi-lbl {{ font-size: 11px; text-transform: uppercase; color: #64748B; font-weight: 600; margin-top: 4px; }}
            .section {{ padding: 20px; }}
            .section-title {{ font-size: 15px; font-weight: 700; margin-bottom: 12px; color: #0F172A; display: flex; align-items: center; }}
            table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
            th {{ background-color: #F8FAFC; color: #475569; font-weight: 700; padding: 10px; text-align: left; border-bottom: 2px solid #CBD5E1; }}
            .footer {{ background: #F8FAFC; padding: 16px; text-align: center; font-size: 12px; color: #94A3B8; border-top: 1px solid #E2E8F0; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>{type_titles.get(report_type, 'BÁO CÁO KHO THỰC TẾ')}</h1>
                <p>CÔNG TY CP PHÂN BÓN & HÓA CHẤT CẦN THƠ — PHÂN XƯỞNG CƠ ĐIỆN | {now_str}</p>
            </div>

            <div class="kpi-row">
                <div class="kpi-cell">
                    <div class="kpi-val" style="color: #2563EB;">{kpis['total_items']:,}</div>
                    <div class="kpi-lbl">Tổng mã vật tư</div>
                </div>
                <div class="kpi-cell">
                    <div class="kpi-val" style="color: #10B981;">{kpis['safe_items']:,}</div>
                    <div class="kpi-lbl">Tồn an toàn</div>
                </div>
                <div class="kpi-cell">
                    <div class="kpi-val" style="color: #F59E0B;">{kpis['warning_items']:,}</div>
                    <div class="kpi-lbl">Sắp hết hàng</div>
                </div>
                <div class="kpi-cell">
                    <div class="kpi-val" style="color: #EF4444;">{kpis['out_of_stock_items']:,}</div>
                    <div class="kpi-lbl">Đang hết hàng</div>
                </div>
            </div>

            <!-- PHẦN 1: CẢNH BÁO VẬT TƯ CẦN ĐẶT HÀNG NGAY -->
            <div class="section" style="border-bottom: 1px solid #E2E8F0;">
                <div class="section-title" style="color: #DC2626;">
                    🚨 DANH SÁCH VẬT TƯ CHẠM NGƯỠNG CẢNH BÁO (CẦN MUA BỔ SUNG)
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>Tên Vật Tư</th>
                            <th>Quy Cách</th>
                            <th style="text-align: center;">ĐVT</th>
                            <th style="text-align: right;">Tồn Thực Tế</th>
                            <th style="text-align: right;">Định Mức Min</th>
                            <th style="text-align: center;">Trạng Thái</th>
                        </tr>
                    </thead>
                    <tbody>
                        {alert_rows_html}
                    </tbody>
                </table>
            </div>

            <!-- PHẦN 2: TOP TIÊU HAO THEO DÕI KẾ HOẠCH MUA SẮM -->
            <div class="section" style="border-bottom: 1px solid #E2E8F0;">
                <div class="section-title" style="color: #2563EB;">
                    📉 TOP VẬT TƯ TIÊU HAO NHIỀU NHẤT TRONG KỲ
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>Tên Vật Tư Tiêu Hao</th>
                            <th style="text-align: center;">ĐVT</th>
                            <th style="text-align: right;">Tổng Số Lượng Xuất</th>
                        </tr>
                    </thead>
                    <tbody>
                        {top_consumed_rows_html}
                    </tbody>
                </table>
            </div>

            <!-- PHẦN 3: VẬT TƯ TỒN KHO LÂU NGÀY / ĐỀ XUẤT THANH LÝ (CHO BÁO CÁO NĂM/THÁNG) -->
            <div class="section">
                <div class="section-title" style="color: #475569;">
                    📦 VẬT TƯ TỒN KHO LÂU NGÀY CHƯA SỬ DỤNG (XEM XÉT THANH LÝ)
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>Tên Vật Tư</th>
                            <th style="text-align: center;">ĐVT</th>
                            <th style="text-align: right;">Số Lượng Tồn</th>
                            <th>Tình Trạng</th>
                        </tr>
                    </thead>
                    <tbody>
                        {dead_stock_rows_html}
                    </tbody>
                </table>
            </div>

            <div class="footer">
                Hệ thống Agent 2 AI tự động đồng bộ từ Google Sheets xưởng Cơ - Điện CFC.<br>
                Email nhận báo cáo: {DEFAULT_RECIPIENT} | Phát tự động vào lúc 16:00 hàng ngày.
            </div>
        </div>
    </body>
    </html>
    """
    return subject, html


def send_email_report(report_type: str = "daily", recipient: str = None) -> tuple[bool, str]:
    """Thực hiện gửi email báo cáo thực tế qua Gmail SMTP."""
    creds = get_smtp_credentials()
    target_email = recipient or creds.get("recipient") or DEFAULT_RECIPIENT

    if not creds.get("password"):
        return False, "Chưa tìm thấy mật khẩu ứng dụng Gmail (EMAIL_APP_PASSWORD trong .env)."

    print(f"[Agent 2] Đang tải dữ liệu Google Sheet...")
    df_tonghop, df_nhap, df_xuat = load_agent2_data(force_sync=True)
    if df_tonghop.empty:
        return False, "Không có dữ liệu tồn kho để tạo báo cáo."

    subject, html_content = generate_email_html(report_type, df_tonghop, df_nhap, df_xuat)

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"CFC Cơ Điện <{creds['user']}>"
    msg["To"] = target_email

    msg.attach(MIMEText(html_content, "html", "utf-8"))

    try:
        print(f"[Agent 2] Đang kết nối Gmail SMTP ({creds['server']}:{creds['port']})...")
        server = smtplib.SMTP(creds["server"], creds["port"], timeout=20)
        server.starttls()
        server.login(creds["user"], creds["password"])
        server.sendmail(creds["user"], target_email, msg.as_string())
        server.quit()
        success_msg = f"Đã gửi thành công {report_type.upper()} báo cáo tới '{target_email}' lúc {datetime.datetime.now().strftime('%H:%M:%S')}."
        print(f"[Agent 2] ✓ {success_msg}")
        return True, success_msg
    except Exception as e:
        err_msg = f"Lỗi khi gửi email qua Gmail SMTP: {e}"
        print(f"[Agent 2] ✗ {err_msg}")
        return False, err_msg


# ── ĐIỂM CHẠY TỰ ĐỘNG HÓA CLI ──
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Agent 2 - Quản trị kho vật tư xưởng cơ điện CFC")
    parser.add_argument("--sync", action="store_true", help="Đồng bộ ngay Google Sheet về máy")
    parser.add_argument("--daily", action="store_true", help="Gửi báo cáo cuối ngày (16:00)")
    parser.add_argument("--weekly", action="store_true", help="Gửi báo cáo tuần")
    parser.add_argument("--monthly", action="store_true", help="Gửi báo cáo tháng")
    parser.add_argument("--yearly", action="store_true", help="Gửi báo cáo năm thanh lý")
    parser.add_argument("--test-mail", action="store_true", help="Gửi email thử nghiệm ngay")
    args = parser.parse_args()

    if args.sync:
        ok, msg = sync_live_google_sheet()
        print(msg)
    elif args.daily:
        ok, msg = send_email_report("daily")
        print(msg)
    elif args.weekly:
        ok, msg = send_email_report("weekly")
        print(msg)
    elif args.monthly:
        ok, msg = send_email_report("monthly")
        print(msg)
    elif args.yearly:
        ok, msg = send_email_report("yearly")
        print(msg)
    elif args.test_mail:
        ok, msg = send_email_report("daily")
        print(msg)
    else:
        print("Chạy thử nghiệm Agent 2...")
        df_tonghop, df_nhap, df_xuat = load_agent2_data(force_sync=True)
        print(f"Tổng hợp: {len(df_tonghop)} dòng | Nhập: {len(df_nhap)} dòng | Xuất: {len(df_xuat)} dòng.")


# ═══════════════════════════════════════════════════════════════════════════
# KỸ SƯ TRƯỞNG AI CƠ - ĐIỆN CFC (GEMINI 4 ARGON / FLASH ENGINE)
# ═══════════════════════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════════════════════
# BƯỚC 1: BỘ PHÂN TÍCH YÊU CẦU & CHẶN LỖI SAI BẢN CHẤT KỸ THUẬT (AGENT 2)
# ═══════════════════════════════════════════════════════════════════════════

def classify_engineering_intent(user_prompt: str) -> dict:
    """
    Phân tích yêu cầu của kỹ sư, tự động sửa lỗi gõ dấu thực tế (sàng thao tác -> sàn thao tác)
    và phân loại hạng mục kỹ thuật để kích hoạt Blacklist / Whitelist kho xưởng.
    """
    raw_text = str(user_prompt).strip()
    norm_text = unicodedata.normalize("NFC", raw_text).lower()

    # 1. Nhận diện lỗi gõ dấu phổ biến tại xưởng:
    # "sàng thao tác", "sàng công tác", "sàng đứng", "sàng bảo trì" -> SÀN THAO TÁC TĨNH
    is_typo_sang = bool(re.search(r'\bsàng\s+(thao\s*tác|công\s*tác|đứng|làm\s*việc|bảo\s*trì|người\s*đứng)\b', norm_text))
    has_dimensions = bool(re.search(r'\d+\s*(?:m|mm|cm)?\s*[x\*×]\s*\d+', norm_text))

    if is_typo_sang or (("sàng" in norm_text) and has_dimensions and any(k in norm_text for k in ["cao", "rộng", "dài", "khung", "chân", "lan can", "cầu thang"])):
        cleaned_prompt = re.sub(r'\bsàng\s+thao\s*tác\b', 'sàn thao tác', raw_text, flags=re.IGNORECASE)
        cleaned_prompt = re.sub(r'\bsàng\s+công\s*tác\b', 'sàn công tác', cleaned_prompt, flags=re.IGNORECASE)
        cleaned_prompt = re.sub(r'\bsàng\s+đứng\b', 'sàn đứng', cleaned_prompt, flags=re.IGNORECASE)
        category = "KET_CAU_TINH"
        note = "Đã nhận diện & chuẩn hóa: 'Sàn thao tác' (Kết cấu khung sàn thép tĩnh cho người đứng làm việc, KHÔNG PHẢI máy sàng rung)."
        return {
            "category": category,
            "cleaned_prompt": cleaned_prompt,
            "engineering_note": note,
            "has_wheels_requested": any(w in norm_text for w in ["di động", "bánh xe", "bánh đẩy", "xe đẩy"])
        }

    # 2. Phân loại theo hạng mục kỹ thuật
    keywords_static = [
        "sàn", "cầu thang", "lan can", "khung", "bệ đỡ", "giá đỡ", "kệ", "mái che",
        "vách", "bồn chứa", "tuyến ống", "đường ống", "bản mã", "bệ máy tĩnh", "chân đế"
    ]
    if any(k in norm_text for k in keywords_static) and not any(k in norm_text for k in ["băng tải", "vít tải", "gầu tải", "máy nghiền", "máy sàng rung", "bơm"]):
        return {
            "category": "KET_CAU_TINH",
            "cleaned_prompt": raw_text,
            "engineering_note": "Hạng mục: Kết cấu tĩnh & Gia công cơ khí tĩnh (Khung thép, sàn, lan can, bệ đỡ).",
            "has_wheels_requested": any(w in norm_text for w in ["di động", "bánh xe", "bánh đẩy", "xe đẩy"])
        }

    keywords_elec = [
        "tủ điện", "chiếu sáng", "đèn", "cảm biến", "van điện từ", "ben hơi", "xy lanh khí nén",
        "dây điện", "cáp điện", "aptomat", "mcb", "mccb", "contactor", "khởi động từ", "rơle", "inverter", "biến tần"
    ]
    if any(k in norm_text for k in keywords_elec) and not any(k in norm_text for k in ["khung", "sàn", "kết cấu", "thép hình"]):
        return {
            "category": "DIEN_TU_DONG_HOA",
            "cleaned_prompt": raw_text,
            "engineering_note": "Hạng mục: Lắp đặt & Sửa chữa Điện - Khí nén - Tự động hóa.",
            "has_wheels_requested": False
        }

    return {
        "category": "THIET_BI_DONG",
        "cleaned_prompt": raw_text,
        "engineering_note": "Hạng mục: Chế tạo & Sửa chữa Thiết bị cơ điện động (Máy móc có chuyển động).",
        "has_wheels_requested": any(w in norm_text for w in ["di động", "bánh xe", "bánh đẩy", "xe đẩy"])
    }


# ═══════════════════════════════════════════════════════════════════════════
# BƯỚC 2: TRUYỀN DỮ LIỆU TỪ LIVE SHEET VÀO PROMPT (DYNAMIC SHEET GROUNDING)
# ═══════════════════════════════════════════════════════════════════════════

STATIC_BLACKLIST_KEYWORDS = [
    "động cơ", "motor", "mô tơ", "hộp giảm tốc", "hgt", "biến tần", "inverter",
    "dây curoa", "curoa", "puly", "buly", "nhông", "xích", "sên",
    "vòng bi", "bạc đạn", "gối đỡ", "lưới sàng"
]

def build_agent2_dynamic_inventory_table(df_th: pd.DataFrame, intent_info: dict) -> str:
    """
    Chuyển đổi toàn bộ DataFrame hiện tại của Agent 2 (Live Sheet ~475-488 mã)
    thành bảng văn bản Markdown cấu trúc chuẩn, ƯU TIÊN SẮP XẾP VẬT TƯ CÒN TỒN LÊN ĐẦU
    để Kỹ sư Trưởng AI bắt buộc tận dụng vật tư có sẵn trong xưởng trước tiên.
    """
    if df_th is None or df_th.empty:
        return "Hiện chưa có dữ liệu tồn kho xưởng từ Live Sheet."

    category = intent_info.get("category", "THIET_BI_DONG")
    has_wheels = intent_info.get("has_wheels_requested", False)

    # Lọc và phân loại vật tư
    available_items = []
    depleted_items = []

    for idx, r in df_th.iterrows():
        stt_val = r.get("stt")
        try:
            stt_int = int(stt_val)
            code = f"VT-{stt_int:03d}"
        except Exception:
            code = f"VT-{idx+1:03d}"

        name = str(r.get("ten_vat_tu", "")).strip()
        if not name or name == "nan":
            continue

        spec = str(r.get("quy_cach", "")).strip()
        spec_clean = spec if spec and spec != "nan" else "-"
        dvt = str(r.get("dvt", "")).strip() or "Cái"
        
        try:
            stock = float(r.get("ton_thuc_te", 0) or 0)
        except Exception:
            stock = 0.0

        status_str = str(r.get("trang_thai", "")).strip()
        if not status_str or status_str == "nan":
            if stock <= 0:
                status_str = "Đã cạn (Tồn = 0)"
            else:
                status_str = "An toàn"
        else:
            if "CẠN" in status_str.upper() or "HẾT" in status_str.upper() or stock <= 0:
                status_str = "Đã cạn (Tồn = 0)"
            elif "SẮP" in status_str.upper() or "WARNING" in status_str.upper():
                status_str = "Sắp hết"
            else:
                status_str = "An toàn"

        # Kiểm tra Blacklist nếu là Kết cấu tĩnh
        if category == "KET_CAU_TINH":
            if str(r.get("nhom", "")).strip().lower() == "điện":
                continue
            combined_name = f"{name} {spec_clean}".lower()
            if any(blk in combined_name for blk in STATIC_BLACKLIST_KEYWORDS):
                continue
            if not has_wheels and "bánh xe" in combined_name:
                continue

        stock_display = f"{stock:g}"
        row_str = f"| {name} | {spec_clean} | {dvt} | {stock_display} | {status_str} |"
        if stock > 0:
            available_items.append((stock, row_str))
        else:
            depleted_items.append(row_str)

    # Sắp xếp vật tư còn tồn theo số lượng tồn giảm dần để AI thấy nguồn lực dồi dào nhất
    available_items.sort(key=lambda x: x[0], reverse=True)

    lines = [
        "### [PHẦN 1: DANH MỤC VẬT TƯ ĐANG CÒN TỒN KHO THỰC TẾ XƯỞNG (TỒN > 0) — BẮT BUỘC TẬN DỤNG TRƯỚC TIÊN]",
        "| Tên vật tư trong kho xưởng | Quy cách kỹ thuật | ĐVT | Tồn thực tế | Trạng thái kho |",
        "|:---|:---|:---:|:---:|:---:|"
    ]
    for _, row in available_items:
        lines.append(row)

    if depleted_items:
        lines.append("")
        lines.append("### [PHẦN 2: DANH MỤC VẬT TƯ ĐÃ HẾT TỒN KHO HOẶC TỒN = 0 (CHỈ XEM THAM KHẢO)]")
        lines.append("| Tên vật tư trong kho xưởng | Quy cách kỹ thuật | ĐVT | Tồn thực tế | Trạng thái kho |")
        lines.append("|:---|:---|:---:|:---:|:---:|")
        for row in depleted_items:
            lines.append(row)

    return "\n".join(lines)


# ═══════════════════════════════════════════════════════════════════════════
# BƯỚC 4: ĐỒNG BỘ KIỂM TRA CHÉO DỮ LIỆU TỒN KHO THỰC TẾ (DATA VERIFICATION)
# ═══════════════════════════════════════════════════════════════════════════

def extract_verified_items_from_response(ai_text: str, df_th: pd.DataFrame) -> pd.DataFrame:
    """
    Trích xuất danh sách các vật tư được Kỹ sư Trưởng AI đề xuất
    để đối chiếu trực tiếp với DataFrame Live Sheet của Agent 2.
    Định danh hoàn toàn theo TÊN VẬT TƯ và QUY CÁCH KỸ THUẬT (Không dùng Mã VT).
    """
    if not ai_text or df_th is None or df_th.empty:
        return pd.DataFrame()

    ai_text_lower = ai_text.lower()
    matched_indices = []

    # Khớp theo Tên vật tư + Quy cách thực tế trong Live Sheet
    for idx, r in df_th.iterrows():
        name = str(r.get("ten_vat_tu", "")).strip().lower()
        if len(name) < 3 or name in ["nan", "none"]:
            continue
        spec = str(r.get("quy_cach", "")).strip().lower()
        if spec == "nan":
            spec = ""

        # Kiểm tra khớp tên
        if name in ai_text_lower:
            # Nếu có quy cách kỹ thuật, ưu tiên kiểm tra khớp quy cách
            if spec and len(spec) >= 2:
                if spec in ai_text_lower:
                    matched_indices.append(idx)
                elif len(name) >= 6:
                    matched_indices.append(idx)
            else:
                if len(name) >= 5:
                    matched_indices.append(idx)

    if not matched_indices:
        return pd.DataFrame()

    # Khử trùng lặp và giữ nguyên thứ tự xuất hiện
    seen = set()
    unique_indices = [x for x in matched_indices if not (x in seen or seen.add(x))]
    matched_df = df_th.loc[unique_indices].copy()

    stt_col = []
    for idx, r in matched_df.iterrows():
        val = r.get("stt")
        try:
            stt_col.append(int(val))
        except Exception:
            stt_col.append(len(stt_col) + 1)

    out_df = pd.DataFrame({
        "STT": stt_col,
        "Tên vật tư trong kho xưởng": matched_df["ten_vat_tu"],
        "Quy cách kỹ thuật": matched_df["quy_cach"].fillna("-"),
        "ĐVT": matched_df["dvt"],
        "Tồn thực tế (Live Sheet)": matched_df["ton_thuc_te"],
        "Trạng thái kho": matched_df.get("trang_thai_badge", matched_df.get("trang_thai", "An toàn"))
    })
    return out_df.reset_index(drop=True)


# ═══════════════════════════════════════════════════════════════════════════
# BƯỚC 3: MASTER SYSTEM PROMPT KỸ SƯ TRƯỞNG AGENT 2 (NVIDIA NIM AI)
# ═══════════════════════════════════════════════════════════════════════════

SYSTEM_PROMPT_AGENT2_CHIEF_ENGINEER = """
Bạn là KỸ SƯ TRƯỞNG CƠ - ĐIỆN của Công ty Cổ phần Phân bón và Hóa chất Cần Thơ (CFC).
Bạn đang vận hành trên AGENT 2: GIÁM SÁT KHO THỰC TẾ XƯỞNG CƠ - ĐIỆN.

=== NGUYÊN TẮC CỐT LÕI BẮT BUỘC KHI TƯ VẤN & BÓC TÁCH VẬT TƯ ===

1. HIỂU ĐÚNG BẢN CHẤT KỸ THUẬT - KHÔNG GỢI Ý THỪA:
   - Trước khi chọn vật tư, phải xác định thiết bị người dùng yêu cầu là KẾT CẤU TĨNH (Sàn thao tác/sàng thao tác đứng làm việc, cầu thang, lan can, khung đỡ, kệ) hay THIẾT BỊ ĐỘNG (Máy sàng phân bón, băng tải, vít tải, quạt).
   - Nếu là SÀN THAO TÁC (kể cả khi người dùng gõ nhầm là "sàng thao tác" kèm kích thước Dài x Rộng x Cao như 2m x 2m x 2m): Đây chỉ là sàn công tác tĩnh để công nhân đi lại và đứng thao tác kỹ thuật an toàn trên cao. Chỉ bóc tách:
     + Khung cột chịu lực & dầm chính: Thép hình U, I, V hoặc Thép hộp.
     + Mặt sàn đi lại: Thép tấm chống trượt, thép tấm trơn, tấm sàn grating.
     + Lan can bảo vệ & thang: Thép ống tròn kẽm, thép V.
     + Phụ kiện liên kết & hoàn thiện: Bản mã chân cột, bulong/tắc kê, que hàn, đá cắt/đá mài, sơn chống rỉ (Jotun/KCC).
   - TUYỆT ĐỐI KHÔNG đưa Động cơ, Hộp số, Dây curoa, Vòng bi, Lưới sàng hay Bánh xe vào Sàn thao tác tĩnh!

2. QUY TẮC SỐ 1: BẮT BUỘC TẬN DỤNG TOÀN BỘ VẬT TƯ TỒN KHO THỰC TẾ (TỒN > 0) TRƯỚC TIÊN:
   - Bạn BẮT BUỘC phải rà soát kỹ lưỡng bảng "[PHẦN 1: DANH MỤC VẬT TƯ ĐANG CÒN TỒN KHO THỰC TẾ XƯỞNG (TỒN > 0)]" được cấp trong thẻ `<DỮ_LIỆU_KHO_XƯỞNG_AGENT_2>`.
   - Mọi cấu kiện gia công (Cột, Dầm, Sàn, Lan can, Bulong, Que hàn, Sơn, Đá cắt) PHẢI ƯU TIÊN TUYỆT ĐỐI chọn từ các vật tư CÒN TỒN THỰC TẾ trong kho xưởng:
     * Cột chịu lực & Khung dầm chính: Tận dụng Thép U (U120, U100), Thép V (V63, V50), Thép I, Thép hộp vuông/chữ nhật đang có tồn kho.
     * Mặt sàn: Tận dụng Thép tấm (3mm, 5mm...), Thép tấm chống trượt hoặc Grating đang có tồn kho.
     * Lan can an toàn: Tận dụng Thép ống tròn (ống kẽm Φ42, Φ34, Φ49...), Thép V đang có sẵn.
     * Vật tư tiêu hao & liên kết: Tận dụng Que hàn, Bulong mạ/inox (M10, M12, M14...), Đá cắt/mài, Sơn chống rỉ đang có tồn kho xưởng.
   - Định danh vật tư HOÀN TOÀN bằng `Tên vật tư trong kho xưởng`, `Quy cách kỹ thuật`, `ĐVT`, `Tồn thực tế` như trong Live Sheet.
   - TUYỆT ĐỐI KHÔNG ghi cột 'Mã VT' / 'Mã vật tư' trong bất kỳ bảng nào (kho xưởng Agent 2 theo dõi độc lập hoàn toàn với kho gốc Agent 1).

3. QUY TẮC MUA BỔ SUNG: CHỈ ĐỀ XUẤT MUA KHI THỰC SỰ KHÔNG CÓ TRONG KHO:
   - Chỉ được đưa vào mục "3. DANH MỤC VẬT TƯ ĐỀ XUẤT MUA BỔ SUNG" trong 2 trường hợp duy nhất:
     (a) Vật tư hoặc quy cách bắt buộc đó hoàn toàn KHÔNG CÓ trong danh mục xưởng hoặc Tồn kho = 0.
     (b) Số lượng tồn kho thực tế nhỏ hơn số lượng cần gia công theo kích thước.
   - NẾU KHO XƯỞNG ĐÃ ĐÁP ỨNG ĐỦ 100% CÁC HẠNG MỤC CẦN THIẾT: Ghi rõ ràng:
     "✅ Kho xưởng CFC đáp ứng đủ 100% vật tư gia công. KHÔNG CẦN ĐỀ XUẤT MUA THÊM."

4. TÍNH TOÁN KHỐI LƯỢNG KỸ THUẬT THEO KÍCH THƯỚC:
   - Tính toán chiều dài, diện tích và quy đổi số lượng (Cây 6m, Tấm, Kg, Bộ) sát thực tế thi công để xuất kho chính xác.

5. CẤU TRÚC BÁO CÁO KỸ SƯ TRƯỞNG:
   - **1. PHÂN TÍCH KẾT CẤU & PHƯƠNG ÁN GIA CÔNG:** (Ngắn gọn, rõ giải pháp kết cấu).
   - **2. BẢNG BÓC TÁCH VẬT TƯ TẬN DỤNG TỒN KHO XƯỞNG CFC (LIVE SHEET):**
     | STT | Hạng mục công việc | Tên vật tư trong kho xưởng | Quy cách kỹ thuật | ĐVT | Tồn thực tế | SL đề xuất xuất | Trạng thái kho |
   - **3. DANH MỤC VẬT TƯ THIẾU CẦN MUA BỔ SUNG & LƯU Ý AN TOÀN:** (Nếu đủ 100% thì ghi rõ không cần mua thêm).
"""


def query_nvidia_nim_chief_engineer(
    full_prompt: str,
    system_instruction: str,
    nvidia_key: str = "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",
    model: str = "nvidia/nemotron-3-ultra-550b-a55b"
) -> str:
    """Gọi mô hình AI NVIDIA NIM chuyên gia bóc tách kỹ thuật tiếng Việt chuẩn 100% cho Agent 2."""
    import requests
    url = "https://integrate.api.nvidia.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {nvidia_key}",
        "Content-Type": "application/json"
    }
    candidate_models = [
        model,
        "nvidia/nemotron-3-ultra-550b-a55b",
        "deepseek-ai/deepseek-v4.1-flash"
    ]
    models_to_try = []
    for m in candidate_models:
        if m and m not in models_to_try:
            models_to_try.append(m)

    for m in models_to_try:
        payload = {
            "model": m,
            "messages": [
                {"role": "system", "content": system_instruction},
                {"role": "user", "content": full_prompt + "\n\nQUAN TRỌNG: Trả lời HOÀN TOÀN bằng TIẾNG VIỆT, tuân thủ đúng bảng bóc tách tận dụng kho xưởng."}
            ],
            "temperature": 0.1,
            "max_tokens": 4096
        }
        try:
            resp = requests.post(url, headers=headers, json=payload, timeout=12)
            if resp.status_code == 200:
                data = resp.json()
                msg = data.get("choices", [{}])[0].get("message", {})
                content = msg.get("content")
                if content and len(str(content).strip()) > 50:
                    return str(content).strip()
                reasoning = msg.get("reasoning_content")
                if reasoning and len(str(reasoning).strip()) > 50:
                    r_text = str(reasoning).strip()
                    m_start = re.search(r'(?:\*\*|\#\#|\d+\.\s*\*\*)\s*(?:1\.\s*)?PHÂN TÍCH', r_text, re.IGNORECASE)
                    if m_start:
                        return r_text[m_start.start():].strip()
        except Exception:
            continue
    return ""


def query_ai_chief_engineer(user_prompt: str, df_th: pd.DataFrame, api_key: str = "", model_name: str = "gemini-flash-lite-latest", uploaded_file=None, nvidia_api_key: str = "") -> str:
    """Gọi Kỹ sư Trưởng AI tư vấn kỹ thuật và bóc tách vật tư tận dụng kho xưởng (Tốc độ siêu nhanh 2-4s với Gemini Flash Lite, dự phòng NVIDIA NIM)."""
    import google.generativeai as genai
    from PIL import Image
    import base64

    # 1. Phân tích intent & Blacklist/Whitelist
    intent_info = classify_engineering_intent(user_prompt)

    # 2. Tạo bảng tồn kho Live Sheet có cấu trúc, ưu tiên hàng còn tồn
    inventory_table_str = build_agent2_dynamic_inventory_table(df_th, intent_info)

    full_prompt = f"""
Dưới đây là toàn bộ DỮ LIỆU TỒN KHO THỰC TẾ TẠI XƯỞNG được đồng bộ trực tiếp từ Google Sheet của Agent 2:
<DỮ_LIỆU_KHO_XƯỞNG_AGENT_2>
{inventory_table_str}
</DỮ_LIỆU_KHO_XƯỞNG_AGENT_2>

---
LƯU Ý PHÂN LOẠI KỸ THUẬT:
{intent_info['engineering_note']}

YÊU CẦU KỸ THUẬT / GIA CÔNG CỦA KỸ SƯ:
{intent_info['cleaned_prompt']}

Hãy phân tích, tư vấn kỹ thuật cặn kẽ và bóc tách danh mục vật tư theo đúng các nguyên tắc bắt buộc của Kỹ sư Trưởng Cơ - Điện (Ưu tiên tận dụng 100% vật tư còn tồn trong kho xưởng CFC trước)!
"""

    # 1. ƯU TIÊN SỐ 1 VỀ TỐC ĐỘ: GOOGLE GEMINI FLASH LITE (Phản hồi siêu tốc 2-5 giây)
    fallback_b64 = "QVEuQWI4Uk42TGNFLWlhR0Jrc1RCdjdVWWNuNVJqMk5fSVBYWlhMaE5PY2NDM0sza3dLZXc="
    default_key = base64.b64decode(fallback_b64).decode("utf-8")
    effective_key = api_key or os.environ.get("GOOGLE_API_KEY", "") or default_key

    if effective_key:
        try:
            genai.configure(api_key=effective_key)
            content_parts = [full_prompt]
            if uploaded_file is not None:
                try:
                    pil_img = Image.open(uploaded_file)
                    content_parts.append(pil_img)
                except Exception:
                    pass

            candidate_models = [
                model_name,
                "gemini-flash-lite-latest",
                "gemini-3.1-flash-lite",
                "gemini-3.5-flash",
                "gemini-2.5-flash",
                "gemini-1.5-flash"
            ]
            models_to_try = []
            for m in candidate_models:
                if m and m not in models_to_try:
                    models_to_try.append(m)

            for m in models_to_try:
                try:
                    model_inst = genai.GenerativeModel(
                        model_name=m,
                        system_instruction=SYSTEM_PROMPT_AGENT2_CHIEF_ENGINEER
                    )
                    resp = model_inst.generate_content(content_parts, request_options={"timeout": 20})
                    if resp and resp.text:
                        return resp.text
                except Exception:
                    continue
        except Exception:
            pass

    # 2. DỰ PHÒNG: ĐỘNG CƠ NVIDIA NIM (Nếu Gemini gặp sự cố mạng hoặc hết quota)
    effective_nv_key = nvidia_api_key or os.environ.get("NVIDIA_API_KEY", "") or "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"
    if uploaded_file is None and effective_nv_key:
        try:
            nv_res = query_nvidia_nim_chief_engineer(full_prompt, SYSTEM_PROMPT_AGENT2_CHIEF_ENGINEER, nvidia_key=effective_nv_key)
            if nv_res:
                return f"🟢 **[ĐỘNG CƠ DỰ PHÒNG: NVIDIA NIM AI]**\n\n{nv_res}"
        except Exception:
            pass

    return "⚠️ Không thể kết nối với mô hình AI Kỹ sư trưởng. Vui lòng kiểm tra lại kết nối mạng hoặc thử lại sau vài giây."


# ═══════════════════════════════════════════════════════════════════════════
# STREAMLIT UI: AGENT 2 — SINGLE-VIEWPORT INDUSTRIAL DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
@st.fragment
def render_agent2_ui(api_key: str = "", model_name: str = "gemini-flash-lite-latest", is_split: bool = False, nvidia_api_key: str = ""):
    import streamlit as st
    import html as py_html
    import time
    import datetime

    target_email = "linhb1702638@gmail.com"

    viewport_lock = """
        html, body {
            overflow-x: hidden !important;
            overflow-y: auto !important;
            height: auto !important;
        }

        [data-testid="stAppViewContainer"] {
            overflow-y: auto !important;
            height: auto !important;
            min-height: 100vh !important;
            background: linear-gradient(180deg, #F8FAFC 0%, #F1F5F9 100%) !important;
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, sans-serif !important;
            color: #0F172A !important;
        }

        [data-testid="stMainBlockContainer"], section.main {
            overflow: visible !important;
            height: auto !important;
        }

        .block-container {
            padding: 8px 16px 20px 16px !important;
            max-width: 100% !important;
            height: auto !important;
            display: flex !important;
            flex-direction: column !important;
            overflow: visible !important;
        }

        .ai-result-scrollbox,
        [data-testid="stMarkdownContainer"] {
            color: #0F172A !important;
            line-height: 1.6 !important;
        }

        .ai-result-scrollbox p,
        .ai-result-scrollbox li,
        .ai-result-scrollbox span,
        .ai-result-scrollbox div,
        [data-testid="stMarkdownContainer"] p,
        [data-testid="stMarkdownContainer"] li,
        [data-testid="stMarkdownContainer"] span {
            color: #1E293B !important;
            font-size: 0.9rem !important;
        }

        .ai-result-scrollbox h1,
        .ai-result-scrollbox h2,
        .ai-result-scrollbox h3,
        .ai-result-scrollbox h4,
        .ai-result-scrollbox strong,
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        [data-testid="stMarkdownContainer"] strong {
            color: #0F172A !important;
            font-weight: 700 !important;
        }

        .ai-result-scrollbox table,
        [data-testid="stMarkdownContainer"] table {
            width: 100% !important;
            border-collapse: collapse !important;
            margin: 14px 0 !important;
            font-size: 0.85rem !important;
            background: #FFFFFF !important;
            border: 1px solid #CBD5E1 !important;
            border-radius: 10px !important;
            overflow: hidden !important;
        }

        .ai-result-scrollbox th,
        [data-testid="stMarkdownContainer"] th {
            background: #F1F5F9 !important;
            color: #0F172A !important;
            font-weight: 700 !important;
            padding: 10px 14px !important;
            border: 1px solid #CBD5E1 !important;
            text-align: left !important;
        }

        .ai-result-scrollbox td,
        [data-testid="stMarkdownContainer"] td {
            padding: 9px 14px !important;
            border: 1px solid #E2E8F0 !important;
            color: #1E293B !important;
            background: #FFFFFF !important;
        }

        .ai-result-scrollbox tr:nth-child(even) td,
        [data-testid="stMarkdownContainer"] tr:nth-child(even) td {
            background: #F8FAFC !important;
        }

        .ai-result-scrollbox::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        .ai-result-scrollbox::-webkit-scrollbar-track {
            background: #F1F5F9;
            border-radius: 8px;
        }
        .ai-result-scrollbox::-webkit-scrollbar-thumb {
            background: #CBD5E1;
            border-radius: 8px;
        }
        .ai-result-scrollbox::-webkit-scrollbar-thumb:hover {
            background: #94A3B8;
        }
    """ if not is_split else """
        [data-testid="stAppViewContainer"], .main {
            background: linear-gradient(180deg, #F8FAFC 0%, #F1F5F9 100%) !important;
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Text", "Segoe UI", Roboto, sans-serif !important;
            color: #0F172A !important;
        }
    """

    st.markdown("<style>" + viewport_lock + """
        [data-testid="stHeader"] {
            display: none !important;
        }

        h1, h2, h3, h4, .section-header {
            font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", sans-serif !important;
            color: #0F172A !important;
            font-weight: 800 !important;
            letter-spacing: -0.02em !important;
        }

        .cfc-kpi-card {
            background: #FFFFFF !important;
            border: 1px solid #E2E8F0 !important;
            border-radius: 20px !important;
            padding: 18px 22px !important;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04) !important;
            transition: transform 0.25s cubic-bezier(0.25, 0.1, 0.25, 1), box-shadow 0.25s ease, border-color 0.25s ease !important;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            text-align: left;
            height: 100%;
            cursor: pointer;
            position: relative;
        }

        .cfc-kpi-card:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 24px rgba(0, 0, 0, 0.08) !important;
            border-color: #2563EB !important;
        }

        .cfc-kpi-title {
            font-size: 11.5px !important;
            font-weight: 700 !important;
            color: #64748B !important;
            text-transform: uppercase !important;
            letter-spacing: 0.05em !important;
            display: flex;
            align-items: center;
            gap: 6px;
            margin-bottom: 6px;
        }

        .cfc-kpi-value-row {
            display: flex;
            align-items: baseline;
            gap: 4px;
        }

        .cfc-kpi-number {
            font-size: 26px !important;
            font-weight: 800 !important;
            color: #0F172A !important;
            letter-spacing: -0.025em !important;
            line-height: 1.1;
        }

        .cfc-kpi-unit {
            font-size: 12px !important;
            font-weight: 600 !important;
            color: #64748B !important;
        }

        /* 4. Streamlit Interactive 5 KPI Buttons */
        div.st-key-btn_filter_all > button,
        div.st-key-btn_filter_safe > button,
        div.st-key-btn_filter_warning > button,
        div.st-key-btn_filter_out > button,
        div.st-key-btn_filter_issues > button {
            background: #FFFFFF !important;
            border: 1px solid #E2E8F0 !important;
            border-radius: 18px !important;
            padding: 12px 16px !important;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.03) !important;
            transition: transform 0.25s cubic-bezier(0.25, 0.1, 0.25, 1), box-shadow 0.25s ease, border-color 0.25s ease !important;
            height: 76px !important;
            display: flex !important;
            flex-direction: column !important;
            align-items: flex-start !important;
            justify-content: center !important;
            text-align: left !important;
            white-space: pre-line !important;
            cursor: pointer !important;
            color: #1E293B !important;
        }

        div.st-key-btn_filter_all > button:hover,
        div.st-key-btn_filter_safe > button:hover,
        div.st-key-btn_filter_warning > button:hover,
        div.st-key-btn_filter_out > button:hover,
        div.st-key-btn_filter_issues > button:hover {
            transform: translateY(-2px) !important;
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.08) !important;
            border-color: #2563EB !important;
            background: #F8FAFC !important;
        }

        div[class*="st-key-btn_filter_"] > button p {
            margin: 0 !important;
            line-height: 1.25 !important;
        }
        div[class*="st-key-btn_filter_"] > button p:first-child {
            font-size: 11px !important;
            font-weight: 700 !important;
            text-transform: uppercase !important;
            letter-spacing: 0.05em !important;
            color: #64748B !important;
        }
        div[class*="st-key-btn_filter_"] > button p:last-child {
            font-size: 26px !important;
            font-weight: 800 !important;
            color: #0F172A !important;
            letter-spacing: -0.025em !important;
            line-height: 1.1 !important;
        }

        /* Specific Status Active Variations for KPI Buttons */
        div.st-key-btn_filter_all > button[kind="primary"],
        div.st-key-btn_filter_safe > button[kind="primary"],
        div.st-key-btn_filter_warning > button[kind="primary"],
        div.st-key-btn_filter_out > button[kind="primary"],
        div.st-key-btn_filter_issues > button[kind="primary"] {
            border-color: #0F172A !important;
            box-shadow: 0 0 0 2px #0F172A, 0 6px 18px rgba(15, 23, 42, 0.15) !important;
            background: #FFFFFF !important;
            transform: translateY(-1px) !important;
        }
        div[class*="st-key-btn_filter_"] > button[kind="primary"] p:last-child {
            color: #0F172A !important;
        }

        /* 5. Streamlit Interactive Buttons & Inputs */
        .stButton > button {
            border-radius: 9999px !important;
            font-weight: 600 !important;
            border: 1px solid #E2E8F0 !important;
            background: #FFFFFF !important;
            color: #334155 !important;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04) !important;
            transition: all 0.2s cubic-bezier(0.25, 0.1, 0.25, 1) !important;
        }

        .stButton > button:hover {
            transform: translateY(-1px) !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08) !important;
            border-color: #CBD5E1 !important;
            background: #F8FAFC !important;
            color: #0F172A !important;
        }

        .stButton > button[kind="primary"] {
            background: #0F172A !important;
            color: #FFFFFF !important;
            border: 1px solid #0F172A !important;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.2) !important;
        }

        .stButton > button[kind="primary"]:hover {
            background: #2563EB !important;
            border-color: #2563EB !important;
            color: #FFFFFF !important;
            box-shadow: 0 6px 18px rgba(37, 99, 235, 0.25) !important;
        }

        /* 6. Clean Table / DataFrame Container */
        [data-testid="stDataFrame"] {
            border: 1px solid #E2E8F0 !important;
            border-radius: 18px !important;
            background: #FFFFFF !important;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04) !important;
            overflow: hidden !important;
        }

        /* 7. Input & Text Area */
        .stTextInput > div > div > input {
            border-radius: 9999px !important;
            border: 1.5px solid #E2E8F0 !important;
            background: #FFFFFF !important;
            padding: 10px 18px !important;
            font-size: 14px !important;
            color: #0F172A !important;
            transition: all 0.2s ease !important;
        }
        .stTextArea > div > div > textarea {
            border-radius: 16px !important;
            border: 1.5px solid #E2E8F0 !important;
            background: #FFFFFF !important;
            padding: 12px 18px !important;
            font-size: 14px !important;
            color: #0F172A !important;
        }

        .stTextInput > div > div > input:focus,
        .stTextArea > div > div > textarea:focus {
            border-color: #2563EB !important;
            box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12) !important;
        }

        /* 8. Custom 5px Scrollbar */
        .agent2-scroll::-webkit-scrollbar,
        [data-testid="stDataFrame"] ::-webkit-scrollbar {
            width: 5px;
            height: 5px;
        }
        .agent2-scroll::-webkit-scrollbar-track,
        [data-testid="stDataFrame"] ::-webkit-scrollbar-track {
            background: #F1F5F9;
        }
        .agent2-scroll::-webkit-scrollbar-thumb,
        [data-testid="stDataFrame"] ::-webkit-scrollbar-thumb {
            background: #CBD5E1;
            border-radius: 9999px;
        }
        .agent2-scroll::-webkit-scrollbar-thumb:hover,
        [data-testid="stDataFrame"] ::-webkit-scrollbar-thumb:hover {
            background: #94A3B8;
        }

        /* 9. Rounded Pills & Tabs */
        .stTabs [data-baseweb="tab-highlight"],
        .stTabs [data-baseweb="tab-border"] {
            display: none !important;
        }
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px !important;
            background: transparent !important;
            border: none !important;
            box-shadow: none !important;
            padding: 4px 0 !important;
            margin-bottom: 0.85rem !important;
        }

        .stTabs [data-baseweb="tab"] {
            height: 44px !important;
            padding: 8px 24px !important;
            border-radius: 9999px !important;
            margin-right: 8px !important;
            background: #FFFFFF !important;
            border: 1px solid #E2E8F0 !important;
            color: #64748B !important;
            font-weight: 600 !important;
            font-size: 0.9rem !important;
            transition: all 0.2s ease !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03) !important;
        }

        .stTabs [data-baseweb="tab"]:hover {
            color: #0F172A !important;
            border-color: #CBD5E1 !important;
            background: #F8FAFC !important;
        }

        .stTabs [data-baseweb="tab"][aria-selected="true"],
        .stTabs [aria-selected="true"] {
            background: #0F172A !important;
            border: 1.5px solid #0F172A !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.2) !important;
            font-weight: 700 !important;
        }
    </style>
    """, unsafe_allow_html=True)

    # 0. TTL Check: Tự động xóa kết quả tư vấn AI sau 30 phút (1800s)
    if "ag2_ai_response" in st.session_state:
        ai_ts = st.session_state.get("ag2_ai_timestamp", 0)
        if ai_ts > 0 and (time.time() - ai_ts) > 1800:
            st.session_state.pop("ag2_ai_response", None)
            st.session_state.pop("ag2_ai_last_prompt", None)
            st.session_state.pop("ag2_ai_timestamp", None)

    # 1. Nạp dữ liệu có auto-refresh 1 giờ (@st.cache_data(ttl=3600))
    force_sync = st.session_state.pop("ag2_force_sync", False)
    with st.spinner("Đang kết nối kho dữ liệu trực tiếp Google Sheet..."):
        df_th, df_nh, df_xt, last_sync_str = load_data_from_sheets(force_sync=force_sync)

    df_th = process_inventory_dataframe(df_th)

    if df_th.empty:
        st.error("Chưa có dữ liệu từ Google Sheet. Vui lòng bấm nút 'Đồng bộ Google Sheet' bên dưới.")
        if st.button("🔄 Đồng bộ ngay Google Sheet"):
            st.session_state["ag2_force_sync"] = True
            st.rerun(scope="fragment")
        return

    # 2. Header Bar (Apple Store Style)
    col_h1, col_h2 = st.columns([5.5, 2.5], vertical_alignment="center")
    with col_h1:
        st.markdown(f"""
        <div style="display:flex; align-items:center; gap:12px; background:#FFFFFF; border:1px solid #E2E8F0; border-radius:16px; padding:14px 20px; border-radius:20px; box-shadow:0 4px 20px rgba(0,0,0,0.04);">
            <div style="font-size:1.35rem;">⚡</div>
            <div>
                <div style="font-size:1.02rem; font-weight:700; color:#0F172A; letter-spacing:-0.02em;">
                    AGENT 2: GIÁM SÁT TỒN KHO
                </div>
                <div style="font-size:0.75rem; color:#64748B; margin-top:1px;">
                    Đồng bộ trực tiếp Google Sheet • Tự động gửi cảnh báo & báo cáo về <strong>{target_email}</strong> lúc <strong>16:00</strong>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    with col_h2:
        st.markdown(f"""
        <div style="display:flex; align-items:center; justify-content:flex-end; gap:8px;">
            <span style="font-size:0.73rem; color:#059669; background:#ECFDF5; border:1px solid #10B981; padding:5px 12px; border-radius:980px; font-weight:600; box-shadow:0 0 10px rgba(46,204,113,0.2);">
                ● LIVE SHEET ({last_sync_str})
            </span>
            <span style="font-size:0.73rem; color:#334155; background:#F1F5F9; border:1px solid #CBD5E1; padding:5px 12px; border-radius:980px; font-weight:600;">
                ⏰ CRON: 16:00
            </span>
        </div>
        """, unsafe_allow_html=True)

    kpis = get_inventory_kpis(df_th, df_xt)

    # 3. Main Workspace: Cột Trái (290px) + Cột Phải (Flex-1)
    col_left, col_right = st.columns([1, 1.6], gap="small") if is_split else st.columns([1.1, 3.4], gap="small")

    # ── CỘT TRÁI: KỸ SƯ TRƯỞNG AI (ARGON 4) + LỌC DANH MỤC + TIỆN ÍCH ──
    with col_left:
        st.markdown("""
        <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:16px; padding:14px 18px; border-radius:20px; margin-bottom:10px; box-shadow:0 4px 20px rgba(0,0,0,0.04);">
            <div style="display:flex; align-items:center; justify-content:space-between;">
                <span style="font-size:0.86rem; font-weight:700; color:#0F172A; letter-spacing:-0.02em;">🧠 KỸ SƯ TRƯỞNG AI CƠ-ĐIỆN</span>
                <span style="font-size:0.65rem; background:#10B981; color:#0F172A; padding:2px 8px; border-radius:980px; font-weight:700;">⚡ NVIDIA NIM AI</span>
            </div>
            <div style="font-size:0.73rem; color:#64748B; margin-top:3px; line-height:1.35;">
                Tư vấn kỹ thuật • Bóc tách bản vẽ / công trình • Tận dụng kho sẵn có
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.form(key="ag2_form", clear_on_submit=False):
            user_ai_inquiry = st.text_input(
                "Yêu cầu kỹ thuật / Gia công công trình (Gõ xong nhấn Enter):",
                placeholder="VD: Sửa chữa mái nhà dột rộng 100m2, bệ đỡ động cơ...",
                key="ag2_ai_inquiry_input",
                label_visibility="collapsed"
            )

            uploaded_doc = st.file_uploader(
                "📷 Tải ảnh bản vẽ / tem nhãn kỹ thuật:",
                type=["png", "jpg", "jpeg", "webp"],
                key="ag2_ai_image_upload",
                help="Tải ảnh chụp bản vẽ gia công, tem nhãn động cơ, thông số thiết bị...",
                label_visibility="collapsed"
            )

            c_ask, c_clr = st.columns([3, 1])
            with c_ask:
                submit_ask = st.form_submit_button("🚀 Bóc tách & Tư vấn", use_container_width=True, type="primary")
            with c_clr:
                submit_clr = st.form_submit_button("Xóa", use_container_width=True, help="Xóa câu hỏi và kết quả")

        if submit_clr:
            st.session_state.pop("ag2_ai_response", None)
            st.session_state.pop("ag2_ai_last_prompt", None)
            st.session_state.pop("ag2_ai_timestamp", None)
            st.rerun(scope="fragment")

        if submit_ask:
            if user_ai_inquiry.strip() or uploaded_doc is not None:
                prompt_text = user_ai_inquiry.strip() if user_ai_inquiry.strip() else "Phân tích bản vẽ kỹ thuật / hình ảnh đính kèm và bóc tách vật tư tận dụng kho."
                with st.spinner("Đang phân tích kỹ thuật & đối chiếu vật tư kho xưởng..."):
                    ans = query_ai_chief_engineer(
                        user_prompt=prompt_text,
                        df_th=df_th,
                        api_key=api_key,
                        model_name=model_name,
                        uploaded_file=uploaded_doc,
                        nvidia_api_key=nvidia_api_key
                    )
                    st.session_state["ag2_ai_response"] = ans
                    st.session_state["ag2_ai_last_prompt"] = prompt_text
                    st.session_state["ag2_ai_timestamp"] = time.time()
                    st.rerun(scope="fragment")
            else:
                st.warning("Vui lòng nhập bài toán kỹ thuật hoặc tải ảnh bản vẽ gia công.")

        st.markdown("<hr style='margin:8px 0; border:none; border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)

        search_query = st.text_input(
            "Tra cứu nhanh danh sách:",
            placeholder="🔍 Lọc nhanh bảng (tên, quy cách...)",
            key="ag2_quick_search_input",
            label_visibility="collapsed"
        )

        st.markdown("<div style='font-size:0.8rem; font-weight:700; color:#1E293B; margin:6px 0;'>⚡ THAO TÁC NHANH</div>", unsafe_allow_html=True)

        c_sync, c_mail = st.columns(2)
        with c_sync:
            if st.button("🔄 Đồng bộ", use_container_width=True, help="Xóa cache và kéo bản mới nhất từ Google Sheet"):
                st.session_state["ag2_force_sync"] = True
                st.toast("✓ Đang xóa cache và đồng bộ dữ liệu mới nhất từ Google Sheet!", icon="✅")
                time.sleep(0.3)
                st.rerun(scope="fragment")
        with c_mail:
            if st.button("✉ Gửi Daily", use_container_width=True, type="primary", help="Gửi báo cáo cảnh báo tức thì về email"):
                with st.spinner("Đang gửi email..."):
                    ok, msg = send_email_report("daily")
                    if ok:
                        st.toast(f"✓ Đã gửi email báo cáo tới {target_email}!", icon="📧")
                    else:
                        st.error(msg)

        with st.expander("📧 Gửi báo cáo định kỳ khác", expanded=False):
            c_wk, c_mo = st.columns(2)
            with c_wk:
                if st.button("Báo cáo Tuần", use_container_width=True):
                    ok, msg = send_email_report("weekly")
                    if ok: st.toast("✓ Đã gửi báo cáo tuần!", icon="📧")
            with c_mo:
                if st.button("Báo cáo Tháng", use_container_width=True):
                    ok, msg = send_email_report("monthly")
                    if ok: st.toast("✓ Đã gửi báo cáo tháng!", icon="📧")
            if st.button("Báo cáo Năm (Thanh lý)", use_container_width=True):
                ok, msg = send_email_report("yearly")
                if ok: st.toast("✓ Đã gửi báo cáo thanh lý!", icon="📧")

        st.markdown(f"""
        <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:16px; padding:12px 16px; border-radius:18px; margin-top:10px; font-size:0.75rem; color:#64748B; box-shadow:0 2px 10px rgba(0,0,0,0.03);">
            <div>📌 <strong style="color:#0F172A;">Cấu hình tự động:</strong></div>
            <div>• Điểm đích: <code>{target_email}</code></div>
            <div>• Tự động quét: <strong>Mỗi 60 phút</strong></div>
            <div>• Giờ gửi Daily: <strong>16:00 mỗi ngày</strong></div>
        </div>
        """, unsafe_allow_html=True)

    # ── CỘT PHẢI: 4 KPI CARDS (BẤM TRỰC TIẾP ĐỂ LỌC VÀ HIỂN THỊ DỮ LIỆU TẠI CHỖ) ──
    with col_right:
        if "ag2_kpi_filter" not in st.session_state:
            st.session_state["ag2_kpi_filter"] = "ALL"
        if "ag2_active_view" not in st.session_state:
            st.session_state["ag2_active_view"] = "INVENTORY"

        cur_kpi = st.session_state["ag2_kpi_filter"]
        cur_view = st.session_state["ag2_active_view"]

        # Hàng 1: 4 Thẻ KPI Tương tác - Bấm trực tiếp để lọc bảng dữ liệu tương ứng ngay tại chỗ
        col_k1, col_k2, col_k3, col_k4 = st.columns(4, gap="small")

        with col_k1:
            is_sel = (cur_view == "INVENTORY" and cur_kpi == "ALL")
            if st.button(
                f"📦 TỔNG VẬT TƯ\n\n{kpis['total_items']:,} mặt hàng",
                key="btn_filter_all",
                use_container_width=True,
                type="primary" if is_sel else "secondary"
            ):
                st.session_state["ag2_active_view"] = "INVENTORY"
                st.session_state["ag2_kpi_filter"] = "ALL"
                st.rerun(scope="fragment")

        with col_k2:
            is_sel = (cur_view == "INVENTORY" and cur_kpi == "SAFE")
            if st.button(
                f"🟢 AN TOÀN\n\n{kpis['safe']:,} mặt hàng",
                key="btn_filter_safe",
                use_container_width=True,
                type="primary" if is_sel else "secondary"
            ):
                st.session_state["ag2_active_view"] = "INVENTORY"
                st.session_state["ag2_kpi_filter"] = "SAFE"
                st.rerun(scope="fragment")

        with col_k3:
            is_sel = (cur_view == "INVENTORY" and cur_kpi == "WARNING")
            if st.button(
                f"🟡 SẮP HẾT ĐỊNH MỨC\n\n{kpis['warning']:,} mặt hàng",
                key="btn_filter_warning",
                use_container_width=True,
                type="primary" if is_sel else "secondary"
            ):
                st.session_state["ag2_active_view"] = "INVENTORY"
                st.session_state["ag2_kpi_filter"] = "WARNING"
                st.rerun(scope="fragment")

        with col_k4:
            is_sel = (cur_view == "INVENTORY" and cur_kpi == "OUT_OF_STOCK")
            if st.button(
                f"🔴 ĐÃ CẠN (HẾT HÀNG)\n\n{kpis['out_of_stock']:,} mặt hàng",
                key="btn_filter_out",
                use_container_width=True,
                type="primary" if is_sel else "secondary"
            ):
                st.session_state["ag2_active_view"] = "INVENTORY"
                st.session_state["ag2_kpi_filter"] = "OUT_OF_STOCK"
                st.rerun(scope="fragment")

        # Thanh chuyển nhanh sang các nghiệp vụ phụ khác (Nhật ký, Tiêu hao, Báo cáo)
        c_nav1, c_nav2, c_nav3 = st.columns(3, gap="small")
        with c_nav1:
            if st.button("📝 Nhật Ký Xuất Thiết Bị", use_container_width=True, type="primary" if cur_view == "HISTORY" else "secondary"):
                st.session_state["ag2_active_view"] = "HISTORY"
                st.rerun(scope="fragment")
        with c_nav2:
            if st.button("📊 Phân Tích Tiêu Hao", use_container_width=True, type="primary" if cur_view == "ANALYTICS" else "secondary"):
                st.session_state["ag2_active_view"] = "ANALYTICS"
                st.rerun(scope="fragment")
        with c_nav3:
            if st.button("⚙️ Báo Cáo & Cấu Hình", use_container_width=True, type="primary" if cur_view == "CONFIG" else "secondary"):
                st.session_state["ag2_active_view"] = "CONFIG"
                st.rerun(scope="fragment")

        # ── HIỂN THỊ NỘI DUNG THEO VIEW HIỆN TẠI ──
        if cur_view == "INVENTORY":
            # 1. Khung hiển thị kết quả tư vấn AI kèm cơ chế TTL 30 phút và nút Xóa thủ công
            if "ag2_ai_response" in st.session_state:
                ai_ts = st.session_state.get("ag2_ai_timestamp", 0)
                elapsed_min = int((time.time() - ai_ts) / 60) if ai_ts > 0 else 0
                remaining_min = max(0, 30 - elapsed_min)

                ai_hdr_col1, ai_hdr_col2 = st.columns([7.8, 2.2], vertical_alignment="center")
                with ai_hdr_col1:
                    st.markdown(
                        f"<div style='font-size:0.88rem; font-weight:700; color:#0F172A; display:flex; align-items:center; gap:8px;'>"
                        f"<span>💡 KẾT QUẢ TƯ VẤN & BÓC TÁCH KHO TỪ KỸ SƯ TRƯỞNG AI</span>"
                        f"<span style='font-size:0.72rem; color:#64748B; background:rgba(255, 255, 255, 0.06); border:1px solid rgba(143, 168, 155, 0.3); padding:2px 8px; border-radius:12px; font-weight:500;'>⏱️ Tự xóa sau {remaining_min} phút</span>"
                        f"</div>",
                        unsafe_allow_html=True
                    )
                with ai_hdr_col2:
                    if st.button("🗑️ Xóa kết quả", key="btn_clear_ai_result", use_container_width=True, help="Xóa ngay kết quả tư vấn để dọn sạch màn hình"):
                        st.session_state.pop("ag2_ai_response", None)
                        st.session_state.pop("ag2_ai_last_prompt", None)
                        st.session_state.pop("ag2_ai_timestamp", None)
                        st.rerun(scope="fragment")

                last_prompt = st.session_state.get('ag2_ai_last_prompt', '')
                with st.container(border=True):
                    st.markdown(f"<div style='font-size:0.82rem; color:#64748B; margin-bottom:12px; padding-bottom:8px; border-bottom:1px solid #F1F5F9;'>Yêu cầu: <strong style='color:#0F172A;'>'{last_prompt}'</strong></div>", unsafe_allow_html=True)
                    st.markdown(st.session_state["ag2_ai_response"])

                # Đối chiếu trực tiếp Live Sheet (HOÀN TOÀN KHÔNG CÓ CỘT MÃ VT)
                verified_df = extract_verified_items_from_response(st.session_state["ag2_ai_response"], df_th)
                if not verified_df.empty:
                    st.markdown("<div style='font-size:0.82rem; font-weight:700; color:#059669; margin-top:14px; margin-bottom:6px;'>🔍 ĐỐI CHIẾU TRỰC TIẾP TỒN KHO THỰC TẾ TRÊN GOOGLE SHEET CỦA AGENT 2 (THEO TÊN & QUY CÁCH):</div>", unsafe_allow_html=True)
                    st.dataframe(verified_df, use_container_width=True, hide_index=True)
                st.markdown("<hr style='margin:12px 0; border:none; border-top:1px solid #E2E8F0;'>", unsafe_allow_html=True)

            # 2. Lọc bảng theo Thẻ KPI vừa bấm (Sử dụng _status_code nội bộ số hóa chuẩn xác tuyệt đối)
            cur_kpi = st.session_state.get("ag2_kpi_filter", "ALL")
            if cur_kpi == "OUT_OF_STOCK":
                filtered_df = df_th[df_th['_status_code'] == 'OUT_OF_STOCK'].copy()
                kpi_desc = "🔴 ĐÃ CẠN (HẾT HÀNG)"
            elif cur_kpi == "WARNING":
                filtered_df = df_th[df_th['_status_code'] == 'WARNING'].copy()
                kpi_desc = "🟡 SẮP HẾT ĐỊNH MỨC"
            elif cur_kpi == "SAFE":
                filtered_df = df_th[df_th['_status_code'] == 'SAFE'].copy()
                kpi_desc = "🟢 AN TOÀN"

            else:
                filtered_df = df_th.copy()
                kpi_desc = "📦 TẤT CẢ VẬT TƯ"

            # 3. Lọc theo Quick Search (nếu có)
            if search_query.strip():
                sq = search_query.strip().lower()
                filtered_df = filtered_df[
                    filtered_df["ten_vat_tu"].astype(str).str.lower().str.contains(sq, na=False) |
                    filtered_df["quy_cach"].astype(str).str.lower().str.contains(sq, na=False) |
                    filtered_df["ghi_chu"].astype(str).str.lower().str.contains(sq, na=False)
                ]

            # Thông tin chỉ báo bộ lọc
            st.markdown(
                f"<div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;'>"
                f"<span style='font-size:0.82rem; font-weight:600; color:#0F172A;'>Đang lọc theo thẻ: <span style='color:#4CD964; font-weight:700;'>{kpi_desc}</span></span>"
                f"<span style='font-size:0.78rem; color:#64748B;'>Hiển thị: <strong style='color:#0F172A;'>{len(filtered_df):,}</strong> / {len(df_th):,} vật tư</span>"
                f"</div>",
                unsafe_allow_html=True
            )

            # Gán nhãn Huy hiệu trạng thái đẹp mắt
            df_display = filtered_df.copy()
            df_display["trang_thai_badge"] = df_display["trang_thai"].apply(format_status_badge)

            df_table = df_display[[
                "stt", "ten_vat_tu", "quy_cach", "dvt",
                "ton_thuc_te", "dinh_muc_min", "trang_thai_badge", "ghi_chu"
            ]].copy()

            df_table.columns = [
                "STT", "Tên Vật Tư", "Quy Cách Kỹ Thuật", "ĐVT",
                "Tồn Thực Tế", "Định Mức Min", "Trạng Thái Kho", "Ghi Chú"
            ]

            # Hiển thị trực tiếp qua st.dataframe (Cam kết 100% không bị lộ thẻ HTML thô)
            st.dataframe(
                df_table,
                use_container_width=True,
                hide_index=True,
                height=480,
                column_config={
                    "STT": st.column_config.NumberColumn("STT", width="small"),
                    "Tên Vật Tư": st.column_config.TextColumn("Tên Vật Tư", width="large"),
                    "Quy Cách Kỹ Thuật": st.column_config.TextColumn("Quy Cách Kỹ Thuật", width="medium"),
                    "ĐVT": st.column_config.TextColumn("ĐVT", width="small"),
                    "Tồn Thực Tế": st.column_config.NumberColumn("Tồn Kho Thực Tế", format="%.1f", width="small"),
                    "Định Mức Min": st.column_config.NumberColumn("Định Mức Min", format="%.1f", width="small"),
                    "Trạng Thái Kho": st.column_config.TextColumn("Trạng Thái Kho", width="medium"),
                    "Ghi Chú": st.column_config.TextColumn("Ghi Chú", width="medium")
                }
            )

        # ── VIEW 2: NHẬT KÝ XUẤT THIẾT BỊ (SẮP XẾP TỪ MỚI NHẤT ĐẾN CŨ NHẤT) ──
        elif cur_view == "HISTORY":
            st.markdown(f"<div style='font-size:0.78rem; color:#64748B; margin-bottom:4px;'>Nhật ký chi tiết các lần xuất vật tư phục vụ bảo trì thiết bị ({len(df_xt):,} lượt xuất - <strong>Sắp xếp từ mới nhất đến cũ nhất</strong>):</div>", unsafe_allow_html=True)
            if not df_xt.empty:
                def parse_xt_date(val, sheet_name=""):
                    if pd.isna(val) or val is None or str(val).strip() == "":
                        return datetime.date(2000, 1, 1)
                    if isinstance(val, (datetime.date, datetime.datetime)):
                        if "11" in str(sheet_name) and val.month != 11 and val.day <= 12:
                            try:
                                return datetime.date(val.year, 11, val.month)
                            except Exception:
                                pass
                        return val.date() if isinstance(val, datetime.datetime) else val
                    s = str(val).strip()
                    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y"):
                        try:
                            return datetime.datetime.strptime(s, fmt).date()
                        except Exception:
                            pass
                    return datetime.date(2000, 1, 1)

                df_sorted_xt = df_xt.copy()
                df_sorted_xt["_parsed_date"] = [
                    parse_xt_date(r.get("ngay"), r.get("thang", "")) for _, r in df_sorted_xt.iterrows()
                ]
                df_sorted_xt = df_sorted_xt.sort_values(by="_parsed_date", ascending=False)
                df_sorted_xt["ngay_hien_thi"] = df_sorted_xt["_parsed_date"].apply(
                    lambda d: d.strftime("%d/%m/%Y") if d.year > 2000 else ""
                )

                display_xt = df_sorted_xt[["ngay_hien_thi", "ten_vat_tu", "so_luong", "dvt", "nguoi_nhan", "thiet_bi_khu_vuc", "thang"]].copy()
                display_xt.columns = ["Ngày", "Vật Tư", "Số Lượng", "ĐVT", "Người Nhận", "Thiết Bị / Vị Trí", "Tháng"]
                st.dataframe(display_xt, use_container_width=True, hide_index=True, height=480)
            else:
                st.info("Chưa có dữ liệu xuất trong các sheet tháng.")

        # ── VIEW 3: PHÂN TÍCH TIÊU HAO THIẾT BỊ ──
        elif cur_view == "ANALYTICS":
            # 1. Month-Based Dynamic Filter
            available_months = []
            if not df_xt.empty and "thang" in df_xt.columns:
                unique_m = df_xt["thang"].dropna().unique().tolist()
                now_month_num = datetime.datetime.now().month
                curr_str = f"tháng {now_month_num}"
                available_months = sorted(
                    unique_m,
                    key=lambda x: (0 if curr_str in str(x).lower() else 1, str(x)),
                    reverse=False
                )
            if not available_months:
                available_months = ["tháng 11", "tháng 10", "tháng 12"]

            c_fltr1, c_fltr2 = st.columns([2.2, 4.8], vertical_alignment="center")
            with c_fltr1:
                selected_month = st.selectbox(
                    "📅 Chọn chu kỳ phân tích:",
                    options=available_months,
                    index=0,
                    key="analytics_month_selector"
                )

            # Lọc dữ liệu theo tháng đã chọn
            df_m = df_xt[df_xt["thang"] == selected_month].copy() if not df_xt.empty else pd.DataFrame()

            # Date boundary logic: nếu là tháng hiện tại, chỉ lấy đến ngày thực tế hiện tại
            now_dt = datetime.datetime.now()
            is_current_month = (f"tháng {now_dt.month}" in str(selected_month).lower())

            def safe_date_filter(val):
                if val is None:
                    return True
                if isinstance(val, (datetime.datetime, datetime.date)):
                    d = val.date() if isinstance(val, datetime.datetime) else val
                    return d <= now_dt.date()
                s = str(val).strip()
                if not s:
                    return True
                for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m", "%d-%m-%Y"):
                    try:
                        d = datetime.datetime.strptime(s, fmt)
                        if fmt == "%d/%m":
                            d = d.replace(year=now_dt.year)
                        return d.date() <= now_dt.date()
                    except ValueError:
                        pass
                return True

            if is_current_month and not df_m.empty:
                df_m = df_m[df_m["ngay"].apply(safe_date_filter)].copy()

            total_month_qty = df_m["so_luong"].sum() if not df_m.empty else 0
            total_month_rows = len(df_m)

            with c_fltr2:
                st.markdown(
                    f"<div style='font-size:0.78rem; color:#475569; text-align:right; padding-top:6px;'>"
                    f"Chu kỳ: <strong style='color:#2563EB;'>{selected_month.upper()}</strong>"
                    f"{' <span style=\"color:#059669; font-weight:600;\">(● Real-time)</span>' if is_current_month else ''} • "
                    f"Lượt xuất: <strong>{total_month_rows:,}</strong> • "
                    f"Tổng lượng tiêu hao: <strong>{total_month_qty:,.1f}</strong>"
                    f"</div>",
                    unsafe_allow_html=True
                )

            st.markdown("<div style='margin-bottom: 4px;'></div>", unsafe_allow_html=True)

            # 2. Side-by-Side 2 Columns Layout
            col_an_left, col_an_right = st.columns([1, 1], gap="medium")

            # ── CỘT TRÁI: TOP 10 VẬT TƯ TIÊU HAO TRONG THÁNG ──
            with col_an_left:
                st.markdown("""
                <div style="background:#FFFFFF; border:1px solid rgba(226,232,240,0.9); border-radius:14px; padding:10px 14px; margin-bottom:8px; box-shadow:0 2px 6px rgba(15,23,42,0.04);">
                    <div style="font-size:0.84rem; font-weight:700; color:#0F172A; letter-spacing:-0.01em;">
                        🔥 TOP 10 VẬT TƯ TIÊU HAO TRONG THÁNG
                    </div>
                </div>
                """, unsafe_allow_html=True)

                if not df_m.empty:
                    top_mat = df_m.groupby(["ten_vat_tu", "dvt"])["so_luong"].sum().reset_index()
                    top_mat = top_mat.sort_values(by="so_luong", ascending=False).head(10).reset_index(drop=True)
                    top_mat["rank"] = [f"#{i+1}" for i in range(len(top_mat))]

                    if total_month_qty > 0:
                        top_mat["ty_le"] = (top_mat["so_luong"] / total_month_qty * 100).map(lambda x: f"{x:.1f}%")
                    else:
                        top_mat["ty_le"] = "0%"

                    top_mat_display = top_mat[["rank", "ten_vat_tu", "dvt", "so_luong", "ty_le"]].copy()
                    top_mat_display.columns = ["Top Rank", "Tên Vật Tư", "ĐVT", "Tổng Tiêu Hao Trong Tháng", "Tỷ Lệ"]

                    st.dataframe(
                        top_mat_display,
                        use_container_width=True,
                        hide_index=True,
                        height=360,
                        column_config={
                            "Top Rank": st.column_config.TextColumn("Top", width="small"),
                            "Tên Vật Tư": st.column_config.TextColumn("Tên Vật Tư", width="large"),
                            "ĐVT": st.column_config.TextColumn("ĐVT", width="small"),
                            "Tổng Tiêu Hao Trong Tháng": st.column_config.NumberColumn("Tiêu Hao", format="%.1f", width="small"),
                            "Tỷ Lệ": st.column_config.TextColumn("Tỷ Lệ", width="small")
                        }
                    )
                else:
                    st.info(f"Chưa có phát sinh dữ liệu xuất trong {selected_month}.")

            # ── CỘT PHẢI: TOP 10 THIẾT BỊ VỚI SƠ ĐỒ CÂY DRILL-DOWN ──
            with col_an_right:
                st.markdown("""
                <div style="background:#FFFFFF; border:1px solid rgba(226,232,240,0.9); border-radius:14px; padding:10px 14px; margin-bottom:8px; box-shadow:0 2px 6px rgba(15,23,42,0.04);">
                    <div style="font-size:0.84rem; font-weight:700; color:#0F172A; letter-spacing:-0.01em;">
                        ⚙️ TOP 10 THIẾT BỊ / HẠNG MỤC SỬ DỤNG (DRILL-DOWN)
                    </div>
                </div>
                """, unsafe_allow_html=True)

                if not df_m.empty:
                    df_m_clean = df_m.copy()
                    df_m_clean["thiet_bi_khu_vuc"] = df_m_clean["thiet_bi_khu_vuc"].replace("", "Chưa ghi nhận thiết bị")

                    eq_grp = df_m_clean.groupby("thiet_bi_khu_vuc")["so_luong"].sum().reset_index()
                    eq_grp = eq_grp.sort_values(by="so_luong", ascending=False).head(10).reset_index(drop=True)

                    tree_container = st.container(height=360)
                    with tree_container:
                        for rank_idx, eq_row in eq_grp.iterrows():
                            eq_name = eq_row["thiet_bi_khu_vuc"]
                            eq_tot = eq_row["so_luong"]
                            sub_df = df_m_clean[df_m_clean["thiet_bi_khu_vuc"] == eq_name].copy()
                            num_distinct = sub_df["ten_vat_tu"].nunique()

                            expander_label = f"⚙️ [{rank_idx+1}] {eq_name} (Tổng xuất: {eq_tot:,.1f} vật tư - {num_distinct} mã)"
                            with st.expander(expander_label, expanded=(rank_idx == 0)):
                                def fmt_date_str(val):
                                    if isinstance(val, (datetime.datetime, datetime.date)):
                                        return val.strftime("%d/%m")
                                    return str(val or "-").strip()

                                sub_df["ngay_str"] = sub_df["ngay"].apply(fmt_date_str)
                                sub_display = sub_df[["ten_vat_tu", "dvt", "so_luong", "ngay_str", "nguoi_nhan"]].copy()
                                sub_display.columns = ["Tên Vật Tư", "ĐVT", "Số Lượng Xuất", "Ngày Xuất", "Người Nhận"]
                                st.dataframe(
                                    sub_display,
                                    use_container_width=True,
                                    hide_index=True,
                                    height=min(180, 36 + len(sub_display) * 32),
                                    column_config={
                                        "Tên Vật Tư": st.column_config.TextColumn("Tên Vật Tư", width="medium"),
                                        "ĐVT": st.column_config.TextColumn("ĐVT", width="small"),
                                        "Số Lượng Xuất": st.column_config.NumberColumn("Số Lượng", format="%.1f", width="small"),
                                        "Ngày Xuất": st.column_config.TextColumn("Ngày", width="small"),
                                        "Người Nhận": st.column_config.TextColumn("Người Nhận", width="small")
                                    }
                                )
                else:
                    st.info(f"Chưa có phát sinh dữ liệu thiết bị trong {selected_month}.")

        # ── VIEW 4: BÁO CÁO & CẤU HÌNH TỰ ĐỘNG HÓA ──
        elif cur_view == "CONFIG":
            col_cfg1, col_cfg2 = st.columns([1.5, 2.5])
            with col_cfg1:
                st.markdown("""
                <div style="background:#FFFFFF; border:1px solid #E2E8F0; border-radius:8px; padding:12px;">
                    <div style="font-weight:700; color:#0F172A; font-size:0.9rem; margin-bottom:8px;">⚙️ TRẠNG THÁI CRON SYSTEM</div>
                    <div style="font-size:0.8rem; color:#334155; line-height:1.6;">
                        <div>• Tác vụ Windows: <code>CFC_Agent2_Daily_16h</code></div>
                        <div>• Trạng thái: <strong style="color:#059669;">ĐANG KÍCH HOẠT (Enabled)</strong></div>
                        <div>• Tần suất: <strong>Hàng ngày lúc 16:00 (4:00 PM)</strong></div>
                        <div>• Hộp thư nhận: <code>linhb1702638@gmail.com</code></div>
                        <div>• SMTP: <code>smtp.gmail.com:587 (TLS)</code></div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
                report_choice = st.selectbox(
                    "Xem trước mẫu báo cáo email:",
                    ["daily", "weekly", "monthly", "yearly"],
                    format_func=lambda x: {
                        "daily": "1. Daily Digest (Cảnh báo cạn kho cuối ngày)",
                        "weekly": "2. Weekly Report (Tiêu hao tuần & Đề xuất mua)",
                        "monthly": "3. Monthly Report (Tổng hợp tiêu hao & thiết bị)",
                        "yearly": "4. Yearly Report (Báo cáo vật tư tồn chết thanh lý)"
                    }[x]
                )

            with col_cfg2:
                st.markdown(f"<div style='font-weight:700; color:#0F172A; font-size:0.88rem; margin-bottom:6px;'>👁️ XEM TRƯỚC GIAO DIỆN EMAIL HTML ({report_choice.upper()}):</div>", unsafe_allow_html=True)
                subj, email_html = generate_email_html(report_choice, df_th, df_nh, df_xt)
                st.caption(f"Tiêu đề: **{subj}**")
                st.components.v1.html(email_html, height=380, scrolling=True)
