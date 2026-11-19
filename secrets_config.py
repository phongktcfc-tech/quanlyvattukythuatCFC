"""
NẠP KHÓA BÍ MẬT TẬP TRUNG (KHÔNG BAO GIỜ GHI KHÓA THẲNG VÀO CODE)

Thứ tự tìm một khóa:
  1. Biến môi trường (os.environ)
  2. st.secrets  (Streamlit Cloud: Settings → Secrets; máy local: .streamlit/secrets.toml)
  3. File .env trong thư mục dự án
  4. File .env dùng chung cũ: D:/Agent EmailAssistant/.env (nếu tồn tại)

Các khóa dùng trong hệ thống:
  GOOGLE_API_KEY, NVIDIA_API_KEY,
  EMAIL_ADDRESS, EMAIL_APP_PASSWORD, RECIPIENT_EMAIL, SMTP_SERVER, SMTP_PORT
"""

import os
import re
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ENV_FILES = [BASE_DIR / ".env", Path("D:/Agent EmailAssistant/.env")]

_env_cache: dict | None = None


def _load_env_files() -> dict:
    global _env_cache
    if _env_cache is not None:
        return _env_cache
    values: dict = {}
    for env_file in ENV_FILES:
        try:
            if not env_file.exists():
                continue
            with open(env_file, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        k = k.strip()
                        if k and k not in values:
                            values[k] = v.strip().strip("\"'")
        except Exception as e:
            print(f"[secrets] Lỗi đọc {env_file}: {e}")
    _env_cache = values
    return values


# Tên các khóa đã bị tự động làm sạch (do có ký tự lạ) — giao diện dùng để cảnh báo người quản trị. Không chứa giá trị khóa.
SANITIZED_KEYS: set = set()
_SKIP_SANITIZE = {"SMTP_SERVER"}


def clean_secret_value(name: str, val: str) -> str:
    """
    Làm sạch giá trị khóa: bỏ dấu nháy bao quanh, khoảng trắng, xuống dòng, ký tự không in được/không phải ASCII
    (dấu cách không ngắt NBSP, ký tự zero-width... thường lẫn vào khi copy-dán và gây lỗi 'Illegal header value').
    Không khóa/mật khẩu hợp lệ nào chứa các ký tự đó (mật khẩu ứng dụng Gmail 'abcd efgh ijkl mnop' vẫn dùng được khi bỏ dấu cách).
    """
    if not isinstance(val, str):
        val = str(val)
    original = val
    if name not in _SKIP_SANITIZE:
        val = val.strip().strip("\"'")
        val = re.sub(r"[^\x21-\x7E]", "", val)
    # Chỉ báo động khi có ký tự "lạ" (không phải ASCII/ký tự điều khiển) hoặc xuống dòng giữa khóa.
    # Dấu cách thường (mật khẩu ứng dụng Gmail) và khoảng trắng/xuống dòng ở hai đầu là bình thường.
    if re.search(r"[^\x20\x09\x0a\x0d\x21-\x7E]", original) or re.search(r"[\r\n]", original.strip(" \t\r\n")):
        SANITIZED_KEYS.add(name)
    return val


def get_secret(name: str, default: str = "") -> str:
    """Trả về giá trị khóa bí mật (đã làm sạch), hoặc `default` nếu không cấu hình ở đâu cả."""
    val = os.environ.get(name, "")
    if val:
        return clean_secret_value(name, val)
    try:
        import streamlit as st
        if name in st.secrets:
            val = str(st.secrets[name])
            if val:
                return clean_secret_value(name, val)
    except Exception:
        pass
    val = _load_env_files().get(name, default) or default
    return clean_secret_value(name, val) if val else val
