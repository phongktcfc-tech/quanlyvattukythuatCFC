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


def get_secret(name: str, default: str = "") -> str:
    """Trả về giá trị khóa bí mật, hoặc `default` nếu không cấu hình ở đâu cả."""
    val = os.environ.get(name, "")
    if val:
        return val
    try:
        import streamlit as st
        if name in st.secrets:
            val = str(st.secrets[name])
            if val:
                return val
    except Exception:
        pass
    return _load_env_files().get(name, default) or default
