"""Test làm sạch khóa bí mật. Chạy: python tests/test_secrets_config.py"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import secrets_config as sc


def main():
    k = "AQ.Ab8Rtest-key_123"
    # Khóa sạch: giữ nguyên, không báo động
    sc.SANITIZED_KEYS.clear()
    assert sc.clean_secret_value("GOOGLE_API_KEY", k) == k
    assert not sc.SANITIZED_KEYS

    # Các ký tự lạ hay lẫn khi copy-dán → phải loại và báo động
    bad = {
        "xuống dòng cuối": k + "\n",
        "xuống dòng giữa": k[:5] + "\r\n" + k[5:],
        "NBSP": " " + k,
        "zero-width": k[:6] + "​" + k[6:],
        "nháy thông minh": "“" + k + "”",
        "dấu nháy thường": '"' + k + '"',
        "tab": k + "\t",
    }
    for label, v in bad.items():
        sc.SANITIZED_KEYS.clear()
        out = sc.clean_secret_value("GOOGLE_API_KEY", v)
        assert out == k, f"{label}: ra {out!r}"
        assert out.isascii() and out.isprintable() and " " not in out
    # Riêng "dấu nháy thường"/xuống dòng/khoảng trắng thông thường ở hai đầu không phải ký tự lạ → không báo động
    sc.SANITIZED_KEYS.clear(); sc.clean_secret_value("GOOGLE_API_KEY", "  " + k + "\n")
    assert not sc.SANITIZED_KEYS
    sc.SANITIZED_KEYS.clear(); sc.clean_secret_value("GOOGLE_API_KEY", " " + k)
    assert "GOOGLE_API_KEY" in sc.SANITIZED_KEYS

    # Mật khẩu ứng dụng Gmail dạng 'abcd efgh ijkl mnop' → bỏ dấu cách, không báo động
    sc.SANITIZED_KEYS.clear()
    assert sc.clean_secret_value("EMAIL_APP_PASSWORD", "abcd efgh ijkl mnop") == "abcdefghijklmnop"
    assert not sc.SANITIZED_KEYS

    # SMTP_SERVER không bị đụng
    assert sc.clean_secret_value("SMTP_SERVER", "smtp.gmail.com") == "smtp.gmail.com"
    print("TẤT CẢ TEST LÀM SẠCH KHÓA ĐẠT")


if __name__ == "__main__":
    main()
