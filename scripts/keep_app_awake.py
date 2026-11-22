"""Mở app Streamlit bằng trình duyệt ẩn như người dùng thật; nếu app đang ngủ thì bấm nút đánh thức.

Chạy: python scripts/keep_app_awake.py [URL]
Mã thoát: 0 = app đã chạy, 1 = không đánh thức/tải được app.
"""
import sys
import time

from playwright.sync_api import sync_playwright

APP_URL = sys.argv[1] if len(sys.argv) > 1 else "https://quanlyvattucfc.streamlit.app"
WAKE_BUTTON_TEXTS = ("get this app back up", "wake up")
LOAD_TIMEOUT_S = 240
ATTEMPTS = 3


def app_is_running(page) -> bool:
    """App chạy khi trang có khung ứng dụng Streamlit hoặc thân app hiện tiêu đề."""
    if page.locator('[data-testid="stApp"]').count() > 0:
        return True
    for frame in page.frames:
        try:
            if frame.locator('[data-testid="stApp"]').count() > 0:
                return True
        except Exception:
            continue
    return False


def click_wake_button(page) -> bool:
    for frame in [page, *page.frames]:
        for text in WAKE_BUTTON_TEXTS:
            try:
                btn = frame.get_by_role("button", name=text, exact=False)
                if btn.count() > 0:
                    btn.first.click(timeout=5000)
                    print(f"[keep-alive] Đã bấm nút đánh thức ({text!r})")
                    return True
            except Exception:
                continue
    return False


def try_once(browser) -> bool:
    page = browser.new_page()
    page.goto(APP_URL, wait_until="domcontentloaded", timeout=60000)
    deadline = time.time() + LOAD_TIMEOUT_S
    woke = False
    while time.time() < deadline:
        if app_is_running(page):
            time.sleep(20)  # giữ kết nối vài giây để Streamlit ghi nhận có người truy cập
            print(f"[keep-alive] App đang chạy (đã phải đánh thức: {woke})")
            return True
        if click_wake_button(page):
            woke = True
        time.sleep(5)
    print("[keep-alive] Hết thời gian chờ, app chưa lên")
    page.close()
    return False


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            for i in range(1, ATTEMPTS + 1):
                print(f"[keep-alive] Lần thử {i}/{ATTEMPTS}: {APP_URL}")
                if try_once(browser):
                    return 0
        finally:
            browser.close()
    return 1


if __name__ == "__main__":
    sys.exit(main())
