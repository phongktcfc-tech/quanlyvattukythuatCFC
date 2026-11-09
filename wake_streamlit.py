import os
import sys
import time
from playwright.sync_api import sync_playwright

APP_URL = os.environ.get("STREAMLIT_APP_URL", "https://quanlyvattucfc.streamlit.app")

def main():
    print(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] Dang ket noi toi Streamlit Cloud: {APP_URL}")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        try:
            response = page.goto(APP_URL, timeout=90000, wait_until="domcontentloaded")
            print(f"Status HTTP: {response.status if response else 'N/A'}")
            
            # Doi 5 giay de trang render
            page.wait_for_timeout(5000)
            
            title = page.title()
            print(f"Tieu de trang: {title}")
            
            wake_btn = page.locator('button:has-text("get this app back up")')
            
            if wake_btn.count() > 0 and wake_btn.first.is_visible():
                print(">>> PHAT HIEN APP DANG NGU (Zzzz)! Dang click nut danh thuc...")
                wake_btn.first.click()
                print("Da click nut danh thuc! Cho container khoi dong lai...")
                page.wait_for_timeout(25000)
            else:
                print("App dang thuc hoat dong binh thuong, khong co man hinh sleep.")
            
            # Cho app load giao dien chinh Streamlit de reset bo dem inactivity
            try:
                page.wait_for_selector('div[data-testid="stAppViewContainer"]', timeout=30000)
                print(">>> XAC NHAN: Container Streamlit da tai xong toan bo giao dien!")
            except Exception as e:
                print(f"Giao dien chinh chua xuat hien ngay: {e}")
            
            # Giu them 5s de websocket gui heartbeat
            page.wait_for_timeout(5000)
            print(">>> HOAN TAT: Session da duoc ghi nhan thanh cong! Streamlit Cloud reset timer sleep.")
            
        except Exception as e:
            print(f"Loi khi truy cap: {e}")
            sys.exit(1)
        finally:
            browser.close()

if __name__ == "__main__":
    main()
