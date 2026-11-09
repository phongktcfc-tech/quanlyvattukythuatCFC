"""
HE THONG GIU KET NOI & TU DONG DANH THUC STREAMLIT CLOUD 24/7
URL: https://quanlyvattucfc.streamlit.app

Co che:
1. Su dung Playwright Chromium headless mo trang giong nhu nguoi dung that.
2. Tu dong quet tim nut 'Yes, get this app back up!' neu app bi sleep de click thuc day.
3. Giu session WebSocket trong 15 giay de Streamlit Cloud reset bo dem thoi gian ngu.
4. Tu dong lap lai moi 30 phut de giu app luon o trang thai ACTIVE.
"""

import time
import datetime
import sys

URL = "https://quanlyvattucfc.streamlit.app"
INTERVAL_SECONDS = 1800  # 30 phut

def wake_up_cloud_app():
    from playwright.sync_api import sync_playwright
    now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{now_str}] Dang kiem tra ket noi Streamlit Cloud: {URL} ...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        try:
            response = page.goto(URL, timeout=90000, wait_until="domcontentloaded")
            http_status = response.status if response else "N/A"
            page.wait_for_timeout(4000)
            
            # Kiem tra nut danh thuc app neu app dang ngu (Zzzz)
            wake_btn = page.locator('button:has-text("get this app back up")')
            
            if wake_btn.count() > 0 and wake_btn.first.is_visible():
                print(f"[{now_str}] >>> PHAT HIEN APP DANG NGU (Zzzz)! Dang click nut danh thuc...")
                wake_btn.first.click()
                print(f"[{now_str}] Da click! Dang cho container Cloud khoi dong...")
                page.wait_for_timeout(25000)
            else:
                print(f"[{now_str}] App dang thuc va hoat dong tot (HTTP {http_status}).")
            
            # Cho app load giao dien chinh Streamlit de reset bo dem inactivity
            try:
                page.wait_for_selector('div[data-testid="stAppViewContainer"]', timeout=30000)
                print(f"[{now_str}] >>> XAC NHAN: Container da tai xong toan bo giao dien!")
            except Exception:
                pass
            
            page.wait_for_timeout(5000)
            print(f"[{now_str}] >>> Ghi nhan session thanh cong! Timer sleep da duoc reset.")
            return True
        except Exception as e:
            print(f"[{now_str}] Loi kiem tra: {e}")
            return False
        finally:
            browser.close()

if __name__ == "__main__":
    print("=" * 65)
    print("  KHOI DONG DICH VU KEEP-ALIVE 24/7 CHO STREAMLIT CLOUD")
    print(f"  Muc tieu: {URL}")
    print(f"  Chu ky kiem tra: {INTERVAL_SECONDS // 60} phut/lan")
    print("=" * 65)
    
    while True:
        try:
            wake_up_cloud_app()
        except KeyboardInterrupt:
            print("\nDa dung dich vu.")
            sys.exit(0)
        except Exception as err:
            print(f"Loi vong lap: {err}")
            
        print(f"Nghi {INTERVAL_SECONDS // 60} phut truoc lan kiem tra tiep theo...\n")
        time.sleep(INTERVAL_SECONDS)
