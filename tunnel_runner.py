import subprocess
import re
import time
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
CLOUDFLARED_EXE = BASE_DIR / "cloudflared.exe"
PUBLIC_URL_FILE = BASE_DIR / "public_url.txt"

def run_tunnel():
    if not CLOUDFLARED_EXE.exists():
        print(f"Lỗi: Không tìm thấy {CLOUDFLARED_EXE}")
        return

    cmd = [str(CLOUDFLARED_EXE), "tunnel", "--url", "http://localhost:8501"]
    print("Đang khởi động Cloudflare Quick Tunnel...")
    
    process = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        encoding="utf-8",
        errors="replace"
    )

    url_found = False
    url_pattern = re.compile(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com")

    for line in iter(process.stdout.readline, ""):
        line_clean = line.strip()
        if line_clean:
            # print(line_clean)
            match = url_pattern.search(line_clean)
            if match:
                url = match.group(0)
                PUBLIC_URL_FILE.write_text(url, encoding="utf-8")
                print("\n" + "=" * 60)
                print(f"🚀 PUBLIC INTERNET URL ĐÃ SẴN SÀNG:")
                print(f"👉 {url}")
                print("=" * 60 + "\n")
                url_found = True
                sys.stdout.flush()

    process.wait()

if __name__ == "__main__":
    while True:
        try:
            run_tunnel()
        except Exception as e:
            print(f"Tunnel error: {e}")
        print("Đang thử kết nối lại Cloudflare Tunnel sau 5 giây...")
        time.sleep(5)
