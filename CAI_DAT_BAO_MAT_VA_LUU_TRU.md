# Cài đặt khóa bí mật & lưu trữ bền vững

## 1. Khóa bí mật (bắt buộc làm)
Các khóa cũ từng nằm trong code và đã lên GitHub → coi như **đã lộ**. Cần:
1. **Thu hồi và tạo khóa mới**: Google AI Studio (Gemini key), NVIDIA NIM key, Gmail App Password (myaccount.google.com → Bảo mật → Mật khẩu ứng dụng).
2. Điền khóa mới vào `.streamlit/secrets.toml` (chạy local, file này đã bị `.gitignore`) — mẫu ở `.streamlit/secrets.toml.example`.
3. Streamlit Cloud: App → Settings → Secrets → dán cùng nội dung.

Thứ tự code tìm khóa: biến môi trường → `st.secrets` → `.env` (xem `secrets_config.py`).

> Khi dán khóa vào Secrets, dán đúng một dòng, nằm trong dấu nháy kép, không xuống dòng. Khóa lẫn ký tự lạ (xuống dòng, dấu cách không ngắt…) làm Google báo `Illegal header value` và AI không chạy. App đã tự làm sạch các ký tự này và hiện cảnh báo ở thanh bên, nhưng vẫn nên dán lại cho sạch.
Ô nhập khóa trên thanh bên giờ để trống; nếu để trống thì dùng khóa trên máy chủ.

> `.streamlit/secrets.toml` hiện đang chứa **khóa cũ** (để app chạy được cho tới khi anh đổi). Hãy thay bằng khóa mới.

## 2. Lưu dữ liệu bền vững trên Streamlit Cloud (khuyến nghị)
Không bật thì app chạy như cũ nhưng dữ liệu vẫn mất mỗi lần Cloud khởi động lại.
1. GitHub → Settings → Developer settings → Fine-grained tokens → tạo token: **chỉ repo này**, quyền **Contents: Read and write**.
2. Thêm vào Secrets: `GITHUB_TOKEN = "..."` (tùy chọn `GITHUB_REPO`, `GITHUB_DATA_BRANCH`).
3. App tự tạo nhánh `data-store` (nhánh riêng, không làm Cloud redeploy) và lưu: `materials_memory.db`, `Master_Inventory.xlsx`, `Processed_Reports/`.
   Khi hai phía cùng sửa, dữ liệu được **hợp nhất** (không ghi đè); bản local bị thay luôn được sao lưu ở `backup/persist/`.

## 3. Chạy kiểm thử
```bash
python tests/test_persistent_store.py   # lưu trữ bền vững (không dùng mạng)
python tests/test_agent2.py             # Agent 2: đồng bộ an toàn + cảnh báo bố cục Sheet
python tests/benchmark_matcher.py       # đo độ chính xác dò mã Agent 1 trên toàn kho (~15 phút)
python -c "import sku_matching_engine as s; s.run_unit_tests('Master_Inventory/Master_Inventory.xlsx')"
```
