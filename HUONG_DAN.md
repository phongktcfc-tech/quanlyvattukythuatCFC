# 🏭 Hệ Thống Quản Lý & Đối Chiếu Vật Tư Kỹ Thuật Nhà Máy

## Hướng Dẫn Sử Dụng

> [!IMPORTANT]
> Cần có **Google AI API Key** (miễn phí) để sử dụng. Lấy tại: https://aistudio.google.com/apikey

---

## 📂 Cấu trúc thư mục

```
D:\QuanLyVatTuKyThuat\
├── Master_Inventory\          ← Kho vật tư gốc (Master_Inventory.xlsx)
├── Daily_Input\               ← File nhập hằng ngày (tự lưu khi upload)
├── Processed_Reports\         ← Báo cáo đối chiếu đã xử lý
├── .streamlit\config.toml     ← Cấu hình server LAN
├── app.py                     ← Ứng dụng chính
├── requirements.txt           ← Thư viện Python
└── start_server.bat           ← Script khởi động nhanh
```

---

## 🚀 Cách khởi động

### Cách 1: Double-click file bat
Nhấp đúp vào `start_server.bat`

### Cách 2: Chạy lệnh trong terminal
```bash
cd D:\QuanLyVatTuKyThuat
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

### Truy cập
| Từ đâu | Địa chỉ |
|--------|---------|
| Máy chủ (máy này) | `http://localhost:8501` |
| Máy khác trong LAN | `http://<IP_máy_chủ>:8501` |

> [!TIP]
> Xem IP máy chủ bằng lệnh `ipconfig` trong CMD. Tìm dòng **IPv4 Address** (VD: `192.168.1.100`)

---

## 📥 Tab 1: Nhập & Đối Chiếu Vật Tư

Hỗ trợ 3 loại đầu vào:

| Định dạng | Mô tả |
|-----------|-------|
| **Excel** (.xlsx/.xls) | Bảng danh sách vật tư |
| **PDF** | Phiếu nhập kho, hóa đơn (text hoặc scan) |
| **Hình ảnh** (.jpg/.png) | Chụp phiếu giao hàng → Gemini Vision đọc |

**Quy trình:**
1. Upload file → Nhấn "🚀 Bắt đầu đối chiếu"
2. AI tự động: Trích xuất → Chuẩn hóa tên → Đối chiếu kho → Phân loại
3. Kết quả hiển thị bảng + tự động lưu Excel vào `Processed_Reports\`

**Phân loại kết quả:**
- 🟢 **KHỚP** — Đã có mã trong kho, hiển thị tên chuẩn + tồn kho
- 🟡 **MỚI** — Chưa có trong kho, AI đề xuất mã mới
- 🔴 **CẦN_XÁC_NHẬN** — Thiếu thông tin, cần bổ sung

---

## 💬 Tab 2: Hỏi Kỹ Sư AI

Chat trực tiếp với AI có kiến thức chuyên sâu:
- Tra mã vật tư: *"bạc đạn 6205 còn mấy cái?"*
- Hỏi quy cách: *"motor 10 ngựa là bao nhiêu kW?"*
- Tìm tương đương: *"thay thế vòng bi SKF 6308 bằng hãng nào?"*
- Giải tiếng lóng: *"mỡ bò EP2 là gì?"*

---

## 📊 Tab 3: Kho Gốc (Master)

- Xem toàn bộ danh sách vật tư gốc
- Tìm kiếm nhanh (mã, tên, quy cách, tiếng lóng)
- Lọc theo nhóm thiết bị
- Cảnh báo vật tư dưới mức tồn kho tối thiểu
- Upload cập nhật file Master mới (tự động backup file cũ)

---

## 📋 Tab 4: Báo Cáo Đã Xử Lý

- Danh sách tất cả báo cáo đối chiếu đã tạo
- Xem nhanh nội dung + tải xuống Excel

---

## ⚙️ Cấu hình nâng cao

### Mở firewall cho LAN (chạy CMD với quyền Admin):
```cmd
netsh advfirewall firewall add rule name="Streamlit_VatTu_8501" dir=in action=allow protocol=TCP localport=8501
```

### Đổi Model AI
Trong sidebar ứng dụng, chọn model phù hợp:
- `gemini-2.0-flash` — Nhanh, tiết kiệm quota
- `gemini-2.5-flash` — Cân bằng tốc độ/chất lượng  
- `gemini-2.5-pro` — Chính xác nhất
