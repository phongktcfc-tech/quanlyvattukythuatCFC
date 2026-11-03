"""Tạo file Master_Inventory.xlsx mẫu với dữ liệu vật tư kỹ thuật nhà máy."""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Kho_Vat_Tu"

# ── Header ──
headers = [
    "Mã vật tư", "Tên vật tư (chuẩn)", "Quy cách / Thông số kỹ thuật",
    "Đơn vị tính", "Nhóm thiết bị", "Vị trí lắp đặt / Kho",
    "Tồn kho hiện tại", "Tồn kho tối thiểu", "Nhà cung cấp",
    "Ghi chú / Tên gọi khác (tiếng lóng)"
]

header_font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)

for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border

# ── Dữ liệu mẫu – thiết bị nhà máy Cơ-Điện ──
data = [
    ["VB-6205-2RS", "Vòng bi 6205-2RS", "d=25mm D=52mm B=15mm, 2 mặt phớt cao su", "Cái", "Vòng bi", "Kho A1-01", 120, 30, "SKF / NSK", "bi bạc đạn 6205, bạc đạn 25"],
    ["VB-6308-ZZ", "Vòng bi 6308-ZZ", "d=40mm D=90mm B=23mm, 2 mặt chắn kim loại", "Cái", "Vòng bi", "Kho A1-02", 45, 10, "FAG / NTN", "bạc đạn 6308, bi 40"],
    ["VB-22216-EK", "Vòng bi tang trống 22216 EK", "d=80mm D=140mm B=33mm, lỗ côn 1:12", "Cái", "Vòng bi", "Kho A1-03", 12, 4, "SKF", "bi tang trống 80, bi lồng"],
    ["VB-UC210", "Gối đỡ vòng bi UC210", "Trục d=50mm, gối đứng P210", "Bộ", "Gối đỡ", "Kho A1-04", 18, 6, "ASAHI / FYH", "gối đỡ 50, bộ gối bi"],
    ["DX-3V-B68", "Dây đai thang 3V/B68", "Dài 68 inch, profile 3V/B", "Sợi", "Dây đai", "Kho A2-01", 24, 8, "Optibelt / Gates", "dây curoa B68, curoa 3V"],
    ["DX-5V-2120", "Dây đai thang 5V/2120", "Dài 2120mm, profile 5V", "Sợi", "Dây đai", "Kho A2-02", 10, 4, "Optibelt", "curoa 5V lớn"],
    ["BR-M5-Z24-SCM", "Bánh răng module m=5 Z=24", "m=5 Z=24 D=130mm, vật liệu SCM440, tôi cải thiện HRC28-32", "Cái", "Bánh răng / Hộp số", "Kho B1-01", 6, 2, "Gia công CNC nội bộ", "nhông m5 z24, bánh răng module 5"],
    ["BR-M4-Z36-S45C", "Bánh răng module m=4 Z=36", "m=4 Z=36 D=152mm, vật liệu S45C, HRC25-30", "Cái", "Bánh răng / Hộp số", "Kho B1-02", 4, 2, "Cơ khí Trường Hải", "nhông m4 z36"],
    ["HGT-R77-I30", "Hộp giảm tốc R77 tỷ số i=30", "Trục ra Ø45mm, công suất 5.5kW, tỷ số truyền 30:1", "Bộ", "Hộp giảm tốc", "Kho B2-01", 3, 1, "SEW / Bonfiglioli", "hộp số R77, reducer R77"],
    ["HGT-K67-I25", "Hộp giảm tốc K67 tỷ số i=25", "Trục ra Ø35mm, công suất 3kW, tỷ số truyền 25:1", "Bộ", "Hộp giảm tốc", "Kho B2-02", 2, 1, "SEW", "hộp số K67"],
    ["DC-7R5-4P", "Động cơ điện 7.5kW 4 cực", "7.5kW, 1450rpm, 380V/3P/50Hz, IP55, chân đế B3", "Cái", "Động cơ điện", "Kho C1-01", 5, 2, "ABB / Siemens / WEG", "mô tơ 10 ngựa, motor 7.5kw"],
    ["DC-5R5-4P", "Động cơ điện 5.5kW 4 cực", "5.5kW, 1450rpm, 380V/3P/50Hz, IP55, mặt bích B5", "Cái", "Động cơ điện", "Kho C1-02", 4, 2, "ABB / Siemens", "mô tơ 7.5 ngựa, motor 5.5kw"],
    ["BT-B800-EP300", "Băng tải cao su B=800mm EP300/3", "Rộng 800mm, 3 lớp mành EP300, dày 12mm (6+3+3)", "Mét", "Băng tải", "Kho D1-01", 150, 50, "Yokohama / Bridgestone", "băng cao su 800, dây băng 800"],
    ["BT-B650-EP200", "Băng tải cao su B=650mm EP200/2", "Rộng 650mm, 2 lớp mành EP200, dày 8mm (4+2+2)", "Mét", "Băng tải", "Kho D1-02", 80, 30, "Bando", "băng 650, dây băng nhỏ"],
    ["RO-108-L600", "Con lăn Ø108 x 600mm", "Ống thép Ø108mm, dài 600mm, trục Ø20, vòng bi 6204", "Cái", "Con lăn", "Kho D2-01", 30, 10, "Nội bộ / CTy TNHH ABC", "ru lô 108, roller 600"],
    ["XT-M16-150-GR88", "Bu lông M16x150 Gr.8.8", "M16x150mm, lục giác, Grade 8.8, mạ kẽm", "Cái", "Bu lông / Ốc vít", "Kho E1-01", 200, 50, "Đại lý bulong HCM", "bu lông 16 dài 15, bulong M16"],
    ["XT-M20-100-GR109", "Bu lông M20x100 Gr.10.9", "M20x100mm, lục giác, Grade 10.9, đen", "Cái", "Bu lông / Ốc vít", "Kho E1-02", 80, 20, "Đại lý bulong HCM", "bulong 20 cường độ cao"],
    ["SR-50x70x10-NBR", "Phớt chặn dầu 50x70x10 NBR", "Trục 50mm, ngoài 70mm, rộng 10mm, chất liệu NBR", "Cái", "Phớt / Gioăng", "Kho F1-01", 40, 10, "NOK / TTO", "phớt 50, seal 50x70"],
    ["GK-VITON-DN80", "Gioăng chịu nhiệt Viton DN80", "DN80, chịu nhiệt 250°C, vật liệu FKM/Viton", "Cái", "Phớt / Gioăng", "Kho F1-02", 15, 5, "Garlock / Klinger", "ron viton 80, gioăng nhiệt"],
    ["BM-EP2-15KG", "Mỡ bôi trơn EP2 xô 15kg", "NLGI Grade 2, gốc Lithium, chịu tải EP", "Xô", "Dầu mỡ", "Kho G1-01", 8, 3, "Shell Alvania / SKF LGMT2", "mỡ bò EP2, mỡ vàng"],
    ["TL-GLAND-M25", "Ốc siết cáp M25 (cable gland)", "Đồng thau mạ Ni, kẹp cáp Ø13-18mm, IP68", "Cái", "Tủ điện / Phụ kiện", "Kho H1-01", 50, 20, "Hawke / Hummel", "rắc co M25, ốc xiết cáp"],
    ["TL-ACB-3P-800A", "Aptomat khối MCCB 3P 800A", "3 cực, 800A, Icu=50kA, điện áp 415V", "Cái", "Tủ điện / Aptomat", "Kho H2-01", 2, 1, "Schneider / LS", "CB 800A, áp tô mát 800"],
    ["TL-CONTACTOR-LC1D65", "Contactor LC1D65M7", "65A, cuộn coil 220VAC, 3P+1NO+1NC", "Cái", "Tủ điện / Contactor", "Kho H2-02", 6, 2, "Schneider", "khởi động từ 65A, công tắc tơ"],
    ["TL-OLR-48-65A", "Rơ le nhiệt 48-65A", "Dải chỉnh 48-65A, lắp contactor LC1D65", "Cái", "Tủ điện / Rơ le", "Kho H2-03", 4, 2, "Schneider LRD3361", "role nhiệt 65A, thermal relay"],
    ["SL-LUOI-5x5-INOX304", "Lưới sàng 5x5mm Inox 304", "Lỗ vuông 5x5mm, dây Ø1.2mm, Inox SUS304, khổ 1.5m", "m²", "Sàng rung", "Kho I1-01", 20, 5, "Nhập TQ / Hàn Quốc", "lưới rây 5ly, lưới sàng 5mm"],
    ["SL-LUOI-10x10-MN", "Lưới sàng 10x10mm thép Mn", "Lỗ vuông 10x10mm, dây Ø3mm, thép Mn chịu mài mòn", "m²", "Sàng rung", "Kho I1-02", 15, 5, "Nội bộ", "lưới sàng 10ly, lưới thép mangan"],
    ["LX-SPRING-SR3624", "Lò xo sàng rung SR3624", "OD=120mm, ID=60mm, H=200mm, tải 5 tấn, thép SUP9", "Cái", "Sàng rung", "Kho I2-01", 16, 4, "KYB / Nội bộ", "lò xo sàng, nhíp sàng rung"],
    ["PK-CERAMICS-150x100", "Tấm lót ceramic 150x100x25mm", "Al₂O₃ 92%, chịu mài mòn, dán epoxy", "Tấm", "Phụ kiện chống mòn", "Kho J1-01", 200, 50, "Nhập TQ / Đức", "gạch ceramic, tấm chống mòn"],
    ["XL-CAO-SU-NR-5MM", "Tấm cao su lót NR 5mm", "Cao su tự nhiên NR, dày 5mm, khổ 1.2m, Shore A 60±5", "m²", "Phụ kiện chống mòn", "Kho J1-02", 30, 10, "Nội bộ", "cao su lót, tấm su đen"],
    ["PK-HAN-QUE-E7016-4MM", "Que hàn E7016 Ø4.0mm", "AWS E7016, Ø4.0mm, hàn thép CT3-SCM, hộp 5kg", "Hộp", "Vật tư hàn", "Kho K1-01", 25, 10, "Kiswel / Hyundai", "que hàn 4 ly, que hàn xanh"],
]

for r, row in enumerate(data, 2):
    for c, val in enumerate(row, 1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center", wrap_text=True)

# ── Điều chỉnh độ rộng cột ──
col_widths = [16, 32, 48, 10, 20, 18, 14, 14, 24, 36]
for i, w in enumerate(col_widths, 1):
    ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

ws.auto_filter.ref = f"A1:J{len(data)+1}"
ws.freeze_panes = "A2"

wb.save("D:\\QuanLyVatTuKyThuat\\Master_Inventory\\Master_Inventory.xlsx")
print("✅ Đã tạo Master_Inventory.xlsx thành công!")
