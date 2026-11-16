# Bản vá `auto_match_gate` cho Agent 1 — ĐÃ ÁP DỤNG (chủ dự án duyệt 16/11/2026)

Vị trí nằm trong vùng `🔒 [FROZEN CORE START]` của `app.py` (hàm `run_agent1_background_worker`).
Đã áp dụng với `strict_exact=True`. Số liệu đo toàn kho (828 ca đổi thông số, `tests/benchmark_matcher.py`):
khớp nhầm 8,6% → 0,1%; 12,6% số ca chuyển sang CẦN_XÁC_NHẬN để kỹ sư chọn tay.
Tài liệu dưới đây giữ lại để đối chiếu; bản áp dụng thực tế dùng `decide_auto_match(search_str, best, strict_exact=True)`
và thêm điều kiện `_auto_decision == "KHỚP"` cho nhánh EXACT.

## 1. Thêm import (đầu `app.py`, cạnh `from sku_matching_engine import ...`)
```python
from auto_match_gate import decide_auto_match
```

## 2. Thay đoạn quyết định tự khớp (trong vùng FROZEN, ngay sau dòng `_search_norm = ...`)

Đoạn gốc:
```python
_is_generic_ok = (
    best and best.get("tier") == "GENERIC_MATCH" and best.get("score", 0) >= 65
    and any(kw in _best_ten_norm for kw in _search_norm.split() if len(kw) >= 3)
)
if best and ((best.get("tier") == "EXACT_MATCH" and best.get("score", 0) >= 85) or _is_generic_ok):
```

Đoạn thay:
```python
_auto_decision, _auto_note = decide_auto_match(search_str, best)
_is_generic_ok = (
    best and best.get("tier") == "GENERIC_MATCH" and best.get("score", 0) >= 65
    and any(kw in _best_ten_norm for kw in _search_norm.split() if len(kw) >= 3)
    and _auto_decision == "KHỚP"
)
if best and _auto_decision == "CẦN_XÁC_NHẬN" and not (r.get("trang_thai") == "KHỚP" and r.get("ma_vat_tu_cu") and str(r.get("ma_vat_tu_cu")).strip() not in ["null", "-", "None", ""]):
    # Có ứng viên gần giống nhưng lệch thông số → để kỹ sư chọn, KHÔNG tự cấp mã mới, KHÔNG tự khớp
    r["trang_thai"] = "CẦN_XÁC_NHẬN"
    r["ma_vat_tu_cu"] = None
    r["ten_trong_kho"] = "-"
    r["ton_kho"] = 0
    r["ma_de_xuat"] = "-"
    r["ghi_chu"] = f"⚠️ CẦN XÁC NHẬN: {_auto_note}"
elif best and ((best.get("tier") == "EXACT_MATCH" and best.get("score", 0) >= 85) or _is_generic_ok):
```
(phần `elif ... else` phía sau giữ nguyên.)

## Giao diện đã sẵn sàng
Trạng thái `CẦN_XÁC_NHẬN` đã có sẵn trong giao diện: bộ lọc riêng, đếm số dòng, hộp chọn 5 mã gợi ý + nút "Áp dụng",
và nút "Duyệt mã mới" không động tới các dòng này (nên không tự cấp mã nhầm).
