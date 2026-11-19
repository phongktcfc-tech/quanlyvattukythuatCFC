"""
Chạy 25 test engine bị ghim (sku_matching_engine.run_unit_tests) trên BẢN KHO CHUẨN cũ.

Các test đó mã hóa kỳ vọng theo bản kho tại thời điểm 03/11/2026. Từ khi kho được đồng bộ với file ERP mới,
một số vật tư từng "chưa có trong kho" (kỳ vọng MỚI) nay đã có mã chính thức trong ERP nên engine khớp đúng —
vì vậy chạy thẳng trên kho hiện tại sẽ báo FAIL hợp lý. File này giữ bộ test có giá trị bằng dữ liệu chuẩn cố định.
Chạy: python tests/test_engine_baseline.py
"""
import contextlib
import io
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import sku_matching_engine as s

BASELINE = ROOT / "tests" / "fixtures" / "Master_Inventory_baseline_20261103.xlsx"


def main():
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        s.run_unit_tests(str(BASELINE))
    out = buf.getvalue()
    assert "FAIL" not in out and "25/25" in out, out[-1500:]
    print("TẤT CẢ 25 TEST ENGINE ĐẠT TRÊN KHO CHUẨN")


if __name__ == "__main__":
    main()
