"""Test Agent 2: đồng bộ an toàn + cảnh báo bố cục Sheet. Chạy: python tests/test_agent2.py (không dùng mạng thật)."""
import io
import sys
import tempfile
import urllib.request
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import openpyxl
import agent2_manager as a2

REAL = Path(__file__).resolve().parent.parent / "Agent2_Data" / "Kho_CoDien_Live.xlsx"


class FakeResp(io.BytesIO):
    def __enter__(self): return self
    def __exit__(self, *a): return False


def main():
    # 1. Bố cục thật của Sheet phải sạch cảnh báo
    wb = openpyxl.load_workbook(REAL, data_only=True)
    assert a2.validate_sheet_layout(wb) == [], f"Sheet thật bị báo lỗi: {a2.validate_sheet_layout(wb)}"

    # 2. Chèn 1 cột vào Tổng Hợp → phải phát hiện
    ws = wb["Tổng Hợp Vật Tư"]
    ws.insert_cols(3)
    warns = a2.validate_sheet_layout(wb)
    assert warns and any("Tổng Hợp Vật Tư" in w for w in warns), "không phát hiện chèn cột"

    # 3. Thiếu hẳn sheet tổng hợp → phải cảnh báo
    wb2 = openpyxl.Workbook(); wb2.active.title = "Khác"
    assert any("Không tìm thấy" in w for w in a2.validate_sheet_layout(wb2))

    # 4. Đồng bộ: file cũ tốt KHÔNG bị hỏng khi tải về trang HTML / dữ liệu cụt / lỗi mạng
    with tempfile.TemporaryDirectory() as t:
        target = Path(t) / "live.xlsx"
        good = REAL.read_bytes()
        target.write_bytes(good)
        orig = urllib.request.urlopen
        try:
            urllib.request.urlopen = lambda *a, **k: FakeResp(b"<html>login</html>" * 100)
            ok, msg = a2.sync_live_google_sheet("x", target)
            assert not ok and target.read_bytes() == good, "HTML làm hỏng file cũ"

            urllib.request.urlopen = lambda *a, **k: FakeResp(b"PK" + b"\0" * 10)
            ok, _ = a2.sync_live_google_sheet("x", target)
            assert not ok and target.read_bytes() == good, "file cụt làm hỏng file cũ"

            def boom(*a, **k): raise OSError("mất mạng")
            urllib.request.urlopen = boom
            ok, _ = a2.sync_live_google_sheet("x", target)
            assert not ok and target.read_bytes() == good, "lỗi mạng làm hỏng file cũ"
            assert not Path(str(target) + ".download").exists(), "còn sót file tạm"

            new = good + b"\0"  # vẫn bắt đầu bằng PK, đủ lớn → được thay
            urllib.request.urlopen = lambda *a, **k: FakeResp(new)
            ok, _ = a2.sync_live_google_sheet("x", target)
            assert ok and target.read_bytes() == new, "không thay file khi tải hợp lệ"
        finally:
            urllib.request.urlopen = orig

    # 5. Trạng thái kho: hai hàm phải cho kết quả giống nhau trên mọi tổ hợp
    import pandas as pd
    cases = [(0, 5), (-1, 5), (3, 5), (5, 5), (6, 5), (10, None), (0, None), (2, 0)]
    for ton, mn in cases:
        s1 = a2.compute_item_status({"Tồn Kho Thực Tế": ton, "Định Mức Tồn Tối Thiểu": mn})[0]
        df = a2.process_inventory_dataframe(pd.DataFrame([{"Tồn Kho Thực Tế": ton, "Định Mức Tồn Tối Thiểu": mn}]))
        assert df["status_code"].iloc[0] == s1, f"hai hàm trạng thái lệch ở {ton},{mn}"

    print("TẤT CẢ TEST AGENT 2 ĐẠT")


if __name__ == "__main__":
    main()
