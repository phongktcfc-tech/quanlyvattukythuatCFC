"""Test cập nhật kho từ file ERP (3 chế độ + bảo vệ). Chạy: python tests/test_master_update.py"""
import io
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT)); os.chdir(ROOT)

import pandas as pd
import app


def erp_file(rows):
    """Dựng file giống báo cáo 'TỔNG HỢP TỒN KHO' của ERP: tiêu đề, header 2 tầng, dòng tổng, dòng dữ liệu, dòng 'Số dòng'."""
    data = [["TỔNG HỢP TỒN KHO"] + [None] * 12, ["Kho: Kho vật tư phòng Kỹ thuật 01"] + [None] * 12,
            [None, "Tên kho", "Mã hàng", "Tên hàng", "ĐVT", "Đầu kỳ", None, "Nhập kho", None, "Xuất kho", None, "Cuối kỳ", None],
            [None] * 5 + ["Số lượng", "Giá trị"] * 4,
            ["Tên kho : Kho vật tư phòng Kỹ thuật 01"] + [None] * 4 + [0] * 8]
    for code, name, sl in rows:
        data.append([None, "Kho vật tư phòng Kỹ thuật 01", code, name, "Cái", sl, sl * 1000, 0, 0, 0, 0, sl, sl * 1000])
    data.append([None, f"Số dòng = {len(rows)}"] + [None] * 11)
    buf = io.BytesIO(); pd.DataFrame(data).to_excel(buf, header=False, index=False)
    return buf.getvalue()


def master_df(rows):
    return pd.DataFrame([{"Tên kho": "Kho vật tư phòng Kỹ thuật 01", "Mã hàng": c, "Tên hàng": n, "ĐVT": "Cái",
                          "Đầu kỳ SL": s, "Đầu kỳ GT": s * 10, "Nhập kho SL": 0, "Nhập kho GT": 0, "Xuất kho SL": 0,
                          "Xuất kho GT": 0, "Cuối kỳ SL": s, "Cuối kỳ GT": s * 10} for c, n, s in rows])


def run(master_rows, file_rows, mode):
    with tempfile.TemporaryDirectory() as t:
        mf = Path(t) / "master.xlsx"
        master_df(master_rows).to_excel(mf, index=False)
        app.MASTER_FILE = mf
        app.get_all_approved_materials = lambda: []          # không lẫn mã đã duyệt trong DB thật
        ok, msg = app.update_master_inventory(erp_file(file_rows), "erp.xlsx", mode=mode)
        assert ok, msg
        m = pd.read_excel(mf, dtype={"Mã hàng": str}).set_index("Mã hàng")
        return m, msg


def main():
    master = [("03.0001", "Van bi", 5), ("03.0002", "Ống đồng", 3), ("03.0003", "Dây điện", 0), ("03.0004", "Bạc đạn 6205", 7),
              ("03.0005", "Contactor", 4), ("03.0006", "Aptomat", 6)]
    erp = [("03.0001", "Van bi", 9), ("03.0002", "Ống đồng", 3), ("03.0004", "Bạc đạn 6205", 7), ("03.0005", "Contactor", 4),
           ("03.0007", "Mã mới ERP", 2)]

    # sync: giữ danh mục, mã vắng trong file → tồn 0, mã trong file cập nhật, mã mới được thêm
    m, msg = run(master, erp, "sync")
    assert set(m.index) == {c for c, _, _ in master} | {"03.0007"}, "sync làm mất hoặc thừa mã"
    assert m.loc["03.0001", "Cuối kỳ SL"] == 9 and m.loc["03.0007", "Cuối kỳ SL"] == 2
    assert m.loc["03.0006", "Cuối kỳ SL"] == 0 and m.loc["03.0006", "Cuối kỳ GT"] == 0, "mã vắng trong file chưa về 0"
    assert m.loc["03.0003", "Tên hàng"] == "Dây điện", "danh mục bị mất tên"
    assert "đặt tồn = 0" in msg and "4 mã có tồn" not in msg

    # replace: chỉ còn đúng các mã trong file
    m, _ = run(master, erp, "replace")
    assert set(m.index) == {c for c, _, _ in erp}

    # merge: hành vi cũ, mã vắng trong file giữ nguyên số liệu
    m, _ = run(master, erp, "merge")
    assert m.loc["03.0006", "Cuối kỳ SL"] == 6

    # bảo vệ: file nhỏ hơn nhiều so với kho → sync tự lùi về merge, KHÔNG đặt tồn = 0
    big = [(f"03.{i:04d}", f"Vật tư {i}", 5) for i in range(1, 101)]
    m, msg = run(big, [("03.0001", "Vật tư 1", 8), ("03.0002", "Vật tư 2", 5)], "sync")
    assert m.loc["03.0001", "Cuối kỳ SL"] == 8 and m.loc["03.0050", "Cuối kỳ SL"] == 5, "bảo vệ không hoạt động"
    assert "KHÔNG đặt tồn = 0" in msg

    # mã trùng trong kho cũ được loại
    dup = master + [("03.0001", "Van bi (trùng)", 1)]
    m, _ = run(dup, erp, "sync")
    assert not m.index.duplicated().any()

    # đếm mã có tồn
    assert app.count_in_stock(pd.DataFrame({"Cuối kỳ SL": [0, 1, 2.5, None]})) == 2
    print("TẤT CẢ TEST CẬP NHẬT KHO ĐẠT")


if __name__ == "__main__":
    main()
