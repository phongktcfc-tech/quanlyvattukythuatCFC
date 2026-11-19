"""Test cổng chống khớp nhầm thông số. Chạy: python tests/test_auto_match_gate.py"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from auto_match_gate import decide_auto_match, extract_spec_numbers


def cand(ten, score=90.0, tier="EXACT_MATCH", reasons=None, ma="X"):
    return {"ma": ma, "ten": ten, "score": score, "tier": tier, "reasons": reasons or []}


def main():
    # Số thông số: bỏ số dính chữ (2RS, NN102WM), giữ số kèm đơn vị
    assert extract_spec_numbers("Bạc đạn 6205-2RS") == {"6205"}
    assert extract_spec_numbers("Động cơ 7,5kW 380V") == {"7.5", "380"}
    assert extract_spec_numbers("Hộp nối đôi Nanoco NN102WM") == set()

    s = lambda q, b: decide_auto_match(q, b, strict_exact=True)[0]

    # Đúng thông số → KHỚP
    assert s("Thép I 200 (6m/cây)", cand("Thép I 200 (6m/cây)")) == "KHỚP"
    assert s("Bạc đạn 6205 2RS", cand("Bạc đạn 6205-2RS")) == "KHỚP"
    # Lệch thông số, dù điểm cao → CẦN_XÁC_NHẬN (các ca thật từ bộ đo)
    assert s("Thép I 202 (6m/cây)", cand("Thép I 200 (6m/cây)")) == "CẦN_XÁC_NHẬN"
    assert s("Dung môi Thinner 25", cand("Dung môi Thinner 024", 95.0)) == "CẦN_XÁC_NHẬN"
    assert s("Rơle 28 - 25A", cand("Rơle 18 - 25A", 85.0)) == "CẦN_XÁC_NHẬN"
    # Chế độ thường không chặn tầng EXACT
    assert decide_auto_match("Thép I 202", cand("Thép I 200"))[0] == "KHỚP"
    # GENERIC (65-84): chặn khi lệch số, cho khi đủ số + có từ khóa chung
    g = lambda ten, **k: cand(ten, 70.0, "GENERIC_MATCH", **k)
    assert decide_auto_match("Van hơi DN 26 phi 34", g("Van hơi DN 25 phi 34"))[0] == "CẦN_XÁC_NHẬN"
    assert decide_auto_match("Van hơi DN 25 phi 34 inox", g("Van hơi DN 25 phi 34"))[0] == "KHỚP"
    # Có cảnh báo xung đột → không bao giờ tự khớp
    assert s("A 10", cand("A 10", 95.0, reasons=["⚠️ Khác thông số"])) == "CẦN_XÁC_NHẬN"
    # Không có ứng viên / điểm thấp → MỚI
    assert s("abc", None) == "MỚI"
    assert s("abc", cand("xyz", 30.0, "SIMILAR_MISMATCH")) == "MỚI"
    # Mã do AI tự chọn: phải có trong kho và không lệch thông số
    from auto_match_gate import check_ai_claimed_match
    names = {"03.0372": "Thép I 200 (6m/cây)", "02.1298": "Biến dòng 300/5A", "03.2035": "Băng keo điện"}
    assert check_ai_claimed_match("Thép I 200 (6m/cây)", "03.0372", names)[0]
    assert not check_ai_claimed_match("Thép I 202 (6m/cây)", "03.0372", names)[0]      # AI "sửa lỗi đánh máy"
    assert not check_ai_claimed_match("Biến dòng 301/5A", "02.1298", names)[0]
    assert check_ai_claimed_match("Băng keo điện bản rộng 5cm (20m/cuộn)", "03.2035", names)[0]  # kho tên chung, không có số
    assert not check_ai_claimed_match("Bạc đạn 6205", "99.9999", names)[0]               # mã AI bịa
    print("TẤT CẢ TEST CỔNG CHỐNG KHỚP NHẦM ĐẠT")


if __name__ == "__main__":
    main()
