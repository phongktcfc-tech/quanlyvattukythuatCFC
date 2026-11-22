"""Test lớp 'ngôn ngữ của thợ' (fi 10 = 100mm, ly = mm, tên chuẩn hóa AI) + cổng. Chạy: python tests/test_worker_language.py"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import worker_language as w
from auto_match_gate import decide_auto_match, extract_spec_numbers
from sku_matching_engine import UniversalTechnicalSKUMatcher as M
import tests.benchmark_matcher as b


def main():
    # 1. Chuẩn hóa cách gọi
    assert w.canon("Đá cắt fi 10") == "đá cắt 100mm"
    assert w.canon("Đá mài fi 400 (thô)") == "đá mài 400mm (thô)"
    assert w.canon("Đá cắt fi 300") == "đá cắt 300mm"
    assert w.canon("M.Khoan 8,5 ly") == "mũi khoan 8.5mm"
    assert w.canon("Bulong phi 10") == "bulong phi 10", "phi 10 của bulong phải GIỮ NGUYÊN (mm), chỉ nhóm đá/đĩa mới quy đổi cm"
    assert w.canon("Mũi khoan thép gió 8.5mm") == "mũi khoan thép 8.5mm"
    assert extract_spec_numbers("Đá cắt fi 10") == extract_spec_numbers("Đá cắt sắt 100mm") == {"100"}

    # 2. Gate không bị nới lỏng: lệch thông số thật vẫn bị chặn
    best = {"ten": "Bạc đạn 6205", "ma": "03.1", "score": 80.0, "tier": "GENERIC_MATCH", "reasons": []}
    assert decide_auto_match("Bạc đạn 6206", best)[0] == "CẦN_XÁC_NHẬN"
    best = {"ten": "Đá cắt fi 10", "ma": "02.1284", "score": 55.0, "tier": "SIMILAR_MISMATCH", "reasons": []}
    assert decide_auto_match("Đá cắt sắt 100mm", best)[0] == "CẦN_XÁC_NHẬN", "ca gần giống phải để kỹ sư xem, không cấp mã mới"
    assert decide_auto_match("Đá cắt sắt 150mm", best)[0] == "MỚI", "khác cỡ thì đúng là vật tư mới"
    best = {"ten": "Mũi đục fi 30", "ma": "03.7628", "score": 50.0, "tier": "SIMILAR_MISMATCH", "reasons": []}
    assert decide_auto_match("Phe gài trục D30", best)[0] == "MỚI"

    # 3. Trên kho thật
    df = b.load_master()
    prim = M(df)

    def primary(q):
        r = prim.match_item(q)
        return [{"score": c["score"], "ma": c["ma_hang"], "ten": c["ten_hang"], "tier": c.get("tier"), "ton": c["cuoi_ky_sl"],
                 "ly_do": "", "reasons": c.get("reasons", [])} for c in r["top_candidates"]]

    def run(t_goc, t_chuan, q_cach):
        p = primary(f"{t_goc} {q_cach}".strip())
        c, _ = w.resolve_candidates(p, t_goc, t_chuan, q_cach, df, M)
        return c[0]

    c = run("Mũi khoan Nachi 8.5mm", "Mũi khoan thép gió", "8.5mm")
    assert c["ma"] == "03.9008" and c["tier"] == "EXACT_MATCH", c
    assert decide_auto_match("Mũi khoan Nachi 8.5mm", c, strict_exact=True)[0] == "KHỚP"

    c = run("Đá cắt sắt", "Đá cắt sắt", "100mm")
    assert c["ma"] == "02.1284", c
    assert c["ten"] == "Đá cắt fi 10", "phải hiển thị tên GỐC trong kho"
    assert decide_auto_match("Đá cắt sắt 100mm", c, strict_exact=True)[0] in ("KHỚP", "CẦN_XÁC_NHẬN")

    # Những vật tư thật sự không có trong kho vẫn phải là MỚI, không bị kéo nhầm
    for t_goc, t_chuan, q in [("Thước cặp mitutoyo 200mm", "Thước cặp cơ khí", "200mm"), ("Phe gài D30", "Phe gài trục", "D30"),
                              ("Xích 60-1R", "Xích công nghiệp", "60-1R (1 dây)"), ("Mũi khoan 9.5mm", "Mũi khoan thép", "9.5mm")]:
        c = run(t_goc, t_chuan, q)
        d = decide_auto_match(f"{t_goc} {q}", c, strict_exact=True)[0] if c else "MỚI"
        assert d != "KHỚP", f"{t_goc}: bị khớp nhầm về {c.get('ma')} {c.get('ten')}"
    print("TẤT CẢ TEST NGÔN NGỮ CỦA THỢ ĐẠT")


if __name__ == "__main__":
    main()
