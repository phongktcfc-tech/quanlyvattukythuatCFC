# -*- coding: utf-8 -*-
"""
CỔNG CHỐNG KHỚP NHẦM THÔNG SỐ CHO AGENT 1 (module độc lập, KHÔNG nằm trong vùng FROZEN CORE)

Mục đích: quyết định có được TỰ ĐỘNG KHỚP ứng viên tốt nhất của sku_matching_engine hay không.
Module này chỉ ĐỌC kết quả của engine, không sửa engine. app.py hiện CHƯA dùng module này
(xem patches/auto_match_gate_app.patch.md — chờ chủ dự án duyệt vì vùng gọi đang bị ghim FROZEN).
"""
import re
from typing import Any, Dict, Optional, Tuple

from worker_language import canon

_SPEC_NUM_RE = re.compile(
    r'(?<![\w.,])(\d+(?:[.,]\d+)?)(?:\s?(?:mm|ly|kw|hp|a|v|w|m|l|cm|kg|bar|inch))?(?![\w])',
    re.IGNORECASE,
)


def extract_spec_numbers(text: str) -> set:
    """Tập các số thông số đứng riêng (có thể kèm đơn vị): 6205, 13mm, 7.5kw... Bỏ qua số dính chữ như '2RS', 'NN102WM'."""
    out = set()
    for m in _SPEC_NUM_RE.finditer(canon(str(text or ""))):  # quy đổi cách gọi của thợ (fi 10 = 100mm, ly = mm) trước khi lấy số
        s = m.group(1).replace(",", ".")
        try:
            out.add(str(float(s)).rstrip("0").rstrip(".") if "." in s else str(int(s)))
        except ValueError:
            continue
    return out


def check_ai_claimed_match(query: str, code: str, name_by_code: Dict[str, str]) -> Tuple[bool, str]:
    """
    Kiểm tra mã kho do AI tự chọn cho một dòng phiếu (AI báo KHỚP kèm mã).
      - Mã phải tồn tại trong kho (chống AI bịa mã).
      - Mọi số thông số trong tên mã kho phải xuất hiện trong dòng phiếu (chống "sửa lỗi đánh máy" 202→200, 301/5A→300/5A).
    Trả về (hợp_lệ, lý_do_nếu_không).
    """
    code = str(code or "").strip()
    if code not in name_by_code:
        return False, f"Mã {code} AI chọn không có trong kho — cần kiểm tra lại"
    stock_name = str(name_by_code[code])
    c_nums = extract_spec_numbers(stock_name)
    q_nums = extract_spec_numbers(query)
    if c_nums and not c_nums.issubset(q_nums):
        missing = ", ".join(sorted(c_nums - q_nums))
        return False, f"AI chọn {code} '{stock_name}' nhưng lệch thông số (kho có {missing} mà chứng từ không có)"
    return True, ""


def decide_auto_match(query: str, best: Optional[Dict[str, Any]], strict_exact: bool = False) -> Tuple[str, str]:
    """
    Trả về (quyết_định, ghi_chú) với quyết_định ∈ {"KHỚP", "MỚI", "CẦN_XÁC_NHẬN"}.
    `best` là ứng viên đầu bảng dạng dict có các khóa: ten, ma, score, tier, reasons
    (đúng định dạng get_candidate_matches() trong app.py trả về).

      - EXACT_MATCH (>=85): KHỚP, trừ khi có cảnh báo xung đột (⚠️) → CẦN_XÁC_NHẬN.
      - GENERIC_MATCH (65-84): chỉ KHỚP khi (1) không có cảnh báo xung đột, (2) có từ khóa chung,
        (3) mọi số thông số trong tên kho đều xuất hiện trong chứng từ (không lệch 6205 ↔ 6206, 21 ↔ 27...).
        Không đạt (3) → CẦN_XÁC_NHẬN (người duyệt chọn trong gợi ý), KHÔNG tự cấp mã mới.
      - strict_exact=True: tầng EXACT cũng áp điều kiện (3) (đánh đổi: nhiều dòng phải xác nhận tay hơn).
      - Còn lại: MỚI.
    """
    if not best:
        return "MỚI", ""
    tier = best.get("tier")
    score = float(best.get("score", 0) or 0)
    reasons = " ".join(str(r) for r in (best.get("reasons") or []))
    cand_name = str(best.get("ten", ""))
    conflict = "⚠️" in reasons

    if tier == "EXACT_MATCH" and score >= 85:
        if conflict:
            return "CẦN_XÁC_NHẬN", "Điểm cao nhưng có cảnh báo xung đột thông số — cần kỹ sư xác nhận"
        if strict_exact:
            # Chế độ nghiêm ngặt: tầng EXACT cũng phải chứa đủ số thông số của tên kho
            c_nums = extract_spec_numbers(cand_name)
            q_nums = extract_spec_numbers(query)
            if c_nums and not c_nums.issubset(q_nums):
                missing = ", ".join(sorted(c_nums - q_nums))
                return "CẦN_XÁC_NHẬN", f"Gợi ý {best.get('ma')} nhưng lệch thông số (kho có {missing} mà chứng từ không có)"
        return "KHỚP", ""

    if tier == "GENERIC_MATCH" and score >= 65:
        q_norm = str(query or "").lower().replace(",", ".").replace("*", "x")
        c_norm = cand_name.lower().replace(",", ".").replace("*", "x")
        shared_kw = any(kw in c_norm for kw in q_norm.split() if len(kw) >= 3)
        if conflict or not shared_kw:
            return "MỚI", ""
        c_nums = extract_spec_numbers(cand_name)
        q_nums = extract_spec_numbers(query)
        if c_nums and not c_nums.issubset(q_nums):
            missing = ", ".join(sorted(c_nums - q_nums))
            return "CẦN_XÁC_NHẬN", f"Gợi ý {best.get('ma')} nhưng lệch thông số (kho có {missing} mà chứng từ không có)"
        return "KHỚP", ""

    # Gợi ý số 1 chưa đủ điểm tự khớp nhưng RẤT giống (đủ số thông số của tên kho + chung từ khóa chủng loại):
    # đừng cấp mã mới ngay (dễ trùng mã) — để kỹ sư xem và chọn.
    if score >= 50 and not conflict:
        c_nums = extract_spec_numbers(cand_name)
        q_nums = extract_spec_numbers(query)
        q_words = {w for w in re.findall(r"[^\W\d_]{3,}", canon(query))}
        c_words = {w for w in re.findall(r"[^\W\d_]{3,}", canon(cand_name))}
        if c_nums and c_nums.issubset(q_nums) and (q_words & c_words):
            return "CẦN_XÁC_NHẬN", f"Gợi ý {best.get('ma')} '{cand_name}' có thể là cùng vật tư (trùng thông số {', '.join(sorted(c_nums))}) — kỹ sư xem lại trước khi cấp mã mới"

    return "MỚI", ""
