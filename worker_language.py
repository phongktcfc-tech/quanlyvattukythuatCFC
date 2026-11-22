# -*- coding: utf-8 -*-
"""
NGÔN NGỮ CỦA THỢ / KỸ THUẬT VIÊN — chuẩn hóa cách gọi dân dã về một dạng chung trước khi so mã kho.
(Module độc lập, KHÔNG sửa sku_matching_engine; app.py chỉ gọi resolve_candidates khi cách tìm thường không ra mã.)

Ví dụ thực tế trong kho CFC:
  "Đá cắt fi 10"      = đá cắt đường kính 10 cm = 100 mm  (phiếu giao hàng ghi "đá cắt sắt 100mm")
  "M.Khoan 8,5 ly"    = Mũi khoan 8,5 mm                  (ly / li = mm)
  "Mũi khoan thép gió Nachi 8.5mm" ↔ kho "Mũi khoan thép 8.5mm" (thép gió = thép, hãng không phải thông số)
"""
import re
from typing import Any, Callable, Dict, List, Optional, Tuple

import pandas as pd

# Nhóm vật tư dạng đĩa: "fi N" với N nhỏ nghĩa là N cm (kho ghi "Đá cắt fi 10", "Đá mài fi 10" = 100 mm)
_DISC_RE = re.compile(r"\b(đá\s*cắt|đá\s*mài|lưỡi\s*cắt|đĩa\s*cắt|đĩa\s*mài|đá\s*nhám)\b")
_FI_RE = re.compile(r"(?:\bfi|\bphi|[øφ])\s*(\d+(?:\.\d+)?)(?!\w)")
_LY_RE = re.compile(r"(\d+(?:\.\d+)?)\s*(?:ly|li)\b")
_CM_RE = re.compile(r"(\d+(?:\.\d+)?)\s*cm\b")
DISC_FI_CM_MAX = 40  # fi 40 trở xuống trên đá cắt/mài được hiểu là cm (fi 300, fi 400 là mm)


def _fmt(v: float) -> str:
    return str(int(round(v))) if abs(v - round(v)) < 1e-9 else str(round(v, 2))


def canon(text: str) -> str:
    """Đưa cách gọi của thợ về dạng chuẩn (chữ thường, dấu thập phân là '.', thông số theo mm)."""
    s = str(text or "").lower().replace(",", ".")
    s = re.sub(r"\bm\.\s*khoan\b", "mũi khoan", s)
    s = _LY_RE.sub(lambda m: f"{m.group(1)}mm", s)
    s = _CM_RE.sub(lambda m: f"{_fmt(float(m.group(1)) * 10)}mm", s)
    if _DISC_RE.search(s):
        def _disc(m):
            n = float(m.group(1))
            return f"{_fmt(n * 10 if n <= DISC_FI_CM_MAX else n)}mm"
        s = _FI_RE.sub(_disc, s)
    s = re.sub(r"\bthép\s+gió\b", "thép", s)
    return re.sub(r"\s+", " ", s).strip()


_CANON_MATCHER_CACHE: Dict[Any, Any] = {}


def _canonical_master(df_master: pd.DataFrame) -> pd.DataFrame:
    df = df_master.copy()
    df["Tên hàng"] = df["Tên hàng"].astype(str).map(canon)
    return df


def get_canonical_matcher(df_master: pd.DataFrame, matcher_cls):
    """Matcher thứ hai dựng trên tên kho đã chuẩn hóa (cache theo nội dung tên hàng)."""
    key = (len(df_master), int(pd.util.hash_pandas_object(df_master["Tên hàng"].astype(str), index=False).sum() % (1 << 62)))
    m = _CANON_MATCHER_CACHE.get(key)
    if m is None:
        _CANON_MATCHER_CACHE.clear()
        m = matcher_cls(_canonical_master(df_master))
        _CANON_MATCHER_CACHE[key] = m
    return m


def _to_candidates(match_res: Dict[str, Any], name_by_code: Dict[str, str], note: str) -> List[Dict[str, Any]]:
    out = []
    for c in match_res.get("top_candidates", []) or []:
        code = str(c["ma_hang"])
        reasons = list(c.get("reasons") or [])
        reasons.append(note)
        out.append({
            "score": c["score"], "pct_str": f"[{int(c['score'])}%]", "ma": code,
            "ten": name_by_code.get(code, str(c["ten_hang"])),  # hiển thị tên GỐC trong kho
            "dvt": str(c["dvt"]), "ton": c["cuoi_ky_sl"], "same_cat": True,
            "ly_do": ", ".join(reasons), "tier": c.get("tier", "SIMILAR_MISMATCH"),
            "don_gia_bq": c.get("don_gia_bq", 0.0), "price_diff_pct": c.get("price_diff_pct"),
            "data_alerts": c.get("data_alerts", []), "reasons": reasons,
        })
    return out


def _is_good(cands: List[Dict[str, Any]]) -> bool:
    if not cands:
        return False
    b = cands[0]
    return (b.get("tier") == "EXACT_MATCH" and b.get("score", 0) >= 85) or (b.get("tier") == "GENERIC_MATCH" and b.get("score", 0) >= 65)


def resolve_candidates(
    primary: List[Dict[str, Any]],
    t_goc: str, t_chuan: str, q_cach: str,
    df_master: pd.DataFrame, matcher_cls,
    input_price: float = 0.0, top_k: int = 5,
) -> Tuple[List[Dict[str, Any]], str]:
    """
    primary = kết quả tìm thường (theo tên gốc + quy cách). Nếu đã đủ tốt (EXACT/GENERIC) → giữ nguyên.
    Nếu không, thử lại theo ngôn ngữ của thợ: (1) tên gốc đã chuẩn hóa đơn vị; (2) tên chuẩn hóa của AI + quy cách.
    Chỉ thay kết quả khi cách thử lại đạt EXACT/GENERIC và điểm cao hơn rõ. Trả về (danh_sách_ứng_viên, ghi_chú).
    """
    if _is_good(primary):
        return primary, ""
    try:
        matcher = get_canonical_matcher(df_master, matcher_cls)
        name_by_code = dict(zip(df_master["Mã hàng"].astype(str).str.strip(), df_master["Tên hàng"].astype(str)))
        base_score = float(primary[0]["score"]) if primary else 0.0
        variants = [
            (canon(f"{t_goc} {q_cach}"), "đối chiếu theo cách gọi của thợ (đã quy đổi đơn vị/viết tắt)"),
        ]
        if str(t_chuan or "").strip():
            variants.append((canon(f"{t_chuan} {q_cach}"), "đối chiếu theo tên chuẩn hóa của AI"))
        best_c: Optional[List[Dict[str, Any]]] = None
        for text, note in variants:
            if not text:
                continue
            cands = _to_candidates(matcher.match_item(text, input_price=input_price, top_k=top_k), name_by_code, note)
            if _is_good(cands) and cands[0]["score"] >= base_score + 5 and (best_c is None or cands[0]["score"] > best_c[0]["score"]):
                best_c = cands
        if best_c:
            return best_c, best_c[0]["ly_do"].split(", ")[-1]
    except Exception as e:  # lỗi ở lớp phụ không được làm hỏng đối chiếu chính
        print(f"[worker_language] bỏ qua do lỗi: {e}")
    return primary, ""
