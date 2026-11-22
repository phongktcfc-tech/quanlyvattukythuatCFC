"""
BỘ ĐO ĐỘ CHÍNH XÁC AGENT 1 — chạy trên kho thật, mô phỏng đúng quyết định tự khớp của app.
Chạy:  python -u tests/benchmark_matcher.py [--limit N] [--max-show 12] [--json out.json]

Mỗi câu hỏi chạy bộ dò mã MỘT lần, rồi áp dụng 3 chính sách quyết định lên cùng kết quả:
  current : logic đang chạy trong app.py (KHỚP nếu EXACT>=85, hoặc GENERIC>=65 chung 1 từ khóa >=3 ký tự)
  gate    : auto_match_gate.decide_auto_match — chặn GENERIC khi lệch số thông số
  strict  : như gate nhưng áp cả cho EXACT (strict_exact=True)

Các phép đo (tập = các mã có tên duy nhất trong kho):
 A. Tự khớp chính tên mình           → phải KHỚP đúng mã.            (cao là tốt)
 C. Tên bị đổi cách viết vô hại      → phải KHỚP đúng mã.            (cao là tốt)
    (chữ thường, '-'→' ', ','→'.', thêm khoảng trắng; GIỮ NGUYÊN mọi số)
 B. Đổi MỘT số thông số (6205→6206)  → KHÔNG được KHỚP về mã gốc.     (thấp là tốt)
    Dòng bị chặn thành CẦN_XÁC_NHẬN được đếm riêng (an toàn, nhưng tốn công duyệt tay).
"""
import argparse
import json
import random
import re
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import pandas as pd
from sku_matching_engine import UniversalTechnicalSKUMatcher
from auto_match_gate import decide_auto_match

MASTER = ROOT / "Master_Inventory" / "Master_Inventory.xlsx"
NUM_RE = re.compile(r"(?<![\w.,])(\d{2,5})(?![\w])")
POLICIES = ["current", "gate", "strict"]


def load_master():
    df = pd.read_excel(MASTER, dtype={"Mã hàng": str, "Tên hàng": str})
    df = df.dropna(subset=["Mã hàng", "Tên hàng"]).copy()
    df["Mã hàng"] = df["Mã hàng"].astype(str).str.strip()
    df["Tên hàng"] = df["Tên hàng"].astype(str).str.strip()
    return df


def perturb(name, rng):
    nums = list(NUM_RE.finditer(name))
    if not nums:
        return None
    m = rng.choice(nums)
    new = int(m.group(1)) + rng.choice([1, 2, 3, 5, 10])
    return name[:m.start(1)] + str(new) + name[m.end(1):]


def harmless_variant(name):
    v = name.lower().replace(" - ", " ").replace("-", " ").replace(",", ".")
    return re.sub(r"\s+", "  ", v).strip()


def as_app_best(res):
    cands = res.get("top_candidates") or []
    if not cands:
        return None
    c = cands[0]
    return {"ma": str(c["ma_hang"]), "ten": str(c["ten_hang"]), "score": c["score"],
            "tier": c.get("tier"), "reasons": c.get("reasons", [])}


def policy_current(query, best):
    if not best:
        return "MỚI"
    if best["tier"] == "EXACT_MATCH" and best["score"] >= 85:
        return "KHỚP"
    q = query.lower().replace(",", ".").replace("*", "x")
    c = best["ten"].lower().replace(",", ".").replace("*", "x")
    if best["tier"] == "GENERIC_MATCH" and best["score"] >= 65 and any(k in c for k in q.split() if len(k) >= 3):
        return "KHỚP"
    return "MỚI"


WORKER = {"on": False, "df": None}


def decisions(matcher, query):
    res = matcher.match_item(query)
    best = as_app_best(res)
    if WORKER["on"]:  # mô phỏng thêm bước resolve_candidates (ngôn ngữ của thợ) như app.py
        from worker_language import resolve_candidates
        prim = [{"score": c["score"], "ma": str(c["ma_hang"]), "ten": str(c["ten_hang"]), "tier": c.get("tier"),
                 "ton": c["cuoi_ky_sl"], "ly_do": "", "reasons": c.get("reasons", [])} for c in (res.get("top_candidates") or [])]
        cands, _ = resolve_candidates(prim, query, "", "", WORKER["df"], UniversalTechnicalSKUMatcher)
        if cands:
            c = cands[0]
            best = {"ma": c["ma"], "ten": c["ten"], "score": c["score"], "tier": c.get("tier"), "reasons": c.get("reasons", [])}
    return best, {
        "current": policy_current(query, best),
        "gate": decide_auto_match(query, best)[0],
        "strict": decide_auto_match(query, best, strict_exact=True)[0],
    }


def run(limit, max_show, out_json):
    df = load_master()
    WORKER["df"] = df
    matcher = UniversalTechnicalSKUMatcher(df)
    rng = random.Random(20261116)
    counts = df["Tên hàng"].str.lower().value_counts()
    unique = df[df["Tên hàng"].str.lower().map(counts) == 1]
    if limit:
        unique = unique.sample(n=min(limit, len(unique)), random_state=1)

    A = {p: 0 for p in POLICIES}; C = {p: 0 for p in POLICIES}
    B_false = {p: 0 for p in POLICIES}; B_confirm = {p: 0 for p in POLICIES}
    nA = nB = 0
    ex_false = {p: [] for p in POLICIES}
    t0 = time.time()
    for i, (_, r) in enumerate(unique.iterrows(), 1):
        code, name = r["Mã hàng"], r["Tên hàng"]
        best, d = decisions(matcher, name)
        nA += 1
        for p in POLICIES:
            A[p] += d[p] == "KHỚP" and bool(best) and best["ma"] == code
        best, d = decisions(matcher, harmless_variant(name))
        for p in POLICIES:
            C[p] += d[p] == "KHỚP" and bool(best) and best["ma"] == code
        q = perturb(name, rng)
        if q:
            best, d = decisions(matcher, q)
            nB += 1
            for p in POLICIES:
                if d[p] == "KHỚP" and best and best["ma"] == code:
                    B_false[p] += 1
                    ex_false[p].append({"code": code, "original": name, "mutated": q, "score": best["score"]})
                if d[p] == "CẦN_XÁC_NHẬN":
                    B_confirm[p] += 1
        if i % 100 == 0:
            print(f"  ...{i}/{len(unique)} ({time.time() - t0:.0f}s)", flush=True)

    pct = lambda a, b: round(100 * a / b, 1) if b else None
    summary = {"master_items": len(df), "A_C_total": nA, "B_total": nB}
    for p in POLICIES:
        summary[p] = {
            "A_selfmatch_pct": pct(A[p], nA),
            "C_harmless_variant_pct": pct(C[p], nA),
            "B_false_match_pct": pct(B_false[p], nB),
            "B_false_match_n": B_false[p],
            "B_sent_to_confirm_n": B_confirm[p],
        }
    print("=" * 74)
    print(json.dumps(summary, ensure_ascii=False, indent=1))
    print("=" * 74)
    for p in ("current", "strict"):
        print(f"\n-- Khớp nhầm còn sót với chính sách '{p}' (tối đa {max_show}):")
        for x in ex_false[p][:max_show]:
            print(f"  [{x['code']}] '{x['original']}' → '{x['mutated']}' (score={x['score']})")
    if out_json:
        Path(out_json).write_text(json.dumps({"summary": summary, "false": ex_false}, ensure_ascii=False, indent=1), encoding="utf-8")
    return summary


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--max-show", type=int, default=12)
    ap.add_argument("--json", default=None)
    ap.add_argument("--worker", action="store_true", help="áp thêm lớp ngôn ngữ của thợ (worker_language.resolve_candidates)")
    a = ap.parse_args()
    WORKER["on"] = a.worker
    run(a.limit, a.max_show, a.json)
