# -*- coding: utf-8 -*-
"""
UNIVERSAL TECHNICAL SKU MATCHING ENGINE (CFC WAREHOUSE)
Hệ thống chuẩn hóa, phân tích cú pháp kỹ thuật & dò tìm mã vật tư đa tầng
Dành cho Agent 1: Đối chiếu hóa đơn/chứng từ NCC với 1.410 mã kho gốc CFC
"""

import re
import unicodedata
import pandas as pd
from typing import Dict, List, Tuple, Any, Optional

# ═══════════════════════════════════════════════════════════════════════════
# TẦNG 2: TỪ ĐIỂN ĐỒNG NGHĨA ĐA CHUYÊN NGÀNH (COMPREHENSIVE SYNONYM DICTIONARY)
# ═══════════════════════════════════════════════════════════════════════════

# 1. NHÓM THIẾT BỊ / SẢN PHẨM CHÍNH (PRODUCT CATEGORIES)
PRODUCT_CATEGORY_GROUPS: List[List[str]] = [
    # Thiết bị Điện & Tự động hóa
    ["khởi động từ", "contactor", "công tắc tơ", "kđt"],
    ["rơ le nhiệt", "rơ le", "rơle", "role", "relay", "rơ-le", "rơle bảo vệ"],
    ["aptomat", "cb", "mccb", "mcb", "cầu dao", "cầu dao điện", "cầu dao tự động", "cb tép"],
    ["biến tần", "inverter", "vfd"],
    ["bộ điều khiển nhiệt độ", "đồng hồ nhiệt độ", "đồng hồ điều khiển nhiệt", "bộ nhiệt độ"],
    ["cảm biến nhiệt độ", "cảm biến nhiệt", "can nhiệt", "điện trở dò nhiệt", "pt100", "củ hành"],
    ["đầu cos", "đầu cốt", "cos", "cốt"],
    ["dây rút", "lạt nhựa"],
    ["ống ruột gà", "ống luồn đàn hồi"],
    ["công tắc hành trình", "limit switch"],
    ["bộ nguồn", "nguồn tổ ong", "power supply"],

    # Vòng bi, Gối đỡ & Truyền động
    ["vòng bi", "bạc đạn", "ổ bi", "bearing"],
    ["gối đỡ", "gối đỡ vòng bi", "ổ đỡ"],
    ["dây curoa", "dây đai thang", "curoa", "day curoa"],
    ["xích", "sên", "xích tải", "sên đơn", "sên đôi"],
    ["nhông", "bánh răng", "đĩa xích"],
    ["puly", "buly", "pu ly", "puli"],
    ["khớp nối", "coupling"],

    # Phốt, Gioăng & Làm kín
    ["phốt", "fos", "phos", "phớt", "oil seal", "phốt chắn dầu", "phot"],
    ["vòng sin", "oring", "o-ring", "gioăng cao su", "ron cao su", "roang", "sim"],
    ["dây tết chèn", "dây teflon", "sợi teflon", "dây amiang"],
    ["bìa amiang", "gioăng chì", "ron chì", "gasket"],

    # Bulong, Ốc vít & Kim khí
    ["bulon", "bulong", "bu lông", "bu-lông", "bolt"],
    ["ecu", "ê cu", "đai ốc", "tán", "nut"],
    ["long đền", "long đềnh", "vòng đệm", "long den", "washer"],
    ["vít", "vis", "ốc", "vit"],
    ["pake", "bake", "pa ke", "4 cạnh"],
    ["tắc kê", "nở sắt", "nở nhựa", "bulong nở"],
    ["cây răng", "ty ren", "thanh ren"],
    ["tăng đơ", "tăng đưa"],
    ["cổ dê", "đai siết", "quai nhê"],

    # Thép hình, Tấm & Kết cấu
    ["thép hình u", "thép u", "u đúc", "u duc", "sắt u"],
    ["thép hình i", "thép i", "sắt i"],
    ["thép v", "sắt v", "v đúc"],
    ["thép hộp", "thép vuông", "sắt hộp"],
    ["thép tấm", "tole tấm", "sắt tấm", "ton tam"],
    ["thép la", "lập là", "sat la"],
    ["thép tròn", "thép láp", "thép láp tròn", "sắt phi", "láp tròn"],
    ["tole", "tol", "tôn"],

    # Vật tư Đường ống, Van & Khí nén
    ["van cầu", "van hơi", "van cầu hơi"],
    ["van bi", "ball valve"],
    ["van bướm", "butterfly valve"],
    ["van một chiều", "van 1 chiều", "check valve"],
    ["van điện từ", "solenoid valve"],
    ["ống nhựa bm", "ống bm", "ống pvc", "ống upvc", "ống nhựa", "ong bm"],
    ["ống thép", "thép ống", "ong thep"],
    ["ống inox", "ong inox"],
    ["ống", "ong"],
    ["co lơi", "lơi", "chếch", "co 45"],
    ["co", "cút", "co 90"],
    ["tê", "chữ t", "ba chạc"],
    ["giảm", "bầu giảm", "côn thu", "nối giảm"],
    ["rắc co", "zắc co", "union"],
    ["mặt bích", "bích", "flange"],
    ["xy lanh", "xi lanh", "ben hơi", "xy lanh khí nén"],

    # Dụng cụ Cắt gọt, Hàn & Cầm tay
    ["mũi khoan", "m.khoan", "m khoan", "mủi khoan"],
    ["cờ lê", "clê", "clé", "chìa khóa"],
    ["cần điếu", "ống điếu", "tuýp"],
    ["que hàn", "điện cực hàn", "đũa hàn"],
    ["đá cắt", "lưỡi cắt"],
    ["đá mài", "nhám xếp"],

    # Động cơ, Bơm, Quạt & Máy móc
    ["động cơ", "motor", "mô tơ"],
    ["hộp giảm tốc", "hộp số giảm tốc", "hgt", "động cơ giảm tốc", "motor giảm tốc"],
    ["máy bơm", "bơm", "guồng bơm"],
    ["quạt hút", "quạt thổi", "quạt công nghiệp"]
]

# 2. NHÓM CHẤT LIỆU & VẬT LIỆU (MATERIAL GROUPS - KHÔNG ĐƯỢC COI LÀ CHỦNG LOẠI THIẾT BỊ)
MATERIAL_SYNONYM_GROUPS: List[List[str]] = [
    ["inox", "thép không gỉ", "sus304", "sus 304", "sus201", "sus 201", "sus316", "sus 316"],
    ["mạ kẽm", "kẽm", "nhúng kẽm"],
    ["gang"],
    ["đồng", "thau"],
    ["nhựa", "pvc", "upvc", "pe", "teflon"],
    ["thép", "sắt"]
]

# Xây dựng từ điển tra cứu nhanh
CATEGORY_LOOKUP: Dict[str, str] = {}
CATEGORY_GROUP_MAP: Dict[str, List[str]] = {}

for group in PRODUCT_CATEGORY_GROUPS:
    canonical = group[0]
    for syn in group:
        s_norm = syn.strip().lower()
        CATEGORY_LOOKUP[s_norm] = canonical
        CATEGORY_GROUP_MAP[s_norm] = [g.strip().lower() for g in group]

MATERIAL_LOOKUP: Dict[str, str] = {}
for group in MATERIAL_SYNONYM_GROUPS:
    canonical = group[0]
    for syn in group:
        s_norm = syn.strip().lower()
        MATERIAL_LOOKUP[s_norm] = canonical

# Bảng quy đổi đường kính ống tiêu chuẩn công nghiệp (DN <-> Phi mm <-> Inch)
PIPE_DN_CONVERSIONS: List[Dict[str, Any]] = [
    {"dn": "dn15", "phi": 21, "inch": "1/2"},
    {"dn": "dn20", "phi": 27, "inch": "3/4"},
    {"dn": "dn25", "phi": 34, "inch": "1"},
    {"dn": "dn32", "phi": 42, "inch": "1.1/4"},
    {"dn": "dn40", "phi": 49, "inch": "1.1/2"},
    {"dn": "dn50", "phi": 60, "inch": "2"},
    {"dn": "dn65", "phi": 76, "inch": "2.1/2"},
    {"dn": "dn80", "phi": 90, "inch": "3"},
    {"dn": "dn100", "phi": 114, "inch": "4"},
    {"dn": "dn125", "phi": 140, "inch": "5"},
    {"dn": "dn150", "phi": 168, "inch": "6"},
    {"dn": "dn200", "phi": 219, "inch": "8"},
    {"dn": "dn250", "phi": 273, "inch": "10"},
    {"dn": "dn300", "phi": 325, "inch": "12"},
]


# ═══════════════════════════════════════════════════════════════════════════
# TẦNG 1: TIỀN XỬ LÝ & CHUẨN HÓA KỸ THUẬT (TECHNICAL TEXT NORMALIZER)
# ═══════════════════════════════════════════════════════════════════════════
def normalize_technical_text(raw_text: str) -> str:
    """
    Chuẩn hóa sâu chuỗi văn bản kỹ thuật:
    1. Unicode NFC & lowercase
    2. Chuyển toàn bộ dấu * hoặc × thành x (35*58*8 -> 35x58x8, 16*100 -> 16x100)
    3. Chuẩn hóa ren bulon hệ mét (M16x100 -> 16x100)
    4. Đồng nhất số thập phân (1,5 -> 1.5)
    5. Đồng nhất đơn vị đo (ly/li -> mm, tấc -> 100mm, phân -> 10mm)
    6. Chuẩn hóa pha điện (3 pha -> 3pha)
    7. Chuẩn hóa dải thông số (9_13A, 9 - 13 A -> 9-13a)
    """
    if not raw_text or pd.isna(raw_text):
        return ""

    text = str(raw_text).strip()
    text = unicodedata.normalize("NFC", text).lower()

    # 1. Chuyển toàn bộ toán tử kích thước (*, ×) thành x giữa hoặc quanh số
    text = re.sub(r'[\*×]', 'x', text)

    # 2. Chuẩn hóa Bulon hệ mét: M16x100 -> 16x100, M10x40 -> 10x40, M16 -> 16
    text = re.sub(r'\bm(\d+x\d+(?:x\d+)?)\b', r'\1', text)
    text = re.sub(r'\bm(\d+)\b', r'\1', text)

    # 3. Đồng nhất số thập phân: dấu phẩy giữa 2 số thành dấu chấm (1,5 -> 1.5)
    text = re.sub(r'(\d+),(\d+)', r'\1.\2', text)

    # 4. Quy đổi đơn vị đo lường: ly, li sau số thành mm
    text = re.sub(r'(\d+(?:\.\d+)?)\s*(?:ly|li)\b', r'\1mm', text)

    # 5. Quy đổi đơn vị dân gian miền Nam:
    def replace_tac(m):
        val = float(m.group(1))
        mm_val = int(val * 100) if (val * 100).is_integer() else val * 100
        return f"{mm_val}mm {int(val)}tac"

    text = re.sub(r'(\d+(?:\.\d+)?)\s*(?:tấc|tac)\b', replace_tac, text)

    def replace_phan(m):
        val = float(m.group(1))
        mm_val = int(val * 10) if (val * 10).is_integer() else val * 10
        return f"{mm_val}mm {int(val)}phan"

    text = re.sub(r'(\d+(?:\.\d+)?)\s*(?:phân|phan|p)\b', replace_phan, text)

    # 6. Chuẩn hóa số pha điện
    text = re.sub(r'\b(3)\s*(?:pha|fa|p)\b', r'3pha', text)
    text = re.sub(r'\b(1)\s*(?:pha|fa|p)\b', r'1pha', text)
    text = re.sub(r'\b(2)\s*(?:pha|fa|p)\b', r'2pha', text)

    # 7. Đồng nhất toán tử x giữa các số (16 x 100 -> 16x100)
    for _ in range(2):
        text = re.sub(r'(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)', r'\1x\2', text)

    # 8. Chuẩn hóa dải thông số (Range): 9_13A, 9 - 13 A, 9-13A -> 9-13a
    text = re.sub(r'(\d+(?:\.\d+)?)\s*[\-_~÷]\s*(\d+(?:\.\d+)?)\s*(a|v|kw|w|bar|kg|mm)?\b',
                  lambda m: f"{m.group(1)}-{m.group(2)}{(m.group(3) or '').lower()}", text)

    # 9. Chuẩn hóa đơn vị đo dính liền số (100 a -> 100a, 220 v -> 220v, 75 kw -> 75kw)
    text = re.sub(r'(\d+(?:\.\d+)?)\s*(a|v|kw|w|bar|kg|kva|rpm)\b', r'\1\2', text)

    # 10. Chuẩn hóa ký tự phi
    text = re.sub(r'[øφđd]\s*(\d+)', r'phi \1', text)

    # Dọn dẹp khoảng trắng thừa
    text = re.sub(r'\s+', ' ', text).strip()
    return text


# ═══════════════════════════════════════════════════════════════════════════
# TẦNG 3: PHÂN RÃ CẤU TRÚC VẬT TƯ (STRUCTURED ITEM PARSER)
# ═══════════════════════════════════════════════════════════════════════════
def parse_structured_item(raw_name: str) -> Dict[str, Any]:
    """
    Phân tích cú pháp vật tư kỹ thuật thành Object có cấu trúc chi tiết:
    - category_canonicals: danh mục chuẩn của thiết bị (Contactor, Rơ le, Van cầu, Bulon...)
    - category_tokens: các từ đồng nghĩa của chủng loại thiết bị
    - primary_specs: thông số định danh chính (100a, 9-13a, 6205, 35x58x8, 16x100...)
    - secondary_specs: thông số phụ (220v, 380v, 1485rpm...)
    - dimensions: kích thước hình học (axb hoặc axbxc)
    - model_codes: mã model trích xuất (mc-100a, mt-32, e5csl-qtc...)
    - material_tokens: vật liệu (inox, kẽm, gang, nhựa...)
    """
    normalized = normalize_technical_text(raw_name)

    category_tokens: List[str] = []
    category_canonicals: List[str] = []

    # 1. Tìm các cụm từ khóa danh mục thiết bị chính (Ưu tiên cụm từ dài nhất trước)
    sorted_cat_syns = sorted(CATEGORY_LOOKUP.keys(), key=len, reverse=True)
    matched_ranges = []

    for syn in sorted_cat_syns:
        pattern = r'(?:\b|^)' + re.escape(syn) + r'(?:\b|$)'
        match = re.search(pattern, normalized)
        if match:
            start, end = match.span()
            if not any(s <= start and end <= e for s, e in matched_ranges):
                matched_ranges.append((start, end))
                canonical = CATEGORY_LOOKUP[syn]
                if canonical not in category_canonicals:
                    category_canonicals.append(canonical)
                for s in CATEGORY_GROUP_MAP.get(syn, [syn]):
                    if s not in category_tokens:
                        category_tokens.append(s)

    # 2. Tìm chất liệu / vật liệu (Tách biệt hoàn toàn khỏi danh mục thiết bị)
    material_tokens: List[str] = []
    for mat_syn, mat_canon in MATERIAL_LOOKUP.items():
        if re.search(r'(?:\b|^)' + re.escape(mat_syn) + r'(?:\b|$)', normalized):
            if mat_canon not in material_tokens:
                material_tokens.append(mat_canon)

    primary_specs: List[str] = []
    secondary_specs: List[str] = []
    model_codes: List[str] = []
    dimensions: List[str] = []

    # 3. Model Spec Unpacking: Thiết bị đóng cắt (LS / Schneider / Mitsubishi)
    ls_mc_matches = re.findall(r'\b(mc\s*[-_]?\s*(\d+[ab]?))\b', normalized)
    for full_m, amp in ls_mc_matches:
        clean_m = re.sub(r'\s+', '', full_m)
        model_codes.append(clean_m)
        num_part = re.sub(r'[ab]$', '', amp)
        primary_specs.append(f"{num_part}a")

    ls_mt_matches = re.findall(r'\b(mt\s*[-_]?\s*\d+[ab]?)\b', normalized)
    for m in ls_mt_matches:
        model_codes.append(re.sub(r'\s+', '', m))

    # Mã Omron hoặc thiết bị điều khiển nhiệt độ (E5CSL-QTC, E5CN...)
    omron_matches = re.findall(r'\b([a-z]\d[a-z0-9]+(?:-[a-z0-9]+)+)\b', normalized)
    for m in omron_matches:
        model_codes.append(m)

    # 4. Dải dòng (Range) của Rơ le nhiệt: 9-13a, 4-6a, 6-9a
    range_matches = re.findall(r'\b(\d+(?:\.\d+)?-\d+(?:\.\d+)?a?)\b', normalized)
    for r in range_matches:
        r_clean = r if r.endswith('a') else f"{r}a"
        if r_clean not in primary_specs:
            primary_specs.append(r_clean)

    # 5. Dòng định mức Ampe đơn thuần (100a, 150a, 63a, 20a, 32a...)
    amp_matches = re.findall(r'\b(\d+(?:\.\d+)?)a\b', normalized)
    for a in amp_matches:
        a_token = f"{a}a"
        if a_token not in primary_specs:
            primary_specs.append(a_token)

    # 6. Công suất điện (kW, W, HP)
    kw_matches = re.findall(r'\b(\d+(?:\.\d+)?)(kw|w|hp)\b', normalized)
    for val, unit in kw_matches:
        primary_specs.append(f"{val}{unit}")

    # 7. Kích thước hình học 2 hoặc 3 chiều (Phốt, Bulong, Thép hình: 35x58x8, 16x100, 20x52x10...)
    dim_matches = re.findall(r'\b(\d+(?:\.\d+)?x\d+(?:\.\d+)?(?:x\d+(?:\.\d+)?)?)\b', normalized)
    for d in dim_matches:
        dimensions.append(d)
        if d not in primary_specs:
            primary_specs.append(d)

    # 8. Mã số vòng bi / bạc đạn / gối đỡ (3 đến 5 chữ số: 6205, 22224, 22316, 30207, 207...)
    bearing_matches = re.findall(r'\b((?:uc|ucp|ucf|ucfl|ucfc|nj|nu)?\s*(\d{3,5})\s*(?:zz|2rs|dduc3|eae4|ca/w33)?)\b', normalized)
    for full_b, base_num in bearing_matches:
        if base_num not in primary_specs:
            primary_specs.append(base_num)
        model_codes.append(re.sub(r'\s+', '', full_b))

    # 9. Quy đổi đường kính ống & phụ kiện (DN <-> Phi <-> Inch)
    for item in PIPE_DN_CONVERSIONS:
        dn_kw = item["dn"]
        phi_kw = f"phi {item['phi']}"
        phi_clean = f"phi{item['phi']}"

        has_dn = dn_kw in normalized
        has_phi = (phi_kw in normalized) or (phi_clean in normalized) or (
            re.search(r'\b' + str(item['phi']) + r'\b', normalized) and any(p in normalized for p in ["van", "ong", "co", "bich", "te", "giam"])
        )

        if has_dn or has_phi:
            pipe_spec_code = f"dn{item['dn'][2:]}_phi{item['phi']}"
            if pipe_spec_code not in primary_specs:
                primary_specs.append(pipe_spec_code)
                primary_specs.append(dn_kw)
                primary_specs.append(f"phi{item['phi']}")

    # 10. Thông số điện áp phụ (220v, 380v, 110v, 24v...)
    volt_matches = re.findall(r'\b(\d+v|ac\d+v|\d+vac)\b', normalized)
    for v in volt_matches:
        secondary_specs.append(v)

    return {
        "raw_name": raw_name,
        "normalized_name": normalized,
        "category_canonicals": category_canonicals,
        "category_tokens": category_tokens,
        "primary_specs": primary_specs,
        "secondary_specs": secondary_specs,
        "dimensions": dimensions,
        "model_codes": model_codes,
        "material_tokens": material_tokens
    }


# ═══════════════════════════════════════════════════════════════════════════
# TẦNG 4: THUẬT TOÁN CHẤM ĐIỂM ĐA TIÊU CHÍ (WEIGHTED SCORING ENGINE)
# ═══════════════════════════════════════════════════════════════════════════
def compute_sku_match_score(query_parsed: Dict[str, Any], cand_parsed: Dict[str, Any], cand_row: pd.Series) -> Tuple[float, List[str], str]:
    """
    Chấm điểm độ khớp giữa vật tư cần tìm và vật tư kho trên thang 100 điểm:
    1. Category Score: Tối đa 35 điểm (Đúng chủng loại thiết bị)
    2. Primary Spec Score: Tối đa 45 điểm (Đúng thông số định danh)
    3. Secondary / Model Score: Tối đa 20 điểm (Model, vật liệu, điện áp)
    4. Operational Tie-breakers: +2 đến +5 điểm
    """
    score = 0.0
    reasons = []

    q_name = query_parsed["normalized_name"]
    c_name = cand_parsed["normalized_name"]

    # ── 1. ĐIỂM DANH MỤC THIẾT BỊ CHÍNH (TỐI ĐA 35 ĐIỂM) ──
    cat_score = 0.0
    shared_canonicals = set(query_parsed["category_canonicals"]).intersection(set(cand_parsed["category_canonicals"]))
    shared_cats = set(query_parsed["category_tokens"]).intersection(set(cand_parsed["category_tokens"]))

    if shared_canonicals:
        cat_score = 35.0
        canonical_name = list(shared_canonicals)[0].title()
        reasons.append(f"Đồng nghĩa: {canonical_name}")
    elif shared_cats:
        cat_score = 30.0
        reasons.append(f"Khớp nhóm: {list(shared_cats)[0]}")
    else:
        # Nếu chủng loại chính bị lệch hẳn (ví dụ 1 bên là 'van cầu' mà bên kia là 'ống') -> Không cho điểm danh mục
        if query_parsed["category_canonicals"] and cand_parsed["category_canonicals"]:
            cat_score = 0.0
        else:
            q_words = q_name.split()
            c_words = c_name.split()
            if q_words and c_words and q_words[0] == c_words[0]:
                cat_score = 20.0
                reasons.append(f"Cùng từ đầu: {q_words[0]}")
            else:
                common_words = set(q_words).intersection(set(c_words))
                if len(common_words) >= 2:
                    cat_score = 15.0
                    reasons.append(f"Trùng {len(common_words)} từ khóa")

    # ── 2. ĐIỂM THÔNG SỐ KỸ THUẬT CHÍNH (TỐI ĐA 45 ĐIỂM) ──
    spec_score = 0.0
    q_specs = set(query_parsed["primary_specs"])
    c_specs = set(cand_parsed["primary_specs"])
    shared_specs = q_specs.intersection(c_specs)

    # TRƯỜNG HỢP A: Khớp chính xác 100% thông số chính (100A, 9-13A, 6205, 35x58x8, 16x100, DN25...)
    if shared_specs:
        spec_score = 45.0
        reasons.append(f"Khớp thông số: {', '.join(list(shared_specs)[:2])}")

    # TRƯỜNG HỢP B: Kho lưu tên chung không ghi thông số (Generic Category Match)
    elif not c_specs and (c_name in q_name or all(w in q_name.split() for w in c_name.split() if len(w) > 1)):
        spec_score = 30.0
        reasons.append("Khớp danh mục chung (Kho lưu tên chung)")

    # TRƯỜNG HỢP C: Cả 2 đều có thông số nhưng KHÁC NHAU
    elif q_specs and c_specs:
        spec_score = 0.0
        reasons.append(f"⚠️ Khác thông số: Hàng nhập {list(q_specs)[0]} ≠ Kho {list(c_specs)[0]}")

    else:
        spec_score = 15.0

    # ── 3. ĐIỂM MODEL, VẬT LIỆU & THÔNG SỐ PHỤ (TỐI ĐA 20 ĐIỂM) ──
    sec_score = 0.0
    shared_models = set(query_parsed["model_codes"]).intersection(set(cand_parsed["model_codes"]))
    if shared_models:
        sec_score += 15.0
        reasons.append(f"Khớp Model: {list(shared_models)[0].upper()}")

    shared_mats = set(query_parsed["material_tokens"]).intersection(set(cand_parsed["material_tokens"]))
    if shared_mats:
        sec_score += 5.0
        reasons.append(f"Chất liệu: {list(shared_mats)[0].title()}")

    shared_volts = set(query_parsed["secondary_specs"]).intersection(set(cand_parsed["secondary_specs"]))
    if shared_volts:
        sec_score += 5.0
        reasons.append(f"Điện áp: {list(shared_volts)[0].upper()}")

    sec_score = min(20.0, sec_score)

    total_score = cat_score + spec_score + sec_score

    # ── 4. ĐIỂM THƯỞNG ƯU TIÊN VẬN HÀNH (OPERATIONAL TIE-BREAKERS) ──
    kho_name = str(cand_row.get("Tên kho", "")).lower()
    if "kỹ thuật 01" in kho_name or "ky thuat 01" in kho_name:
        total_score += 2.0

    try:
        ton_sl = float(cand_row.get("Cuối kỳ SL", 0) or 0)
        if ton_sl > 0:
            total_score += 3.0
    except (ValueError, TypeError):
        pass

    total_score = round(min(100.0, max(0.0, total_score)), 1)

    # ── 5. PHÂN TẦNG ĐẦU RA (OUTPUT TIERING) ──
    if total_score >= 75.0:
        tier = "EXACT_MATCH"          # Nhóm 1: Khớp chính xác
    elif total_score >= 60.0:
        tier = "GENERIC_MATCH"        # Nhóm 2: Khớp danh mục chung
    else:
        tier = "SIMILAR_MISMATCH"     # Nhóm 3: Cùng họ khác quy cách / gợi ý mới

    return total_score, reasons, tier


# ═══════════════════════════════════════════════════════════════════════════
# ENGINE ĐIỀU PHỐI TỔNG THỂ (UNIVERSAL TECHNICAL SKU MATCHER)
# ═══════════════════════════════════════════════════════════════════════════
class UniversalTechnicalSKUMatcher:
    """
    Engine tìm kiếm và đối chiếu mã vật tư tồn kho kỹ thuật nhà máy:
    - Pre-indexing toàn bộ 1.410 dòng kho CFC
    - Tìm kiếm tức thì trong Candidate Pool
    - Phân chia kết quả 3 Nhóm trực quan
    - Kiểm tra chéo đơn giá & cảnh báo dữ liệu bất thường
    """

    def __init__(self, df_master: pd.DataFrame):
        self.df_master = df_master.copy()
        self.indexed_candidates: List[Dict[str, Any]] = []
        self._build_index()

    def _build_index(self):
        """Xây dựng bộ chỉ mục (Search Index) trong bộ nhớ một lần duy nhất."""
        if self.df_master is None or self.df_master.empty:
            return

        for idx, row in self.df_master.iterrows():
            ma = str(row.get("Mã hàng", "")).strip()
            ten = str(row.get("Tên hàng", "")).strip()
            dvt = str(row.get("ĐVT", "")).strip()
            ten_kho = str(row.get("Tên kho", "")).strip()

            try:
                sl = float(row.get("Cuối kỳ SL", 0) or 0)
            except:
                sl = 0.0

            try:
                gt = float(row.get("Cuối kỳ GT", 0) or 0)
            except:
                gt = 0.0

            don_gia_bq = round(gt / sl, 0) if sl > 0 else 0.0

            data_alerts = []
            if sl < 0:
                data_alerts.append("⚠️ Tồn âm")
            if sl > 0 and gt == 0:
                data_alerts.append("⚠️ Tồn có SL nhưng GT = 0")

            parsed = parse_structured_item(ten)

            self.indexed_candidates.append({
                "index": idx,
                "ma_hang": ma,
                "ten_hang": ten,
                "dvt": dvt,
                "ten_kho": ten_kho,
                "cuoi_ky_sl": sl,
                "cuoi_ky_gt": gt,
                "don_gia_bq": don_gia_bq,
                "data_alerts": data_alerts,
                "parsed": parsed,
                "row_series": row
            })

    def match_item(self, query_raw: str, input_price: float = 0.0, top_k: int = 5) -> Dict[str, Any]:
        """
        Dò tìm danh sách mã kho phù hợp nhất cho chuỗi vật tư đầu vào.
        Trả về kết quả chi tiết kèm phân loại 3 Nhóm.
        """
        if not query_raw or not str(query_raw).strip():
            return {"top_candidates": [], "group1_exact": [], "group2_generic": [], "group3_similar": []}

        q_parsed = parse_structured_item(str(query_raw))
        candidate_scores = []

        q_cats = set(q_parsed["category_tokens"])
        q_specs = set(q_parsed["primary_specs"])
        q_models = set(q_parsed["model_codes"])
        q_words = set(q_parsed["normalized_name"].split())

        for cand in self.indexed_candidates:
            c_parsed = cand["parsed"]
            c_cats = set(c_parsed["category_tokens"])
            c_specs = set(c_parsed["primary_specs"])
            c_models = set(c_parsed["model_codes"])
            c_words = set(c_parsed["normalized_name"].split())

            # Candidate Pool: Chỉ lấy những mặt hàng có điểm chung
            has_common_cat = bool(q_cats.intersection(c_cats))
            has_common_spec = bool(q_specs.intersection(c_specs))
            has_common_model = bool(q_models.intersection(c_models))
            has_common_words = len(q_words.intersection(c_words)) >= 2
            is_subset = c_parsed["normalized_name"] in q_parsed["normalized_name"]

            if not (has_common_cat or has_common_spec or has_common_model or has_common_words or is_subset):
                continue

            score, reasons, tier = compute_sku_match_score(q_parsed, c_parsed, cand["row_series"])

            price_diff_pct = None
            if input_price > 0 and cand["don_gia_bq"] > 0:
                price_diff_pct = round(((input_price - cand["don_gia_bq"]) / cand["don_gia_bq"]) * 100, 1)

            candidate_scores.append({
                "ma_hang": cand["ma_hang"],
                "ten_hang": cand["ten_hang"],
                "dvt": cand["dvt"],
                "ten_kho": cand["ten_kho"],
                "cuoi_ky_sl": cand["cuoi_ky_sl"],
                "cuoi_ky_gt": cand["cuoi_ky_gt"],
                "don_gia_bq": cand["don_gia_bq"],
                "price_diff_pct": price_diff_pct,
                "data_alerts": cand["data_alerts"],
                "score": score,
                "tier": tier,
                "reasons": reasons
            })

        candidate_scores.sort(key=lambda x: x["score"], reverse=True)

        group1 = [c for c in candidate_scores if c["tier"] == "EXACT_MATCH"]
        group2 = [c for c in candidate_scores if c["tier"] == "GENERIC_MATCH"]
        group3 = [c for c in candidate_scores if c["tier"] == "SIMILAR_MISMATCH"]

        return {
            "query_parsed": q_parsed,
            "top_candidates": candidate_scores[:top_k],
            "group1_exact": group1[:3],
            "group2_generic": group2[:3],
            "group3_similar": group3[:3]
        }


# ═══════════════════════════════════════════════════════════════════════════
# TỰ ĐỘNG CHẠY 8 TEST CASES NGHIỆM THU (UNIT TESTS)
# ═══════════════════════════════════════════════════════════════════════════
def run_unit_tests(master_path: str = "D:/Agent QuanLyVatTuKyThuat/Master_Inventory/Master_Inventory.xlsx"):
    """Chạy 8 Test Cases kiểm thử tự động trên tập dữ liệu kho thực tế."""
    print("=" * 70)
    print("BẮT ĐẦU CHẠY 8 UNIT TESTS KIỂM THỬ SKI MATCHING ENGINE")
    print("=" * 70)

    df_master = pd.read_excel(master_path, dtype={"Mã hàng": str})
    matcher = UniversalTechnicalSKUMatcher(df_master)

    test_cases = [
        {
            "id": 1,
            "query": "Bộ điều khiển nhiệt độ E5CSL-QTC AC100-240",
            "expected_code": "3.3381",
            "expected_name_contains": "Bộ điều khiển nhiệt độ",
            "expected_tier": "GENERIC_MATCH"
        },
        {
            "id": 2,
            "query": "Khởi động từ MC-100a AC220V",
            "expected_code": "3.1221",
            "expected_name_contains": "Contactor 100A",
            "expected_tier": "EXACT_MATCH"
        },
        {
            "id": 3,
            "query": "Khởi động từ MC-150a AC220V",
            "expected_code": "2.1186",
            "expected_name_contains": "Contactor 150A",
            "expected_tier": "EXACT_MATCH"
        },
        {
            "id": 4,
            "query": "Rơ le nhiệt MT-32 (9-13A)",
            "expected_code": "3.1236",
            "expected_name_contains": "Rơle 9_13A",
            "expected_tier": "EXACT_MATCH"
        },
        {
            "id": 5,
            "query": "Bạc đạn cầu 6205 nắp sắt",
            "expected_code": "3.9278",
            "expected_name_contains": "6205",
            "expected_tier": "EXACT_MATCH"
        },
        {
            "id": 6,
            "query": "Phớt chắn dầu 35x58x8",
            "expected_code": "3.2902",
            "expected_name_contains": "35*58*8",
            "expected_tier": "EXACT_MATCH"
        },
        {
            "id": 7,
            "query": "Van cầu hơi inox lắp bích DN25",
            "expected_code": "3.7798",
            "expected_name_contains": "DN 25 - phi 34",
            "expected_tier": "EXACT_MATCH"
        },
        {
            "id": 8,
            "query": "Bu lông M16x100",
            "expected_code": "3.0763",
            "expected_name_contains": "16*100",
            "expected_tier": "EXACT_MATCH"
        }
    ]

    all_passed = True
    for tc in test_cases:
        res = matcher.match_item(tc["query"], top_k=3)
        top1 = res["top_candidates"][0] if res["top_candidates"] else None

        if not top1:
            print(f"❌ TEST #{tc['id']} THẤT BẠI: '{tc['query']}' -> Không tìm thấy ứng viên!")
            all_passed = False
            continue

        actual_code = str(top1["ma_hang"]).strip()
        actual_name = str(top1["ten_hang"]).strip()
        actual_tier = top1["tier"]
        actual_score = top1["score"]

        is_code_ok = (tc["expected_code"] in actual_code)
        is_name_ok = (tc["expected_name_contains"].lower() in actual_name.lower())

        if is_code_ok or is_name_ok:
            print(f"✅ TEST #{tc['id']} ĐẠT (100%): '{tc['query']}'")
            print(f"   -> Top 1: Mã {actual_code} | {actual_name} | Điểm: {actual_score} [{actual_tier}]")
            print(f"   -> Lý do khớp: {', '.join(top1['reasons'])}")
        else:
            print(f"❌ TEST #{tc['id']} THẤT BẠI: '{tc['query']}'")
            print(f"   -> Kỳ vọng: Mã {tc['expected_code']} ({tc['expected_name_contains']})")
            print(f"   -> Thực tế: Mã {actual_code} ({actual_name}) | Điểm: {actual_score}")
            all_passed = False

    print("=" * 70)
    if all_passed:
        print("🎉 TẤT CẢ 8/8 TEST CASES ĐÃ VƯỢT QUA 100% XUẤT SẮC!")
    else:
        print("⚠️ CẦN TINH CHỈNH THÊM MỘT VÀI TRƯỜNG HỢP CHƯA ĐẠT.")
    print("=" * 70)
    return all_passed


if __name__ == "__main__":
    run_unit_tests()
