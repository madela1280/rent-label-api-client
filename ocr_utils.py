# ocr_utils.py
# ------------------------------------------------------------
# 전화/이름/주소 규칙 강화:
# - 전화: 010-1234-**** / 010-1234-5678 / 0507-1234-5678 / 05xx-... → 첫 번째만 채택
# - 대여자명: 전화 '바로 앞'의 한글 2~8자
# - 주소: 이름 바로 아래줄 우선, 없으면 전화 아래줄 + 다음 줄 이어붙임(괄호 포함)
# - ROI 실패 시 전체 이미지 폴백
# ------------------------------------------------------------

import os
import re
from datetime import date
from typing import List, Tuple, Optional

from PIL import Image, ImageOps, ImageFilter
import pytesseract

try:
    import cv2
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# =========================
# 1) 전처리 & OCR
# =========================
def _preprocess(img: Image.Image, strong: bool = False) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=220, threshold=2))
        g = g.point(lambda x: 255 if x > 170 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g

def _ocr_text(img: Image.Image, allow_kor: bool = True, psm: int = 6) -> str:
    cfg = f"--oem 3 --psm {psm}"
    lang = "kor+eng" if allow_kor else "eng"
    try:
        return pytesseract.image_to_string(img, config=cfg, lang=lang)
    except Exception:
        return ""

def _resize_image(img: Image.Image, max_w: int = 1400) -> Image.Image:
    w, h = img.size
    if w > max_w:
        scale = max_w / float(w)
        return img.resize((max_w, int(h * scale)))
    return img

# =========================
# 2) 정규식/도우미
# =========================
# 전화: 010 / 05xx (0507 포함) + 3~4 + (4 or ****)
PHONE_RE = re.compile(
    r"(?:\(?0?1[016789]\)?|0?5\d{2})[-\s\.]?\d{3,4}[-\s\.]?(?:\d{4}|\*{4})"
)

INVOICE_RE = re.compile(
    r"(?:운송장|송장|송장번호|운송장번호|운송장\s*No\.?|Invoice|Tracking)\D{0,8}([0-9\-\s]{10,18})",
    re.IGNORECASE,
)
MAYBE_INVOICE_FALLBACK = re.compile(r"\b\d[\d\-\s]{9,17}\d\b")

LABEL_NAME_RE = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR_RE = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _clean(s: str) -> str:
    return re.sub(r"[|\[\]{}<>]+", " ", s).strip()

def _normalize_phone(s: str) -> str:
    m = PHONE_RE.search(s)
    if not m:
        return ""
    ph = m.group()
    ph = re.sub(r"[^\d\*]", "-", ph)
    ph = re.sub(r"-{2,}", "-", ph).strip("-")
    return ph

def _likely_name_immediate(left: str) -> Optional[str]:
    """
    전화 바로 앞에서 한글 2~8자 블록만 추출.
    예: '홍길동 010-1234-5678' → '홍길동'
    """
    left = LABEL_NAME_RE.sub("", left).strip()
    m = re.search(r"([가-힣]{2,8})\s*$", left)
    return m.group(1) if m else None

def _is_address(s: str) -> bool:
    s2 = LABEL_ADDR_RE.sub("", s)
    return any(tok in s2 for tok in ["시", "군", "구", "읍", "면", "동", "리", "로", "길", "번길", "아파트", "빌라", "호"])

# =========================
# 3) ROI 추출
# =========================
def _crop_invoice_roi_cv2(path: str):
    if not HAS_CV2:
        return None
    try:
        img = cv2.imread(path)
        h, w = img.shape[:2]

        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        m1 = cv2.inRange(hsv, (85, 80, 60), (110, 255, 255))
        m2 = cv2.inRange(hsv, (110, 80, 60), (135, 255, 255))
        mask = cv2.bitwise_or(m1, m2)

        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return None
        cnt = max(contours, key=cv2.contourArea)
        x, y, w0, h0 = cv2.boundingRect(cnt)

        x1 = max(0, x - int(0.02 * w))
        y1 = max(0, y - int(0.05 * h))
        x2 = min(w, x + w0 + int(0.72 * w))
        y2 = min(h, y + h0 + int(0.25 * h))

        roi = img[y1:y2, x1:x2]
        if roi.size == 0:
            return None
        return Image.fromarray(cv2.cvtColor(roi, cv2.COLOR_BGR2RGB))
    except Exception:
        return None

def _crop_invoice_roi_ratio(path: str) -> Image.Image:
    im = Image.open(path)
    W, H = im.size
    x1 = int(W * 0.04)
    y1 = int(H * 0.25)
    x2 = int(W * 0.86)
    y2 = int(H * 0.80)
    return im.crop((x1, y1, x2, y2))

def _crop_invoice_roi(path: str) -> Image.Image:
    roi = _crop_invoice_roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _crop_invoice_roi_ratio(path)

# =========================
# 4) 텍스트 → 필드 파싱
# =========================
def _extract_invoice(text: str) -> str:
    m = INVOICE_RE.search(text)
    if m:
        raw = m.group(1)
        digits = re.sub(r"\D", "", raw)
        if 10 <= len(digits) <= 18:
            return digits
    m2 = MAYBE_INVOICE_FALLBACK.search(text.replace(" ", ""))
    if m2:
        digits = re.sub(r"\D", "", m2.group())
        if 10 <= len(digits) <= 18:
            return digits
    return ""

def _parse_fields(text: str) -> dict:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    clean_lines = [_clean(ln) for ln in lines]

    phone, name, addr, invoice = "", "", "", ""
    phone_idx = -1
    name_idx = -1

    # 1) 전화: 처음 발견되는 1개만 (한 줄에 2개 있어도 첫 번째만)
    for i, ln in enumerate(clean_lines):
        ph = _normalize_phone(ln)
        if ph:
            phone = ph
            phone_idx = i
            break

    # 2) 이름: 전화 줄에서 '바로 앞' 한글 2~8자, 실패 시 윗줄
    if phone and phone_idx >= 0:
        ln = clean_lines[phone_idx]
        # 전화 패턴 위치 재탐색 (split 안전)
        m = PHONE_RE.search(ln)
        left = ln[:m.start()].strip() if m else ln
        nm = _likely_name_immediate(left)
        if nm:
            name = nm
            name_idx = phone_idx
        elif phone_idx - 1 >= 0:
            up = LABEL_NAME_RE.sub("", clean_lines[phone_idx - 1]).strip()
            nm2 = _likely_name_immediate(up)
            if nm2:
                name = nm2
                name_idx = phone_idx - 1

    # 3) 주소: 이름 아래줄 우선, 없으면 전화 아래줄
    addr_candidates = []
    if name and name_idx >= 0:
        if name_idx + 1 < len(clean_lines):
            addr_candidates.append(clean_lines[name_idx + 1])
        if name_idx + 2 < len(clean_lines):
            addr_candidates.append(clean_lines[name_idx + 2])
    elif phone_idx >= 0:
        if phone_idx + 1 < len(clean_lines):
            addr_candidates.append(clean_lines[phone_idx + 1])
        if phone_idx + 2 < len(clean_lines):
            addr_candidates.append(clean_lines[phone_idx + 2])

    addr_parts: List[str] = []
    for j, cand in enumerate(addr_candidates[:2]):
        c = LABEL_ADDR_RE.sub("", cand).strip()
        if not c:
            continue
        if j == 0:
            # 첫 줄은 그대로
            addr_parts.append(c)
        else:
            # 두 번째 줄은 주소성 키워드/숫자/괄호가 있으면 이어붙임
            if _is_address(c) or re.search(r"[0-9()]", c):
                addr_parts.append(c)
    addr = " ".join(addr_parts).strip()

    # 4) 송장번호
    invoice = _extract_invoice("\n".join(clean_lines))

    return {"전화번호": phone, "대여자명": name, "주소": addr, "송장번호": invoice}

# =========================
# 5) QR → 기종/기기번호
# =========================
def _map_model_device(qr_text: str) -> Tuple[str, str]:
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM": "심포니", "LT": "락티나", "S": "스윙", "M": "스윙맥스", "F": "프리스타일", "G": "각시밀", "C": "시밀레"}

    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2:
        return MAP.get(m2.group(1), "-"), m2.group(2)

    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1:
        return MAP.get(m1.group(1), "-"), raw

    return "-", ""

# =========================
# 6) OCR 전략
# =========================
def _try_ocr_strategies(img_path: str) -> str:
    roi = _crop_invoice_roi(img_path)
    roi = _resize_image(roi, 1400)

    txt = _ocr_text(_preprocess(roi, strong=False), allow_kor=True, psm=6)
    if len(re.sub(r"\s+", "", txt)) < 18:
        t2 = _ocr_text(_preprocess(roi, strong=True), allow_kor=True, psm=6)
        if len(re.sub(r"\s+", "", t2)) > len(re.sub(r"\s+", "", txt)):
            txt = t2

    parsed = _parse_fields(txt)
    if parsed.get("전화번호") or parsed.get("주소") or parsed.get("송장번호"):
        return txt

    full = Image.open(img_path)
    full = _resize_image(full, 1600)
    txf = _ocr_text(_preprocess(full, strong=False), allow_kor=True, psm=6)
    if len(re.sub(r"\s+", "", txf)) < 18:
        t2 = _ocr_text(_preprocess(full, strong=True), allow_kor=True, psm=6)
        if len(re.sub(r"\s+", "", t2)) > len(re.sub(r"\s+", "", txf)):
            txf = t2
    return txf

# =========================
# 7) 엔트리
# =========================
def make_final_entry(qr_text: str, 송장_image_path: str):
    txt = _try_ocr_strategies(송장_image_path)

    try:
        os.makedirs("_debug", exist_ok=True)
        with open(os.path.join("_debug", "ocr_lines_full.txt"), "w", encoding="utf-8") as f:
            f.write(txt)
    except Exception:
        pass

    parsed = _parse_fields(txt)
    model, device_id = _map_model_device(qr_text)
    ship_date = date.today().isoformat()

    out = {
        "출고일": ship_date,
        "대여자명": parsed.get("대여자명", ""),
        "전화번호": parsed.get("전화번호", ""),
        "주소": parsed.get("주소", ""),
        "기기번호": device_id,
        "기종": model,
        "송장번호": parsed.get("송장번호", ""),
    }

    # 보정: 전화는 있는데 이름이 비었을 때, 같은 줄의 '바로 앞' 재시도
    if out["전화번호"] and not out["대여자명"]:
        for ln in [ln.strip() for ln in txt.splitlines() if ln.strip()]:
            m = PHONE_RE.search(ln)
            if not m:
                continue
            left = ln[:m.start()].strip()
            nm = _likely_name_immediate(left)
            if nm:
                out["대여자명"] = nm
                break

    return out

def make_final_entry_fast(qr_text: str, 송장_image_path: str):
    roi = _crop_invoice_roi(송장_image_path)
    roi = _resize_image(roi, 1000)
    txt = _ocr_text(_preprocess(roi, strong=False), allow_kor=True, psm=6)

    parsed = _parse_fields(txt)
    if not (parsed.get("전화번호") or parsed.get("주소") or parsed.get("송장번호")):
        full = Image.open(송장_image_path)
        full = _resize_image(full, 1200)
        txt2 = _ocr_text(_preprocess(full, strong=False), allow_kor=True, psm=6)
        parsed = _parse_fields(txt2)

    model, device_id = _map_model_device(qr_text)
    ship_date = date.today().isoformat()

    return {
        "출고일": ship_date,
        "대여자명": parsed.get("대여자명", ""),
        "전화번호": parsed.get("전화번호", ""),
        "주소": parsed.get("주소", ""),
        "기기번호": device_id,
        "기종": model,
        "송장번호": parsed.get("송장번호", ""),
    }
