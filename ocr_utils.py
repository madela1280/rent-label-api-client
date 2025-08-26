# ocr_utils.py
# ------------------------------------------------------------
# 요구사항 반영(2025-08-26)
# - 전화: 010-1234-****, 010-1234-5678, 05xx-1234-5678 형식 지원
#   * 라인에 여러 개 있어도 "첫 번째"만 사용
# - 대여자명: 전화 "바로 앞"의 한글 2~8자
# - 주소: 대여자명 바로 아래 줄을 시작으로, 필요 시 다음 줄까지 이어붙임
# - 송장번호: 대여자명 "한 줄 위"에서 12자리(####-####-####) 우선 추출
# - ROI 실패 시 전체 이미지 OCR로 폴백
# ------------------------------------------------------------

import os
import re
from datetime import date
from typing import List, Tuple, Optional

from PIL import Image, ImageOps, ImageFilter
import pytesseract

# 선택: OpenCV가 있으면 ROI 탐지에 활용
try:
    import cv2
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# 시스템별 테서랙트 경로 지정(없으면 기본값)
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
        # 가벼운 이진화
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

def _resize(img: Image.Image, max_w: int = 1400) -> Image.Image:
    w, h = img.size
    if w > max_w:
        s = max_w / float(w)
        return img.resize((max_w, int(h * s)))
    return img

# =========================
# 2) 정규식 & 도우미
# =========================
# 전화: 010 또는 05xx 시작, 마지막 4자리 또는 **** 허용
PHONE_RE = re.compile(
    r"(?:010|05\d{2})[-\s\.]?\d{3,4}[-\s\.]?(?:\d{4}|\*{4})"
)

# 송장번호: 우선 ####-####-#### (정확 12자리)
INVOICE_12_RE = re.compile(r"\b\d{4}[-\s]?\d{4}[-\s]?\d{4}\b")

LABEL_NAME_RE = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR_RE = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _clean(s: str) -> str:
    return re.sub(r"[|\[\]{}<>]+", " ", s).strip()

def _normalize_phone(line: str) -> Optional[re.Match]:
    """해당 라인에서 '첫 번째' 전화 패턴만 반환"""
    return PHONE_RE.search(line)

def _name_left_of_phone(line: str, m: re.Match) -> str:
    """전화 바로 왼쪽의 한글 2~8자 블록만 추출"""
    left = line[:m.start()]
    left = LABEL_NAME_RE.sub("", left).strip()
    # 마지막 한글 블록(2~8) 취득
    k = re.findall(r"[가-힣]{2,8}", left)
    return k[-1] if k else ""

ADDR_TOKENS = ("시","군","구","읍","면","동","리","로","길","번길","아파트","빌라","호","단지")
def _looks_like_address(s: str) -> bool:
    s2 = LABEL_ADDR_RE.sub("", s)
    return any(t in s2 for t in ADDR_TOKENS) or bool(re.search(r"\d|\(|\)", s2))

# =========================
# 3) ROI
# =========================
def _crop_roi_cv2(path: str):
    if not HAS_CV2:
        return None
    try:
        img = cv2.imread(path)
        h, w = img.shape[:2]
        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        # 파란 세로띠 탐지(받는분)
        m1 = cv2.inRange(hsv, (85, 80, 60), (110, 255, 255))
        m2 = cv2.inRange(hsv, (110, 80, 60), (135, 255, 255))
        mask = cv2.bitwise_or(m1, m2)
        cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not cnts:
            return None
        x,y,w0,h0 = cv2.boundingRect(max(cnts, key=cv2.contourArea))
        x1 = max(0, x - int(0.02*w))
        y1 = max(0, y - int(0.05*h))
        x2 = min(w, x + w0 + int(0.72*w))
        y2 = min(h, y + h0 + int(0.25*h))
        roi = img[y1:y2, x1:x2]
        if roi.size == 0: return None
        return Image.fromarray(cv2.cvtColor(roi, cv2.COLOR_BGR2RGB))
    except Exception:
        return None

def _crop_roi_ratio(path: str) -> Image.Image:
    im = Image.open(path)
    W, H = im.size
    # 중앙~하단 넓게 (라벨 위치 일반값)
    x1 = int(W*0.04); y1 = int(H*0.25)
    x2 = int(W*0.90); y2 = int(H*0.82)
    return im.crop((x1,y1,x2,y2))

def _crop_roi(path: str) -> Image.Image:
    roi = _crop_roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _crop_roi_ratio(path)

# =========================
# 4) 텍스트 → 필드
# =========================
def _parse_fields(lines: List[str]) -> dict:
    """
    규칙:
    - 전화: 010-1234-****, 010-1234-5678, 05xx-1234-5678 → 라인에서 '첫 번째'만
    - 대여자명: 전화 '바로 앞'의 한글 2~8자. 없으면 이전 줄에서 같은 규칙 시도
    - 주소: 대여자명 바로 아래 줄을 시작으로, 필요 시 다음 줄까지 이어붙임
    - 송장번호: 대여자명 '한 줄 위'에서 ####-####-#### 우선 추출(없으면 전체에서 보조 탐색)
    """
    clean = [_clean(x) for x in lines]

    phone = ""
    name = ""
    addr  = ""
    invoice = ""

    phone_i = -1
    name_i  = -1

    # 1) 전화(첫 번째만)
    for i, ln in enumerate(clean):
        m = _normalize_phone(ln)
        if not m: 
            continue
        phone = re.sub(r"[^\d\*]", "-", m.group())
        phone = re.sub(r"-{2,}", "-", phone).strip("-")
        phone_i = i
        # 이름: 같은 줄에서 '바로 왼쪽'
        n = _name_left_of_phone(ln, m)
        if n:
            name = n
            name_i = i
        break

    # 같은 줄에서 못 찾았으면 '전화 바로 윗줄'에서 찾기
    if phone and not name and phone_i > 0:
        up = clean[phone_i-1]
        # 전화가 윗줄에도 있을 수 있으니 라벨 제거 후 끝쪽 한글블록
        up2 = LABEL_NAME_RE.sub("", up).strip()
        m2 = re.findall(r"[가-힣]{2,8}", up2)
        if m2:
            name = m2[-1]
            name_i = phone_i-1

    # 2) 주소: 이름 아래줄 우선, 없으면 전화 아래줄
    lines_for_addr_from = name_i if name else phone_i
    if lines_for_addr_from >= 0:
        first = clean[lines_for_addr_from+1] if lines_for_addr_from+1 < len(clean) else ""
        second= clean[lines_for_addr_from+2] if lines_for_addr_from+2 < len(clean) else ""
        first = LABEL_ADDR_RE.sub("", first).strip()
        second= LABEL_ADDR_RE.sub("", second).strip()

        parts = []
        if first:
            parts.append(first)
        if second and (_looks_like_address(second)):
            parts.append(second)
        addr = " ".join([p for p in parts if p]).strip()

    # 3) 송장번호: 이름 한 줄 위 우선
    if name and name_i > 0:
        cand = clean[name_i-1]
        m = INVOICE_12_RE.search(cand.replace(" ", ""))
        if m:
            invoice = m.group(0).replace(" ", "")
    # 전체에서 보조 탐색
    if not invoice:
        for ln in clean:
            m = INVOICE_12_RE.search(ln.replace(" ", ""))
            if m:
                invoice = m.group(0).replace(" ", "")
                break

    return {
        "전화번호": phone,
        "대여자명": name,
        "주소": addr,
        "송장번호": invoice,
    }

# =========================
# 5) QR → 기종/기기번호
# =========================
def _map_model_device(qr_text: str) -> Tuple[str, str]:
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM":"심포니","LT":"락티나","S":"스윙","M":"스윙맥스","F":"프리스타일","G":"각시밀","C":"시밀레"}

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
def _try_ocr(path: str) -> str:
    roi = _crop_roi(path)
    roi = _resize(roi, 1400)

    txt = _ocr_text(_preprocess(roi, strong=False), allow_kor=True, psm=6)
    # 텍스트가 짧으면 강하게 재시도
    if len(re.sub(r"\s+","",txt)) < 18:
        t2 = _ocr_text(_preprocess(roi, strong=True), allow_kor=True, psm=6)
        if len(re.sub(r"\s+","",t2)) > len(re.sub(r"\s+","",txt)):
            txt = t2

    # 내용이 너무 빈약하면 전체 이미지 시도
    parsed = _parse_fields([ln for ln in txt.splitlines() if ln.strip()])
    if not (parsed.get("전화번호") or parsed.get("주소") or parsed.get("송장번호") or parsed.get("대여자명")):
        full = Image.open(path)
        full = _resize(full, 1600)
        t3 = _ocr_text(_preprocess(full, strong=False), allow_kor=True, psm=6)
        if len(re.sub(r"\s+","",t3)) > len(re.sub(r"\s+","",txt)):
            txt = t3
    return txt

# =========================
# 7) 엔트리
# =========================
def make_final_entry(qr_text: str, 송장_image_path: str):
    txt = _try_ocr(송장_image_path)

    # 디버그 저장
    try:
        os.makedirs("_debug", exist_ok=True)
        with open(os.path.join("_debug","ocr_full.txt"),"w",encoding="utf-8") as f:
            f.write(txt)
    except Exception:
        pass

    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    parsed = _parse_fields(lines)

    model, device_id = _map_model_device(qr_text)
    ship_date = date.today().isoformat()

    out = {
        "출고일": ship_date,
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
        "송장번호": parsed.get("송장번호",""),
    }
    return out

def make_final_entry_fast(qr_text: str, 송장_image_path: str):
    """
    빠른 미리보기: ROI → 약한 OCR 1회로 전화/이름/주소/송장 후보만 뽑음
    """
    roi = _crop_roi(송장_image_path)
    roi = _resize(roi, 1100)
    txt = _ocr_text(_preprocess(roi, strong=False), allow_kor=True, psm=6)
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    parsed = _parse_fields(lines)

    model, device_id = _map_model_device(qr_text)
    ship_date = date.today().isoformat()
    return {
        "출고일": ship_date,
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
        "송장번호": parsed.get("송장번호",""),
    }

