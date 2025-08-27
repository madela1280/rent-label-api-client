# ocr_utils.py — 복구 전용 / 단순·안정화 버전
# - 송장번호 전면 제거 (속도 ↑)
# - 전화번호 기준으로 대여자명(왼쪽), 주소(다음 줄 1개)만 추출
# - 특정 금지 번호(010-7394-3535) 무시
# - QR에서만 기종/기기번호 추출(항상 표시 유지)

import os, re
from datetime import date
from typing import List, Tuple
from PIL import Image, ImageOps, ImageFilter
import pytesseract

try:
    import cv2
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# Tesseract 경로 환경변수 우선 사용
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# ---------------- 기본 처리 ----------------
def _preprocess(img: Image.Image, strong: bool=False) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=220, threshold=2))
        g = g.point(lambda x: 255 if x > 170 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g

def _ocr_text(img: Image.Image, psm:int=6) -> str:
    try:
        return pytesseract.image_to_string(img, config=f"--oem 3 --psm {psm}", lang="kor+eng")
    except Exception:
        return ""

def _resize(img: Image.Image, max_w:int=1400) -> Image.Image:
    w,h = img.size
    if w > max_w:
        s = max_w/float(w)
        return img.resize((max_w, int(h*s)))
    return img

# ---------------- 규칙 ----------------
R_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)
ADDR_TOKENS = ("시","군","구","읍","면","동","리","로","길","번길","아파트","빌라","호","단지")
BANNED_PHONES = {"010-7394-3535"}

def _clean(s:str) -> str:
    return re.sub(r"[|\[\]{}<>]+"," ",s).strip()

def _looks_like_address(s:str)->bool:
    s2 = LABEL_ADDR.sub("", s)
    return any(t in s2 for t in ADDR_TOKENS) or bool(re.search(r"\d|\(|\)", s2))

def _parse_fields(lines: List[str]) -> dict:
    """전화번호가 있는 줄을 기준으로: 왼쪽=대여자명(한글 2~8), 다음 줄=주소(1줄만)."""
    lines = [_clean(x) for x in lines if x and x.strip()]
    phone, name, addr = "", "", ""

    for i, ln in enumerate(lines):
        m = R_010.search(ln)
        if not m:
            continue

        # 전화번호 마스킹 (010-가운데-****)
        mid = m.group(2)
        phone = f"010-{mid}-****"
        # 금지 번호는 무시
        if phone in BANNED_PHONES:
            phone = ""

        # 같은 줄 왼쪽에서 대여자명
        left = LABEL_NAME.sub("", ln[:m.start()]).strip()
        k = re.findall(r"[가-힣]{2,8}", left)
        if k:
            name = k[-1]

        # 바로 아래 줄 1줄만 주소 후보
        if i + 1 < len(lines):
            cand = LABEL_ADDR.sub("", lines[i+1]).strip()
            if _looks_like_address(cand) or len(cand) >= 6:
                addr = cand

        break  # 가장 먼저 찾은 전화 한 번만 사용

    return {"대여자명": name, "전화번호": phone, "주소": addr}

# ---------------- QR → 기종/기기번호 ----------------
def _map_model_device(qr_text:str)->Tuple[str,str]:
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

# ---------------- 메인 엔트리 ----------------
def make_final_entry(qr_text:str, img_path:str):
    im = Image.open(img_path)
    im = _resize(im, 1400)
    txt = _ocr_text(_preprocess(im, False), psm=6)
    if len(re.sub(r"\s+","",txt)) < 16:
        # 너무 빈약하면 강처리
        txt2 = _ocr_text(_preprocess(im, True), psm=6)
        if len(re.sub(r"\s+","",txt2)) > len(re.sub(r"\s+","",txt)):
            txt = txt2

    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    parsed = _parse_fields(lines)
    model, device_id = _map_model_device(qr_text)

    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
    }

def make_final_entry_fast(qr_text:str, img_path:str):
    # 프리뷰: 해상도만 줄이고 동일 규칙 적용(빠름)
    im = Image.open(img_path)
    im = _resize(im, 900)
    txt = _ocr_text(_preprocess(im, False), psm=6)
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    parsed = _parse_fields(lines)
    model, device_id = _map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
    }


