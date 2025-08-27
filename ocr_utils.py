# ocr_utils.py — 2025-08-27
# 목적: 전화번호 / 전화번호 '앞' 대여자명 / 전화번호 '아래 한 줄' 주소(첫 줄)만 신속 추출
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

pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

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

def _resize(img: Image.Image, max_w:int=1200) -> Image.Image:
    w,h = img.size
    if w>max_w:
        s = max_w/float(w)
        return img.resize((max_w, int(h*s)))
    return img

# -------- 패턴 --------
R_010  = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
R_05XX = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _clean(s:str) -> str:
    return re.sub(r"[|\[\]{}<>]+"," ",s).strip()

def _first_phone_in_line(line:str):
    m010  = R_010.search(line)
    m05xx = R_05XX.search(line)
    if m010 and m05xx:
        return m010 if m010.start() < m05xx.start() else m05xx
    return m010 or m05xx

def _format_phone(m)->str:
    if not m: return ""
    if m.re is R_010:
        return f"010-{m.group(2)}-{m.group(3)}"
    else:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

def _name_left_of_phone(line:str, m)->str:
    left = line[:m.start()]
    left = LABEL_NAME.sub("", left).strip()
    k = re.findall(r"[가-힣]{2,8}", left)
    return k[-1] if k else ""

def _parse_fields(lines: List[str]) -> dict:
    clean = [_clean(x) for x in lines]

    phone, name, addr = "", "", ""
    phone_i, name_i = -1, -1

    for i, ln in enumerate(clean):
        m = _first_phone_in_line(ln)
        if not m: continue
        phone = _format_phone(m)
        phone_i = i
        n = _name_left_of_phone(ln, m)
        if n: name, name_i = n, i
        break

    if phone and not name and phone_i > 0:
        up = LABEL_NAME.sub("", clean[phone_i-1]).strip()
        k = re.findall(r"[가-힣]{2,8}", up)
        if k: name, name_i = k[-1], phone_i-1

    if phone_i >= 0 and phone_i+1 < len(clean):
        addr = LABEL_ADDR.sub("", clean[phone_i+1]).strip()

    return {"전화번호":phone, "대여자명":name, "주소":addr}

# -------- ROI --------
def _roi_cv2(path:str):
    if not HAS_CV2: return None
    try:
        img = cv2.imread(path); h,w = img.shape[:2]
        # 송장 하단 텍스트 띠 대략 잡기 (중앙~하단)
        y1 = int(h*0.35); y2 = int(h*0.85); x1 = int(w*0.05); x2 = int(w*0.95)
        roi = img[y1:y2, x1:x2]
        if roi.size==0: return None
        return Image.fromarray(cv2.cvtColor(roi, cv2.COLOR_BGR2RGB))
    except Exception:
        return None

def _roi_ratio(path:str)->Image.Image:
    im = Image.open(path); W, H = im.size
    x1 = int(W * 0.05); x2 = int(W * 0.95)
    y1 = int(H * 0.35); y2 = int(H * 0.85)
    return im.crop((x1, y1, x2, y2))

def _crop_roi(path:str)->Image.Image:
    roi = _roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _roi_ratio(path)

# -------- 메인 --------
def _try_ocr(path:str)->str:
    # 1) ROI 약 → 강
    roi = _resize(_crop_roi(path), 1200)
    t = _ocr_text(_preprocess(roi, False), psm=6)
    if len(re.sub(r"\s+","",t)) < 14:
        t2 = _ocr_text(_preprocess(roi, True), psm=6)
        if len(re.sub(r"\s+","",t2)) > len(re.sub(r"\s+","",t)): t = t2
    # 2) 부족하면 전체 이미지 한 번 더(강)
    if len(re.sub(r"\s+","",t)) < 14:
        full = _resize(Image.open(path), 1400)
        t3 = _ocr_text(_preprocess(full, True), psm=6)
        if len(re.sub(r"\s+","",t3)) > len(re.sub(r"\s+","",t)): t = t3
    return t

def make_final_entry(qr_text:str, path:str):
    txt = _try_ocr(path)
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

def make_final_entry_fast(qr_text:str, path:str):
    # ROI만 강처리로 초고속
    roi = _resize(_crop_roi(path), 900)
    txt = _ocr_text(_preprocess(roi, True), psm=6)
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

# -------- QR → 기종/기기번호 --------
def _map_model_device(qr_text:str)->Tuple[str,str]:
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM":"심포니","LT":"락티나","S":"스윙","M":"스윙맥스","F":"프리스타일","G":"각시밀","C":"시밀레"}
    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2: return MAP.get(m2.group(1), "-"), m2.group(2)
    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1: return MAP.get(m1.group(1), "-"), raw
    return "-", ""


