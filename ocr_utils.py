# ocr_utils.py — 2025-08-26
# 전화번호 인식 우선:
#  - 010으로 시작 → 항상 "010-가운데-****" 로 마스킹
#  - 05xx로 시작 → "05xx-가운데-마지막" 12자리 그대로
#  - 한 줄에 여러 개면 "가장 먼저 등장한 것" 하나만 사용
# 대여자명 = 전화번호"바로 앞"의 한글 2~8자(없으면 윗줄에서 추정)
# 주소 = 대여자명 바로 아래 줄만

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

# 외부에서 Tesseract 경로 주입 가능
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# -------------------------------
# 이미지 전처리 / OCR
# -------------------------------
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

# -------------------------------
# 정규식 / 라벨
# -------------------------------
R_010  = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(?:\d{0,4}|\*{4})")
R_05XX = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")

LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _clean(s:str) -> str:
    return re.sub(r"[|\[\]{}<>]+"," ",s).strip()

def _first_phone_in_line(line:str):
    """같은 라인에서 010/05xx 중 '가장 먼저' 등장한 것"""
    m010  = R_010.search(line)
    m05xx = R_05XX.search(line)
    if m010 and m05xx:
        return m010 if m010.start() < m05xx.start() else m05xx
    return m010 or m05xx

def _format_phone(m)->str:
    if not m: return ""
    if m.re is R_010:
        mid = m.group(2)
        return f"010-{mid}-****"
    else:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

def _name_left_of_phone(line:str, m)->str:
    left = line[:m.start()]
    left = LABEL_NAME.sub("", left).strip()
    k = re.findall(r"[가-힣]{2,8}", left)
    return k[-1] if k else ""

# -------------------------------
# 필드 파싱 (경량화: 주소 1줄, 송장 건너뜀)
# -------------------------------
def _parse_fields(lines: List[str]) -> dict:
    clean = [_clean(x) for x in lines]

    phone, name, addr = "", "", ""
    phone_i, name_i = -1, -1

    # 1) 전화번호(첫 발견) + 같은 줄 이름
    for i, ln in enumerate(clean):
        m = _first_phone_in_line(ln)
        if not m:
            continue
        phone = _format_phone(m)
        phone_i = i
        n = _name_left_of_phone(ln, m)
        if n:
            name, name_i = n, i
        break

    # 윗줄에서 이름 추정
    if phone and not name and phone_i > 0:
        up = LABEL_NAME.sub("", clean[phone_i-1]).strip()
        k = re.findall(r"[가-힣]{2,8}", up)
        if k:
            name, name_i = k[-1], phone_i-1

    # 2) 주소: 이름/전화 줄의 바로 아래 1줄만
    base_i = name_i if name else phone_i
    if base_i >= 0 and base_i+1 < len(clean):
        addr = LABEL_ADDR.sub("", clean[base_i+1]).strip()

    # 3) 송장번호: 성능을 위해 읽지 않음(빈칸)
    invoice = ""

    return {"전화번호": phone, "대여자명": name, "주소": addr, "송장번호": invoice}

# -------------------------------
# QR → 기종/기기번호
# -------------------------------
def _map_model_device(qr_text:str)->Tuple[str,str]:
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM":"심포니","LT":"락티나","S":"스윙","M":"스윙맥스","F":"프리스타일","G":"각시밀","C":"시밀레"}
    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2: return MAP.get(m2.group(1), "-"), m2.group(2)
    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1: return MAP.get(m1.group(1), "-"), raw
    return "-", ""

# -------------------------------
# ROI 추출
# -------------------------------
def _roi_cv2(path:str):
    if not HAS_CV2: return None
    try:
        img = cv2.imread(path); h,w = img.shape[:2]
        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        m1 = cv2.inRange(hsv, (85,80,60),(110,255,255))
        m2 = cv2.inRange(hsv, (110,80,60),(135,255,255))
        mask = cv2.bitwise_or(m1,m2)
        cnts,_ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not cnts: return None
        x,y,w0,h0 = cv2.boundingRect(max(cnts,key=cv2.contourArea))
        x1=max(0,x-int(0.02*w)); y1=max(0,y-int(0.05*h))
        x2=min(w,x+w0+int(0.72*w)); y2=min(h,y+h0+int(0.25*h))
        roi = img[y1:y2, x1:x2]
        if roi.size==0: return None
        return Image.fromarray(cv2.cvtColor(roi, cv2.COLOR_BGR2RGB))
    except Exception:
        return None

def _roi_ratio(path:str)->Image.Image:
    """OpenCV 미사용 시: 화면 비율 기반 중앙~하단 영역만 자르기(속도↑)."""
    im = Image.open(path); W, H = im.size
    x1 = int(W * 0.05)
    y1 = int(H * 0.30)
    x2 = int(W * 0.90)
    y2 = int(H * 0.78)
    return im.crop((x1, y1, x2, y2))

def _crop_roi(path:str)->Image.Image:
    roi = _roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _roi_ratio(path)

# -------------------------------
# OCR 본문 (빠른 버전: 재시도 없음)
# -------------------------------
def _try_ocr(path:str)->str:
    roi = _crop_roi(path)
    roi = _resize(roi, 1200)  # 더 작게 → 속도↑
    t = _ocr_text(_preprocess(roi, False), psm=6)
    return t

# -------------------------------
# 공개 API
# -------------------------------
def make_final_entry(qr_text:str, 송장_image_path:str):
    txt = _try_ocr(송장_image_path)
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

    result = {
        "출고일": ship_date,
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
        "송장번호": parsed.get("송장번호",""),  # 정책상 항상 빈칸
    }
    return result

def make_final_entry_fast(qr_text:str, 송장_image_path:str):
    """
    초고속 프리뷰:
    - 중앙~하단 띠 숫자 전용으로 OCR
    - 전화번호만 빠르게 추출
    """
    roi = _crop_roi(송장_image_path)
    roi = _resize(roi, 900)

    # 전화가 주로 있는 중앙~하단 띠만 스캔
    W, H = roi.size
    band = roi.crop((int(W*0.05), int(H*0.40), int(W*0.95), int(H*0.78)))

    # 숫자/하이픈 전용 (빠름)
    cfg_fast = "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789-"
    txt = pytesseract.image_to_string(_preprocess(band, strong=False), config=cfg_fast, lang="eng")
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]

    phone = ""
    for ln in lines:
        m010  = R_010.search(ln)
        m05xx = R_05XX.search(ln)
        # 한 줄에 둘 다 있으면 앞에 나온 쪽 채택
        m = m010 if (m010 and (not m05xx or m010.start() < m05xx.start())) else (m05xx or None)
        if not m:
            continue
        if m.re is R_010:
            mid = m.group(2)  # 가운데 3~4자리
            phone = f"010-{mid}-****"
        else:
            phone = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        break  # "가장 먼저" 하나만 사용

    ship_date = date.today().isoformat()
    model, device_id = _map_model_device(qr_text)

    result = {
        "출고일": ship_date,
        "대여자명": "",
        "전화번호": phone,
        "주소": "",
        "기기번호": device_id,
        "기종": model,
        "송장번호": "",  # 프리뷰에서도 읽지 않음
    }
    return result
