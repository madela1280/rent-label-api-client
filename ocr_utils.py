# ocr_utils.py — 2025-08-26
# 전화 인식 우선:
#  - 010으로 시작 → 항상 "010-가운데-****" 로 마스킹
#  - 05xx로 시작 → "05xx-가운데-마지막" 12자리 그대로
#  - 한 줄에 여러 개면 "가장 먼저 등장한 것" 하나만 사용
# 이름 = 전화 "바로 앞"의 한글 2~8자(없으면 윗줄에서 추정)
# 주소 = 이름 바로 아래 줄부터, 필요 시 다음 줄 이어붙임
# 송장번호 = 이름 한 줄 위에서 ####-####-#### 우선

import os, re
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
    if w>max_w:
        s = max_w/float(w)
        return img.resize((max_w, int(h*s)))
    return img

# -------- 전화/이름/주소/송장 --------
R_010  = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(?:\d{0,4}|\*{4})")
R_05XX = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
R_INVOICE12 = re.compile(r"\b\d{4}[-\s]?\d{4}[-\s]?\d{4}\b")
LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)
ADDR_TOKENS = ("시","군","구","읍","면","동","리","로","길","번길","아파트","빌라","호","단지")

def _clean(s:str) -> str:
    return re.sub(r"[|\[\]{}<>]+"," ",s).strip()

def _looks_like_address(s:str)->bool:
    s2 = LABEL_ADDR.sub("", s)
    return any(t in s2 for t in ADDR_TOKENS) or bool(re.search(r"\d|\(|\)", s2))

def _first_phone_in_line(line:str):
    """같은 라인에서 010/05xx 중 '가장 먼저' 등장한 것 반환"""
    m010  = R_010.search(line)
    m05xx = R_05XX.search(line)
    if m010 and m05xx:
        return m010 if m010.start() < m05xx.start() else m05xx
    return m010 or m05xx

def _format_phone(m)->str:
    """정규식 매치 → 포맷팅 규칙 적용"""
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

def _parse_fields(lines: List[str]) -> dict:
    clean = [_clean(x) for x in lines]

    phone, name, addr, invoice = "", "", "", ""
    phone_i, name_i = -1, -1

    # 1) 전화: 텍스트 순서상 '첫 번째'
    for i, ln in enumerate(clean):
        m = _first_phone_in_line(ln)
        if not m: continue
        phone = _format_phone(m)
        phone_i = i
        # 같은 줄에서 이름 시도
        n = _name_left_of_phone(ln, m)
        if n: name, name_i = n, i
        break

    # 못 찾으면 윗줄에서 이름 추정
    if phone and not name and phone_i > 0:
        up = LABEL_NAME.sub("", clean[phone_i-1]).strip()
        k = re.findall(r"[가-힣]{2,8}", up)
        if k: name, name_i = k[-1], phone_i-1

    # 2) 주소: 이름 줄 기준 아래 1~2줄
    base_i = name_i if name else phone_i
    if base_i >= 0:
        first  = LABEL_ADDR.sub("", clean[base_i+1]).strip() if base_i+1 < len(clean) else ""
        second = LABEL_ADDR.sub("", clean[base_i+2]).strip() if base_i+2 < len(clean) else ""
        parts=[]
        if first: parts.append(first)
        if second and _looks_like_address(second): parts.append(second)
        addr = " ".join(p.strip() for p in parts if p).strip()

    # 3) 송장번호: 이름 '한 줄 위' 우선
    if name and name_i>0:
        up = clean[name_i-1].replace(" ","")
        m = R_INVOICE12.search(up)
        if m: invoice = m.group(0)
    if not invoice:
        for ln in clean:
            m = R_INVOICE12.search(ln.replace(" ",""))
            if m: invoice = m.group(0); break

    return {"전화번호":phone, "대여자명":name, "주소":addr, "송장번호":invoice}

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

# -------- ROI --------
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
    im = Image.open(path); W,H = im.size
    return im.crop((int(W*0.04), int(H*0.25), int(W*0.90), int(H*0.82)))

def _crop_roi(path:str)->Image.Image:
    roi = _roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _roi_ratio(path)

# -------- 메인 --------
def _try_ocr(path:str)->str:
    roi = _crop_roi(path); roi = _resize(roi, 1400)
    t = _ocr_text(_preprocess(roi, False), psm=6)
    if len(re.sub(r"\s+","",t)) < 18:
        t2 = _ocr_text(_preprocess(roi, True), psm=6)
        if len(re.sub(r"\s+","",t2)) > len(re.sub(r"\s+","",t)): t = t2
    # 너무 빈약하면 전체 이미지
    parsed = _parse_fields([ln for ln in t.splitlines() if ln.strip()])
    if not (parsed.get("전화번호") or parsed.get("주소") or parsed.get("대여자명") or parsed.get("송장번호")):
        full = Image.open(path); full = _resize(full, 1600)
        t3 = _ocr_text(_preprocess(full, False), psm=6)
        if len(re.sub(r"\s+","",t3)) > len(re.sub(r"\s+","",t)): t = t3
    return t

def make_final_entry(qr_text:str, 송장_image_path:str):
    txt = _try_ocr(송장_image_path)
    try:
        os.makedirs("_debug", exist_ok=True)
        with open(os.path.join("_debug","ocr_full.txt"),"w",encoding="utf-8") as f: f.write(txt)
    except Exception: pass

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

def make_final_entry_fast(qr_text:str, 송장_image_path:str):
    roi = _crop_roi(송장_image_path); roi = _resize(roi, 1100)
    txt = _ocr_text(_preprocess(roi, False), psm=6)
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
