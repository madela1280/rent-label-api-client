# ocr_utils.py — 2025-08-27 (전화번호 중심 파서, 송장번호 제거, 주소 1줄)
# - 속도: ROI 비율 크롭 + 적정 리사이즈
# - 추출 규칙(전화번호 → 이름 → 주소):
#   1) 전화번호: 화면에서 '가장 먼저' 보이는 010/05xx 하나만
#      - 010-가운데-**** 로 마스킹
#      - 05xx-가운데-마지막 그대로
#   2) 대여자명: 같은 줄에서 전화번호 "왼쪽"의 한글 2~8자 또는 영문 이름(최대 20자)
#      - 같은 줄에 없으면 바로 윗줄에서 추정
#      - '양정희/이종연' 같은 경우는 첫 항목만 사용
#   3) 주소: 이름 줄 '바로 아래 1줄'만 사용 (정확도/속도 목적)
#      - '주소:' 같은 라벨은 제거
# - 송장번호는 전혀 추출하지 않음

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

# --- Tesseract 경로(환경변수 허용) ---
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# --- 전처리/공통 ---
def _preprocess(img: Image.Image, strong: bool=False) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.1, percent=200, threshold=2))
        g = g.point(lambda x: 255 if x > 168 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g

def _ocr_text(img: Image.Image, psm:int=6, lang:str="kor+eng") -> str:
    try:
        return pytesseract.image_to_string(img, config=f"--oem 3 --psm {psm}", lang=lang)
    except Exception:
        return ""

def _resize(img: Image.Image, max_w:int=1300) -> Image.Image:
    w,h = img.size
    if w>max_w:
        s = max_w/float(w)
        return img.resize((max_w, int(h*s)))
    return img

# --- 정규식/라벨 ---
R_010   = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
R_05XX  = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _clean(s:str) -> str:
    return re.sub(r"[|\[\]{}<>]+"," ",s).strip()

# --- ROI: 파란/녹색 검출 있으면 사용, 없으면 화면 비율 ---
def _roi_cv2(path:str):
    if not HAS_CV2: return None
    try:
        import numpy as np  # noqa
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
    # 중앙~하단 폭넓게: 받는분/주소 영역 커버
    im = Image.open(path); W, H = im.size
    x1 = int(W * 0.05)
    y1 = int(H * 0.28)
    x2 = int(W * 0.92)
    y2 = int(H * 0.82)
    return im.crop((x1, y1, x2, y2))

def _crop_roi(path:str)->Image.Image:
    roi = _roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _roi_ratio(path)

# --- 이름/전화/주소 추출(전화 중심) ---
NAME_SPLIT = re.compile(r"[\/\|,\(\)\[\] ]+")
ENG_NAME   = re.compile(r"^[A-Za-z][A-Za-z \-]{1,19}$")
KOR_NAME   = re.compile(r"[가-힣]{2,8}")

def _first_phone_in_line(line:str):
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

def _pick_name_from_text(txt:str)->str:
    # 우선 한글 2~8자, 없으면 영문 이름
    k = KOR_NAME.findall(txt)
    if k:
        # '양정희/이종연' 같은 경우는 가장 마지막 한글 덩어리 선택
        cand = k[-1]
        return cand
    # 영문 이름(간단 허용)
    t = txt.strip()
    if ENG_NAME.match(t):
        return t.strip()
    # 분리 후 각각 검사
    for part in NAME_SPLIT.split(txt):
        part = part.strip()
        if not part: continue
        if KOR_NAME.fullmatch(part): return part
        if ENG_NAME.fullmatch(part): return part
    return ""

def _parse_by_phone(lines: List[str]) -> dict:
    """
    전략:
      - 위에서부터 훑으며 '첫 번째' 전화번호가 있는 라인을 기준 anchor로 삼는다.
      - 같은 줄 왼쪽에서 이름 시도. 없으면 윗줄 전체에서 이름 시도.
      - 주소는 anchor(이름 줄) 바로 아래 1줄만 사용(속도/안정성).
    """
    clean = [_clean(x) for x in lines]

    phone, name, addr = "", "", ""
    phone_i, name_i = -1, -1

    # 1) 전화번호 anchor
    for i, ln in enumerate(clean):
        m = _first_phone_in_line(ln)
        if not m: continue
        phone = _format_phone(m)
        phone_i = i
        # 같은 줄 왼쪽에서 이름 시도
        left = ln[:m.start()]
        left = LABEL_NAME.sub("", left).strip()
        name = _pick_name_from_text(left)
        if name: name_i = i
        break

    # 2) 이름이 비어있으면 윗줄에서 추정
    if phone and not name and phone_i > 0:
        up = LABEL_NAME.sub("", clean[phone_i-1]).strip()
        name = _pick_name_from_text(up)
        if name: name_i = phone_i-1

    # 3) 주소: 이름 줄 기준 아래 1줄만, 라벨 제거
    base_i = name_i if name else phone_i
    if base_i >= 0 and base_i+1 < len(clean):
        addr1 = LABEL_ADDR.sub("", clean[base_i+1]).strip()
        # 영어 노이즈 비율이 과하면 버리고 빈 값으로 둠 (UI에서 '-' 처리)
        hangul = len(re.findall(r"[가-힣]", addr1))
        latin  = len(re.findall(r"[A-Za-z]", addr1))
        if hangul >= latin:
            addr = addr1
        else:
            addr = ""

    return {"전화번호": phone, "대여자명": name, "주소": addr}

# --- 메인 OCR 흐름 ---
def _try_ocr_text(path:str) -> str:
    roi = _crop_roi(path)
    roi = _resize(roi, 1200)
    t = _ocr_text(_preprocess(roi, False), psm=6, lang="kor+eng")
    # 텍스트가 빈약하면 강한 전처리 한 번 더
    if len(re.sub(r"\s+","",t)) < 18:
        t2 = _ocr_text(_preprocess(roi, True), psm=6, lang="kor+eng")
        if len(re.sub(r"\s+","",t2)) > len(re.sub(r"\s+","",t)):
            t = t2
    return t

def make_final_entry(qr_text:str, 송장_image_path:str):
    txt = _try_ocr_text(송장_image_path)
    try:
        os.makedirs("_debug", exist_ok=True)
        with open(os.path.join("_debug","ocr_full.txt"),"w",encoding="utf-8") as f: f.write(txt)
    except Exception:
        pass

    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    parsed = _parse_by_phone(lines)

    # 기종/기기번호: QR 매핑(없으면'-')
    model, device_id = _map_model_device(qr_text)

    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id or "-",
        "기종": model or "-",
        # "송장번호"는 사용하지 않음
    }

def make_final_entry_fast(qr_text:str, 송장_image_path:str):
    """
    초고속 프리뷰:
      - 좁은 ROI + 숫자/하이픈 전용으로 전화번호만 빠르게 찾아서 반환
      - 이름/주소는 정식 OCR에서 보강
    """
    roi = _crop_roi(송장_image_path)
    roi = _resize(roi, 900)

    # 중앙~하단 띠
    W, H = roi.size
    band = roi.crop((int(W*0.05), int(H*0.40), int(W*0.95), int(H*0.78)))

    cfg_fast = "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789-"
    txt = pytesseract.image_to_string(_preprocess(band, strong=False), config=cfg_fast, lang="eng")
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]

    phone = ""
    for ln in lines:
        m = _first_phone_in_line(ln)
        if m:
            phone = _format_phone(m)
            break

    model, device_id = _map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명": "",
        "전화번호": phone,
        "주소": "",
        "기기번호": device_id or "-",
        "기종": model or "-",
    }

# --- QR → 기종/기기번호 ---
def _map_model_device(qr_text:str)->Tuple[str,str]:
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM":"심포니","LT":"락티나","S":"스윙","M":"스윙맥스","F":"프리스타일","G":"각시밀","C":"시밀레"}
    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2: return MAP.get(m2.group(1), "-"), m2.group(2)
    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1: return MAP.get(m1.group(1), "-"), raw
    return "-", raw



