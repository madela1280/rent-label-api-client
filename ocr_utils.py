# ocr_utils.py — 앵커(이름 첫 글자) 기준 초간단·고정 규칙
# - 이름: 한글 최대 5자
# - 전화: 같은 줄 오른쪽에서 010/05** 시작 "하나만"
# - 주소: 이름 줄의 '바로 아랫줄' 1줄만
# - 금지번호: 010-7394-3535 제외
# - 노란 라벨 영역 우선 크롭(가능시)

import os, re
from datetime import date
from typing import Tuple, Dict, Optional
from PIL import Image, ImageOps, ImageFilter
import pytesseract

try:
    import cv2, numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# Tesseract 경로(환경변수 우선)
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# ---------------- 공통 유틸 ----------------
def _resize(img: Image.Image, max_w:int) -> Image.Image:
    w,h = img.size
    if w > max_w:
        s = max_w/float(w)
        return img.resize((max_w, int(h*s)))
    return img

def _preprocess(img: Image.Image, strong: bool=False) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=240, threshold=2))
        g = g.point(lambda x: 255 if x > 165 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g

def _clamp(v,a,b): return max(a, min(b, v))

def _tess_line(img: Image.Image, lang:str, allowlist:str=None) -> str:
    cfg = "--oem 3 --psm 7"
    if allowlist:
        cfg += f" -c tessedit_char_whitelist={allowlist}"
    try:
        return pytesseract.image_to_string(img, config=cfg, lang=lang)
    except Exception:
        return ""

# ---------------- 규칙/정규식 ----------------
# 전화 규칙(요청 고수)
R_PHONE_010 = re.compile(r"(010)[-\s\.]?(\d{4})[-\s\.]?(\d{4}|\*{4})")
R_PHONE_05A = re.compile(r"(05\d{2})[-\s\.]?(\d{3})[-\s\.]?(\d{4})")
R_PHONE_05B = re.compile(r"(05\d{2})[-\s\.]?(\d{4})[-\s\.]?(\d{4})")
BANNED_PHONES = {"010-7394-3535"}

ADDR_TOKENS = ("시","군","구","읍","면","동","로","길","번길","아파트","빌라","호","단지")
STOP_WORDS   = {"주소","아파트","전화","연락처","기종","기기번호","심포니","락티나","스윙","스윙맥스","프리스타일","각시밀","시밀레"}

def _digits_like_fix(s: str) -> str:
    rep = {'O':'0','o':'0','D':'0','l':'1','I':'1','|':'1','!':'1','S':'5','B':'8','Z':'2','‘':'1','’':'1','“':'1','”':'1'}
    return "".join(rep.get(ch, ch) for ch in s)

def _reformat_phone_from_digits(s: str) -> Optional[str]:
    ds = re.sub(r"[^\d]", "", s)
    if len(ds) >= 11 and ds.startswith("010"):
        return f"010-{ds[3:7]}-{ds[7:11]}"
    if len(ds) >= 12 and ds.startswith("05"):
        return f"{ds[:4]}-{ds[4:8]}-{ds[8:12]}"
    if len(ds) == 11 and ds.startswith("05"):
        return f"{ds[:4]}-{ds[4:7]}-{ds[7:11]}"
    return None

def _pick_first_phone(line: str) -> str:
    line = _digits_like_fix(line)
    for pat in (R_PHONE_010, R_PHONE_05A, R_PHONE_05B):
        m = pat.search(line)
        if m:
            v = m.group(0)
            if v == "010-7394-3535":  # 금지
                return ""
            return v
    rf = _reformat_phone_from_digits(line)
    if rf and rf != "010-7394-3535":
        return rf
    return ""

def _address_prefix(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    # 전화 흔적 제거
    for pat in (R_PHONE_010, R_PHONE_05A, R_PHONE_05B):
        s = pat.sub("", s)
    # 괄호 앞까지만
    s = re.split(r"[();]", s)[0].strip()
    return s

# ---------------- 노란 라벨 탐지 ----------------
def _find_yellow_block(pil_img: Image.Image):
    """노란 라벨 영역(BGR HSV 범위)"""
    if not HAS_CV2: return None
    bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    lower = np.array([15,  60, 120], np.uint8)
    upper = np.array([40, 255, 255], np.uint8)
    mask = cv2.inRange(hsv, lower, upper)
    if cv2.countNonZero(mask) < (pil_img.width * pil_img.height) * 0.003:
        return None
    kernel = np.ones((5,5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts: return None
    cnt = max(cnts, key=lambda c: cv2.boundingRect(c)[2]*cv2.boundingRect(c)[3])
    x,y,w,h = cv2.boundingRect(cnt)
    if w*h < (pil_img.width*pil_img.height)*0.01:
        return None
    return (x,y,w,h)

# ---------------- 앵커 기반 ROI ----------------
def _rois_from_anchor(img: Image.Image, ax: float, ay: float):
    """앵커 기준으로 '이름줄', '같은 줄 오른쪽(전화)', '바로 아랫줄(주소 1줄)' ROI 생성"""
    W,H = img.size
    px = _clamp(int(ax * W), 0, W-1)
    py = _clamp(int(ay * H), 0, H-1)

    # 한 줄 높이 추정
    line_h = max(18, int(H * 0.035))

    # 이름: 앵커 x부터 오른쪽(최대 5자 가정) + 세로는 라인 폭만
    name_w = _clamp(int(W * 0.18), 140, 380)
    x1 = _clamp(px - 10, 0, W-1)
    x2 = _clamp(px + name_w, 0, W)
    y1 = _clamp(py - int(line_h*0.6), 0, H-1)
    y2 = _clamp(py + int(line_h*0.6), 0, H)
    name_roi = img.crop((x1, y1, x2, y2))

    # 전화: 같은 줄 오른쪽 끝까지
    phone_roi = img.crop((_clamp(x2 + 6, 0, W-1), y1, W, y2))

    # 주소: '바로 아랫줄'만 (라인 높이만큼 한 줄)
    addr_y1 = _clamp(py + int(line_h*0.95), 0, H-1)
    addr_y2 = _clamp(addr_y1 + line_h, 0, H)
    addr_roi = img.crop((0, addr_y1, W, addr_y2))

    return name_roi, phone_roi, addr_roi

# ---------------- 파싱 ----------------
def _extract(img: Image.Image, anchor: Optional[Tuple[float,float]]) -> Tuple[str,str,str]:
    """반환: (주소, 이름, 전화) — 주소는 한 줄만"""
    if img is None or anchor is None:
        return "", "", ""
    ax, ay = anchor
    name_roi, phone_roi, addr_roi = _rois_from_anchor(img, ax, ay)

    # 이름: 한글 최대 5자
    name_raw = _tess_line(_preprocess(name_roi, False), "kor")
    if len(re.sub(r"\s+","",name_raw)) < 2:
        name_raw = _tess_line(_preprocess(name_roi, True), "kor")
    cand = re.findall(r"[가-힣]{2,}", name_raw)
    name = ""
    for t in cand:
        if t in STOP_WORDS: continue
        name = t[:5]  # 최대 5자
        break

    # 전화: 같은 줄 오른쪽에서 첫 1개
    phone_line = _tess_line(_preprocess(phone_roi, False), "eng", allowlist="0123456789-*OolISBZ|!")
    if len(re.sub(r"\s+","",phone_line)) < 4:
        phone_line = _tess_line(_preprocess(phone_roi, True), "eng", allowlist="0123456789-*OolISBZ|!")
    phone = _pick_first_phone(phone_line)

    # 주소: 바로 아랫줄 1줄
    addr_line = _tess_line(_preprocess(addr_roi, False), "kor")
    if len(re.sub(r"\s+","",addr_line)) < 4:
        addr_line = _tess_line(_preprocess(addr_roi, True), "kor")
    addr_line = _address_prefix(addr_line)
    # 토큰 검증(시/군/구/동/로/길 등 하나라도 포함 시 채택)
    addr = addr_line if any(tok in addr_line for tok in ADDR_TOKENS) else addr_line

    return addr, name, phone

# ---------------- QR → 모델/기기 ----------------
def _map_model_device(qr_text:str)->Tuple[str,str]:
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM":"심포니","LT":"락티나","S":"스윙","M":"스윙맥스","F":"프리스타일","G":"각시밀","C":"시밀레"}
    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2: return MAP.get(m2.group(1), "-"), m2.group(2)
    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1: return MAP.get(m1.group(1), "-"), raw
    return "-", ""

def _final(qr_text:str, address:str, name:str, phone:str)->Dict[str,str]:
    model, device_id = _map_model_device(qr_text)
    # 금지번호 최종 필터
    if phone == "010-7394-3535":
        phone = ""
    return {
      "출고일": date.today().isoformat(),
      "대여자명": name or "",
      "전화번호": phone or "",
      "주소": address or "",
      "기기번호": device_id,
      "기종": model,
    }

# ---------------- 노란 라벨 크롭 ----------------
def _yellow_crop_and_anchor(im: Image.Image, anchor: Optional[Tuple[float,float]]):
    bbox = _find_yellow_block(im)
    if not bbox:
        return im, anchor  # 노란 영역 미검출 → 전체 사용
    x,y,w,h = bbox
    yimg = im.crop((x,y,x+w,y+h))
    if anchor is None:
        return yimg, None
    ax, ay = anchor
    px = _clamp(int(ax * im.size[0]), 0, im.size[0]-1)
    py = _clamp(int(ay * im.size[1]), 0, im.size[1]-1)
    if x <= px <= x+w and y <= py <= y+h:
        rel = ((px - x)/float(w), (py - y)/float(h))
        return yimg, rel
    return im, anchor

# ---------------- 공개 API ----------------
def make_final_entry_fast(qr_text:str, img_path:str, anchor: Optional[Tuple[float,float]]=None)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 1800)  # 프리뷰/빠른 경로
    crop, rel = _yellow_crop_and_anchor(im, anchor)
    addr, name, phone = _extract(crop, rel)
    return _final(qr_text, addr, name, phone)

def make_final_entry(qr_text:str, img_path:str, anchor: Optional[Tuple[float,float]]=None)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2200)  # 정식 경로
    crop, rel = _yellow_crop_and_anchor(im, anchor)
    addr, name, phone = _extract(crop, rel)
    # 부족시 강처리 1회만
    if not (name and phone and addr):
        addr2, name2, phone2 = _extract(_preprocess(crop, True), rel)
        addr = addr or addr2
        name = name or name2
        phone = phone or phone2
    return _final(qr_text, addr, name, phone)






