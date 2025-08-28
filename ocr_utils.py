# ocr_utils.py — 단일 앵커(이름) 기준 추출 고정판
import os, re
from datetime import date
from typing import List, Tuple, Dict, Any, Optional
from PIL import Image, ImageOps, ImageFilter
import pytesseract

try:
    import cv2, numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# Tesseract 경로
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# ---------- 공통 ----------
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
def _clean(s:str)->str: return re.sub(r"[|\[\]{}<>]+"," ", s or "").strip()

# 전화 규칙
R_PHONE_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4}|\*{4})")
R_PHONE_05  = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
BANNED_PHONES = {"010-7394-3535"}

LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
STOP_WORDS = {"주소","아파트","전화","연락처","기종","기기번호","심포니","락티나","스윙","스윙맥스","프리스타일","각시밀","시밀레"}

def _address_prefix(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"\s+", " ", s)
    # 전화 제거
    s = R_PHONE_010.sub("", s)
    s = R_PHONE_05.sub("", s)
    # (,) ; 뒤쪽 날림
    s = re.split(r"[();]", s)[0].strip()
    return s[:30]  # 앞부분만

def _tess_line(img: Image.Image, lang:str, allowlist:str=None) -> str:
    cfg = "--oem 3 --psm 7"
    if allowlist:
        cfg += f" -c tessedit_char_whitelist={allowlist}"
    try:
        return pytesseract.image_to_string(img, config=cfg, lang=lang)
    except Exception:
        return ""

# ---------- 노란 영역 ----------
def _find_yellow_block(pil_img: Image.Image) -> Optional[Tuple[int,int,int,int]]:
    if not HAS_CV2: return None
    bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    lower1 = np.array([15,  60, 120], np.uint8)
    upper1 = np.array([40, 255, 255], np.uint8)
    mask = cv2.inRange(hsv, lower1, upper1)
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

# ---------- ROI 만들기 ----------
def _roi_from_anchor(yimg: Image.Image, anchor_rel_y: float) -> Dict[str, Image.Image]:
    """이름: 왼쪽 0~45%, 전화: 오른쪽 55~100%, 주소: 이름줄 바로 다음줄"""
    W,H = yimg.size
    y = _clamp(int(anchor_rel_y * H), 0, H-1)

    # 라인 높이 추정(대략)
    line_h = max(18, int(H * 0.035))

    # 이름 줄
    y1 = _clamp(y - int(line_h*0.6), 0, H-1)
    y2 = _clamp(y + int(line_h*0.6), 0, H)
    name_box  = (0, y1, int(W*0.45), y2)
    phone_box = (int(W*0.55), y1, W, y2)

    # 주소 줄(이름 아래쪽 한 줄)
    ya = _clamp(y + int(line_h*1.2), 0, H-1)
    yb = _clamp(ya + int(line_h*1.2), 0, H)
    addr_box  = (0, ya, W, yb)

    return {
        "name":  yimg.crop(name_box),
        "phone": yimg.crop(phone_box),
        "addr":  yimg.crop(addr_box),
    }

# ---------- 추출 ----------
def _extract_with_anchor(yimg: Image.Image, anchor_rel_y: Optional[float]) -> Tuple[str,str,str]:
    if yimg is None: return "", "", ""
    if anchor_rel_y is None: return "", "", ""

    rois = _roi_from_anchor(yimg, anchor_rel_y)

    # 이름: 한글 2~4자, 금지어 제외
    name_raw = _tess_line(_preprocess(rois["name"], False), "kor")
    if len(re.sub(r"\s+","",name_raw)) < 2:
        name_raw = _tess_line(_preprocess(rois["name"], True), "kor")
    name_tokens = re.findall(r"[가-힣]{2,4}", LABEL_NAME.sub("", name_raw))
    name_tokens = [t for t in name_tokens if t not in STOP_WORDS]
    name = name_tokens[0] if name_tokens else ""

    # 전화: 오른쪽, allowlist
    phone_raw = _tess_line(_preprocess(rois["phone"], False), "eng", allowlist="0123456789-*")
    if len(re.sub(r"\s+","",phone_raw)) < 5:
        phone_raw = _tess_line(_preprocess(rois["phone"], True), "eng", allowlist="0123456789-*")
    phone = ""
    m1 = R_PHONE_010.search(phone_raw)
    if m1:
        phone = m1.group(0)
        if phone == "010-7394-3535": phone = ""
    if not phone:
        m2 = R_PHONE_05.search(phone_raw)
        if m2: phone = m2.group(0)

    # 주소: 다음 줄, 앞쪽만
    addr_raw = _tess_line(_preprocess(rois["addr"], False), "kor")
    if len(re.sub(r"\s+","",addr_raw)) < 4:
        addr_raw = _tess_line(_preprocess(rois["addr"], True), "kor")
    addr = _address_prefix(addr_raw)

    return addr, name, phone

# ---------- QR → 모델/기기 ----------
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
    return {
      "출고일": date.today().isoformat(),
      "대여자명": name or "",
      "전화번호": phone or "",
      "주소": address or "",
      "기기번호": device_id,
      "기종": model,
    }

# ---------- 외부 API ----------
def _yellow_and_rel(im: Image.Image, anchor: Optional[Tuple[float,float]]):
    bbox = _find_yellow_block(im)
    if not bbox: return im, None
    x,y,w,h = bbox
    yimg = im.crop((x,y,x+w,y+h))
    rel_y = None
    if anchor:
        ax, ay = anchor
        px = _clamp(int(ax * im.size[0]), 0, im.size[0]-1)
        py = _clamp(int(ay * im.size[1]), 0, im.size[1]-1)
        if x <= px <= x+w and y <= py <= y+h:
            rel_y = (py - y) / float(h)
    return yimg, rel_y

def make_final_entry_fast(qr_text:str, img_path:str, anchor: Optional[Tuple[float,float]]=None)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2000)
    yimg, rel_y = _yellow_and_rel(im, anchor)
    addr, name, phone = _extract_with_anchor(yimg, rel_y)
    return _final(qr_text, addr, name, phone)

def make_final_entry(qr_text:str, img_path:str, anchor: Optional[Tuple[float,float]]=None)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2400)
    yimg, rel_y = _yellow_and_rel(im, anchor)
    addr, name, phone = _extract_with_anchor(yimg, rel_y)
    if not (addr and name and phone):
        addr2, name2, phone2 = _extract_with_anchor(_preprocess(yimg, True) if yimg else None, rel_y)
        addr = addr or addr2
        name = name or name2
        phone = phone or phone2
    return _final(qr_text, addr, name, phone)


