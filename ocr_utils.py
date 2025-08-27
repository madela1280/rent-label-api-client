# ocr_utils.py — 노란영역 전용 + 새 규칙 고정
# - 주소: 노란 배경 영역에서 "가장 주소스러운" 긴 줄
# - 이름: 주소 줄 바로 위 라인의 "왼쪽 첫 한글 2~4자"
# - 전화: 4형식 중 "첫 일치"를 '보이는 그대로' 채택(010도 마스킹 안 함)
# - 금지번호 010-7394-3535는 제외
# - QR → 기종/기기번호 매핑 유지

import os, re
from datetime import date
from typing import List, Tuple, Dict, Any, Optional

from PIL import Image, ImageOps, ImageFilter
import pytesseract

# --- Optional OpenCV (노란영역 분리) ---
try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# --- Tesseract 경로 ---
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# ---------------- 전처리 ----------------
def _resize(img: Image.Image, max_w:int) -> Image.Image:
    w, h = img.size
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

def _clean(s:str)->str:
    return re.sub(r"[|\[\]{}<>]+"," ", s or "").strip()

# ---------------- 전화 규칙(보이는 그대로) ----------------
# 010: 010-1234-5678 또는 010-1234-**** (가운데 3~4자리 허용, 끝 4자리 숫자 또는 ****)
R_PHONE_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4}|\*{4})")
# 05xx: 05**-123-4567 또는 05**-1234-5678
R_PHONE_05  = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")

BANNED_PHONES = {"010-7394-3535"}  # 금지번호(정확 매칭 시 제외)

LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _normalize_hyphen(phone_text: str) -> str:
    """비교용 정규화: 숫자와 '*'만 남기고 하이픈 포맷으로 재조립(표시는 원문 유지)"""
    digits = re.sub(r"[^\d\*]", "", phone_text)
    if digits.startswith("010") and (len(digits) == 11 or (len(digits)==10 and "*" in digits[-4:])):
        mid = digits[3:-4]
        last = digits[-4:]
        return f"010-{mid}-{last}"
    if digits.startswith("05") and len(digits) >= 10:
        # 05xx-(3|4)-(4)
        head = digits[:4]
        rest = digits[4:]
        if len(rest) == 7:
            return f"{head}-{rest[:3]}-{rest[3:]}"
        elif len(rest) == 8:
            return f"{head}-{rest[:4]}-{rest[4:]}"
    return phone_text

# ---------------- 주소 전처리 ----------------
def _address_prefix(s: str) -> str:
    s2 = LABEL_ADDR.sub("", s or "").strip()
    if not s2: return ""
    s2 = re.split(r"[(),]", s2)[0].strip()
    s2 = re.sub(r"\s+", " ", s2).strip()
    s2 = (s2.replace("서울특별시","서울").replace("부산광역시","부산").replace("대구광역시","대구")
              .replace("인천광역시","인천").replace("광주광역시","광주").replace("대전광역시","대전")
              .replace("울산광역시","울산").replace("세종특별자치시","세종"))
    return s2

# ---------------- Tesseract data helpers ----------------
def _tess_data(img: Image.Image, psm:int=6) -> List[Dict[str, Any]]:
    try:
        d = pytesseract.image_to_data(img, config=f"--oem 3 --psm {psm}", lang="kor", output_type=pytesseract.Output.DICT)
    except Exception:
        return []
    n = len(d.get("text", []))
    out = []
    for i in range(n):
        txt = (d["text"][i] or "").strip()
        if not txt: continue
        try: conf = float(d["conf"][i])
        except: conf = -1.0
        out.append({
            "text": txt,
            "left": int(d["left"][i]), "top": int(d["top"][i]),
            "width": int(d["width"][i]), "height": int(d["height"][i]),
            "conf": conf,
            "line_id": (int(d["block_num"][i]), int(d["par_num"][i]), int(d["line_num"][i])),
        })
    return out

def _group_lines(words: List[Dict[str, Any]]):
    lines: Dict[Tuple[int,int,int], List[Dict[str, Any]]] = {}
    for w in words:
        lines.setdefault(w["line_id"], []).append(w)
    for k in lines:
        lines[k].sort(key=lambda x: x["left"])
    return lines

def _line_text(words: List[Dict[str, Any]]) -> str:
    return _clean(" ".join(w["text"] for w in words))

# ---------------- 노란 배경 블록 탐지 ----------------
def _find_yellow_block(pil_img: Image.Image) -> Optional[Tuple[int,int,int,int]]:
    """
    HSV에서 노랑(H≈15~40) 범위 마스크 → 가장 큰 컨투어 bbox 반환.
    """
    if not HAS_CV2: return None
    bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)

    lower1 = np.array([15,  60, 120], np.uint8)
    upper1 = np.array([30, 255, 255], np.uint8)
    lower2 = np.array([30,  60, 120], np.uint8)
    upper2 = np.array([40, 255, 255], np.uint8)

    mask = cv2.inRange(hsv, lower1, upper1) | cv2.inRange(hsv, lower2, upper2)
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

# ---------------- 주소 선택(노란영역) ----------------
_ADDR_HINT = ("시","도","군","구","읍","면","동","리","로","길","번길","아파트","호")

def _pick_address_line(lines: Dict[Tuple[int,int,int], List[Dict[str, Any]]]) -> Tuple[str, int]:
    """
    가장 주소스럽고 긴 줄을 선택. 반환: (주소텍스트, 라인인덱스)
    """
    if not lines: return "", -1
    keys = sorted(lines.keys(), key=lambda k: (k[0], k[1], k[2]))
    best = ("", -1, -1.0)  # text, idx, score
    for i, k in enumerate(keys):
        t = _line_text(lines[k])
        score = 0.0
        # 주소 힌트 토큰 가중치
        score += sum(1 for tok in _ADDR_HINT if tok in t) * 2.0
        # 숫자 길이 가중치
        score += len(re.findall(r"\d", t)) * 0.5
        # 전체 길이
        score += min(len(t), 40) * 0.1
        if score > best[2]:
            best = (t, i, score)
    return best[0], best[1]

# ---------------- 이름/전화 추출(노란영역 내부) ----------------
def _extract_from_yellow(yimg: Image.Image) -> Tuple[str, str, str]:
    """
    1) 주소 라인 선택
    2) 이름: 주소 바로 위 라인의 '왼쪽 첫' 한글 2~4자
    3) 전화: 위→아래로 훑어 010/05xx 첫 일치 '그대로' 채택 (금지번호 제외)
    """
    words = _tess_data(_preprocess(yimg, False), psm=6)
    if not words:
        words = _tess_data(_preprocess(yimg, True), psm=6)
        if not words:
            return "", "", ""

    lines = _group_lines(words)
    keys = sorted(lines.keys(), key=lambda k: (k[0], k[1], k[2]))

    # (1) 주소
    addr_text, addr_idx = _pick_address_line(lines)
    address = _address_prefix(addr_text) if addr_text else ""

    # (2) 이름: 주소 바로 위 라인 왼쪽 첫 토큰
    name = ""
    if addr_idx > 0:
        above_text = LABEL_NAME.sub("", _line_text(lines[keys[addr_idx-1]])).strip()
        # 왼쪽→오른쪽 순서로 토큰화: 한글 2~4자
        toks = re.findall(r"[가-힣]{2,4}", above_text)
        name = toks[0] if toks else ""

    # (3) 전화: 첫 일치 '그대로'(원문 substring)
    phone = ""
    for i, k in enumerate(keys):
        t = _line_text(lines[k])  # 공백/특수문자 정리된 라인 텍스트
        raw_line = " ".join(w["text"] for w in lines[k])  # 원문 구성(표시용 substring 목적)
        # 010
        m = R_PHONE_010.search(raw_line)
        if m:
            norm = _normalize_hyphen(m.group(0))
            if norm in BANNED_PHONES:
                pass
            else:
                phone = m.group(0)  # 보이는 그대로
                break
        # 05xx
        m = R_PHONE_05.search(raw_line)
        if m:
            phone = m.group(0)     # 보이는 그대로
            break

    return address, name, phone

# ---------------- QR 파싱 ----------------
def _map_model_device(qr_text:str)->Tuple[str,str]:
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM":"심포니","LT":"락티나","S":"스윙","M":"스윙맥스","F":"프리스타일","G":"각시밀","C":"시밀레"}
    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2:
        return MAP.get(m2.group(1), "-"), m2.group(2)
    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1:
        return MAP.get(m1[1], "-"), raw
    return "-", ""

# ---------------- 최종 포맷 ----------------
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

# ---------------- 공개 엔트리 ----------------
def make_final_entry_fast(qr_text:str, img_path:str)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2000)
    # 노란 영역
    bbox = _find_yellow_block(im)
    yimg = im.crop((bbox[0], bbox[1], bbox[0]+bbox[2], bbox[1]+bbox[3])) if bbox else im
    address, name, phone = _extract_from_yellow(yimg)
    return _final(qr_text, address, name, phone)

def make_final_entry(qr_text:str, img_path:str)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2400)
    bbox = _find_yellow_block(im)
    yimg = im.crop((bbox[0], bbox[1], bbox[0]+bbox[2], bbox[1]+bbox[3])) if bbox else im

    address, name, phone = _extract_from_yellow(yimg)

    # 보강: 빈 항목이 있으면 강처리로 한 번 더 시도
    if not (address and name and phone):
        yimg2 = _preprocess(yimg, True)
        address2, name2, phone2 = _extract_from_yellow(yimg2)
        address = address or address2
        name    = name or name2
        phone   = phone or phone2

    return _final(qr_text, address, name, phone)


