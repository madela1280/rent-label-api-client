# ocr_utils.py — 노란영역 + 좌측 파란 가로바 기준 추출 (정확도 우선)
# 규칙(요청 고정):
# - 노란 배경 영역만 사용
# - 좌측 '파란 두꺼운 가로바' 아래 첫 줄(첫 칸 시작)이 "대여자명"
# - 그 '다음 줄'이 "주소" (주소는 앞부분만 정리)
# - 전화번호는 화면에 보이는 그대로: 05**-3/4-4, 010-1234-5678, 010-1234-**** 중
#   * 첫 번째로 발견된 것만 채택
#   * 금지번호 010-7394-3535 는 제외
# - QR 텍스트에서 기종/기기번호는 기존 규칙 유지

import os, re
from datetime import date
from typing import List, Tuple, Dict, Any, Optional

from PIL import Image, ImageOps, ImageFilter
import pytesseract

# ----- Optional OpenCV -----
try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# ----- Tesseract path -----
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# ================= 공통 유틸 =================
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

def _clean(s:str)->str:
    return re.sub(r"[|\[\]{}<>]+"," ", s or "").strip()

# ================= 패턴/정규식 =================
# 전화 4형식(보이는 그대로 저장)
R_PHONE_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4}|\*{4})")
R_PHONE_05  = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
BANNED_PHONES = {"010-7394-3535"}

LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

STOP_WORDS_FOR_NAME = {"주소","아파트","수령","수취","받는","전화","연락처"}

def _normalize_for_ban(phone_text: str) -> str:
    """금지번호 비교용 정규화(하이픈 통일)"""
    t = re.sub(r"[^\d\*]", "", phone_text)
    if t.startswith("010") and len(t) >= 7:
        mid = t[3:-4]
        last = t[-4:]
        return f"010-{mid}-{last}"
    if t.startswith("05") and len(t) >= 10:
        head = t[:4]
        rest = t[4:]
        if len(rest) == 7:
            return f"{head}-{rest[:3]}-{rest[3:]}"
        if len(rest) == 8:
            return f"{head}-{rest[:4]}-{rest[4:]}"
    return phone_text

def _address_prefix(s: str) -> str:
    s2 = LABEL_ADDR.sub("", s or "").strip()
    if not s2: return ""
    s2 = re.split(r"[(),]", s2)[0].strip()
    s2 = re.sub(r"\s+", " ", s2).strip()
    s2 = (s2.replace("서울특별시","서울").replace("부산광역시","부산").replace("대구광역시","대구")
              .replace("인천광역시","인천").replace("광주광역시","광주").replace("대전광역시","대전")
              .replace("울산광역시","울산").replace("세종특별자치시","세종"))
    return s2

# ================= OCR helpers =================
def _tess_data(img: Image.Image, psm:int=6, lang:str="kor") -> List[Dict[str, Any]]:
    try:
        d = pytesseract.image_to_data(img, config=f"--oem 3 --psm {psm}", lang=lang, output_type=pytesseract.Output.DICT)
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

# ================= 색상 ROI 탐지 =================
def _find_yellow_block(pil_img: Image.Image) -> Optional[Tuple[int,int,int,int]]:
    """노란 영역(라벨 전체)을 HSV로 검출"""
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

def _find_left_blue_bar(yimg: Image.Image) -> Optional[Tuple[int,int,int,int]]:
    """노란 영역 내부 '좌측 파란 두꺼운 가로바' 탐지 (실패해도 치명적이지 않음)."""
    if not HAS_CV2: return None
    bgr = cv2.cvtColor(np.array(yimg), cv2.COLOR_RGB2BGR)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)
    # 파랑(H≈90~135)
    mask = cv2.inRange(hsv, np.array([90,40,40],np.uint8), np.array([135,255,255],np.uint8))
    kernel = np.ones((5,5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts: return None

    H, W = yimg.size[1], yimg.size[0]
    # 좌측 절반에 있고 가로로 긴 사각형 우선
    cands = []
    for c in cnts:
        x,y,w,h = cv2.boundingRect(c)
        if x > W*0.55:  # 너무 오른쪽은 제외
            continue
        if w < h*2:     # 가로로 두껍지 않으면 제외
            continue
        if w*h < (W*H)*0.005:
            continue
        cands.append((y, x, w, h))
    if not cands:
        return None
    cands.sort()  # 위쪽 우선
    y,x,w,h = cands[0]
    return (x,y,w,h)

# ================= 노란영역 내부 파싱 =================
def _extract_from_yellow(yimg: Image.Image) -> Tuple[str, str, str]:
    """
    좌측 파란 가로바 기준:
      - 바 아래 첫 줄(왼쪽 첫 한글 2~4자) → 이름
      - 그 다음 줄 → 주소(앞부분만)
    파란바 탐지 실패 시:
      - 왼쪽 42% 컬럼의 첫 줄 = 이름, 그 다음 줄 = 주소
    전화는 전체(yimg)에서 첫 일치(그대로), 금지번호 제외.
    """
    W, H = yimg.size

    # (A) 파란바/왼쪽 컬럼 위치 확정
    bar = _find_left_blue_bar(yimg)
    if bar:
        bar_x, bar_y, bar_w, bar_h = bar
        left_x1 = max(0, bar_x - int(0.02*W))
        left_x2 = min(W, bar_x + max(bar_w, int(0.40*W)))
        name_band_y1 = bar_y + bar_h
    else:
        left_x1, left_x2 = 0, int(W*0.42)
        name_band_y1 = int(H*0.18)  # 보편적인 위치(대략 파란바 높이 가정)

    # (B) 라인 단위 OCR
    words = _tess_data(_preprocess(yimg, False), psm=6)
    if not words:
        words = _tess_data(_preprocess(yimg, True), psm=6)
        if not words:
            return "", "", ""

    lines = _group_lines(words)
    keys_sorted = sorted(lines.keys(), key=lambda k: (k[0],k[1],k[2]))

    # 보조: 각 라인의 yTop/Bottom, left컬럼 텍스트
    meta = []
    for k in keys_sorted:
        ws = lines[k]
        y_top = min(w["top"] for w in ws)
        y_bot = max(w["top"]+w["height"] for w in ws)
        text_full = " ".join(w["text"] for w in ws)
        text_left = " ".join(w["text"] for w in ws if (w["left"] + w["width"]/2) <= left_x2)
        meta.append((y_top, y_bot, text_full.strip(), text_left.strip(), k))

    # (C) 이름/주소: 파란바 아래에서 탐색
    name = ""
    addr = ""
    name_line_idx = -1

    # 이름: 파란바 바로 아래에 있는 "왼쪽 컬럼"의 첫 한글 2~4자
    # name 후보는 name_band_y1 이후의 첫 라인 중 왼쪽컬럼 텍스트 보유한 라인
    for i, (yt, yb, full, left, k) in enumerate(meta):
        if yt < name_band_y1: 
            continue
        if not left:
            continue
        # 라벨 제거
        left_norm = LABEL_NAME.sub("", _clean(left))
        # 금지단어 스킵
        if any(sw in left_norm for sw in STOP_WORDS_FOR_NAME):
            # 다음 라인으로
            continue
        toks = re.findall(r"[가-힣]{2,4}", left_norm)
        if toks:
            name = toks[0]  # 왼쪽 '첫' 토큰
            name_line_idx = i
            break

    # 주소: 이름 라인의 '다음 라인'을 사용(왼쪽/오른쪽 무관, 전체 라인)
    if name_line_idx >= 0 and name_line_idx + 1 < len(meta):
        addr_line_full = _clean(meta[name_line_idx + 1][2])
        # 주소에서 이름/전화 흔적 제거
        if name and addr_line_full.startswith(name):
            addr_line_full = addr_line_full[len(name):].strip()
        addr_line_full = R_PHONE_010.sub("", addr_line_full)
        addr_line_full = R_PHONE_05.sub("", addr_line_full)
        addr = _address_prefix(addr_line_full)

    # (D) 파란바/왼쪽 접근이 실패했다면 폴백(주소스러운 줄 선택 → 바로 위가 이름)
    if not name or not addr:
        # 주소 후보: 숫자와 주소 토큰이 많은 줄
        best_idx, best_score = -1, -1.0
        ADDR_HINT = ("시","도","군","구","읍","면","동","리","로","길","번길","아파트","호")
        for i,(yt,yb,full,left,k) in enumerate(meta):
            t = _clean(full)
            sc = sum(1 for tok in ADDR_HINT if tok in t) * 2.0 + len(re.findall(r"\d", t)) * 0.5 + min(len(t),40)*0.1
            if sc > best_score:
                best_idx, best_score = i, sc
        if best_idx >= 0:
            addr = _address_prefix(meta[best_idx][2])
            # 이름: 그 위 라인의 왼쪽 첫 토큰
            if best_idx > 0:
                left_norm = LABEL_NAME.sub("", _clean(meta[best_idx-1][3]))
                toks = re.findall(r"[가-힣]{2,4}", left_norm)
                if toks:
                    name = toks[0]

    # (E) 전화: yimg 전체에서 '첫 일치', 보이는 그대로. 금지번호 제외.
    phone = ""
    for i,(yt,yb,full,left,k) in enumerate(meta):
        raw = meta[i][2]
        m = R_PHONE_010.search(raw)
        if m:
            norm = _normalize_for_ban(m.group(0))
            if norm not in BANNED_PHONES:
                phone = m.group(0)
                break
        m = R_PHONE_05.search(raw)
        if m:
            phone = m.group(0)
            break

    return addr or "", name or "", phone or ""

# ================= QR → 기종/기기번호 =================
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

# ================= 최종 포맷 =================
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

# ================= 공개 엔트리 =================
def make_final_entry_fast(qr_text:str, img_path:str)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2000)

    # 노란영역
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

    # 부족할 땐 강처리 후 1회 재시도
    if not (address and name and phone):
        yimg2 = _preprocess(yimg, True)
        address2, name2, phone2 = _extract_from_yellow(yimg2)
        address = address or address2
        name    = name or name2
        phone   = phone or phone2

    return _final(qr_text, address, name, phone)

