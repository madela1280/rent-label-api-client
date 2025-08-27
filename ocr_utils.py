# ocr_utils.py — 규칙 고정(주소=파란/녹색 박스 최댓글씨, 이름=주소 첫글자 아래, 전화=규칙4종, 010 마스킹, 금지번호 제외)
# - OpenCV가 있으면 파란/녹색 박스를 HSV로 탐지해서 주소 ROI로 사용
# - 없으면 전체 OCR에서 "가장 큰 폰트 라인(화면 중앙 가중)"을 주소로 사용
# - 이름은 주소 첫 글자의 x범위를 기준으로, 바로 아래 얇은 수직 밴드에서 한글 2~4자 추출
# - 전화는 이미지 전체에서 010/05xx 규칙만 허용, 010은 항상 010-1234-****로 저장, 05xx는 원문 유지
# - 금지번호 010-7394-3535는 절대 저장하지 않음
# - QR 텍스트에서 기종/기기번호 추출

import os, re
from datetime import date
from typing import List, Tuple, Dict, Any, Optional

from PIL import Image, ImageOps, ImageFilter
import pytesseract

# --- Optional OpenCV ---
try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# --- Tesseract path ---
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# ---------------- 공통 전처리 ----------------
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

# ---------------- 정규식/규칙 ----------------
# 전화 규칙 4종
R_010_FULL   = re.compile(r"(010)[-\s\.]?(\d{4})[-\s\.]?(\d{4})")
R_010_344    = re.compile(r"(010)[-\s\.]?(\d{3})[-\s\.]?(\d{4})")  # 보정용(드물게 3-4)
R_05_3_4     = re.compile(r"(05\d{2})[-\s\.]?(\d{3})[-\s\.]?(\d{4})")
R_05_4_4     = re.compile(r"(05\d{2})[-\s\.]?(\d{4})[-\s\.]?(\d{4})")

BANNED_PHONES = {"010-7394-3535"}  # 금지번호 (절대 저장하지 않음)

LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _mask_010(m: re.Match) -> str:
    # 010은 항상 010-1234-**** 로 저장
    mid = m.group(2)
    if len(mid) == 3:  # 3-4 케이스는 가운데 0 padding 없이 사용
        mid = f"{mid}"
    return f"010-{mid}-****"

def _format_05(m: re.Match) -> str:
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

def _address_prefix(s: str) -> str:
    # 주소는 "식별 가능한 앞부분"만
    s2 = LABEL_ADDR.sub("", s or "").strip()
    if not s2: return ""
    # 괄호/쉼표 앞까지만
    s2 = re.split(r"[(),]", s2)[0].strip()
    s2 = re.sub(r"\s+", " ", s2).strip()
    # 광역시 축약
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

# ---------------- 박스(파란/녹색) 탐지 ----------------
def _find_colored_band_bbox(pil_img: Image.Image) -> Optional[Tuple[int,int,int,int]]:
    """파란 또는 녹색 큰 박스 영역 탐지 (있으면 우선 사용)."""
    if not HAS_CV2: return None
    bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)

    # 파랑
    blue_mask1 = cv2.inRange(hsv, np.array([90, 40, 40], np.uint8), np.array([135, 255, 255], np.uint8))
    # 녹색
    green_mask = cv2.inRange(hsv, np.array([35, 40, 40], np.uint8),  np.array([85, 255, 255], np.uint8))
    mask = cv2.bitwise_or(blue_mask1, green_mask)

    if cv2.countNonZero(mask) < (pil_img.width * pil_img.height) * 0.005:
        return None

    kernel = np.ones((5,5), np.uint8)
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    cnts, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cnts: return None
    # 가장 넓은 컨투어 선택
    cnt = max(cnts, key=lambda c: cv2.boundingRect(c)[2]*cv2.boundingRect(c)[3])
    x,y,w,h = cv2.boundingRect(cnt)
    if w*h < (pil_img.width*pil_img.height)*0.02:
        return None
    return (x,y,w,h)

# ---------------- 주소/이름 추출 ----------------
def _extract_address_and_name(img: Image.Image) -> Tuple[str, str]:
    """
    1) 색상 박스가 있으면 그 내부의 "가장 큰 폰트 라인"을 주소로.
       - 주소 라인의 "첫 글자 x~폭" 기준으로, 바로 아래 얇은 수직 밴드에서 이름(한글 2~4자)
    2) 없으면 전체에서 "가장 큰 폰트 라인(중앙 가중)"을 주소로, 그 바로 아래 라인을 이름으로.
    """
    # 1) 색 박스 우선
    bbox = _find_colored_band_bbox(img)
    if bbox:
        x,y,w,h = bbox
        addr_roi = img.crop((x, y, x+w, y+h))
        # 박스 내부 라인들 중 폰트 높이(중앙 가중) 최댓값 라인 선택
        words = _tess_data(_preprocess(addr_roi, False), psm=6)
        lines = _group_lines(words)
        if lines:
            scored = []
            W, H = addr_roi.size
            for k, ws in lines.items():
                # 라인별 median height + 중앙 가중
                h_med = sorted([w["height"] for w in ws])[len(ws)//2]
                cx = sum(w["left"]+w["width"]/2 for w in ws)/len(ws)
                cy = sum(w["top"]+w["height"]/2 for w in ws)/len(ws)
                center = 1.0 - (abs(cx - W/2)/(W/2))*0.3 - (abs(cy - H/2)/(H/2))*0.3
                scored.append((h_med*center, k))
            scored.sort(reverse=True)
            addr_line = lines[scored[0][1]]
            addr_text = _line_text(addr_line)
            address = _address_prefix(addr_text)

            # 주소 첫 글자의 x범위 추정
            first_word = addr_line[0]
            x1 = x + first_word["left"]
            x2 = x + first_word["left"] + max(first_word["width"], 80)  # 최소 폭 80px

            # 이름 ROI: 박스 바로 아래 y~y+h*1.2 구간의 수직 밴드
            ny1 = y + h
            ny2 = min(img.height, ny1 + int(h * 1.2))
            name_roi = img.crop((max(0, x1-10), ny1, min(img.width, x2+10), ny2))
            raw = pytesseract.image_to_string(_preprocess(name_roi, False), config="--oem 3 --psm 7", lang="kor")
            raw = _clean(LABEL_NAME.sub("", raw))
            toks = re.findall(r"[가-힣]{2,4}", raw)
            name = toks[-1] if toks else ""
            return address, name

    # 2) 색 박스 실패 → 전체에서 "가장 큰 폰트 라인(중앙 가중)"
    words = _tess_data(_preprocess(img, False), psm=6)
    lines = _group_lines(words)
    if not lines:
        return "", ""

    scored = []
    W, H = img.size
    for k, ws in lines.items():
        h_med = sorted([w["height"] for w in ws])[len(ws)//2]
        cx = sum(w["left"]+w["width"]/2 for w in ws)/len(ws)
        cy = sum(w["top"]+w["height"]/2 for w in ws)/len(ws)
        center = 1.0 - (abs(cx - W/2)/(W/2))*0.3 - (abs(cy - H/2)/(H/2))*0.3
        scored.append((h_med*center, k))
    scored.sort(reverse=True)
    addr_line = lines[scored[0][1]]
    address = _address_prefix(_line_text(addr_line))

    # 이름: 주소 라인 바로 아래 가장 가까운 라인
    y_addr = max(w["top"]+w["height"] for w in addr_line)
    below = []
    for k, ws in lines.items():
        y_top = min(w["top"] for w in ws)
        if y_top >= y_addr:
            below.append((y_top, k))
    below.sort()
    name = ""
    if below:
        name_line = lines[below[0][1]]
        raw = LABEL_NAME.sub("", _line_text(name_line))
        toks = re.findall(r"[가-힣]{2,4}", raw)
        name = toks[-1] if toks else ""

    return address, name

# ---------------- 전화 추출 ----------------
def _extract_phone(img: Image.Image) -> str:
    """
    전화 규칙:
      - 05**-123-4567 / 05**-1234-5678  → 그대로 저장
      - 010-1234-**** / 010-1234-5678   → 저장 시 항상 010-1234-**** 로 변환
      - 금지번호 010-7394-3535는 제외
    우선순위: 010(비금지) > 05xx. 같은 우선순위면 중앙/신뢰도 점수로 선택.
    """
    # 두 단계(보통→강)로 단어 박스 생성
    words = _tess_data(_preprocess(img, False), psm=6)
    if not words:
        words = _tess_data(_preprocess(img, True), psm=6)
        if not words:
            return ""

    lines = _group_lines(words)
    W, H = img.size
    cands: List[Tuple[str, float, int]] = []  # (phone, score, priority)

    for ws in lines.values():
        t = _line_text(ws)

        # 010: full
        for m in R_010_FULL.finditer(t):
            raw_full = f"010-{m.group(2)}-{m.group(3)}"
            if raw_full in BANNED_PHONES:  # 금지
                continue
            ph = _mask_010(m)  # 저장은 항상 마스킹
            cx = sum(w["left"]+w["width"]/2 for w in ws)/len(ws)
            cy = sum(w["top"]+w["height"]/2 for w in ws)/len(ws)
            center = 1.0 - (abs(cx-W/2)/(W/2))*0.4 - (abs(cy-H/2)/(H/2))*0.4
            conf = sum(max(0.0, w["conf"]) for w in ws)/max(1,len(ws))
            cands.append((ph, center + conf/100.0, 2))

        # 010: 3-4 보정
        for m in R_010_344.finditer(t):
            raw_full = f"010-{m.group(2)}-{m.group(3)}"
            if raw_full in BANNED_PHONES:
                continue
            ph = _mask_010(m)
            cx = sum(w["left"]+w["width"]/2 for w in ws)/len(ws)
            cy = sum(w["top"]+w["height"]/2 for w in ws)/len(ws)
            center = 1.0 - (abs(cx-W/2)/(W/2))*0.4 - (abs(cy-H/2)/(H/2))*0.4
            conf = sum(max(0.0, w["conf"]) for w in ws)/max(1,len(ws))
            cands.append((ph, center + conf/100.0, 2))

        # 05xx: 3-4
        for m in R_05_3_4.finditer(t):
            ph = _format_05(m)
            cx = sum(w["left"]+w["width"]/2 for w in ws)/len(ws)
            cy = sum(w["top"]+w["height"]/2 for w in ws)/len(ws)
            center = 1.0 - (abs(cx-W/2)/(W/2))*0.4 - (abs(cy-H/2)/(H/2))*0.4
            conf = sum(max(0.0, w["conf"]) for w in ws)/max(1,len(ws))
            cands.append((ph, center + conf/100.0, 1))

        # 05xx: 4-4
        for m in R_05_4_4.finditer(t):
            ph = _format_05(m)
            cx = sum(w["left"]+w["width"]/2 for w in ws)/len(ws)
            cy = sum(w["top"]+w["height"]/2 for w in ws)/len(ws)
            center = 1.0 - (abs(cx-W/2)/(W/2))*0.4 - (abs(cy-H/2)/(H/2))*0.4
            conf = sum(max(0.0, w["conf"]) for w in ws)/max(1,len(ws))
            cands.append((ph, center + conf/100.0, 1))

    if not cands:
        return ""

    # 우선순위(010>05xx) → 점수 순
    cands.sort(key=lambda x: (x[2], x[1]), reverse=True)
    return cands[0][0]

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
        return MAP.get(m1.group(1), "-"), raw
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
    im = _resize(im, 2000)  # 정확도 위해 충분 해상도
    address, name = _extract_address_and_name(im)
    phone = _extract_phone(im)
    return _final(qr_text, address, name, phone)

def make_final_entry(qr_text:str, img_path:str)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2400)

    address, name = _extract_address_and_name(im)
    phone = _extract_phone(im)

    # 보강: 하나라도 비면 강처리 이미지로 재시도
    if not (address and name and phone):
        im2 = _preprocess(im, True)
        address2, name2 = _extract_address_and_name(im2)
        phone2 = _extract_phone(im2)
        address = address or address2
        name    = name or name2
        phone   = phone or phone2

    return _final(qr_text, address, name, phone)


