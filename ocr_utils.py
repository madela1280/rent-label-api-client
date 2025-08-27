# ocr_utils.py — 노란 배경 블록 전용 추출 + 전화 규칙(4종, 첫번째만) + 010 마스킹 + 금지번호 제외
# 규칙
# - 입력 사진에서 "노란 배경" 블록만 색분리 → 그 영역 안에서만 추출
# - 대여자명: (노란 블록) 전화가 있는 "같은 줄"의 왼쪽에서 한글 2~4자, 없으면 바로 윗줄의 왼쪽
# - 전화번호: 패턴 4종 중 "첫번째로 등장"한 것만 사용
#     · 05**-123-4567  / 05**-1234-5678  → 그대로 저장
#     · 010-1234-5678  / 010-1234-****   → 항상 010-1234-**** 로 저장
#     · 금지번호 010-7394-3535 는 절대 저장하지 않음
# - 주소: (노란 블록) 전화 줄 "아래쪽"에서 처음 만나는 주소형 문장을 앞부분만 정규화(식별용)
# - QR 텍스트에서 기종/기기번호 추출은 기존 규칙 유지

import os, re
from datetime import date
from typing import List, Tuple, Dict, Any, Optional

from PIL import Image, ImageOps, ImageFilter
import pytesseract

# -------- Optional OpenCV (노란 영역 분리에 사용) --------
try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# -------- Tesseract 경로 --------
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# -------- 전처리 --------
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

# -------- 전화 규칙(4종) + 금지 --------
R_010_FULL = re.compile(r"(010)[-\s\.]?(\d{4})[-\s\.]?(\d{4})")
R_010_344  = re.compile(r"(010)[-\s\.]?(\d{3})[-\s\.]?(\d{4})")          # 예외(3-4) 대비
R_05_3_4   = re.compile(r"(05\d{2})[-\s\.]?(\d{3})[-\s\.]?(\d{4})")
R_05_4_4   = re.compile(r"(05\d{2})[-\s\.]?(\d{4})[-\s\.]?(\d{4})")

BANNED_PHONES = {"010-7394-3535"}  # ❗ 금지번호

LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _mask_010(m: re.Match) -> str:
    # 010은 항상 010-1234-**** 로 저장
    mid = m.group(2)
    return f"010-{mid}-****"

def _format_05(m: re.Match) -> str:
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

def _is_banned_010(m: re.Match) -> bool:
    raw = f"010-{m.group(2)}-{m.group(3)}"
    return raw in BANNED_PHONES

# -------- 주소 전처리 --------
def _address_prefix(s: str) -> str:
    s2 = LABEL_ADDR.sub("", s or "").strip()
    if not s2: return ""
    s2 = re.split(r"[(),]", s2)[0].strip()
    s2 = re.sub(r"\s+", " ", s2).strip()
    s2 = (s2.replace("서울특별시","서울").replace("부산광역시","부산").replace("대구광역시","대구")
              .replace("인천광역시","인천").replace("광주광역시","광주").replace("대전광역시","대전")
              .replace("울산광역시","울산").replace("세종특별자치시","세종"))
    return s2

# -------- Tesseract data helpers --------
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

# -------- 노란 배경 블록 탐지 --------
def _find_yellow_block(pil_img: Image.Image) -> Optional[Tuple[int,int,int,int]]:
    """
    HSV에서 노랑(H≈15~40) + 충분한 채도/명도 범위로 마스크 → 가장 큰 컨투어 bbox 반환.
    """
    if not HAS_CV2: return None
    bgr = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    hsv = cv2.cvtColor(bgr, cv2.COLOR_BGR2HSV)

    # 노란색 범위(실내 조명 고려해 넓게 2구간)
    lower1 = np.array([15,  60, 120], np.uint8)
    upper1 = np.array([30, 255, 255], np.uint8)
    lower2 = np.array([30,  60, 120], np.uint8)
    upper2 = np.array([40, 255, 255], np.uint8)

    mask1 = cv2.inRange(hsv, lower1, upper1)
    mask2 = cv2.inRange(hsv, lower2, upper2)
    mask = cv2.bitwise_or(mask1, mask2)

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

# -------- 노란 블록 내부 파싱 --------
def _extract_from_yellow(yimg: Image.Image) -> Tuple[str, str, str]:
    """
    노란 블록 내부에서 (위→아래, 좌→우) 순서로:
      1) 전화: 규칙 4종 중 '첫번째 등장' 채택 (010은 마스킹, 05xx는 그대로, 금지 제외)
      2) 이름: 전화가 있는 '같은 줄' 왼쪽에서 한글 2~4자, 없으면 바로 윗줄 왼쪽
      3) 주소: 전화 줄 '아래쪽'에서 처음 만나는 문장을 주소 앞부분으로 정규화
    반환: (address, name, phone)
    """
    # 충분 해상도 + 가벼운 전처리
    words = _tess_data(_preprocess(yimg, False), psm=6)
    lines = _group_lines(words)
    if not lines:
        # 강처리 1회 보강
        words = _tess_data(_preprocess(yimg, True), psm=6)
        lines = _group_lines(words)
        if not lines:
            return "", "", ""

    # 라인 키 정렬(상→하)
    keys = sorted(lines.keys(), key=lambda k: (k[0], k[1], k[2]))

    phone = ""
    phone_line_idx = -1

    # (1) 전화: 첫번째로 등장한 패턴 1개만 채택
    for i, k in enumerate(keys):
        t = _line_text(lines[k])

        # 010 full
        for m in R_010_FULL.finditer(t):
            if _is_banned_010(m):  # 금지
                continue
            phone = _mask_010(m)
            phone_line_idx = i
            break
        if phone: break

        # 010 3-4
        for m in R_010_344.finditer(t):
            if _is_banned_010(m):  # 금지
                continue
            phone = _mask_010(m)
            phone_line_idx = i
            break
        if phone: break

        # 05xx 3-4
        for m in R_05_3_4.finditer(t):
            phone = _format_05(m)
            phone_line_idx = i
            break
        if phone: break

        # 05xx 4-4
        for m in R_05_4_4.finditer(t):
            phone = _format_05(m)
            phone_line_idx = i
            break
        if phone: break

    # (2) 이름: 전화가 있는 줄의 '왼쪽', 없으면 바로 윗줄
    name = ""
    if phone and phone_line_idx >= 0:
        t = _line_text(lines[keys[phone_line_idx]])
        # 전화 match 위치를 다시 찾아서 그 앞부분만 사용
        left_text = t
        m = None
        for pat in (R_010_FULL, R_010_344, R_05_3_4, R_05_4_4):
            m = pat.search(t)
            if m: break
        if m:
            left_text = t[:m.start()]
        left_text = LABEL_NAME.sub("", left_text).strip()
        toks = re.findall(r"[가-힣]{2,4}", left_text)
        name = toks[-1] if toks else ""
        if not name and phone_line_idx > 0:
            prev_text = LABEL_NAME.sub("", _line_text(lines[keys[phone_line_idx-1]])).strip()
            toks2 = re.findall(r"[가-힣]{2,4}", prev_text)
            name = toks2[-1] if toks2 else ""
    else:
        # 전화가 없더라도 첫 두 줄 중 왼쪽에서 이름 추정(보수적으로)
        if keys:
            t0 = LABEL_NAME.sub("", _line_text(lines[keys[0]])).strip()
            toks0 = re.findall(r"[가-힣]{2,4}", t0)
            if toks0:
                name = toks0[-1]

    # (3) 주소: 전화 줄 아래에서 '처음 만나는 문장'을 앞부분만 반환
    address = ""
    start_idx = phone_line_idx + 1 if phone_line_idx >= 0 else 0
    for j in range(start_idx, len(keys)):
        cand = _line_text(lines[keys[j]])
        cand2 = _address_prefix(cand)
        if cand2:
            address = cand2
            break

    return address, name, phone

# -------- QR → 기종/기기번호 --------
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

# -------- 최종 포맷 --------
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

# -------- 공개 엔트리 (프리뷰/정식) --------
def make_final_entry_fast(qr_text:str, img_path:str)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2000)

    # 노란 블록 찾기
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

    # 보강: 비어있으면 강처리로 한 번 더 시도
    if not (address and name and phone):
        yimg2 = _preprocess(yimg, True)
        # image_to_data는 PIL.Image 필요 → 강처리 이미지는 이미 PIL
        address2, name2, phone2 = _extract_from_yellow(yimg2)
        address = address or address2
        name    = name or name2
        phone   = phone or phone2

    return _final(qr_text, address, name, phone)



