# ocr_utils.py — 삼점 앵커(이름/전화/주소) 정확도 강화판
import os, re
from datetime import date
from typing import List, Tuple, Dict, Any, Optional
from PIL import Image, ImageOps, ImageFilter
import pytesseract

try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# Tesseract 경로
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

def _clean(s:str)->str:
    return re.sub(r"[|\[\]{}<>]+"," ", s or "").strip()

def _clamp(v,a,b): return max(a, min(b, v))

# 전화 규칙
R_PHONE_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4}|\*{4})")
R_PHONE_05  = re.compile(r"(05\d{2})[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
BANNED_PHONES = {"010-7394-3535"}

LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)
STOP_WORDS_FOR_NAME = {
    "주소","아파트","수령","수취","받는","전화","연락처","기종","기기번호",
    "심포니","락티나","스윙","스윙맥스","프리스타일","각시밀","시밀레"
}

def _normalize_for_ban(phone_text: str) -> str:
    t = re.sub(r"[^\d\*]", "", phone_text)
    if t.startswith("010") and len(t) >= 7:
        mid = t[3:-4]; last = t[-4:]
        return f"010-{mid}-{last}"
    if t.startswith("05") and len(t) >= 10:
        head = t[:4]; rest = t[4:]
        if len(rest)==7: return f"{head}-{rest[:3]}-{rest[3:]}"
        if len(rest)==8: return f"{head}-{rest[:4]}-{rest[4:]}"
    return phone_text

def _address_prefix(s: str) -> str:
    s2 = LABEL_ADDR.sub("", s or "").strip()
    if not s2: return ""
    s2 = re.split(r"[(),;]", s2)[0].strip()
    s2 = re.sub(r"\s+", " ", s2).strip()
    s2 = (s2.replace("서울특별시","서울").replace("부산광역시","부산").replace("대구광역시","대구")
              .replace("인천광역시","인천").replace("광주광역시","광주").replace("대전광역시","대전")
              .replace("울산광역시","울산").replace("세종특별자치시","세종"))
    return s2

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

# ---------------- 노란 박스 탐지 ----------------
def _find_yellow_block(pil_img: Image.Image) -> Optional[Tuple[int,int,int,int]]:
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

# ---------------- 추출 로직 ----------------
def _extract_from_yellow(yimg: Image.Image, anchors_rel: Optional[Dict[str,float]]) -> Tuple[str, str, str]:
    """anchors_rel: {'name': y(0..1), 'phone': y(0..1), 'addr': y(0..1)} — yimg 좌표계 기준"""
    W,H = yimg.size
    words = _tess_data(_preprocess(yimg, False), psm=6)
    if not words:
        words = _tess_data(_preprocess(yimg, True), psm=6)
        if not words: return "", "", ""

    lines = _group_lines(words)
    keys = sorted(lines.keys(), key=lambda k: (k[0],k[1],k[2]))

    meta = []
    for k in keys:
        ws = lines[k]
        y_top = min(w["top"] for w in ws)
        y_bot = max(w["top"]+w["height"] for w in ws)
        left_words  = [w for w in ws if (w["left"] + w["width"]/2) <= W*0.45]
        right_words = [w for w in ws if (w["left"] + w["width"]/2) >= W*0.55]
        meta.append({
            "yt":y_top, "yb":y_bot,
            "full":" ".join(w["text"] for w in ws).strip(),
            "left":" ".join(w["text"] for w in left_words).strip(),
            "right":" ".join(w["text"] for w in right_words).strip(),
            "left_words": left_words,
        })

    def _pick_line_by_y(y_norm: float, require_left=False) -> int:
        target = _clamp(y_norm,0.0,1.0) * H
        best_i, best_d = -1, 10**9
        for i,m in enumerate(meta):
            if require_left and not m["left"]: continue
            yc = (m["yt"]+m["yb"])/2.0
            d = abs(yc - target)
            if d < best_d:
                best_d, best_i = d, i
        return best_i

    name, addr, phone = "", "", ""
    name_idx = -1

    # 1) 이름: name 앵커 우선(왼쪽 컬럼), 금지어 제외
    if anchors_rel and "name" in anchors_rel:
        cand_i = _pick_line_by_y(anchors_rel["name"], require_left=True)
        if cand_i >= 0:
            left_norm = LABEL_NAME.sub("", _clean(meta[cand_i]["left"]))
            toks = re.findall(r"[가-힣]{2,4}", left_norm)
            toks = [t for t in toks if t not in STOP_WORDS_FOR_NAME]
            if toks:
                name = toks[0]; name_idx = cand_i

    # 보조: 앵커가 없거나 실패 시, 왼쪽 컬럼에서 위쪽부터 합리적인 후보
    if not name:
        for i,m in enumerate(meta):
            if not m["left"]: continue
            left_norm = LABEL_NAME.sub("", _clean(m["left"]))
            toks = re.findall(r"[가-힣]{2,4}", left_norm)
            toks = [t for t in toks if t not in STOP_WORDS_FOR_NAME]
            if toks:
                name = toks[0]; name_idx = i; break

    # 2) 주소: 이름 라인의 '다음 줄' 우선 → 없으면 addr 앵커 근처
    def _is_addrish(t:str)->bool:
        return any(tok in t for tok in ("시","도","군","구","읍","면","동","리","로","길","번길","아파트","호")) or bool(re.search(r"\d", t))
    if name_idx >= 0 and name_idx+1 < len(meta):
        cand = _clean(meta[name_idx+1]["full"])
        if name and cand.startswith(name): cand = cand[len(name):].strip()
        cand = R_PHONE_010.sub("", cand); cand = R_PHONE_05.sub("", cand)
        if _is_addrish(cand): addr = _address_prefix(cand)

    if not addr and anchors_rel and "addr" in anchors_rel:
        i2 = _pick_line_by_y(anchors_rel["addr"])
        if i2 >= 0:
            cand = _clean(meta[i2]["full"])
            cand = R_PHONE_010.sub("", cand); cand = R_PHONE_05.sub("", cand)
            if _is_addrish(cand): addr = _address_prefix(cand)

    # 3) 전화: phone 앵커 줄의 오른쪽 컬럼 우선 → 없으면 전체 검색
    def _search_phone_line(i: int) -> Optional[str]:
        raw = meta[i]["right"] or meta[i]["full"]
        m1 = R_PHONE_010.search(raw)
        if m1:
            norm = _normalize_for_ban(m1.group(0))
            if norm not in BANNED_PHONES: return m1.group(0)
        m2 = R_PHONE_05.search(raw)
        if m2: return m2.group(0)
        return None

    if anchors_rel and "phone" in anchors_rel:
        pi = _pick_line_by_y(anchors_rel["phone"])
        if pi >= 0:
            ph = _search_phone_line(pi)
            if ph: phone = ph

    if not phone:
        # 전체에서 첫 일치
        for i in range(len(meta)):
            ph = _search_phone_line(i)
            if ph: phone = ph; break

    return _address_prefix(addr) if addr else "", name or "", phone or ""

# ---------------- QR → 기종/기기번호 ----------------
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

# ---------------- 앵커 변환 & 엔트리 ----------------
def _yellow_and_anchors(im: Image.Image, anchors: Optional[Dict[str,Tuple[float,float]]]):
    bbox = _find_yellow_block(im)
    if not bbox:
        return im, None
    x,y,w,h = bbox
    yimg = im.crop((x,y,x+w,y+h))
    rel = {}
    if anchors:
        for k,(ax,ay) in anchors.items():
            px = _clamp(int(ax * im.size[0]), 0, im.size[0]-1)
            py = _clamp(int(ay * im.size[1]), 0, im.size[1]-1)
            if x <= px <= x+w and y <= py <= y+h:
                rel[k] = (py - y) / float(h)
    return yimg, rel if rel else None

def make_final_entry_fast(qr_text:str, img_path:str, anchors: Optional[Dict[str,Tuple[float,float]]]=None)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2000)
    yimg, rel = _yellow_and_anchors(im, anchors)
    address, name, phone = _extract_from_yellow(yimg, {k:v for k,v in (rel or {}).items()})
    return _final(qr_text, address, name, phone)

def make_final_entry(qr_text:str, img_path:str, anchors: Optional[Dict[str,Tuple[float,float]]]=None)->Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 2400)
    yimg, rel = _yellow_and_anchors(im, anchors)
    address, name, phone = _extract_from_yellow(yimg, {k:v for k,v in (rel or {}).items()})
    if not (address and name and phone):
        address2, name2, phone2 = _extract_from_yellow(_preprocess(yimg, True), {k:v for k,v in (rel or {}).items()})
        address = address or address2
        name    = name or name2
        phone   = phone or phone2
    return _final(qr_text, address, name, phone)

