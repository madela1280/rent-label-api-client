# ocr_utils.py — 정확도 우선 버전
# - 전체 OCR + 위치 데이터(image_to_data)로 "전화 줄"을 먼저 찾고
#   같은 줄의 왼쪽에서 이름(한글 2~4자), 다음 줄에서 주소 앞부분(식별용)만 추출
# - 프리뷰: 가벼운 전처리 1회
# - 정식: 전처리(보통/강) 2회 중 "전화 검출 점수"가 높은 쪽 선택
# - QR에서 기종/기기번호 유지

import os, re
from datetime import date
from typing import List, Tuple, Dict, Any
from PIL import Image, ImageOps, ImageFilter
import pytesseract

pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# -------- 전처리 --------
def _preprocess(img: Image.Image, strong: bool=False) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=240, threshold=2))
        g = g.point(lambda x: 255 if x > 165 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g

def _resize(img: Image.Image, max_w:int) -> Image.Image:
    w, h = img.size
    if w > max_w:
        s = max_w / float(w)
        return img.resize((max_w, int(h*s)))
    return img

# -------- 규칙/정규식 --------
R_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)
ADDR_TOKENS = ("시","도","군","구","읍","면","동","리","로","길","번길","호")
BANNED_PHONES = {"010-7394-3535"}

def _clean(s:str) -> str:
    return re.sub(r"[|\[\]{}<>]+"," ", s).strip()

def _norm_phone(m: re.Match) -> str:
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

def _address_prefix(s: str) -> str:
    s2 = LABEL_ADDR.sub("", s or "").strip()
    if not s2: return ""
    s2 = re.split(r"[(),]", s2)[0].strip()
    mnum = re.search(r"\d+", s2)
    cut = None
    if mnum:
        cut = mnum.end()
        if cut < len(s2) and s2[cut] == ' ':
            cut += 1
    head = s2[:cut] if cut else s2
    head = re.sub(r"\s+", " ", head).strip()
    head = head.replace("서울특별시", "서울").replace("부산광역시","부산").replace("대구광역시","대구") \
               .replace("인천광역시","인천").replace("광주광역시","광주").replace("대전광역시","대전") \
               .replace("울산광역시","울산").replace("세종특별자치시","세종")
    if len(head) > 30:
        head = head[:30].rstrip()
    if not any(t in head for t in ("구","군","시")) and not re.search(r"\d", head):
        return ""
    return head

# -------- OCR with boxes --------
def _tess_data(img: Image.Image, psm:int=6) -> List[Dict[str, Any]]:
    """
    pytesseract.image_to_data 결과를 파싱해서 단어 단위 정보 리스트 반환.
    각 항목: {text, left, top, width, height, conf, line_id}
    line_id는 (block_num, par_num, line_num) 튜플로 구성.
    """
    try:
        raw = pytesseract.image_to_data(img, config=f"--oem 3 --psm {psm}", lang="kor", output_type=pytesseract.Output.DICT)
    except Exception:
        return []

    n = len(raw.get("text", []))
    out = []
    for i in range(n):
        txt = (raw["text"][i] or "").strip()
        if not txt:
            continue
        try:
            conf = float(raw["conf"][i])
        except:
            conf = -1.0
        item = {
            "text": txt,
            "left": int(raw["left"][i]),
            "top": int(raw["top"][i]),
            "width": int(raw["width"][i]),
            "height": int(raw["height"][i]),
            "conf": conf,
            "line_id": (int(raw["block_num"][i]), int(raw["par_num"][i]), int(raw["line_num"][i])),
        }
        out.append(item)
    return out

def _group_lines(words: List[Dict[str, Any]]) -> Dict[Tuple[int,int,int], List[Dict[str, Any]]]:
    lines: Dict[Tuple[int,int,int], List[Dict[str, Any]]] = {}
    for w in words:
        lines.setdefault(w["line_id"], []).append(w)
    # 좌->우 정렬
    for k in lines:
        lines[k].sort(key=lambda x: x["left"])
    return lines

def _line_text(words: List[Dict[str, Any]]) -> str:
    return _clean(" ".join(w["text"] for w in words))

def _select_best_phone(line_text: str) -> Tuple[str, int, int]:
    """
    줄 텍스트에서 전화번호 후보를 찾아 정규화 반환.
    (전화문자열, start_index, end_index) — 없으면 ("", -1, -1)
    """
    best = ("", -1, -1)
    for m in R_010.finditer(line_text):
        phone = _norm_phone(m)
        if phone in BANNED_PHONES:
            continue
        # 간단 점수: 가운데 3~4자리 길이 선호, 하이픈 포함 형태 선호
        score = (1 if len(m.group(2)) == 4 else 0) + (1 if "-" in line_text[m.start():m.end()] else 0)
        if best[0] == "" or score > ((1 if len(best[0].split("-")[1])==4 else 0)):
            best = (phone, m.start(), m.end())
    return best

def _extract_from_lines(lines: Dict[Tuple[int,int,int], List[Dict[str, Any]]]) -> Dict[str, str]:
    """
    줄 단위로 전화 → 이름(왼쪽) → 주소(다음 줄) 추출
    """
    # 라인 키 정렬(문서 상단→하단)
    keys = sorted(lines.keys(), key=lambda k: (k[0], k[1], k[2]))
    for idx, key in enumerate(keys):
        words = lines[key]
        tline = _line_text(words)
        phone, p_s, p_e = _select_best_phone(tline)
        if not phone:
            continue

        # 이름: 같은 줄의 왼쪽 영역 텍스트에서 한글 2~4자 토큰 최우선
        left_text = tline[:p_s].strip()
        left_text = LABEL_NAME.sub("", left_text)
        name_tokens = re.findall(r"[가-힣]{2,4}", left_text)
        name = name_tokens[-1] if name_tokens else ""

        # 주소: 다음 줄(있다면)에서 앞부분 식별용만
        addr = ""
        if idx + 1 < len(keys):
            next_words = lines[keys[idx+1]]
            next_text = _line_text(next_words)
            if next_text:
                addr = _address_prefix(next_text)

        return {"대여자명": name, "전화번호": phone, "주소": addr}

    return {"대여자명": "", "전화번호": "", "주소": ""}

# -------- QR 파싱 --------
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

# -------- 엔트리 포맷 --------
def _final(qr_text: str, parsed: Dict[str,str]) -> Dict[str,str]:
    model, device_id = _map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
    }

# -------- 공개 함수 --------
def make_final_entry_fast(qr_text: str, img_path: str) -> Dict[str,str]:
    im = Image.open(img_path)
    im = _resize(im, 1400)  # 프리뷰도 정확도 위해 충분 해상도
    im_p = _preprocess(im, False)
    words = _tess_data(im_p, psm=6)
    lines = _group_lines(words)
    parsed = _extract_from_lines(lines)
    return _final(qr_text, parsed)

def make_final_entry(qr_text: str, img_path: str) -> Dict[str,str]:
    """
    보강 단계: 보통/강 전처리 두 번 돌려서
    '전화번호 검출 여부'를 최우선으로 선택
    """
    im = Image.open(img_path)
    im = _resize(im, 2000)

    # pass1: 보통
    im1 = _preprocess(im, False)
    w1 = _tess_data(im1, psm=6)
    l1 = _group_lines(w1)
    p1 = _extract_from_lines(l1)

    # pass2: 강
    im2 = _preprocess(im, True)
    w2 = _tess_data(im2, psm=6)
    l2 = _group_lines(w2)
    p2 = _extract_from_lines(l2)

    # 선택 기준: 전화 검출 우선, 그 다음 이름/주소 길이
    score1 = (1 if p1.get("전화번호") else 0) + (1 if p1.get("대여자명") else 0) + (1 if p1.get("주소") else 0)
    score2 = (1 if p2.get("전화번호") else 0) + (1 if p2.get("대여자명") else 0) + (1 if p2.get("주소") else 0)
    parsed = p2 if score2 > score1 else p1

    return _final(qr_text, parsed)

