# ocr_utils.py — 이름/전화 필수, 주소는 앞부분(식별용)만 추출 버전
# - 속도: 프리뷰는 저해상도 + 가벼운 전처리, 정식은 보강 단계만 강처리
# - 전화: 전체 숫자 유지(마스킹 제거), 특정 금지 번호는 공란 처리
# - 이름: 전화번호가 있는 "같은 줄"의 왼쪽에서 한글 2~4글자 토큰을 우선 추출
# - 주소: 전화줄 "다음 줄"에서 앞부분만(예: 서울 강남구 개포동 12) 잘라 식별용으로 반환
# - QR에서 기종/기기번호는 그대로 유지

import os, re
from datetime import date
from typing import List, Tuple
from PIL import Image, ImageOps, ImageFilter
import pytesseract

# ---------------- 환경 ----------------
try:
    import cv2  # 현재 로직에서는 사용하지 않지만, 설치 유무만 체크
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# Tesseract 경로 환경변수 우선 사용
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# ---------------- OCR 유틸 ----------------
def _preprocess(img: Image.Image, strong: bool=False) -> Image.Image:
    """가벼운 전처리: 속도 우선 / strong일 때만 이진화 강화."""
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=220, threshold=2))
        # 이진화: 너무 공격적으로 하지 않음 (주소 한글 깨짐 방지)
        g = g.point(lambda x: 255 if x > 165 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g

def _ocr_text(img: Image.Image, psm:int=6) -> str:
    """한국어 위주 인식: eng 제외로 속도 향상/오인식 감소"""
    try:
        # 다단/문단 → 6, 한 줄 → 7
        return pytesseract.image_to_string(img, config=f"--oem 3 --psm {psm}", lang="kor")
    except Exception:
        return ""

def _resize(img: Image.Image, max_w:int=1400) -> Image.Image:
    w,h = img.size
    if w > max_w:
        s = max_w/float(w)
        return img.resize((max_w, int(h*s)))
    return img

# ---------------- 규칙 ----------------
# 010-1234-5678, 010 1234 5678, 010.1234.5678 모두 허용
R_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")
LABEL_NAME = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

# 주소 토큰(앞부분 식별용)
ADDR_TOKENS = ("시","도","군","구","읍","면","동","리","로","길","번길","호")
BANNED_PHONES = {"010-7394-3535"}  # 금지 번호는 공란 처리

def _clean(s:str) -> str:
    return re.sub(r"[|\[\]{}<>]+"," ",s).strip()

def _norm_phone(m: re.Match) -> str:
    """정규화: 010-1234-5678 형태(마스킹 제거, 전부 보존)."""
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

def _extract_name_from_left(text_left:str) -> str:
    """
    왼쪽 영역에서 가장 끝쪽 한글 토큰(2~4글자)을 이름으로.
    너무 긴 토큰(5+), 숫자/영문 포함 토큰은 제외.
    """
    s = LABEL_NAME.sub("", text_left).strip()
    # 공백/구분자 기준 토큰화
    toks = re.findall(r"[가-힣]{2,4}", s)
    if toks:
        return toks[-1]  # 가장 오른쪽(끝) 토큰이 이름일 확률이 높음
    return ""

def _looks_like_address_line(s:str) -> bool:
    s2 = LABEL_ADDR.sub("", s).strip()
    if not s2: return False
    # 주소 앞부분 특징: 토큰(시/구/동 등) + 숫자 일부 동반
    return (any(t in s2 for t in ADDR_TOKENS) or bool(re.search(r"\d", s2)))

def _address_prefix(s:str) -> str:
    """
    주소 라인에서 '앞부분 식별용'만 잘라 반환.
    예: '서울특별시 강남구 개포동 12...' → '서울 강남구 개포동 12'
    - 시/도/구/군/동 등 행정 토큰과 첫 숫자까지 포함
    - 너무 길어지지 않게 18~24자 정도로 클립 (가독/일관성)
    """
    s2 = LABEL_ADDR.sub("", s).strip()
    if not s2:
        return ""

    # 괄호/상세 시작 전까지만
    s2 = re.split(r"[(),]", s2)[0].strip()

    # '…동 12' 같은 형태로 첫 숫자 토큰을 찾으면 그 숫자까지 포함
    mnum = re.search(r"\d+", s2)
    cut = None
    if mnum:
        # 숫자 다음에 오는 단어 경계까지만
        cut = mnum.end()
        # 숫자 뒤 공백 하나까지 포함
        if cut < len(s2) and s2[cut] == ' ':
            cut += 1

    head = s2[:cut] if cut else s2

    # 행정구역 토큰 정규화(시/도/구/동 앞뒤 공백 정리)
    head = re.sub(r"\s+", " ", head).strip()

    # '서울특별시' → '서울' 간단 줄임(가독)
    head = head.replace("서울특별시", "서울").replace("부산광역시","부산").replace("대구광역시","대구") \
               .replace("인천광역시","인천").replace("광주광역시","광주").replace("대전광역시","대전") \
               .replace("울산광역시","울산").replace("세종특별자치시","세종")

    # 너무 길면 앞 22자 정도로 클립
    if len(head) > 24:
        head = head[:24].rstrip()

    # 최소 식별: 구/군/시 중 하나라도 포함되면 OK
    if not any(t in head for t in ("구","군","시")) and not re.search(r"\d", head):
        # 식별력 부족하면 빈값
        return ""
    return head

def _parse_fields(lines: List[str]) -> dict:
    """전화번호가 있는 줄을 기준으로: 왼쪽=이름, 다음 줄=주소(식별용 앞부분만)."""
    lines = [_clean(x) for x in lines if x and x.strip()]
    phone, name, addr = "", "", ""

    for i, ln in enumerate(lines):
        m = R_010.search(ln)
        if not m:
            continue

        # 전화번호 정규화 (마스킹 없음)
        phone = _norm_phone(m)
        if phone in BANNED_PHONES:
            phone = ""

        # 같은 줄 왼쪽에서 이름 후보 추출
        left = ln[:m.start()]
        name = _extract_name_from_left(left)

        # 바로 아래 줄 1줄만 주소 후보
        if i + 1 < len(lines):
            cand = lines[i+1]
            if _looks_like_address_line(cand):
                addr = _address_prefix(cand)

        break  # 가장 먼저 찾은 전화 한 번만 사용

    return {"대여자명": name, "전화번호": phone, "주소": addr}

# ---------------- QR → 기종/기기번호 ----------------
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

# ---------------- 메인 엔트리 ----------------
def make_final_entry(qr_text:str, img_path:str):
    """정식 OCR: 텍스트가 빈약하면 강처리 보강."""
    im = Image.open(img_path)
    im = _resize(im, 1400)
    txt = _ocr_text(_preprocess(im, False), psm=6)
    if len(re.sub(r"\s+","",txt)) < 16:
        # 너무 빈약하면 강처리(한 번만)
        txt2 = _ocr_text(_preprocess(im, True), psm=6)
        if len(re.sub(r"\s+","",txt2)) > len(re.sub(r"\s+","",txt)):
            txt = txt2

    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    parsed = _parse_fields(lines)
    model, device_id = _map_model_device(qr_text)

    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
    }

def make_final_entry_fast(qr_text:str, img_path:str):
    """프리뷰: 해상도 낮게 + 가벼운 전처리, 1회 OCR만."""
    im = Image.open(img_path)
    im = _resize(im, 900)  # 프리뷰는 더 작게
    txt = _ocr_text(_preprocess(im, False), psm=6)
    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
    parsed = _parse_fields(lines)
    model, device_id = _map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
    }



