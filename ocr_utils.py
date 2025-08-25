# ocr_utils.py
# ------------------------------------------------------------
# 기존 서버 흐름을 그대로 유지합니다.
# - make_final_entry(qr_text, image_path) 시그니처/키 이름 동일
# - 의존성: Pillow, pytesseract (OpenCV는 선택적)
# - TESSERACT_CMD 환경변수로 실행 경로 지정 가능(없으면 기본값)
# ------------------------------------------------------------

import os
import re
from datetime import date
from typing import List, Tuple

from PIL import Image, ImageOps, ImageFilter
import pytesseract

# (선택) OpenCV가 있으면 ROI 탐지에 활용, 없으면 비율 기반으로 동작
try:
    import cv2
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

# 시스템별 테서랙트 경로 지정(없으면 기본값 사용)
pytesseract.pytesseract.tesseract_cmd = os.getenv("TESSERACT_CMD", pytesseract.pytesseract.tesseract_cmd)


# =========================
# 1) 전처리 & OCR 유틸
# =========================
def _preprocess(img: Image.Image, strong: bool = False) -> Image.Image:
    """OCR 안정화를 위한 기본 전처리"""
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=220, threshold=2))
        # 가벼운 이진화
        g = g.point(lambda x: 255 if x > 170 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g


def _ocr_text(img: Image.Image, allow_kor: bool = True) -> str:
    """한/영 OCR"""
    cfg = "--oem 3 --psm 6"
    lang = "kor+eng" if allow_kor else "eng"
    return pytesseract.image_to_string(img, config=cfg, lang=lang)


def _resize_image(img: Image.Image, max_w: int = 1600) -> Image.Image:
    """너무 큰 이미지는 축소하여 OCR 속도 개선"""
    w, h = img.size
    if w > max_w:
        scale = max_w / float(w)
        return img.resize((max_w, int(h * scale)))
    return img


# =========================
# 2) 정규식/정규화
# =========================
INVOICE_RE = re.compile(r"(\d{4})[- ]?(\d{4})[- ]?(\d{4})")
PHONE_RE   = re.compile(r"((?:01[016789]|05\d{2})[- ]?\d{3,4}[- ]?(?:\d{4}|\*{4}))")

def _normalize_invoice(s: str) -> str:
    s = s.replace("—", "-").replace("–", "-")
    m = INVOICE_RE.search(s)
    if not m:
        return ""
    # 050x-xxxx-xxxx 처럼 첫 그룹이 05로 시작하면 송장으로 취급하지 않음
    if m.group(1).startswith("05"):
        return ""
    return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"


def _likely_name(s: str) -> bool:
    """순수 한글 2~4자 → 이름으로 가정"""
    return bool(re.fullmatch(r"[가-힣]{2,4}", s))


def _is_address(s: str) -> bool:
    """주소 판별 간단 로직"""
    return any(tok in s for tok in ["시", "군", "구", "읍", "면", "동", "리", "로", "길"])


def _normalize_phone(s: str) -> str:
    m = PHONE_RE.search(s)
    if not m:
        return ""
    return m.group().replace(" ", "").replace("--", "-")


# =========================
# 3) ROI 추출
# =========================
def _crop_invoice_roi_cv2(path: str):
    """OpenCV 사용: 파란 '받는분' 세로 띠 기준으로 우측 베이지 라벨 넓게 포함."""
    if not HAS_CV2:
        return None
    try:
        img = cv2.imread(path)
        h, w = img.shape[:2]

        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        # 청록~파랑 범위 (스티커 좌측 '받는분' 띠 색)
        m1 = cv2.inRange(hsv, (85, 80, 60), (110, 255, 255))
        m2 = cv2.inRange(hsv, (110, 80, 60), (135, 255, 255))
        mask = cv2.bitwise_or(m1, m2)

        contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not contours:
            return None

        cnt = max(contours, key=cv2.contourArea)
        x, y, w0, h0 = cv2.boundingRect(cnt)

        # 파란띠 기준으로 오른쪽/하단 넓게 확장(샘플 레이아웃 기준)
        x1 = max(0, x - int(0.02 * w))
        y1 = max(0, y - int(0.05 * h))
        x2 = min(w, x + w0 + int(0.72 * w))
        y2 = min(h, y + h0 + int(0.25 * h))

        roi = img[y1:y2, x1:x2]
        if roi.size == 0:
            return None

        return Image.fromarray(cv2.cvtColor(roi, cv2.COLOR_BGR2RGB))
    except Exception:
        return None


def _crop_invoice_roi_ratio(path: str) -> Image.Image:
    """OpenCV 미사용 시: 화면 비율 기반으로 중앙~하단 영역을 넓게 자름."""
    im = Image.open(path)
    W, H = im.size
    # 샘플 라벨 기준 경험값: 좌 4% ~ 우 86%, 상 28% ~ 하 76%
    x1 = int(W * 0.04)
    y1 = int(H * 0.28)
    x2 = int(W * 0.86)
    y2 = int(H * 0.76)
    return im.crop((x1, y1, x2, y2))


def _crop_invoice_roi(path: str) -> Image.Image:
    roi = _crop_invoice_roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _crop_invoice_roi_ratio(path)


# =========================
# 4) 텍스트 → 필드 파싱
# =========================
def _parse_fields(lines: List[str]) -> dict:
    invoice, phone, name, addr = "", "", "", ""

    for i, ln in enumerate(lines):
        # 송장번호
        if not invoice:
            m = re.search(r'(\d{4})[-\s]?(\d{4})[-\s]?(\d{4})', ln)
            if m and not m.group(1).startswith("05"):   # 0507-xxxx-xxxx 제외
                invoice = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

        # 전화번호
        if not phone:
            m = re.search(r'(01[016789]|05\d{2})[-\s]?\d{3,4}[-\s]?(?:\d{4}|\*{4})', ln)
            if m:
                phone = m.group().replace(" ", "").replace("--", "-")
                # 이름: 전화 앞부분
                left = ln.split(m.group())[0].strip()
                left = re.sub(r"[\[\](){}|,.;:]+", "", left).strip()
                if _likely_name(left):
                    name = left
                # 주소: 전화 있는 줄 다음 줄
                if i + 1 < len(lines):
                    nxt = lines[i + 1].strip()
                    if _is_address(nxt):
                        addr = nxt

    # 이름 보정
    if not name:
        for ln in lines:
            if _likely_name(ln):
                name = ln.strip()
                break

    # 주소 보정
    if not addr:
        addr_candidates = [ln.strip() for ln in lines if _is_address(ln)]
        if addr_candidates:
            addr = max(addr_candidates, key=len)

    return {
        "송장번호": invoice,
        "전화번호": phone,
        "대여자명": name,
        "주소": addr,
    }


# =========================
# 5) QR → 기종/기기번호 매핑
# =========================
def _map_model_device(qr_text: str) -> Tuple[str, str]:
    """
    - SM123456 → 심포니 / 123456
    - LT123456 → 락티나 / 123456
    - S/M/F/G/C 시작 → 첫글자 매핑, 기기번호는 원문 그대로
    """
    raw = (qr_text or "").strip()
    u = re.sub(r"[^A-Z0-9]", "", raw.upper())
    MAP = {"SM": "심포니", "LT": "락티나",
           "S": "스윙", "M": "스윙맥스", "F": "프리스타일", "G": "각시밀", "C": "시밀레"}

    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2:
        return MAP.get(m2.group(1), "-"), m2.group(2)

    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1:
        return MAP.get(m1.group(1), "-"), raw

    return "-", ""


# =========================
# 6) 메인 엔트리 (기존 시그니처 유지)
# =========================
def make_final_entry(qr_text: str, 송장_image_path: str):
    # ROI 자르기
    roi = _crop_invoice_roi(송장_image_path)
    roi = _resize_image(roi, 1600)

    # 디버그 저장 (ROI 이미지)
    try:
        os.makedirs("_debug", exist_ok=True)
        roi.save(os.path.join("_debug", "roi.jpg"))
    except Exception:
        pass

    # OCR: 약하게 → strong은 결과 짧을 때만
    txt = _ocr_text(_preprocess(roi, strong=False))
    if len(re.sub(r"\s+", "", txt)) < 20:   # 20자 미만이면 strong OCR
        txt = _ocr_text(_preprocess(roi, strong=True))

    # 디버그 저장 (OCR 라인)
    try:
        with open(os.path.join("_debug", "ocr_lines.txt"), "w", encoding="utf-8") as f:
            f.write(txt)
    except Exception:
        pass

    lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]

    # 필드 파싱
    parsed = _parse_fields(lines)

    # QR → 기종/기기번호
    model, device_id = _map_model_device(qr_text)

    # 출고일(서버 날짜)
    ship_date = date.today().isoformat()

    # 결과
    out = {
        "출고일": ship_date,
        "대여자명": parsed.get("대여자명", ""),
        "전화번호": parsed.get("전화번호", ""),
        "주소": parsed.get("주소", ""),
        "기기번호": device_id,
        "기종": model,
        "송장번호": parsed.get("송장번호", ""),
    }

    # 보정: 전화는 있는데 이름이 비었을 경우
    if out["전화번호"] and not out["대여자명"]:
        for ln in lines:
            if out["전화번호"] in ln:
                left = ln.split(out["전화번호"])[0].strip()
                left = left.split("/")[-1].strip()
                left = re.sub(r"[\[\](){}|,.;:]+", "", left).strip()
                if _likely_name(left):
                    out["대여자명"] = left
                    break

    return out

def make_final_entry_fast(qr_text: str, 송장_image_path: str):
    """
    초고속 프리뷰:
    - ROI 자름 → 가로 1024로 축소
    - 숫자 전용 OCR(영문만, whitelist)로 '송장번호/전화' 우선 추출
    - 송장번호는 '05'로 시작하면 제외
    - 이름/주소는 비워둘 수 있음(필요 최소만 빠르게)
    """
    # ROI + 축소
    roi = _crop_invoice_roi(송장_image_path)
    roi = _resize_image(roi, 1024)

    # 숫자 전용 OCR (빠름)
    cfg_fast = "--oem 3 --psm 6 -c tessedit_char_whitelist=0123456789-"
    txt_num = pytesseract.image_to_string(_preprocess(roi, strong=False), config=cfg_fast, lang="eng")
    lines_num = [ln.strip() for ln in txt_num.splitlines() if ln.strip()]

    # 기본 값
    ship_date = date.today().isoformat()
    model, device_id = _map_model_device(qr_text)
    invoice, phone = "", ""

    # 숫자만으로 송장/전화 추출
    for ln in lines_num:
        if not phone:
            m = re.search(r'(01[016789]|05\d{2})[-]?\d{3,4}[-]?(?:\d{4}|\*{4})', ln)
            if m:
                phone = re.sub(r'\s+', '', m.group()).replace('--', '-')
        if not invoice:
            m = re.search(r'(\d{4})[-]?(\d{4})[-]?(\d{4})', ln)
            if m and not m.group(1).startswith('05'):  # 05로 시작하면 송장 제외
                invoice = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        if phone and invoice:
            break

    return {
        "출고일": ship_date,
        "대여자명": "",
        "전화번호": phone,
        "주소": "",
        "기기번호": device_id,
        "기종": model,
        "송장번호": invoice,
    }


