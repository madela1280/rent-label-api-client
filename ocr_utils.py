# ocr_utils.py
# ------------------------------------------------------------
# 교체본 (기존 흐름/엔드포인트/키 그대로)
# - make_final_entry(qr_text, image_path): 정식 OCR (정밀)
# - make_final_entry_fast(qr_text, image_path): 프리뷰 OCR (빠름)
# - 이름/전화/주소/송장번호 모두 추출 (송장번호 재활성화)
# - ROI 실패 시 전체 이미지 백업 OCR → 빈값 방지
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
pytesseract.pytesseract.tesseract_cmd = os.getenv(
    "TESSERACT_CMD",
    pytesseract.pytesseract.tesseract_cmd
)

# =========================
# 1) 전처리 & OCR 유틸
# =========================
def _preprocess(img: Image.Image, strong: bool = False) -> Image.Image:
    """OCR 안정화를 위한 기본 전처리(빠르고 안전하게)"""
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=220, threshold=2))
        # 가벼운 이진화
        g = g.point(lambda x: 255 if x > 170 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g


def _ocr_text(img: Image.Image, allow_kor: bool = True, psm: int = 6) -> str:
    """한/영 OCR"""
    cfg = f"--oem 3 --psm {psm}"
    lang = "kor+eng" if allow_kor else "eng"
    try:
        return pytesseract.image_to_string(img, config=cfg, lang=lang)
    except Exception:
        return ""


def _resize_image(img: Image.Image, max_w: int = 1400) -> Image.Image:
    """너무 큰 이미지는 축소하여 OCR 속도/안정성 개선"""
    w, h = img.size
    if w > max_w:
        scale = max_w / float(w)
        return img.resize((max_w, int(h * scale)))
    return img

# =========================
# 2) 정규식/도우미
# =========================
PHONE_RE = re.compile(
    r"(?:\(?0?1[016789]\)?|0?5\d{2})[-\s\.]?\d{3,4}[-\s\.]?(?:\d{4}|\*{4})"
)
INVOICE_RE = re.compile(
    r"(?:운송장|송장|송장번호|운송장번호|운송장 No\.?|Invoice|Tracking)\D{0,8}([0-9\-\s]{10,18})",
    re.IGNORECASE,
)
MAYBE_INVOICE_FALLBACK = re.compile(r"\b\d[\d\-\s]{9,17}\d\b")

LABEL_NAME_RE = re.compile(r"^(받는.?|수령인|수취인|이름)\s*[:：]?\s*", re.I)
LABEL_ADDR_RE = re.compile(r"^(주소|배달지|배송지)\s*[:：]?\s*", re.I)

def _clean(s: str) -> str:
    return re.sub(r"[|\[\]{}<>]+", " ", s).strip()

def _normalize_phone(s: str) -> str:
    m = PHONE_RE.search(s)
    if not m:
        return ""
    ph = m.group()
    ph = re.sub(r"[^\d\*]", "-", ph)
    ph = re.sub(r"-{2,}", "-", ph)
    # 괄호/0 접두 정리
    ph = ph.replace("(-", "-").replace("-)", "-").strip("-")
    return ph

def _likely_name(s: str) -> bool:
    """순수 한글 2~5자 → 이름 후보"""
    s = re.sub(LABEL_NAME_RE, "", s).strip()
    return bool(re.fullmatch(r"[가-힣]{2,5}", s))

def _is_address(s: str) -> bool:
    """주소 판별 간단 로직"""
    s = re.sub(LABEL_ADDR_RE, "", s)
    return any(tok in s for tok in ["시", "군", "구", "읍", "면", "동", "리", "로", "길", "번길", "아파트", "빌라", "호"])

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
    # 경험값: 좌 4% ~ 우 86%, 상 25% ~ 하 80%
    x1 = int(W * 0.04)
    y1 = int(H * 0.25)
    x2 = int(W * 0.86)
    y2 = int(H * 0.80)
    return im.crop((x1, y1, x2, y2))


def _crop_invoice_roi(path: str) -> Image.Image:
    roi = _crop_invoice_roi_cv2(path) if HAS_CV2 else None
    return roi if roi is not None else _crop_invoice_roi_ratio(path)

# =========================
# 4) 텍스트 → 필드 파싱
# =========================
def _extract_invoice(text: str) -> str:
    """송장번호: 레이블 기반 1순위, 없으면 숫자 10~18자리 후보 1개"""
    text1 = text.replace(" ", "")
    # 레이블 근접 매칭
    m = INVOICE_RE.search(text)
    if m:
        raw = m.group(1)
        digits = re.sub(r"\D", "", raw)
        if 10 <= len(digits) <= 18:
            return digits
    # 폴백: 전체에서 10~18자리 연속 숫자
    m2 = MAYBE_INVOICE_FALLBACK.search(text1)
    if m2:
        digits = re.sub(r"\D", "", m2.group())
        if 10 <= len(digits) <= 18:
            return digits
    return ""

def _parse_fields(text: str) -> dict:
    """
    규칙:
    - 전화번호: 010-1234-**** / 010-1234-5678 / 05xx-1234-5678 / (010)1234-5678
    - 대여자명: 전화번호 '같은 줄'에서 앞쪽 텍스트(레이블 제거) → 없으면 바로 윗줄에서 한글 2~5자
    - 주소: 전화번호 줄 '아래쪽' 1~2줄 이어붙이기(레이블 제거)
    - 송장번호: 레이블+숫자 또는 10~18자리 숫자 패턴
    """
    # 라인 전처리
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    clean_lines = [_clean(ln) for ln in lines]

    phone, name, addr, invoice = "", "", "", ""
    phone_idx = -1

    # 1) 전화/송장 먼저 찾기 (가장 먼저 나오는 1개 사용)
    for i, ln in enumerate(clean_lines):
        if not phone:
            ph = _normalize_phone(ln)
            if ph:
                phone = ph
                phone_idx = i
        if not invoice:
            # 송장은 전체 텍스트 기준으로 뒤에 별도 탐색하므로 여기서는 건너뛰기
            pass

    # 2) 이름: 전화 줄의 '앞' 텍스트 or 윗줄
    if phone and phone_idx >= 0:
        ln = clean_lines[phone_idx]
        try:
            left = ln.split(phone)[0].strip()
        except Exception:
            # phone 포맷이 바뀌며 split이 실패할 수 있으니 재탐색
            m = PHONE_RE.search(ln)
            left = ln[:m.start()].strip() if m else ""
        left = LABEL_NAME_RE.sub("", left).strip(" :·,.;|")
        if _likely_name(left):
            name = left
        elif phone_idx - 1 >= 0:
            up = LABEL_NAME_RE.sub("", clean_lines[phone_idx - 1]).strip()
            if _likely_name(up):
                name = up

    # 3) 주소: 전화 줄 아래 1~2줄
    if phone_idx >= 0:
        cand = []
        if phone_idx + 1 < len(clean_lines):
            l1 = LABEL_ADDR_RE.sub("", clean_lines[phone_idx + 1]).strip()
            if l1:
                cand.append(l1)
        if phone_idx + 2 < len(clean_lines):
            l2 = LABEL_ADDR_RE.sub("", clean_lines[phone_idx + 2]).strip()
            if l2 and (_is_address(l2) or re.search(r"\d", l2)):
                cand.append(l2)
        addr = " ".join(cand).strip()

    # 4) 송장번호: 전체 텍스트 기준
    invoice = _extract_invoice("\n".join(clean_lines))

    return {
        "전화번호": phone,
        "대여자명": name,
        "주소": addr,
        "송장번호": invoice,
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
    MAP = {
        "SM": "심포니",
        "LT": "락티나",
        "S": "스윙",
        "M": "스윙맥스",
        "F": "프리스타일",
        "G": "각시밀",
        "C": "시밀레",
    }

    m2 = re.match(r"^(SM|LT)(\d{2,})$", u)
    if m2:
        return MAP.get(m2.group(1), "-"), m2.group(2)

    m1 = re.match(r"^([SMFLGC])[A-Z0-9]*$", u)
    if m1:
        return MAP.get(m1.group(1), "-"), raw

    return "-", ""

# =========================
# 6) 공통 OCR 시도 (ROI → 실패 시 전체)
# =========================
def _try_ocr_strategies(img_path: str) -> str:
    """
    빠르고 안정적인 2단계 시도:
    1) ROI(약하게 psm6) → 부족하면 strong psm6
    2) 여전히 부족하면 '원본 전체'(약하게 psm6) → strong psm6
    """
    # 1) ROI
    roi = _crop_invoice_roi(img_path)
    roi = _resize_image(roi, 1400)

    txt = _ocr_text(_preprocess(roi, strong=False), allow_kor=True, psm=6)
    if len(re.sub(r"\s+", "", txt)) < 18:  # 너무 짧으면 강하게
        txt2 = _ocr_text(_preprocess(roi, strong=True), allow_kor=True, psm=6)
        if len(re.sub(r"\s+", "", txt2)) > len(re.sub(r"\s+", "", txt)):
            txt = txt2

    # ROI에서 연락처/주소/송장 추출 성공 여부 확인
    parsed = _parse_fields(txt)
    enough = parsed.get("전화번호") or parsed.get("주소") or parsed.get("송장번호")
    if enough:
        return txt

    # 2) 전체 이미지 재시도
    full = Image.open(img_path)
    full = _resize_image(full, 1600)

    txt_full = _ocr_text(_preprocess(full, strong=False), allow_kor=True, psm=6)
    if len(re.sub(r"\s+", "", txt_full)) < 18:
        txt_full2 = _ocr_text(_preprocess(full, strong=True), allow_kor=True, psm=6)
        if len(re.sub(r"\s+", "", txt_full2)) > len(re.sub(r"\s+", "", txt_full)):
            txt_full = txt_full2

    return txt_full

# =========================
# 7) 메인 엔트리 (정식 OCR)
# =========================
def make_final_entry(qr_text: str, 송장_image_path: str):
    # OCR 수행
    txt = _try_ocr_strategies(송장_image_path)

    # 디버그 저장
    try:
        os.makedirs("_debug", exist_ok=True)
        with open(os.path.join("_debug", "ocr_lines_full.txt"), "w", encoding="utf-8") as f:
            f.write(txt)
    except Exception:
        pass

    # 필드 파싱
    parsed = _parse_fields(txt)

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

    # 보정: 전화는 있는데 이름이 비었을 경우 (전화 포함된 줄의 전화 '앞' 텍스트 재시도)
    if out["전화번호"] and not out["대여자명"]:
        for ln in [ln.strip() for ln in txt.splitlines() if ln.strip()]:
            if out["전화번호"].replace("-", "") in re.sub(r"\D", "", ln):
                left = ln.split(out["전화번호"].split("-")[0])[0].strip()
                left = LABEL_NAME_RE.sub("", left).split("/")[-1].strip(" ,.;:|")
                if _likely_name(left):
                    out["대여자명"] = left
                    break

    return out

# =========================
# 8) 프리뷰 (빠름)
# =========================
def make_final_entry_fast(qr_text: str, 송장_image_path: str):
    """
    프리뷰:
    - ROI 기준 psm6 약한 OCR 1회
    - 연락처/주소/송장 빠르게 파싱 (부족하면 전체 이미지 1회만 추가 시도)
    - 1~3초 내 반환 목표
    """
    # 1) ROI 1패스
    roi = _crop_invoice_roi(송장_image_path)
    roi = _resize_image(roi, 1000)
    txt = _ocr_text(_preprocess(roi, strong=False), allow_kor=True, psm=6)

    parsed = _parse_fields(txt)
    # 2) 폴백: 아무것도 못 찾으면 전체 1패스
    if not (parsed.get("전화번호") or parsed.get("주소") or parsed.get("송장번호")):
        full = Image.open(송장_image_path)
        full = _resize_image(full, 1200)
        txt2 = _ocr_text(_preprocess(full, strong=False), allow_kor=True, psm=6)
        parsed = _parse_fields(txt2)

    model, device_id = _map_model_device(qr_text)
    ship_date = date.today().isoformat()

    return {
        "출고일": ship_date,
        "대여자명": parsed.get("대여자명", ""),
        "전화번호": parsed.get("전화번호", ""),
        "주소": parsed.get("주소", ""),
        "기기번호": device_id,
        "기종": model,
        "송장번호": parsed.get("송장번호", ""),
    }
