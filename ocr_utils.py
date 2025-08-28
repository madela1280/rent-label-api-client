# ocr_utils.py — CLOVA OCR 버전 (General, KR)
# - 프리뷰/정식 모두 이 로직 사용
# - 앵커(이름 첫글자) 주변만 잘라서 CLOVA로 1회 호출 → 라인/좌표 기반 추출
# - 규칙: 이름<=5자(한글), 전화는 010 또는 05** 시작 1개, 주소는 이름 라인 바로 아래 1줄
# - 금지번호: 010-7394-3535 (무시)
import os, io, re, time, base64, uuid
from typing import Tuple, List, Dict, Any, Optional
from datetime import date
import requests
from PIL import Image, ImageOps, ImageFilter

# ---------------- 환경 변수 (Render에 설정) ----------------
CLOVA_URL    = os.getenv("NCP_OCR_URL", "")
CLOVA_SECRET = os.getenv("NCP_OCR_SECRET", "")

# ---------------- 공통 유틸 ----------------
BANNED_PHONES = {"010-7394-3535"}
R_010 = re.compile(r"(010)[-\s]?(?P<m>\d{3,4})[-\s]?(?P<e>\d{0,4}\*{0,4})")
R_050X = re.compile(r"(05\d{2})[-\s]?(\d{3,4})[-\s]?(\d{4})")
KOREAN_NAME = re.compile(r"[가-힣]{2,5}")

ADDR_TOKENS = ("시","군","구","읍","면","동","리","로","길","번길","아파트","빌라","호","단지")

def _clamp(v, lo, hi): return max(lo, min(hi, v))

def _jpeg_bytes(img: Image.Image, max_w: int = 1000, quality: int = 72) -> bytes:
    # 가로 축 기준 리사이즈 + 약간의 샤픈
    w, h = img.size
    if w > max_w:
        s = max_w / float(w)
        img = img.resize((max_w, int(h*s)))
    img = ImageOps.autocontrast(img)
    img = img.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()

def _post_clova(img_bytes: bytes) -> Dict[str, Any]:
    if not CLOVA_URL or not CLOVA_SECRET:
        raise RuntimeError("CLOVA env missing (NCP_OCR_URL / NCP_OCR_SECRET)")
    b64 = base64.b64encode(img_bytes).decode("ascii")
    payload = {
        "version": "V2",
        "requestId": str(uuid.uuid4()),
        "timestamp": int(time.time() * 1000),
        "images": [{
            "format": "jpg",
            "name": "roi",
            "data": b64,
        }]
    }
    r = requests.post(
        CLOVA_URL,
        json=payload,
        headers={"Content-Type":"application/json","X-OCR-SECRET": CLOVA_SECRET},
        timeout=8,
    )
    r.raise_for_status()
    return r.json()

def _group_lines(fields: List[Dict[str, Any]], y_tol: int = 16) -> List[List[Dict[str, Any]]]:
    """bbox 중심 y 기준으로 라인 묶기"""
    items = []
    for f in fields:
        verts = f.get("boundingPoly", {}).get("vertices", [])
        if len(verts) < 4: continue
        ys = [v.get("y",0) for v in verts]
        xs = [v.get("x",0) for v in verts]
        cy = sum(ys)/len(ys)
        cx = sum(xs)/len(xs)
        items.append({"text": f.get("inferText","").strip(), "cx": cx, "cy": cy, "x": min(xs)})
    # cy 기준 정렬
    items.sort(key=lambda t: (int(t["cy"]), int(t["x"])))
    # 라인 그룹핑
    lines: List[List[Dict[str,Any]]] = []
    for it in items:
        if not lines: lines.append([it]); continue
        if abs(it["cy"] - sum(x["cy"] for x in lines[-1]) / len(lines[-1])) <= y_tol:
            lines[-1].append(it)
        else:
            lines.append([it])
    # 각 라인을 좌→우 정렬
    for ln in lines: ln.sort(key=lambda t: t["x"])
    return lines

def _line_text(ln: List[Dict[str,Any]]) -> str:
    return " ".join(t["text"] for t in ln if t["text"])

def _looks_like_addr(s: str) -> bool:
    s2 = s.replace(":", " ").strip()
    return any(tok in s2 for tok in ADDR_TOKENS)

def _pick_phone_from_text(s: str) -> str:
    # 010 우선, 없으면 050x
    m = R_010.search(s)
    if m:
        mid = m.group("m"); end = m.group("e")
        ph = f"010-{mid}-{end}" if end else f"010-{mid}-****"
        if ph in BANNED_PHONES:
            return ""
        return ph
    m2 = R_050X.search(s)
    if m2:
        return f"{m2.group(1)}-{m2.group(2)}-{m2.group(3)}"
    return ""

def _extract_by_anchor(fields: List[Dict[str,Any]], anchor_y_img: int) -> Dict[str,str]:
    """앵커 y와 가장 가까운 라인을 '이름 라인'으로 간주"""
    if not fields:
        return {"대여자명":"", "전화번호":"", "주소":""}

    lines = _group_lines(fields)
    # 각 라인의 평균 y
    ys = [int(sum(t["cy"] for t in ln)/len(ln)) for ln in lines]
    # 앵커 y에 가장 가까운 라인 찾기
    idx = min(range(len(ys)), key=lambda i: abs(ys[i] - anchor_y_img))
    name_line = lines[idx]
    name_text = _line_text(name_line)

    # 이름: 라인 좌측에서 한글 2~5자 추출 → 5자 초과시 5자까지만
    name = ""
    left_chunk = name_text.split()[0] if name_text else ""
    mname = KOREAN_NAME.findall(left_chunk)
    if mname:
        name = mname[0][:5]

    # 같은 라인 우측에서 전화 우선 탐색
    phone = _pick_phone_from_text(name_text)

    # 주소: 바로 아래 라인 1줄만
    addr = ""
    if idx + 1 < len(lines):
        cand = _line_text(lines[idx+1])
        # 주소/전화 뒤섞임 방지: 전화 제거 후 판단
        cand_no_phone = R_010.sub("", R_050X.sub("", cand)).strip(" .;:")
        if _looks_like_addr(cand_no_phone) or len(cand_no_phone) >= 6:
            addr = cand_no_phone

    # 라인에서 전화 못 찾으면 그 아래 라인들에서 최초 1개
    if not phone:
        for j in range(idx, min(idx+3, len(lines))):
            p = _pick_phone_from_text(_line_text(lines[j]))
            if p:
                phone = p; break

    # 금지번호 제거
    if phone in BANNED_PHONES:
        phone = ""

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
def _prepare_roi(path: str, anchor: Optional[Tuple[float,float]]) -> Tuple[Image.Image, int]:
    """
    원본에서 앵커 y를 기준으로 상하 여백만 포함해 잘라서 반환.
    anchor = (ax, ay) with 0..1 (프론트에서 전달)
    반환: (ROI image, anchor_y_in_roi)
    """
    im = Image.open(path).convert("RGB")
    w, h = im.size
    if not anchor:
        return im, int(h*0.5)

    ax, ay = anchor
    y = int(_clamp(ay, 0, 1) * h)

    # 노란 영역에 맞춰 상하 폭을 좁게 잡는다 (아래쪽을 더 많이 포함)
    top  = _clamp(y - int(h*0.12), 0, h-1)
    bot  = _clamp(y + int(h*0.28), top+1, h)
    roi  = im.crop((0, top, w, bot))
    return roi, (y - top)

def _run_clova(path: str, anchor: Optional[Tuple[float,float]]) -> Dict[str,str]:
    roi_img, anchor_y_in_roi = _prepare_roi(path, anchor)
    jpg = _jpeg_bytes(roi_img, max_w=1000, quality=72)
    j = _post_clova(jpg)
    fields = (j.get("images") or [{}])[0].get("fields") or []
    # 앵커는 원본 y 기준이 아니라 ROI 내부 y로 전달
    return _extract_by_anchor(fields, anchor_y_in_roi)

def make_final_entry(qr_text:str, img_path:str, anchor: Optional[Tuple[float,float]]=None):
    parsed = _run_clova(img_path, anchor)
    model, device_id = _map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
    }

def make_final_entry_fast(qr_text:str, img_path:str, anchor: Optional[Tuple[float,float]]=None):
    # 프리뷰도 동일 엔진 사용(속도 충분). 필요 시 max_w만 900으로 줄여도 됨.
    parsed = _run_clova(img_path, anchor)
    model, device_id = _map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명": parsed.get("대여자명",""),
        "전화번호": parsed.get("전화번호",""),
        "주소": parsed.get("주소",""),
        "기기번호": device_id,
        "기종": model,
    }







