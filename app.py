import os
import shutil
import hashlib
import urllib.parse
import uuid

from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse
from pydantic import BaseModel
from starlette.middleware.sessions import SessionMiddleware

import httpx
import requests
import msal

from ocr_utils import make_final_entry
from excel_utils import append_row_to_excel
APP_VERSION = os.getenv("APP_VERSION", "2025-08-25-02")

# -------------------------------
# FastAPI & Session
# -------------------------------
app = FastAPI()

app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET", "change-me"),
    same_site="none",   # ← 로그인 도메인(ms) → 우리 도메인 콜백 시 쿠키 전달
    https_only=True,    # ← SameSite=None이면 Secure(HTTPS) 필수
    max_age=3600,
    session_cookie="session",
)

from fastapi.middleware.cors import CORSMiddleware

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],   # 필요하면 특정 도메인만 지정
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------------------------------
# ENV & Constants
# -------------------------------
# ✅ 과거 값 개입 차단: 기본값은 하드코딩하되, 필요 시 환경변수로 덮어쓸 수 있게(운영이 유연해짐)
CLIENT_ID = os.getenv("CLIENT_ID", "41745db3-a5c5-4e6e-acd7-fc4ce18b1999")
TENANT_ID = os.getenv("TENANT_ID", "405ba8a3-73ff-4423-8925-d9eda360cfa7")  # GUID 또는 yourtenant.onmicrosoft.com
CLIENT_SECRET = os.getenv("CLIENT_SECRET")  # 반드시 설정 필요
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://rent-label-api-client-docker.onrender.com/callback")

# ✅ OIDC + Graph 권장 스코프 (로그인 식별을 위해 openid/profile/email은 필수로 넣자)
SCOPES = [
    "User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"
]

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

GRAPH = "https://graph.microsoft.com/v1.0"

# -------------------------------
# MSAL App 생성
# -------------------------------
def _build_msal_app():
    if not CLIENT_SECRET:
        raise RuntimeError("CLIENT_SECRET env is missing.")
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )

# -------------------------------
# 로그인 (Azure OAuth2 - MSAL)
# -------------------------------
@app.get("/login")
def login(request: Request):
    request.session["state"] = str(uuid.uuid4())
    nonce = str(uuid.uuid4())

    # ✅ 실제 authorize URL을 얻어 로그/디버그에 활용
    auth_url = _build_msal_app().get_authorization_request_url(
    scopes=["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"],
    state=request.session["state"],
    redirect_uri=REDIRECT_URI,
    prompt="login",
    response_mode="query",
)

    sep = "&" if "?" in auth_url else "?"
    return RedirectResponse(f"{auth_url}{sep}nonce={nonce}")

# -------------------------------
# 콜백 (인증 코드 → 토큰 교환)
# -------------------------------
@app.get("/callback")
async def callback(request: Request):
    if request.query_params.get("state") != request.session.get("state"):
        return JSONResponse({"error": "state mismatch"}, status_code=400)

    code = request.query_params.get("code")
    if not code:
        return JSONResponse(status_code=400, content={"error": "Authorization code missing"})

    result = _build_msal_app().acquire_token_by_authorization_code(
        code,
        scopes=["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"],
        redirect_uri=REDIRECT_URI,
    )

    if "access_token" not in result:
        return JSONResponse({"error": "Token acquire failed", "details": result}, status_code=400)

    # refresh_token 파일 저장(있으면)
    try:
        with open("refresh_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("refresh_token", ""))
    except Exception:
        pass

    # ✅ access_token 파일 저장 (우리는 이걸로 Graph 호출)
    try:
        with open("access_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("access_token", ""))
    except Exception:
        pass

       # ✅ 세션에 사용자 + 토큰 저장
    claims = result.get("id_token_claims", {}) or {}
    request.session.clear()
    request.session["user"] = {
        "name": claims.get("name"),
        "upn": claims.get("preferred_username"),
        "oid": claims.get("oid"),
    }
    request.session["tokens"] = {
        "access_token": result.get("access_token", ""),
        "refresh_token": result.get("refresh_token", ""),
    }

    return RedirectResponse("/me")

# --- /me: 세션에서 user만 확인 ---
@app.get("/me")
def me(request: Request):
    user = request.session.get("user")
    if not user:
        return RedirectResponse("/login")
    return JSONResponse({"status": "ok", "user": user})

# === 진단용: 런타임 Azure 설정/로그인 URL 확인 (강화) ===
from hashlib import sha256
from fastapi.responses import PlainTextResponse

@app.get("/__debug/azure")
def dbg_azure():
    sec = os.getenv("CLIENT_SECRET") or ""
    return {
        "client_id": CLIENT_ID,
        "tenant_id": TENANT_ID,
        "authority": AUTHORITY,
        "redirect_uri": REDIRECT_URI,
        "scopes": SCOPES,
        "secret_len": len(sec),
        "secret_fp": sha256(sec.encode()).hexdigest()[:12],
    }

@app.get("/login-url", response_class=PlainTextResponse)
def login_url():
    url = _build_msal_app().get_authorization_request_url(
        scopes=SCOPES, state="debug", redirect_uri=REDIRECT_URI, prompt="select_account", response_mode="query"
    )
    return url

# ✅ 추가: 토큰으로 실제 테넌트/사용자 확인 (누가/어느 디렉터리인지 1방에 증명)
@app.get("/whoami")
def whoami(request: Request):
    tokens = request.session.get("tokens")
    if not tokens:
        return RedirectResponse("/login")
    headers = {"Authorization": f"Bearer {tokens['access_token']}"}
    try:
        me = requests.get(f"{GRAPH}/me", headers=headers).json()
        org = requests.get(f"{GRAPH}/organization", headers=headers).json()
        return {"me": me, "organization": org}
    except Exception as e:
        return {"error": str(e)}

from uuid import uuid4

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@app.get("/", response_class=HTMLResponse)
def root():
    with open(os.path.join(BASE_DIR, "index.html"), "r", encoding="utf-8") as f:
        return HTMLResponse(f.read(), media_type="text/html; charset=utf-8")

@app.get("/__ping")
def ping(): return {"ping": str(uuid4())}

@app.get("/manifest.webmanifest", response_class=FileResponse)
def manifest():
    return FileResponse(os.path.join(BASE_DIR, "manifest.webmanifest"))

@app.get("/sw.js", response_class=FileResponse)
def sw():
    return FileResponse(os.path.join(BASE_DIR, "sw.js"))

# 콜백 경로 변형까지 모두 수용
@app.get("/callback/")
async def callback_slash(request: Request):
    return await callback(request)

@app.get("/login/callback")
async def callback_login_path(request: Request):
    return await callback(request)

@app.get("/login/callback/")
async def callback_login_path_slash(request: Request):
    return await callback(request)

# --- Graph 호출 테스트: refresh_token으로 access_token 갱신 후 /me 조회 ---
SCOPES_GRAPH = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]

from typing import Optional
from fastapi import Request

def _get_access_token(request: Optional[Request] = None):
    if request is not None:
        tok = (request.session.get("tokens") or {}).get("access_token")
        if tok:
            return tok
    try:
        with open("access_token.txt", "r", encoding="utf-8") as f:
            t = f.read().strip()
        return t if t else None
    except Exception:
        return None

@app.get("/graph/me")
def graph_me(request: Request):
    token = _get_access_token(request)
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(f"{GRAPH}/me", headers=headers)
    return JSONResponse({"status": r.status_code, "json": r.json()})

@app.get("/onedrive")
def onedrive(request: Request):
    token = _get_access_token(request)
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get("https://graph.microsoft.com/v1.0/me/drive/root/children", headers=headers)
    return JSONResponse({"status": r.status_code, "json": r.json()})

from fastapi import Body

FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

@app.post("/excel/append")
def excel_append(
    row: list = Body(...)
):
    """
    row 예시:
    ["2025-08-12","홍길동","010-1234-5678","서울시 강남구 ...","SM123456","시밀레 S6","송장번호123"]
    """
    token = _get_access_token()
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # 1) 파일 찾기
    search = requests.get(
        f"{GRAPH}/me/drive/root/search(q='{FILE_NAME}')?$top=1",
        headers=headers
    ).json()
    items = search.get("value", [])
    if not items or items[0]["name"] != FILE_NAME:
        return JSONResponse({"error": "file_not_found", "details": FILE_NAME}, status_code=404)

    item_id = items[0]["id"]

    # 2) 현재 사용 범위 조회 → 다음 행 계산
    used = requests.get(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/usedRange",
        headers=headers
    ).json()
    address = used.get("address") or f"{SHEET_NAME}!A1:A1"
    # address 예: "유축기출고!A1:G12" → 끝행 12 추출
    try:
        last_row = int(address.split("!")[1].split(":")[1][1:])
    except Exception:
        last_row = 1
    next_row = last_row + 1
    target = f"A{next_row}:G{next_row}"  # 열 수는 필요에 맞게 조정

    # 3) 쓰기
    resp = requests.patch(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
        headers=headers,
        json={"values": [row]},
    )
    if resp.status_code != 200:
        return JSONResponse({"error": "write_failed", "status": resp.status_code, "text": resp.text}, status_code=500)

    return {"status": "ok", "range": target, "written": row}

# --- 사진 + OCR + OneDrive 엑셀 쓰기 ---
@app.post("/process-ocr/")
async def process_ocr(
    qr_text: str = Form(""),
    image: UploadFile = File(...),
    dry: int = Form(0)
):
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as f:
        shutil.copyfileobj(image.file, f)
    try:
        # 1) OCR
        result = make_final_entry(qr_text, temp_path)

        # 2) 엑셀 행
        row = [
            result.get("출고일", ""),
            result.get("대여자명", ""),
            result.get("전화번호", ""),
            result.get("주소", ""),
            result.get("기기번호", ""),
            result.get("기종", ""),
            result.get("송장번호", ""),
        ]

        # 3) 미리보기면 저장 생략
        if dry:
            return {"status": "preview", "data": result}

        # 4) 저장
        ok, info = write_row_to_onedrive(row)
        if not ok:
            return {"status": "ocr_ok_but_write_failed", "data": result, "write_error": info}
        return {"status": "success", "data": result, "write_info": info}
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

# === OneDrive에 한 줄 쓰는 헬퍼 ===
def write_row_to_onedrive(row):
    token = _get_access_token()
    if not token:
        return False, {"error": "no_access_token"}

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
    SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

    # 1) 파일 찾기
    search = requests.get(
        f"{GRAPH}/me/drive/root/search(q='{FILE_NAME}')?$top=1", headers=headers
    ).json()
    items = search.get("value", [])
    if not items or items[0]["name"] != FILE_NAME:
        return False, {"error": "file_not_found", "file": FILE_NAME}
    item_id = items[0]["id"]

    # 2) 사용범위 조회 → 다음 행 계산
    used = requests.get(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/usedRange",
        headers=headers,
    ).json()
    address = used.get("address") or f"{SHEET_NAME}!A1:A1"
    try:
        last_row = int(address.split("!")[1].split(":")[1][1:])
    except Exception:
        last_row = 1
    next_row = last_row + 1
    target = f"A{next_row}:G{next_row}"

    # 3) 쓰기
    resp = requests.patch(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
        headers=headers,
        json={"values": [row]},
    )
    if resp.status_code != 200:
        return False, {"error": "write_failed", "status": resp.status_code, "text": resp.text}
    return True, {"range": target}

# --- OCR 디버그: 원문/신뢰도 확인용 ---
from pytesseract import image_to_data, Output
from PIL import Image, ImageOps, ImageFilter

def _preprocess_for_ocr(p: str) -> Image.Image:
    # 가벼운 전처리: 흑백 → 대비/샤픈 → 이진화
    img = Image.open(p).convert("L")
    img = ImageOps.autocontrast(img)
    img = img.filter(ImageFilter.UnsharpMask(radius=1.0, percent=150, threshold=3))
    # 간단 이진화
    img = img.point(lambda x: 255 if x > 180 else 0, mode="1")
    return img.convert("L")

@app.post("/ocr/debug")
async def ocr_debug(image: UploadFile = File(...)):
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as f:
        shutil.copyfileobj(image.file, f)
    try:
        img = _preprocess_for_ocr(temp_path)
        cfg = "--oem 3 --psm 6 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-:/()#"
        data = image_to_data(img, config=cfg, output_type=Output.DICT, lang="eng+kor")

        out = []
        n = min(len(data["text"]), 200)
        for i in range(n):
            txt = (data["text"][i] or "").strip()
            if not txt:
                continue
            out.append({
                "text": txt,
                "conf": float(data["conf"][i]),
                "bbox": [int(data["left"][i]), int(data["top"][i]), int(data["width"][i]), int(data["height"][i])]
            })

        return {"count": len(out), "items": out}
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.post("/excel/upload")
def excel_upload(payload: dict = Body(...)):
    # 프론트에서 보낸 값 → 엑셀 한 행으로 변환
    row = [
        payload.get("shipDate", ""),
        payload.get("name", ""),
        payload.get("phone", ""),
        payload.get("addr", ""),
        payload.get("deviceId", ""),
        payload.get("model", ""),
        payload.get("invoice", ""),
    ]

    ok, info = write_row_to_onedrive(row)  # 이미 위에서 정의된 헬퍼 사용
    if not ok:
        return JSONResponse({"status": "error", "write_error": info}, status_code=500)
    return {"status": "ok", "range": info.get("range")}

@app.get("/__version")
def version():
    return {"version": APP_VERSION}

if __name__ == "__main__":
    import uvicorn, os
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))




