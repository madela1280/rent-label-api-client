import os
import shutil
import hashlib
import uuid

from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File, Body
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse, PlainTextResponse
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware

import requests
import msal
from uuid import uuid4
from typing import Optional, Dict, Any

from ocr_utils import make_final_entry, make_final_entry_fast

APP_VERSION = os.getenv("APP_VERSION", "2025-08-26-01")

# -------------------------------
# FastAPI & Session
# -------------------------------
app = FastAPI()

app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET", "change-me"),
    same_site="none",
    https_only=True,
    max_age=3600,
    session_cookie="session",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# -------------------------------
# ENV & Constants
# -------------------------------
CLIENT_ID = os.getenv("CLIENT_ID", "41745db3-a5c5-4e6e-acd7-fc4ce18b1999")
TENANT_ID = os.getenv("TENANT_ID", "405ba8a3-73ff-4423-8925-d9eda360cfa7")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://rent-label-api-client-docker.onrender.com/callback")

# ★ offline_access + openid/profile 추가 (리프레시/사일런트)
SCOPES = [
    "User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All",
    "offline_access", "openid", "profile"
]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH = "https://graph.microsoft.com/v1.0"

# -------------------------------
# MSAL App + Token Cache (영구)
# -------------------------------
CACHE_PATH = os.getenv("MSAL_CACHE_PATH", "msal_cache.bin")

def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception:
            pass
    return cache

def _save_cache(cache: msal.SerializableTokenCache):
    if cache.has_state_changed:
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(cache.serialize())

def _msal_app():
    if not CLIENT_SECRET:
        raise RuntimeError("CLIENT_SECRET env is missing.")
    cache = _load_cache()
    app_ = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET, token_cache=cache
    )
    return app_, cache

def _acquire_token_silent() -> Optional[str]:
    app_, cache = _msal_app()
    accounts = app_.get_accounts()
    if accounts:
        result = app_.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]
    return None

def _get_access_token(request: Optional[Request] = None):
    # 1) MSAL 캐시에서 사일런트
    tok = _acquire_token_silent()
    if tok:
        return tok
    # 2) 레거시 파일(최후 수단)
    try:
        with open("access_token.txt", "r", encoding="utf-8") as f:
            t = (f.read() or "").strip()
            return t or None
    except:
        return None

# -------------------------------
# 로그인 & 콜백
# -------------------------------
@app.get("/login")
def login(request: Request):
    # 더 이상 prompt="login" 강제하지 않음 (반복 로그인 방지)
    request.session["state"] = str(uuid.uuid4())
    app_, cache = _msal_app()
    auth_url = app_.get_authorization_request_url(
        scopes=SCOPES,
        state=request.session["state"],
        redirect_uri=REDIRECT_URI,
        response_mode="query",
    )
    return RedirectResponse(auth_url)

@app.get("/callback")
async def callback(request: Request):
    if request.query_params.get("state") != request.session.get("state"):
        return JSONResponse({"error": "state mismatch"}, status_code=400)
    code = request.query_params.get("code")
    if not code:
        return JSONResponse({"error": "Authorization code missing"}, status_code=400)

    app_, cache = _msal_app()
    result = app_.acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri=REDIRECT_URI)
    if "access_token" not in result:
        return JSONResponse({"error": "Token acquire failed", "details": result}, status_code=400)

    # 토큰 캐시 영구 저장
    _save_cache(cache)

    # 레거시 파일도 유지(호환)
    try:
        with open("refresh_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("refresh_token", "") or "")
    except: pass
    try:
        with open("access_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("access_token", "") or "")
    except: pass

    claims = result.get("id_token_claims", {}) or {}
    request.session.clear()
    request.session["user"] = {
        "name": claims.get("name"),
        "upn": claims.get("preferred_username"),
        "oid": claims.get("oid"),
    }
    return RedirectResponse("/me")

@app.get("/me")
def me(request: Request):
    # 사일런트 시도
    if _acquire_token_silent():
        user = request.session.get("user") or {}
        return JSONResponse({"status": "ok", "user": user})
    # 없으면 로그인 안내
    return RedirectResponse("/login")

# -------------------------------
# Static
# -------------------------------
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

# -------------------------------
# Graph Helper & Excel
# -------------------------------
GRAPH = "https://graph.microsoft.com/v1.0"
FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

_DRIVE_ITEM_ID_CACHE = {"name": None, "id": None}
def _get_drive_item_id(headers, file_name):
    if _DRIVE_ITEM_ID_CACHE["name"] == file_name and _DRIVE_ITEM_ID_CACHE["id"]:
        return _DRIVE_ITEM_ID_CACHE["id"]
    search = requests.get(
        f"{GRAPH}/me/drive/root/search(q='{file_name}')?$top=1", headers=headers
    ).json()
    items = search.get("value", [])
    if not items or items[0]["name"] != file_name:
        return None
    _DRIVE_ITEM_ID_CACHE["name"] = file_name
    _DRIVE_ITEM_ID_CACHE["id"] = items[0]["id"]
    return _DRIVE_ITEM_ID_CACHE["id"]

def write_row_to_onedrive(row):
    """
    워크북 세션(persistChanges=True) + 쓰기 후 즉시 읽어 검증.
    """
    token = _get_access_token()
    if not token:
        return False, {"error":"no_access_token"}

    base_headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    item_id = _get_drive_item_id(base_headers, FILE_NAME)
    if not item_id:
        return False, {"error":"file_not_found","file":FILE_NAME}

    meta = requests.get(
        f"{GRAPH}/me/drive/items/{item_id}?$select=webUrl,name,parentReference", headers=base_headers
    ).json()
    web_url = meta.get("webUrl")

    # 세션 생성
    sess = requests.post(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/createSession",
        headers=base_headers, json={"persistChanges": True}
    )
    if sess.status_code not in (200, 201):
        return False, {"error":"session_create_failed","status":sess.status_code,"text":sess.text}
    sid = sess.json().get("id")
    headers = {**base_headers, "workbook-session-id": sid}

    try:
        used = requests.get(
            f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/usedRange",
            headers=headers
        ).json()
        address = used.get("address") or f"{SHEET_NAME}!A1:A1"
        try:
            last_row = int(address.split("!")[1].split(":")[1][1:])
        except Exception:
            last_row = 1
        next_row = last_row + 1
        target = f"A{next_row}:G{next_row}"

        wr = requests.patch(
            f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
            headers=headers, json={"values":[row]}
        )
        if wr.status_code != 200:
            return False, {"error":"write_failed","status":wr.status_code,"text":wr.text,"range":target}

        rd = requests.get(
            f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
            headers=headers
        )
        if rd.status_code != 200:
            return False, {"error":"verify_read_failed","status":rd.status_code,"text":rd.text,"range":target}
        values = (rd.json() or {}).get("values") or []
        verified = bool(values and values[0][:len(row)] == row)
        if not verified:
            return False, {"error":"verify_mismatch","range":target,"read_back":values}

        return True, {"range":target, "file_webUrl": web_url, "sheet": SHEET_NAME}
    finally:
        pass

# -------------------------------
# OCR + Excel
# -------------------------------
@app.post("/process-ocr/")
async def process_ocr(
    qr_text: str = Form(""),
    image: UploadFile = File(...),
    dry: int = Form(0),          # 1=초고속 프리뷰, 0=정식 OCR
    no_write: int = Form(0)      # 1=정식 OCR 하되 "쓰기 없이"
):
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as f:
        shutil.copyfileobj(image.file, f)
    try:
        if dry:
            result = make_final_entry_fast(qr_text, temp_path)
            return {"status": "preview", "data": result}

        result = make_final_entry(qr_text, temp_path)
        if no_write:
            return {"status": "review", "data": result}

        row = [
            result.get("출고일", ""),
            result.get("대여자명", ""),
            result.get("전화번호", ""),
            result.get("주소", ""),
            result.get("기기번호", ""),
            result.get("기종", ""),
            result.get("송장번호", ""),
        ]
        ok, info = write_row_to_onedrive(row)
        if not ok:
            return {"status":"ocr_ok_but_write_failed","data":result,"write_error":info}
        return {"status":"success","data":result,"write_info":info}
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.post("/preview-ocr")
async def preview_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as f: shutil.copyfileobj(image.file, f)
    try:
        return {"status":"preview","data":make_final_entry_fast(qr_text,temp_path)}
    finally:
        if os.path.exists(temp_path): os.remove(temp_path)

# -------------------------------
# Save result (프론트 인식결과 그대로 저장)
# -------------------------------
@app.post("/save-result")
def save_result(data: Dict[str, Any] = Body(...)):
    def g(*keys, default=""):
        for k in keys:
            v = data.get(k)
            if v not in (None, ""):
                return v
        return default

    row = [
        g("출고일", "shipDate"),
        g("대여자명", "name"),
        g("전화번호", "phone"),
        g("주소", "addr"),
        g("기기번호", "deviceId"),
        g("기종", "model"),
        g("송장번호", "invoice"),
    ]
    ok, info = write_row_to_onedrive(row)
    if not ok:
        return JSONResponse({"status": "write_failed", "write_error": info, "row": row}, status_code=500)
    return {"status": "success", "write_info": info}

# -------------------------------
# Misc
# -------------------------------
@app.get("/__version")
def version(): return {"version": APP_VERSION}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))






