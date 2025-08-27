import os
import shutil
import hashlib
import urllib.parse
import uuid

from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File, Body
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse, PlainTextResponse
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware

import requests
_HTTP = requests.Session()  # 성능개선: 세션(Keep-Alive)
import msal
from uuid import uuid4
from typing import Optional, Dict, Any

from ocr_utils import make_final_entry, make_final_entry_fast

APP_VERSION = os.getenv("APP_VERSION", "2025-08-26-rt1")

# -------------------------------
# FastAPI & Session
# -------------------------------
app = FastAPI()

# 세션 30일 유지 (모바일 브라우저 재방문 시 유지)
app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET", "change-me"),
    same_site="lax",
    https_only=True,
    max_age=60*60*24*30,   # 30 days
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
# ENV & Constants (기본 설정)
# -------------------------------
CLIENT_ID = os.getenv("CLIENT_ID", "41745db3-a5c5-4e6e-acd7-fc4ce18b1999")
TENANT_ID = os.getenv("TENANT_ID", "405ba8a3-73ff-4423-8925-d9eda360cfa7")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://rent-label-api-client-docker.onrender.com/callback")

SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH = "https://graph.microsoft.com/v1.0"

TOKEN_DIR = os.getenv("TOKEN_DIR", ".")
ACCESS_PATH = os.path.join(TOKEN_DIR, "access_token.txt")
REFRESH_PATH = os.path.join(TOKEN_DIR, "refresh_token.txt")

os.makedirs(TOKEN_DIR, exist_ok=True)

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

def _save_token_files(access_token: str, refresh_token: str | None):
    try:
        if access_token:
            with open(ACCESS_PATH, "w", encoding="utf-8") as f:
                f.write(access_token)
    except:
        pass
    try:
        if refresh_token:
            with open(REFRESH_PATH, "w", encoding="utf-8") as f:
                f.write(refresh_token)
    except:
        pass

def _read_file(path: str) -> Optional[str]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read().strip() or None
    except:
        return None

def _refresh_access_token() -> Optional[str]:
    """refresh_token으로 access_token 재발급"""
    rt = _read_file(REFRESH_PATH)
    if not rt:
        return None
    try:
        result = _build_msal_app().acquire_token_by_refresh_token(rt, scopes=SCOPES)
        if "access_token" in result:
            _save_token_files(result.get("access_token", ""), result.get("refresh_token"))
            return result.get("access_token")
        return None
    except Exception:
        return None

def _get_access_token(request: Optional[Request] = None) -> Optional[str]:
    """
    1) 세션에 access_token 있으면 사용
    2) 파일(access_token.txt) 읽어 사용
    3) 없거나 만료시 refresh_token으로 재발급
    """
    if request is not None:
        tok = (request.session.get("tokens") or {}).get("access_token")
        if tok:
            return tok

    tok = _read_file(ACCESS_PATH)
    if tok:
        return tok

    # 파일에 access가 없으면 refresh 시도
    new_tok = _refresh_access_token()
    return new_tok

def _graph_get(url: str, token: Optional[str]):
    """401/403 나오면 한 번만 새 토큰으로 재시도"""
    if not token:
        return None, 401
    h = {"Authorization": f"Bearer {token}"}
    resp = _HTTP.get(url, headers=h)
    if resp.status_code in (401, 403):
        # refresh 후 재시도
        new_tok = _refresh_access_token()
        if not new_tok:
            return resp, resp.status_code
        h2 = {"Authorization": f"Bearer {new_tok}"}
        resp2 = _HTTP.get(url, headers=h2)
        return resp2, resp2.status_code
    return resp, resp.status_code

def _graph_patch(url: str, token: Optional[str], json_body: dict):
    if not token:
        return None, 401
    h = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = _HTTP.patch(url, headers=h, json=json_body)
    if resp.status_code in (401, 403):
        new_tok = _refresh_access_token()
        if not new_tok:
            return resp, resp.status_code
        h2 = {"Authorization": f"Bearer {new_tok}", "Content-Type": "application/json"}
        resp2 = _HTTP.patch(url, headers=h2, json=json_body)
        return resp2, resp2.status_code
    return resp, resp.status_code

# -------------------------------
# 로그인 & 콜백
# -------------------------------
@app.get("/login")
def login(request: Request):
    request.session["state"] = str(uuid.uuid4())
    nonce = str(uuid.uuid4())
    auth_url = _build_msal_app().get_authorization_request_url(
        scopes=SCOPES,
        state=request.session["state"],
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
        response_mode="query",
    )
    sep = "&" if "?" in auth_url else "?"
    return RedirectResponse(f"{auth_url}{sep}nonce={nonce}")

@app.get("/callback")
async def callback(request: Request):
    if request.query_params.get("state") != request.session.get("state"):
        return JSONResponse({"error": "state mismatch"}, status_code=400)
    code = request.query_params.get("code")
    if not code:
        return JSONResponse({"error": "Authorization code missing"}, status_code=400)

    result = _build_msal_app().acquire_token_by_authorization_code(
        code, scopes=SCOPES, redirect_uri=REDIRECT_URI,
    )
    if "access_token" not in result:
        return JSONResponse({"error": "Token acquire failed", "details": result}, status_code=400)

    # 파일 + 세션 모두 저장 (재방문 시 파일 사용)
    _save_token_files(result.get("access_token", ""), result.get("refresh_token"))
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

@app.get("/me")
def me(request: Request):
    user = request.session.get("user")
    if not user:
        return RedirectResponse("/login")
    return JSONResponse({"status": "ok", "user": user})

# -------------------------------
# Debug / Login 상태 점검
# -------------------------------
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
        "secret_fp": hashlib.sha256(sec.encode()).hexdigest()[:12],
        "has_access_file": bool(_read_file(ACCESS_PATH)),
        "has_refresh_file": bool(_read_file(REFRESH_PATH)),
    }

@app.get("/__login_status")
def login_status():
    tok = _get_access_token()
    if not tok:
        return {"logged_in": False}
    # /me 호출로 유효성 확인(401이면 내부적으로 refresh 후 재시도)
    resp, status = _graph_get(f"{GRAPH}/me", tok)
    return {"logged_in": status == 200, "status": status}

@app.get("/login-url", response_class=PlainTextResponse)
def login_url():
    return _build_msal_app().get_authorization_request_url(
        scopes=SCOPES, state="debug", redirect_uri=REDIRECT_URI,
        prompt="select_account", response_mode="query"
    )

@app.get("/whoami")
def whoami(request: Request):
    token = _get_access_token(request)
    resp, status = _graph_get(f"{GRAPH}/me", token)
    if status != 200:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    me = resp.json()
    org_resp, _ = _graph_get(f"{GRAPH}/organization", _get_access_token())
    org = org_resp.json() if org_resp is not None else {}
    return {"me": me, "organization": org}

# -------------------------------
# Static files
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

# 콜백 경로 분기
@app.get("/callback/")
async def callback_slash(request: Request):
    return await callback(request)

@app.get("/login/callback")
async def callback_login(request: Request):
    return await callback(request)

@app.get("/login/callback/")
async def callback_login2(request: Request):
    return await callback(request)

# -------------------------------
# OneDrive / Excel Helper
# -------------------------------
FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

_DRIVE_ITEM_ID_CACHE = {"name": None, "id": None}
def _get_drive_item_id(headers, file_name):
    if _DRIVE_ITEM_ID_CACHE["name"] == file_name and _DRIVE_ITEM_ID_CACHE["id"]:
        return _DRIVE_ITEM_ID_CACHE["id"]
    search = _HTTP.get(
        f"{GRAPH}/me/drive/root/search(q='{file_name}')?$top=1", headers=headers
    )
    if search.status_code in (401,403):
        # refresh + 재시도
        new_tok = _refresh_access_token()
        if not new_tok:
            return None
        headers = {"Authorization": f"Bearer {new_tok}"}
        search = _HTTP.get(
            f"{GRAPH}/me/drive/root/search(q='{file_name}')?$top=1", headers=headers
        )

    items = search.json().get("value", [])
    if not items or items[0]["name"] != file_name:
        return None
    _DRIVE_ITEM_ID_CACHE["name"] = file_name
    _DRIVE_ITEM_ID_CACHE["id"] = items[0]["id"]
    return _DRIVE_ITEM_ID_CACHE["id"]

def write_row_to_onedrive(row):
    """401/403 나오면 자동 갱신 후 1회 재시도"""
    token = _get_access_token()
    if not token:
        return False, {"error":"no_access_token"}

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    item_id = _get_drive_item_id(headers, FILE_NAME)
    if not item_id:
        return False, {"error":"file_not_found","file":FILE_NAME}

    # usedRange
    used_resp, status = _graph_get(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/usedRange",
        token
    )
    if status != 200:
        return False, {"error":"used_range_failed","status":status,"text":getattr(used_resp, "text", "")}
    used = used_resp.json()

    address = used.get("address") or f"{SHEET_NAME}!A1:A1"
    try:
        last_row = int(address.split("!")[1].split(":")[1][1:])
    except Exception:
        last_row = 1

    next_row = last_row + 1
    target = f"A{next_row}:G{next_row}"

    # write
    patch_url = f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')"
    resp, status = _graph_patch(patch_url, token, {"values":[row]})
    if status != 200:
        return False, {"error":"write_failed","status":status,"text":getattr(resp, "text", "")}
    return True, {"range":target}

# -------------------------------
# API: Excel Append / OCR
# -------------------------------
@app.post("/excel/append")
def excel_append(row: list = Body(...)):
    ok, info = write_row_to_onedrive(row)
    if not ok:
        if info.get("error") == "no_access_token":
            return JSONResponse(info, status_code=401)
        return JSONResponse(info, status_code=500)
    return {"status": "ok", **info}

@app.post("/process-ocr/")
async def process_ocr(
    qr_text: str = Form(""),
    image: UploadFile = File(...),
    dry: int = Form(0),          # 1=미리보기(고속 처리), 0=정식 OCR
    no_write: int = Form(0)      # 1=정식 OCR 되지만 "기록 없이" 결과만 반환(후속 검증용)
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

        # 엑셀 기록
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
            # 토큰 없음이면 401로 명확히 반환(프론트가 /login 유도)
            if info.get("error") == "no_access_token":
                return JSONResponse({"status":"ocr_ok_but_write_failed","data":result,"write_error":info}, status_code=401)
            return {"status": "ocr_ok_but_write_failed", "data": result, "write_error": info}
        return {"status": "success", "data": result, "write_info": info}
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
# Save result (수동 입력 저장)
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
        g("출고일", "shipDate", "출고일자"),
        g("대여자명", "name", "받는분"),
        g("전화번호", "phone"),
        g("주소", "addr"),
        g("기기번호", "deviceId"),
        g("기종", "model"),
        g("송장번호", "invoice", "운송장번호"),
    ]
    ok, info = write_row_to_onedrive(row)
    if not ok:
        if info.get("error") == "no_access_token":
            return JSONResponse({"status": "write_failed", "write_error": info, "row": row}, status_code=401)
        return JSONResponse({"status": "write_failed", "write_error": info, "row": row}, status_code=500)
    return {"status": "success", "write_info": info, "row": row}

# -------------------------------
# Misc
# -------------------------------
@app.get("/__version")
def version(): return {"version": APP_VERSION}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))









