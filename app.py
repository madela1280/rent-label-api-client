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
_HTTP = requests.Session()
import msal
from uuid import uuid4
from typing import Optional, Dict, Any

from ocr_utils import make_final_entry, make_final_entry_fast

APP_VERSION = os.getenv("APP_VERSION", "2025-08-27-09")

# -------------------------------
# FastAPI & Session
# -------------------------------
app = FastAPI()

app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET", "change-me"),
    same_site="lax",
    https_only=True,
    max_age=60*60*24*30,  # 30일
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
CLIENT_ID = os.getenv("CLIENT_ID", "")
TENANT_ID = os.getenv("TENANT_ID", "")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://rent-label-api-client-docker.onrender.com/callback")

SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH = "https://graph.microsoft.com/v1.0"

# -------------------------------
# MSAL App
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
    # 바로 홈으로 (무한루프/체감 혼란 방지)
    return RedirectResponse("/")

@app.get("/me")
def me(request: Request):
    user = request.session.get("user")
    if not user:
        return RedirectResponse("/login")
    return JSONResponse({"status": "ok", "user": user})

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
# Graph Helper
# -------------------------------
def _get_access_token(request: Optional[Request] = None):
    if request is not None:
        tok = (request.session.get("tokens") or {}).get("access_token")
        if tok: return tok
    return None

# -------------------------------
# Excel Append
# -------------------------------
FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

_DRIVE_ITEM_ID_CACHE = {"name": None, "id": None}
def _get_drive_item_id(headers, file_name):
    if _DRIVE_ITEM_ID_CACHE["name"] == file_name and _DRIVE_ITEM_ID_CACHE["id"]:
        return _DRIVE_ITEM_ID_CACHE["id"]
    search = _HTTP.get(
        f"{GRAPH}/me/drive/root/search(q='{file_name}')?$top=1", headers=headers
    ).json()
    items = search.get("value", [])
    if not items or items[0]["name"] != file_name:
        return None
    _DRIVE_ITEM_ID_CACHE["name"] = file_name
    _DRIVE_ITEM_ID_CACHE["id"] = items[0]["id"]
    return _DRIVE_ITEM_ID_CACHE["id"]

def write_row_to_onedrive(row, request: Optional[Request] = None):
    token = _get_access_token(request)
    if not token: return False, {"error":"no_access_token"}
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    item_id = _get_drive_item_id(headers, FILE_NAME)
    if not item_id:
        return False, {"error":"file_not_found","file":FILE_NAME}

    used = _HTTP.get(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/usedRange", headers=headers
    ).json()
    address = used.get("address") or f"{SHEET_NAME}!A1:A1"
    try: last_row = int(address.split("!")[1].split(":")[1][1:])
    except: last_row = 1
    next_row = last_row + 1
    target = f"A{next_row}:F{next_row}"  # 6개 컬럼

    resp = _HTTP.patch(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
        headers=headers, json={"values":[row]}
    )
    if resp.status_code != 200:
        return False, {"error":"write_failed","status":resp.status_code,"text":resp.text}
    return True, {"range":target}

# -------------------------------
# OCR + Excel (강력 예외 처리)
# -------------------------------
@app.post("/process-ocr/")
async def process_ocr(
    request: Request,
    qr_text: str = Form(""),
    image: UploadFile = File(...),
    dry: int = Form(0),
    no_write: int = Form(0)
):
    temp_path = f"temp_{image.filename}"
    try:
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(image.file, f)
    except Exception as e:
        return JSONResponse({"status":"error","step":"save_upload","error":str(e)}, status_code=500)

    try:
        try:
            result = make_final_entry_fast(qr_text, temp_path) if dry else make_final_entry(qr_text, temp_path)
        except Exception as e:
            return JSONResponse({"status":"error","step":"ocr","error":str(e)}, status_code=500)

        if no_write:
            return {"status": "review", "data": result}

        row = [
            result.get("출고일", ""),
            result.get("대여자명", ""),
            result.get("전화번호", ""),
            result.get("주소", ""),
            result.get("기기번호", ""),
            result.get("기종", ""),
        ]
        ok, info = write_row_to_onedrive(row, request)
        if not ok:
            code = 401 if info.get("error")=="no_access_token" else 500
            return JSONResponse({"status":"error","step":"onedrive","detail":info,"data":result}, status_code=code)

        return {"status": "success", "data": result, "write_info": info}
    finally:
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except:
            pass

@app.post("/preview-ocr")
async def preview_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    temp_path = f"temp_{image.filename}"
    try:
        with open(temp_path, "wb") as f: shutil.copyfileobj(image.file, f)
    except Exception as e:
        return JSONResponse({"status":"error","step":"save_upload","error":str(e)}, status_code=500)

    try:
        try:
            data = make_final_entry_fast(qr_text, temp_path)
        except Exception as e:
            return JSONResponse({"status":"error","step":"ocr_preview","error":str(e)}, status_code=500)
        return {"status":"preview","data":data}
    finally:
        try:
            if os.path.exists(temp_path): os.remove(temp_path)
        except:
            pass

# -------------------------------
# 수동 저장
# -------------------------------
@app.post("/save-result")
def save_result(request: Request, data: Dict[str, Any] = Body(...)):
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
    ]
    ok, info = write_row_to_onedrive(row, request)
    if not ok:
        status = 401 if info.get("error") == "no_access_token" else 500
        return JSONResponse({"status": "error", "step":"onedrive", "write_error": info, "row": row}, status_code=status)
    return {"status": "success", "write_info": info}

# -------------------------------
# Misc
# -------------------------------
@app.get("/__version")
def version(): return {"version": APP_VERSION}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))













