import os
import shutil
import hashlib
import urllib.parse
import uuid

from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File, Body
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse, PlainTextResponse
from pydantic import BaseModel
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware

import requests
import msal
from uuid import uuid4
from typing import Optional, Dict, Any

from ocr_utils import make_final_entry, make_final_entry_fast

APP_VERSION = os.getenv("APP_VERSION", "2025-08-25-06")

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

SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
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
        prompt="login",
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

    try:
        with open("refresh_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("refresh_token", ""))
    except: pass
    try:
        with open("access_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("access_token", ""))
    except: pass

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
# Debug endpoints
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
    }

@app.get("/login-url", response_class=PlainTextResponse)
def login_url():
    return _build_msal_app().get_authorization_request_url(
        scopes=SCOPES, state="debug", redirect_uri=REDIRECT_URI,
        prompt="select_account", response_mode="query"
    )

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

# 콜백 경로 변형 대응
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
SCOPES_GRAPH = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]

def _get_access_token(request: Optional[Request] = None):
    """세션 > 파일(access_token.txt) 순으로 토큰 조회"""
    if request is not None:
        tok = (request.session.get("tokens") or {}).get("access_token")
        if tok: return tok
    try:
        with open("access_token.txt", "r", encoding="utf-8") as f:
            return f.read().strip() or None
    except:
        return None

@app.get("/graph/me")
def graph_me(request: Request):
    token = _get_access_token(request)
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    return requests.get(f"{GRAPH}/me", headers={"Authorization": f"Bearer {token}"}).json()

@app.get("/onedrive")
def onedrive(request: Request):
    token = _get_access_token(request)
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    return requests.get("https://graph.microsoft.com/v1.0/me/drive/root/children",
                        headers={"Authorization": f"Bearer {token}"}).json()

# -------------------------------
# Excel Helper
# -------------------------------
FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

# 전역 캐시
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
    Workbook 세션을 열고(persistChanges=True) 쓰기 → 즉시 읽기 검증까지 수행.
    """
    token = _get_access_token()
    if not token:
        return False, {"error": "no_access_token"}

    base_headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # 파일 찾기
    item_id = _get_drive_item_id(base_headers, FILE_NAME)
    if not item_id:
        return False, {"error": "file_not_found", "file": FILE_NAME}

    meta = requests.get(
        f"{GRAPH}/me/drive/items/{item_id}?$select=webUrl,name,parentReference",
        headers=base_headers
    ).json()
    web_url = meta.get("webUrl")

    # 세션 생성
    sess = requests.post(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/createSession",
        headers=base_headers, json={"persistChanges": True}
    )
    if sess.status_code not in (200, 201):
        return False, {"error": "session_create_failed", "status": sess.status_code, "text": sess.text}
    sid = sess.json().get("id")
    headers = {**base_headers, "workbook-session-id": sid}

    try:
        # 다음 행 계산
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

        # 쓰기
        wr = requests.patch(
            f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
            headers=headers, json={"values": [row]}
        )
        if wr.status_code != 200:
            return False, {"error": "write_failed", "status": wr.status_code, "text": wr.text, "range": target}

        # 검증
        rd = requests.get(
            f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
            headers=headers
        )
        if rd.status_code != 200:
            return False, {"error": "verify_read_failed", "status": rd.status_code, "text": rd.text, "range": target}
        values = (rd.json() or {}).get("values") or []
        verified = bool(values and values[0][:len(row)] == row)
        if not verified:
            return False, {"error": "verify_mismatch", "range": target, "read_back": values}

        return True, {"range": target, "file_webUrl": web_url, "sheet": SHEET_NAME}
    finally:
        # 세션 종료는 선택 사항
        pass

# -------------------------------
# Excel Append (남겨둠: 호환용)
# -------------------------------
@app.post("/excel/append")
def excel_append(row: list = Body(...)):
    ok, info = write_row_to_onedrive(row)
    if not ok:
        return JSONResponse({"error": "write_failed", "details": info}, status_code=500)
    return {"status": "ok", "range": info.get("range"), "written": row}

# -------------------------------
# OCR + Excel
# -------------------------------
@app.post("/process-ocr/")
async def process_ocr(
    qr_text: str = Form(""),
    image: UploadFile = File(...),
    dry: int = Form(0),          # 1=초고속 프리뷰, 0=정식 OCR
    no_write: int = Form(0)      # 1=정식 OCR 하되 "쓰기 없이" 결과만 반환
):
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as f:
        shutil.copyfileobj(image.file, f)
    try:
        if dry:
            # 빠른 미리보기
            result = make_final_entry_fast(qr_text, temp_path)
            return {"status": "preview", "data": result}

        # 정식 OCR
        result = make_final_entry(qr_text, temp_path)

        if no_write:
            return {"status": "review", "data": result}

        # 실제 쓰기
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
# OneDrive Save Result (신규)
# -------------------------------
@app.post("/save-result")
def save_result(data: Dict[str, Any] = Body(...)):
    """
    프론트에서 보내온 인식결과를 그대로 엑셀에 저장
    - 영문/한글 키 모두 허용
    - 저장 검증까지 수행
    """
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





