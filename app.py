# app.py  (붙여넣기 전체 코드)

import os
import shutil
import hashlib
import uuid
from uuid import uuid4
from typing import Optional, Dict, Any

from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File, Body
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse, PlainTextResponse
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware

import requests
import msal
from msal import SerializableTokenCache

from ocr_utils import make_final_entry, make_final_entry_fast

APP_VERSION = os.getenv("APP_VERSION", "2025-08-25-restore-flow-01")

# ================================
# FastAPI & Sessions
# ================================
app = FastAPI()

app.add_middleware(
    SessionMiddleware,
    secret_key=os.getenv("SESSION_SECRET", "change-me"),
    same_site="none",
    https_only=True,
    max_age=60 * 60 * 24 * 30,  # 30 days
    session_cookie="session",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ================================
# ENV & MS Graph
# ================================
CLIENT_ID = os.getenv("CLIENT_ID", "41745db3-a5c5-4e6e-acd7-fc4ce18b1999")
TENANT_ID = os.getenv("TENANT_ID", "405ba8a3-73ff-4423-8925-d9eda360cfa7")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://rent-label-api-client-docker.onrender.com/callback")

SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH = "https://graph.microsoft.com/v1.0"

CACHE_FILE = "msal_cache.bin"
ACCOUNT_FILE = "msal_account.json"  # 저장: {"home_account_id": "..."}

FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ================================
# MSAL helpers (1회 로그인 → 이후 자동)
# ================================
def _load_cache() -> SerializableTokenCache:
    cache = SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())
    return cache

def _save_cache(cache: SerializableTokenCache):
    if cache.has_state_changed:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            f.write(cache.serialize())

def _build_msal_app(cache: Optional[SerializableTokenCache] = None):
    if not CLIENT_SECRET:
        raise RuntimeError("CLIENT_SECRET env is missing.")
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )

def _read_saved_account_home_id() -> Optional[str]:
    path = os.path.join(BASE_DIR, ACCOUNT_FILE)
    if not os.path.exists(path):
        return None
    try:
        import json
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f).get("home_account_id")
    except:
        return None

def _write_saved_account_home_id(home_id: str):
    import json
    with open(os.path.join(BASE_DIR, ACCOUNT_FILE), "w", encoding="utf-8") as f:
        json.dump({"home_account_id": home_id}, f)

def _acquire_token_silent() -> Optional[str]:
    """
    최초 1회 로그인 후에는 이 함수가 캐시에서 토큰을 자동 갱신해줌.
    """
    cache = _load_cache()
    app_msal = _build_msal_app(cache)
    home_id = _read_saved_account_home_id()
    if not home_id:
        return None

    accounts = app_msal.get_accounts()
    account = next((a for a in accounts if a.get("home_account_id") == home_id), None)
    if not account:
        return None

    result = app_msal.acquire_token_silent(SCOPES, account=account)
    _save_cache(cache)
    if result and "access_token" in result:
        return result["access_token"]
    return None

def _get_access_token(request: Optional[Request] = None) -> Optional[str]:
    # 1) 세션에 있으면 사용
    if request is not None:
        tokens = request.session.get("tokens")
        if tokens and tokens.get("access_token"):
            return tokens["access_token"]
    # 2) 캐시에서 조용히 갱신(권장)
    tok = _acquire_token_silent()
    if tok:
        return tok
    # 3) 과거 호환(텍스트 파일에 저장해둔 토큰)
    try:
        with open("access_token.txt", "r", encoding="utf-8") as f:
            return f.read().strip() or None
    except:
        return None

# ================================
# Auth endpoints
# ================================
@app.get("/login")
def login(request: Request):
    # 매번 강제 로그인 제거 (prompt 기본값)
    state = str(uuid.uuid4())
    request.session["state"] = state
    nonce = str(uuid.uuid4())

    cache = _load_cache()
    app_msal = _build_msal_app(cache)
    auth_url = app_msal.get_authorization_request_url(
        scopes=SCOPES,
        state=state,
        redirect_uri=REDIRECT_URI,
        response_mode="query",
        # prompt 생략 → 이미 로그인된 계정이면 브라우저가 자동 통과(테넌트 정책에 따름)
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

    cache = _load_cache()
    app_msal = _build_msal_app(cache)
    result = app_msal.acquire_token_by_authorization_code(
        code, scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    _save_cache(cache)

    if "access_token" not in result:
        return JSONResponse({"error": "Token acquire failed", "details": result}, status_code=400)

    # 계정 식별자 저장(한 번만 로그인하면 이후 silent)
    account = app_msal.get_accounts()[0] if app_msal.get_accounts() else None
    if account and account.get("home_account_id"):
        _write_saved_account_home_id(account["home_account_id"])

    # 구형 호환: 텍스트 파일에도 저장(필요시)
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
    # 캐시에 유효 계정이 있으면 로그인 없이도 OK로 응답
    user = request.session.get("user")
    token = _get_access_token(request)
    if token:
        return JSONResponse({"status": "ok", "user": user or {"note": "silent-login-active"}})
    # 토큰이 전혀 없으면 로그인 필요
    return RedirectResponse("/login")

# 디버그/도우미
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
        "cache_exists": os.path.exists(CACHE_FILE),
        "account_saved": bool(_read_saved_account_home_id()),
    }

@app.get("/login-url", response_class=PlainTextResponse)
def login_url():
    cache = _load_cache()
    return _build_msal_app(cache).get_authorization_request_url(
        scopes=SCOPES, state="debug", redirect_uri=REDIRECT_URI, response_mode="query"
    )

@app.get("/whoami")
def whoami(request: Request):
    token = _get_access_token(request)
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    headers = {"Authorization": f"Bearer {token}"}
    try:
        me = requests.get(f"{GRAPH}/me", headers=headers).json()
        org = requests.get(f"{GRAPH}/organization", headers=headers).json()
        return {"me": me, "organization": org}
    except Exception as e:
        return {"error": str(e)}

# ================================
# Static
# ================================
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

# URL 변형 대응
@app.get("/callback/")
async def callback_slash(request: Request): return await callback(request)
@app.get("/login/callback")
async def callback_login(request: Request): return await callback(request)
@app.get("/login/callback/")
async def callback_login2(request: Request): return await callback(request)

# ================================
# Graph helpers
# ================================
def _get_drive_item_id(headers, file_name):
    # 간단 캐시
    if not hasattr(_get_drive_item_id, "_cache"):
        _get_drive_item_id._cache = {}
    if file_name in _get_drive_item_id._cache:
        return _get_drive_item_id._cache[file_name]

    search = requests.get(
        f"{GRAPH}/me/drive/root/search(q='{file_name}')?$top=1", headers=headers
    ).json()
    items = search.get("value", [])
    if not items or items[0]["name"] != file_name:
        return None
    _get_drive_item_id._cache[file_name] = items[0]["id"]
    return items[0]["id"]

def write_row_to_onedrive(row):
    """
    - 워크북 세션을 열고(write-through) 그 세션으로만 작업
    - 쓰기 후 즉시 읽어 검증까지 하고 결과(파일 링크, 시트, 범위) 반환
    """
    token = _get_access_token()
    if not token:
        return False, {"error": "no_access_token"}

    base_headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # 1) 대상 파일 찾기 (+ 링크)
    item_id = _get_drive_item_id(base_headers, FILE_NAME)
    if not item_id:
        return False, {"error": "file_not_found", "file": FILE_NAME}

    meta = requests.get(
        f"{GRAPH}/me/drive/items/{item_id}?$select=webUrl,name,parentReference",
        headers=base_headers
    ).json()
    web_url = meta.get("webUrl")

    # 2) 워크북 세션 생성 (persistChanges=True)
    sess = requests.post(
        f"{GRAPH}/me/drive/items/{item_id}/workbook/createSession",
        headers=base_headers, json={"persistChanges": True}
    )
    if sess.status_code not in (200, 201):
        return False, {"error": "session_create_failed", "status": sess.status_code, "text": sess.text}
    sid = sess.json().get("id")
    headers = {**base_headers, "workbook-session-id": sid}

    try:
        # 3) 다음 행 계산
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

        # 4) 쓰기
        wr = requests.patch(
            f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
            headers=headers, json={"values": [row]}
        )
        if wr.status_code != 200:
            return False, {"error": "write_failed", "status": wr.status_code, "text": wr.text, "range": target}

        # 5) 검증(바로 읽어 확인)
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
        # 세션 종료는 옵션(미종료해도 서버가 정리). 명시 종료 원하면 주석 해제:
        # requests.post(f"{GRAPH}/me/drive/items/{item_id}/workbook/closeSession",
        #               headers=headers)
        pass

# ================================
# OCR endpoints
# ================================
@app.post("/preview-ocr")
async def preview_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    """
    촬영 직후 빠른 미리보기 → 프론트 '인식결과' 즉시 표시용.
    """
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as f:
        shutil.copyfileobj(image.file, f)
    try:
        data = make_final_entry_fast(qr_text, temp_path)
        return {"status": "preview", "data": data}
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

@app.post("/save-result")
def save_result(data: Dict[str, Any] = Body(...)):
    """
    ✅ 재-OCR 하지 않음. 프론트가 미리보기 결과(data)를 그대로 보내면 엑셀 저장.
    기대 key: 출고일, 대여자명, 전화번호, 주소, 기기번호, 기종, 송장번호
    """
    row = [
        data.get("출고일", ""),
        data.get("대여자명", ""),
        data.get("전화번호", ""),
        data.get("주소", ""),
        data.get("기기번호", ""),
        data.get("기종", ""),
        data.get("송장번호", ""),
    ]
    ok, info = write_row_to_onedrive(row)
    if not ok:
        return JSONResponse({"status": "write_failed", "data": data, "write_error": info}, status_code=500)
    return {"status": "success", "data": data, "write_info": info}

# (호환용) 예전 엔드포인트 유지
@app.post("/process-ocr/")
async def process_ocr(
    qr_text: str = Form(""),
    image: UploadFile = File(...),
    dry: int = Form(0),     # 1 = 빠른 미리보기
    no_write: int = Form(0) # 1 = 결과만
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
            return {"status": "ocr_ok_but_write_failed", "data": result, "write_error": info}
        return {"status": "success", "data": result, "write_info": info}
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

# ================================
# Excel direct append (유지)
# ================================
@app.post("/excel/append")
def excel_append(row: list = Body(...)):
    ok, info = write_row_to_onedrive(row)
    if not ok:
        return JSONResponse({"error": info}, status_code=500)
    return {"status": "ok", "range": info.get("range"), "written": row}

# ================================
# Misc
# ================================
@app.get("/onedrive")
def onedrive_list(request: Request):
    token = _get_access_token(request)
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    return requests.get(f"{GRAPH}/me/drive/root/children",
                        headers={"Authorization": f"Bearer {token}"}).json()

@app.get("/graph/me")
def graph_me(request: Request):
    token = _get_access_token(request)
    if not token:
        return JSONResponse({"error": "no_access_token"}, status_code=401)
    return requests.get(f"{GRAPH}/me", headers={"Authorization": f"Bearer {token}"}).json()

@app.get("/__version")
def version(): return {"version": APP_VERSION}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))





