import os, shutil, uuid, time, requests
from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File, Body
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware

import msal
from typing import Optional, Dict, Any
from uuid import uuid4
from ocr_utils import make_final_entry, make_final_entry_fast

_HTTP = requests.Session()
APP_VERSION = os.getenv("APP_VERSION", "2025-08-27-stable")

# ------------------------------- FastAPI & Session -------------------------------
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
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"],
)

# ------------------------------- ENV / Graph -------------------------------
CLIENT_ID     = os.getenv("CLIENT_ID")
TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SCOPES    = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All", "offline_access", "openid", "profile"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH     = "https://graph.microsoft.com/v1.0"

FILE_NAME  = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")
TABLE_NAME = os.getenv("TABLE_NAME", "출고내역")  # 있으면 rows/add, 없으면 range 패치

# ------------------------------- MSAL helpers -------------------------------
def _build_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

def _save_tokens(result:dict):
    open("refresh_token.txt","w",encoding="utf-8").write(result.get("refresh_token","") or "")
    open("access_token.txt","w",encoding="utf-8").write(result.get("access_token","") or "")

def _load_refresh_token()->Optional[str]:
    p = "refresh_token.txt"
    return open(p,"r",encoding="utf-8").read().strip() if os.path.exists(p) else None

def _load_access_token()->Optional[str]:
    p = "access_token.txt"
    return open(p,"r",encoding="utf-8").read().strip() if os.path.exists(p) else None

def _get_access_token(force_refresh: bool=False)->Optional[str]:
    rtok = _load_refresh_token()
    if force_refresh and rtok:
        r = _build_msal_app().acquire_token_by_refresh_token(rtok, scopes=SCOPES)
        if "access_token" in r:
            _save_tokens(r)
            return r["access_token"]
    if rtok:
        r = _build_msal_app().acquire_token_by_refresh_token(rtok, scopes=SCOPES)
        if "access_token" in r:
            _save_tokens(r)
            return r["access_token"]
    return _load_access_token()

# ------------------------------- Graph helpers -------------------------------
_DRIVE_ITEM_ID: Optional[str] = None
_SESSION_ID: Optional[str] = None

def _headers(token: str, with_session: bool=True) -> Dict[str,str]:
    h = {"Authorization": f"Bearer {token}", "Content-Type":"application/json"}
    if with_session and _SESSION_ID:
        h["workbook-session-id"] = _SESSION_ID
    return h

def _get_drive_item_id(token: str) -> str:
    global _DRIVE_ITEM_ID
    if _DRIVE_ITEM_ID:
        return _DRIVE_ITEM_ID
    r = _HTTP.get(f"{GRAPH}/me/drive/root/search(q='{FILE_NAME}')?$top=1",
                  headers=_headers(token, with_session=False), timeout=30)
    r.raise_for_status()
    items = r.json().get("value", [])
    if not items or items[0]["name"] != FILE_NAME:
        raise RuntimeError(f"file_not_found:{FILE_NAME}")
    _DRIVE_ITEM_ID = items[0]["id"]
    return _DRIVE_ITEM_ID

def _ensure_session(token: str) -> str:
    global _SESSION_ID
    if _SESSION_ID:
        return _SESSION_ID
    url = f"{GRAPH}/me/drive/items/{_get_drive_item_id(token)}/workbook/createSession"
    r = _HTTP.post(url, headers=_headers(token, with_session=False), json={"persist": True}, timeout=30)
    r.raise_for_status()
    _SESSION_ID = r.json().get("id")
    return _SESSION_ID

def _write_row_table(token: str, row: list) -> requests.Response:
    _ensure_session(token)
    url = f"{GRAPH}/me/drive/items/{_get_drive_item_id(token)}/workbook/worksheets('{SHEET_NAME}')/tables('{TABLE_NAME}')/rows/add"
    return _HTTP.post(url, headers=_headers(token), json={"values":[row]}, timeout=60)

def _write_row_range(token: str, row: list) -> requests.Response:
    _ensure_session(token)
    # usedRange로 마지막 행 계산 → 다음 행 A..F 패치
    used = _HTTP.get(
        f"{GRAPH}/me/drive/items/{_get_drive_item_id(token)}/workbook/worksheets('{SHEET_NAME}')/usedRange",
        headers=_headers(token), timeout=30
    )
    used.raise_for_status()
    addr = used.json().get("address") or f"{SHEET_NAME}!A1:A1"
    try:
        last_row = int(addr.split("!")[1].split(":")[1][1:])
    except Exception:
        last_row = 1
    next_row = last_row + 1
    target = f"A{next_row}:F{next_row}"
    url = f"{GRAPH}/me/drive/items/{_get_drive_item_id(token)}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')"
    return _HTTP.patch(url, headers=_headers(token), json={"values":[row]}, timeout=60)

def write_row_to_onedrive(row: list, token: Optional[str]=None) -> (bool, Dict[str,Any]):
    """
    1) 테이블 '출고내역' rows/add 시도
    2) 없으면 usedRange 기반 A..F 패치로 폴백
    반환: (ok, info|error)
    """
    tok = token or _get_access_token()
    if not tok:
        return False, {"error":"no_access_token"}

    # 테이블 우선
    try:
        r = _write_row_table(tok, row)
        if r.status_code in (200, 201):
            return True, {"mode":"table","status":r.status_code}
        if r.status_code in (404, 400):
            # 테이블 없거나 잘못됨 → 폴백
            r2 = _write_row_range(tok, row)
            if r2.status_code == 200:
                return True, {"mode":"range","status":200}
            return False, {"error":"write_failed","status":r2.status_code,"text":r2.text}
        if r.status_code == 401:
            return False, {"error":"unauthorized","status":401,"text":r.text}
        return False, {"error":"write_failed","status":r.status_code,"text":r.text}
    except requests.HTTPError as e:
        st = getattr(e.response, "status_code", 0) if e.response else 0
        if st == 401:
            return False, {"error":"unauthorized","status":401,"text":str(e)}
        return False, {"error":"exception","text":str(e)}
    except Exception as e:
        return False, {"error":"exception","text":str(e)}

# ------------------------------- Auth Routes -------------------------------
def _redirect_uri(request: Request) -> str:
    # 현재 요청 기준으로 콜백 URI 생성 (http→https 강제)
    uri = str(request.url_for("callback"))
    if uri.startswith("http://"):
        uri = "https://" + uri[len("http://"):]
    return uri

@app.get("/login")
def login(request: Request):
    request.session["state"] = str(uuid.uuid4())
    auth_url = _build_msal_app().get_authorization_request_url(
        scopes=SCOPES, state=request.session["state"],
        redirect_uri=_redirect_uri(request), prompt="select_account", response_mode="query",
    )
    return RedirectResponse(auth_url)

@app.get("/callback")
async def callback(request: Request):
    if request.query_params.get("state") != request.session.get("state"):
        return JSONResponse({"error":"state mismatch"}, status_code=400)
    code = request.query_params.get("code")
    if not code:
        return JSONResponse({"error":"Authorization code missing"}, status_code=400)

    result = _build_msal_app().acquire_token_by_authorization_code(
        code, scopes=SCOPES, redirect_uri=_redirect_uri(request)
    )
    if "access_token" not in result:
        return JSONResponse({"error":"Token acquire failed","details":result}, status_code=400)

    _save_tokens(result)
    request.session["tokens"] = {"access_token": result["access_token"]}
    return RedirectResponse("/")

# ------------------------------- OCR Routes -------------------------------
@app.post("/preview-ocr")
async def preview_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    p = f"temp_{image.filename}"
    with open(p,"wb") as f: shutil.copyfileobj(image.file,f)
    try:
        result = make_final_entry_fast(qr_text, p)
        return {"status":"preview","data":result}
    finally:
        if os.path.exists(p): os.remove(p)

@app.post("/process-ocr/")
async def process_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    p = f"temp_{image.filename}"
    with open(p,"wb") as f: shutil.copyfileobj(image.file,f)
    try:
        result = make_final_entry(qr_text, p)
        # 저장은 여기서 하지 않음(확인 버튼에서 저장)
        return {"status":"ok","data":result}
    finally:
        if os.path.exists(p): os.remove(p)

@app.post("/save-result")
def save_result(data: Dict[str, Any] = Body(...)):
    # 행 구성 (A..F)
    row = [
        data.get("출고일",""),
        data.get("대여자명",""),
        data.get("전화번호",""),
        data.get("주소",""),
        data.get("기기번호",""),
        data.get("기종",""),
    ]

    token = _get_access_token()
    if not token:
        return JSONResponse({"status":"write_failed","write_error":{"error":"no_access_token"}}, status_code=401)

    ok, info = write_row_to_onedrive(row, token=token)
    if ok:
        return {"status":"success","write_info":info}

    # 401 → 토큰 강제 리프레시 후 1회 재시도
    if info.get("error") in ("unauthorized",) or info.get("status") == 401:
        token2 = _get_access_token(force_refresh=True)
        if not token2:
            return JSONResponse({"status":"write_failed","write_error":{"error":"no_access_token"}}, status_code=401)
        ok2, info2 = write_row_to_onedrive(row, token=token2)
        if ok2:
            return {"status":"success","write_info":info2}
        return JSONResponse({"status":"write_failed","write_error":info2}, status_code=500)

    return JSONResponse({"status":"write_failed","write_error":info}, status_code=500)

# ------------------------------- Static / Misc -------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@app.get("/", response_class=HTMLResponse)
def root():
    with open(os.path.join(BASE_DIR,"index.html"),"r",encoding="utf-8") as f:
        return HTMLResponse(f.read(), media_type="text/html; charset=utf-8")

@app.get("/__ping") 
def ping(): return {"ping": str(uuid4())}

@app.get("/manifest.webmanifest", response_class=FileResponse)
def manifest(): return FileResponse(os.path.join(BASE_DIR,"manifest.webmanifest"))

@app.get("/sw.js", response_class=FileResponse)
def sw(): return FileResponse(os.path.join(BASE_DIR,"sw.js"))

@app.get("/__version")
def version(): return {"version":APP_VERSION}

if __name__=="__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT",10000)))



















