import os, shutil, hashlib, uuid, requests, time
from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File, Body, BackgroundTasks
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware

import msal
from uuid import uuid4
from typing import Optional, Dict, Any
from ocr_utils import make_final_entry, make_final_entry_fast

_HTTP = requests.Session()
APP_VERSION = os.getenv("APP_VERSION", "2025-08-27-restore-fast5")

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
    allow_origins=["*"], allow_credentials=True,
    allow_methods=["*"], allow_headers=["*"],
)

# -------------------------------
# ENV
# -------------------------------
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://rent-label-api-client-docker.onrender.com/callback")

SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH = "https://graph.microsoft.com/v1.0"

FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")

# -------------------------------
# MSAL Helper
# -------------------------------
def _build_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )

def _save_tokens(result:dict):
    with open("refresh_token.txt","w",encoding="utf-8") as f:
        f.write(result.get("refresh_token",""))
    with open("access_token.txt","w",encoding="utf-8") as f:
        f.write(result.get("access_token",""))

def _load_refresh_token()->Optional[str]:
    if os.path.exists("refresh_token.txt"):
        return open("refresh_token.txt","r",encoding="utf-8").read().strip()
    return None

def _load_access_token()->Optional[str]:
    if os.path.exists("access_token.txt"):
        return open("access_token.txt","r",encoding="utf-8").read().strip()
    return None

def _get_access_token()->Optional[str]:
    tok = _load_access_token()
    if tok: return tok
    rtok = _load_refresh_token()
    if not rtok: return None
    result = _build_msal_app().acquire_token_by_refresh_token(rtok, scopes=SCOPES)
    if "access_token" in result:
        _save_tokens(result)
        return result["access_token"]
    return None

# -------------------------------
# Graph Helper (짧은 타임아웃 + 1회 재시도)
# -------------------------------
_DRIVE_ITEM_ID_CACHE = {"name": None, "id": None}

def _http_get(url, headers, timeout=4):
    try:
        return _HTTP.get(url, headers=headers, timeout=timeout)
    except requests.RequestException as e:
        raise e

def _http_patch(url, headers, json, timeout=4):
    try:
        return _HTTP.patch(url, headers=headers, json=json, timeout=timeout)
    except requests.RequestException as e:
        raise e

def _with_retry(fn, *args, **kwargs):
    # 1회 재시도 (총 2번)
    try:
        return fn(*args, **kwargs)
    except Exception:
        time.sleep(0.5)
        return fn(*args, **kwargs)

def _get_drive_item_id(headers, file_name):
    if _DRIVE_ITEM_ID_CACHE["name"] == file_name and _DRIVE_ITEM_ID_CACHE["id"]:
        return _DRIVE_ITEM_ID_CACHE["id"]
    resp = _with_retry(_http_get, f"{GRAPH}/me/drive/root/search(q='{file_name}')?$top=1", headers, 4)
    search = resp.json()
    items = search.get("value", [])
    if not items or items[0]["name"] != file_name: return None
    _DRIVE_ITEM_ID_CACHE["name"] = file_name
    _DRIVE_ITEM_ID_CACHE["id"] = items[0]["id"]
    return items[0]["id"]

def write_row_to_onedrive(row):
    token = _get_access_token()
    if not token: return False, {"error":"no_access_token"}
    headers = {"Authorization": f"Bearer {token}", "Content-Type":"application/json"}
    item_id = _get_drive_item_id(headers, FILE_NAME)
    if not item_id: return False, {"error":"file_not_found","file":FILE_NAME}

    used = _with_retry(_http_get, f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/usedRange", headers, 4).json()
    addr = used.get("address") or f"{SHEET_NAME}!A1:A1"
    try: last_row = int(addr.split("!")[1].split(":")[1][1:])
    except: last_row = 1
    next_row = last_row+1
    target = f"A{next_row}:F{next_row}"

    resp = _with_retry(
        _http_patch,
        f"{GRAPH}/me/drive/items/{item_id}/workbook/worksheets('{SHEET_NAME}')/range(address='{target}')",
        headers,
        {"values":[row]},
        4
    )
    if resp.status_code!=200:
        return False, {"error":"write_failed","status":resp.status_code,"text":resp.text}
    return True, {"range":target}

# -------------------------------
# In-Memory Job Queue
# -------------------------------
JOBS: Dict[str, Dict[str, Any]] = {}

def _background_write(job_id: str, row):
    try:
        JOBS[job_id]["status"] = "running"
        ok, info = write_row_to_onedrive(row)
        if ok:
            JOBS[job_id]["status"] = "success"
            JOBS[job_id]["result"] = info
        else:
            JOBS[job_id]["status"] = "failed"
            JOBS[job_id]["error"]  = info
    except Exception as e:
        JOBS[job_id]["status"] = "failed"
        JOBS[job_id]["error"]  = {"error":"exception","text":str(e)}

# -------------------------------
# Routes: Login
# -------------------------------
@app.get("/login")
def login(request: Request):
    request.session["state"] = str(uuid.uuid4())
    auth_url = _build_msal_app().get_authorization_request_url(
        scopes=SCOPES, state=request.session["state"],
        redirect_uri=REDIRECT_URI, prompt="select_account", response_mode="query",
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
        code, scopes=SCOPES, redirect_uri=REDIRECT_URI
    )
    if "access_token" not in result:
        return JSONResponse({"error":"Token acquire failed","details":result}, status_code=400)

    _save_tokens(result)
    request.session["tokens"] = {"access_token": result["access_token"]}
    return RedirectResponse("/")

# -------------------------------
# OCR Endpoints
# -------------------------------
@app.post("/process-ocr/")
async def process_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    temp_path = f"temp_{image.filename}"
    with open(temp_path,"wb") as f: shutil.copyfileobj(image.file,f)
    try:
        result = make_final_entry(qr_text, temp_path)
        row = [
            result.get("출고일",""),
            result.get("대여자명",""),
            result.get("전화번호",""),
            result.get("주소",""),
            result.get("기기번호",""),
            result.get("기종",""),
        ]
        # 여기서는 기존대로 동기 저장 시도 (정식 OCR 단계)
        ok, info = write_row_to_onedrive(row)
        if not ok:
            return {"status":"ocr_ok_but_write_failed","data":result,"write_error":info}
        return {"status":"success","data":result,"write_info":info}
    finally:
        if os.path.exists(temp_path): os.remove(temp_path)

@app.post("/preview-ocr")
async def preview_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    temp_path = f"temp_{image.filename}"
    with open(temp_path,"wb") as f: shutil.copyfileobj(image.file,f)
    try:
        result = make_final_entry_fast(qr_text,temp_path)
        return {"status":"preview","data":result}
    finally:
        if os.path.exists(temp_path): os.remove(temp_path)

# -------------------------------
# Save Result (비동기 큐 + 즉시 응답)
# -------------------------------
@app.post("/save-result")
def save_result(background_tasks: BackgroundTasks, data: Dict[str, Any] = Body(...)):
    row = [
        data.get("출고일",""),
        data.get("대여자명",""),
        data.get("전화번호",""),
        data.get("주소",""),
        data.get("기기번호",""),
        data.get("기종",""),
    ]

    # 토큰 체크 선행(없으면 바로 401)
    tok = _get_access_token()
    if not tok:
        return JSONResponse({"status":"write_failed","write_error":{"error":"no_access_token"}, "row":row}, status_code=401)

    # 비동기 큐에 등록
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status":"queued", "created": time.time()}
    background_tasks.add_task(_background_write, job_id, row)

    # 즉시 응답 (클라이언트는 최대 5초만 폴링)
    return {"status":"queued", "job_id": job_id}

@app.get("/job-status")
def job_status(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        return JSONResponse({"status":"not_found"}, status_code=404)
    payload = {"status": job["status"]}
    if job["status"] == "success":
        payload["result"] = job.get("result")
        payload["status_text"] = "완료"
    elif job["status"] == "failed":
        payload["error"] = job.get("error")
        payload["status_text"] = "실패"
    else:
        payload["status_text"] = "진행중"
    return payload

# -------------------------------
# Static / Misc
# -------------------------------
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

















