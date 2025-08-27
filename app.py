import os, shutil, uuid, time, requests
from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File, Body, BackgroundTasks
from fastapi.responses import RedirectResponse, JSONResponse, FileResponse, HTMLResponse
from starlette.middleware.sessions import SessionMiddleware
from fastapi.middleware.cors import CORSMiddleware

import msal
from typing import Optional, Dict, Any
from uuid import uuid4
from ocr_utils import make_final_entry, make_final_entry_fast

_HTTP = requests.Session()
APP_VERSION = os.getenv("APP_VERSION", "2025-08-27-fast-queue")

app = FastAPI()
app.add_middleware(SessionMiddleware, secret_key=os.getenv("SESSION_SECRET","change-me"),
                   same_site="none", https_only=True, max_age=3600, session_cookie="session")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True,
                   allow_methods=["*"], allow_headers=["*"])

CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "https://rent-label-api-client-docker.onrender.com/callback")

SCOPES = ["User.Read", "Files.ReadWrite.All", "Sites.ReadWrite.All"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH = "https://graph.microsoft.com/v1.0"

FILE_NAME   = os.getenv("FILE_NAME", "유축기출고.xlsx")
SHEET_NAME  = os.getenv("WORKSHEET_NAME", "유축기출고")
TABLE_NAME  = os.getenv("TABLE_NAME", "출고내역")

def _msal(): return msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
def _save_tokens(result:dict):
    open("refresh_token.txt","w",encoding="utf-8").write(result.get("refresh_token",""))
    open("access_token.txt","w",encoding="utf-8").write(result.get("access_token",""))
def _load(fname): return open(fname,"r",encoding="utf-8").read().strip() if os.path.exists(fname) else None

def _get_access_token()->Optional[str]:
    tok = _load("access_token.txt")
    if tok: return tok
    rt = _load("refresh_token.txt")
    if not rt: return None
    r = _msal().acquire_token_by_refresh_token(rt, scopes=SCOPES)
    if "access_token" in r:
        _save_tokens(r); return r["access_token"]
    return None

_DRIVE_ITEM_ID = None
_SESSION_ID = None
def _graph_headers(token): 
    h = {"Authorization": f"Bearer {token}", "Content-Type":"application/json"}
    if _SESSION_ID: h["workbook-session-id"] = _SESSION_ID
    return h

def _ensure_item_id(token):
    global _DRIVE_ITEM_ID
    if _DRIVE_ITEM_ID: return _DRIVE_ITEM_ID
    r = _HTTP.get(f"{GRAPH}/me/drive/root/search(q='{FILE_NAME}')?$top=1", headers=_graph_headers(token), timeout=4)
    r.raise_for_status()
    v = r.json().get("value", [])
    if not v or v[0]["name"] != FILE_NAME: raise RuntimeError("file_not_found")
    _DRIVE_ITEM_ID = v[0]["id"]; return _DRIVE_ITEM_ID

def _ensure_session(token):
    global _SESSION_ID
    if _SESSION_ID: return _SESSION_ID
    r = _HTTP.post(f"{GRAPH}/me/drive/items/{_ensure_item_id(token)}/workbook/createSession",
                   headers={"Authorization": f"Bearer {token}", "Content-Type":"application/json"},
                   json={"persist": True}, timeout=4)
    r.raise_for_status()
    _SESSION_ID = r.json().get("id")
    return _SESSION_ID

def _rows_add(token, values):
    # 테이블 rows/add (usedRange/patch 안 씀)
    _ensure_session(token)
    url = f"{GRAPH}/me/drive/items/{_ensure_item_id(token)}/workbook/worksheets('{SHEET_NAME}')/tables('{TABLE_NAME}')/rows/add"
    r = _HTTP.post(url, headers=_graph_headers(token), json={"values":[values]}, timeout=6)
    if r.status_code == 409:  # 세션 만료 등
        # 세션 초기화 후 재시도 1회
        global _SESSION_ID
        _SESSION_ID = None
        _ensure_session(token)
        r = _HTTP.post(url, headers=_graph_headers(token), json={"values":[values]}, timeout=6)
    r.raise_for_status()
    return True

JOBS: Dict[str, Dict[str, Any]] = {}

def _bg_write(job_id: str, row):
    try:
        JOBS[job_id]["status"] = "running"
        tok = _get_access_token()
        if not tok: 
            JOBS[job_id]["status"]="failed"; JOBS[job_id]["error"]={"error":"no_access_token"}; return
        _rows_add(tok, row)
        JOBS[job_id]["status"]="success"
    except Exception as e:
        JOBS[job_id]["status"]="failed"; JOBS[job_id]["error"]={"error":"write_failed","text":str(e)}

@app.get("/login")
def login(request: Request):
    request.session["state"] = str(uuid.uuid4())
    url = _msal().get_authorization_request_url(scopes=SCOPES, state=request.session["state"],
                                                redirect_uri=REDIRECT_URI, prompt="select_account", response_mode="query")
    return RedirectResponse(url)

@app.get("/callback")
async def callback(request: Request):
    if request.query_params.get("state") != request.session.get("state"):
        return JSONResponse({"error":"state mismatch"}, status_code=400)
    code = request.query_params.get("code")
    if not code: return JSONResponse({"error":"no code"}, status_code=400)
    result = _msal().acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri=REDIRECT_URI)
    if "access_token" not in result: return JSONResponse({"error":"token fail","details":result}, status_code=400)
    _save_tokens(result); request.session["tokens"]={"access_token":result["access_token"]}
    return RedirectResponse("/")

@app.post("/preview-ocr")
async def preview_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    # 프리뷰는 OCR만 (저장은 안함)
    p = f"temp_{image.filename}"
    with open(p,"wb") as f: shutil.copyfileobj(image.file,f)
    try:
        data = make_final_entry_fast(qr_text, p)
        return {"status":"preview","data":data}
    finally:
        if os.path.exists(p): os.remove(p)

@app.post("/process-ocr/")
async def process_ocr(qr_text: str = Form(""), image: UploadFile = File(...)):
    # 정식도 OCR만 (저장은 confirm에서)
    p = f"temp_{image.filename}"
    with open(p,"wb") as f: shutil.copyfileobj(image.file,f)
    try:
        data = make_final_entry(qr_text, p)
        return {"status":"ok","data":data}
    finally:
        if os.path.exists(p): os.remove(p)

@app.post("/save-result")
def save_result(background_tasks: BackgroundTasks, data: Dict[str, Any] = Body(...)):
    # 즉시 queued 응답, 백그라운드에서 rows/add
    row = [
        data.get("출고일",""),
        data.get("대여자명",""),
        data.get("전화번호",""),
        data.get("주소",""),
        data.get("기기번호",""),
        data.get("기종",""),
    ]
    tok = _get_access_token()
    if not tok:
        return JSONResponse({"status":"write_failed","write_error":{"error":"no_access_token"}}, status_code=401)

    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"status":"queued","created": time.time()}
    background_tasks.add_task(_bg_write, job_id, row)
    return {"status":"queued","job_id":job_id}

@app.get("/job-status")
def job_status(job_id: str):
    j = JOBS.get(job_id)
    if not j: return JSONResponse({"status":"not_found"}, status_code=404)
    return {"status": j["status"], **({"error":j["error"]} if "error" in j else {})}

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
@app.get("/", response_class=HTMLResponse)
def root():
    return HTMLResponse(open(os.path.join(BASE_DIR,"index.html"),"r",encoding="utf-8").read(),
                        media_type="text/html; charset=utf-8")
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

















