import os
import shutil
import hashlib
import urllib.parse
import uuid

from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File
from fastapi.responses import RedirectResponse, JSONResponse
from pydantic import BaseModel
from starlette.middleware.sessions import SessionMiddleware

import httpx
import requests
import msal

from ocr_utils import make_final_entry
from excel_utils import append_row_to_excel

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
    prompt="select_account",
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
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    # ✅ 토큰 실패 시 상세 에러를 그대로 반환해 ‘무한 추측’ 방지
    if "access_token" not in result:
        return JSONResponse({"error": "Token acquire failed", "details": result}, status_code=400)

    try:
        with open("refresh_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("refresh_token", ""))
    except Exception:
        pass

    request.session["tokens"] = {
        "access_token": result["access_token"],
        "refresh_token": result.get("refresh_token"),
        "expires_in": result.get("expires_in"),
        "id_token_claims": result.get("id_token_claims"),
    }

    return RedirectResponse("/me")

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

@app.get("/__ping")
def ping(): return {"ping": str(uuid4())}

@app.get("/")
def root():
    return {"message": "probe1"}

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

@app.get("/me")
def me(request: Request):
    tokens = request.session.get("tokens")
    if not tokens:
        return RedirectResponse("/login")
    return JSONResponse({"status": "ok", "id_token_claims": tokens.get("id_token_claims")})




