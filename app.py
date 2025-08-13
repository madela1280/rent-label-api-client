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
app.add_middleware(SessionMiddleware, secret_key=os.getenv("SESSION_SECRET", "change-me"))

# -------------------------------
# ENV & Constants
# -------------------------------
# 🔒 과거 값 개입 방지: 하드 고정 (환경변수 무시)
CLIENT_ID = "41745db3-a5c5-4e6e-acd7-fc4ce18b1999"
TENANT_ID = "405ba8a3-73ff-4423-8925-d9eda360cfa7"
CLIENT_SECRET = os.getenv("CLIENT_SECRET")  # 시크릿만 env에서 읽음
REDIRECT_URI = "https://rent-label-api-client-docker.onrender.com/callback"

SCOPES = ["offline_access", "Files.ReadWrite.All", "Sites.ReadWrite.All", "User.Read"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Excel/Graph 관련
FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")  # /onedrive/ping 용(선택)
GRAPH = "https://graph.microsoft.com/v1.0"

# -------------------------------
# MSAL App 생성
# -------------------------------
def _build_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )

# -------------------------------
# 기본/상태
# -------------------------------
@app.get("/")
def root():
    return {"message": "✅ rent-label-api-client is running"}

# -------------------------------
# 테스트 이미지 업로드 → OCR → 엑셀 반영
# -------------------------------
@app.post("/upload-test-image/")
async def upload_test_image(image: UploadFile = File(...)):
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as buffer:
        shutil.copyfileobj(image.file, buffer)
    try:
        result = make_final_entry("TEST_QR", temp_path)
        append_row_to_excel(result)
        return {"status": "success", "data": result}
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

# -------------------------------
# 로그인 (Azure OAuth2 - MSAL)
# -------------------------------
@app.get("/login")
def login(request: Request):
    request.session["state"] = str(uuid.uuid4())
    nonce = str(uuid.uuid4())  # 캐시/이전값 방지
    auth_url = _build_msal_app().get_authorization_request_url(
        scopes=SCOPES,
        state=request.session["state"],
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
        # 불필요한 과거 client_id가 끼어들 여지 제거
        # (msal은 여기서 client_id를 내부 설정(CLIENT_ID)로 사용)
    )
    # 캐시 무효화를 위해 쿼리에 nonce 부착
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

    if "access_token" not in result:
        return JSONResponse({"error": "Token acquire failed", "details": result}, status_code=400)

    # refresh_token 저장 (서버 로컬 파일)
    try:
        with open("refresh_token.txt", "w", encoding="utf-8") as f:
            f.write(result.get("refresh_token", ""))
    except Exception:
        pass

    # 세션 저장
    request.session["tokens"] = {
        "access_token": result["access_token"],
        "refresh_token": result.get("refresh_token"),
        "expires_in": result.get("expires_in"),
        "id_token_claims": result.get("id_token_claims"),
    }

    return RedirectResponse("/me")

@app.get("/me")
def me(request: Request):
    tokens = request.session.get("tokens")
    if not tokens:
        return RedirectResponse("/login")
    return JSONResponse({"status": "ok", "id_token_claims": tokens.get("id_token_claims")})

# -------------------------------
# 임의로 엑셀 특정 행에 쓰는 API (토큰 직접 전달)
# -------------------------------
class ExcelInput(BaseModel):
    access_token: str
    row: list  # 예: ["2025-07-30", "홍길동", "010-1234-5678", "주소", "유축기기종", "기기번호", "송장번호"]

@app.post("/write-excel")
async def write_excel(data: ExcelInput):
    headers = {
        "Authorization": f"Bearer {data.access_token}",
        "Content-Type": "application/json",
    }
    encoded_path = urllib.parse.quote(f"/{FILE_NAME}")
    base_url = f"{GRAPH}/me/drive/root:{encoded_path}:/workbook/worksheets('{WORKSHEET_NAME}')"
    try:
        async with httpx.AsyncClient(timeout=20.0) as client:
            used_range_url = f"{base_url}/usedRange"
            used_res = await client.get(used_range_url, headers=headers)
            used_data = used_res.json()

            if "address" not in used_data:
                return {"error": "Unable to detect used range", "details": used_data}

            address = used_data["address"]
            last_row = int(address.split("!")[1].split(":")[1][1:])
            next_row = last_row + 1
            target_range = f"A{next_row}:G{next_row}"

            range_url = f"{base_url}/range(address='{target_range}')"
            response = await client.patch(range_url, headers=headers, json={"values": [data.row]})

            if response.status_code != 200:
                return {"error": "Failed to write to Excel", "details": response.text}
    except Exception as e:
        return {"error": "Internal Server Error", "details": str(e)}
    return {"status": "success", "row": data.row, "range": target_range}

# -------------------------------
# OCR → 엑셀 반영 (실사용 라우트)
# -------------------------------
@app.post("/process-ocr/")
async def process_ocr(qr_text: str = Form(...), image: UploadFile = File(...)):
    temp_path = f"temp_{image.filename}"
    with open(temp_path, "wb") as buffer:
        shutil.copyfileobj(image.file, buffer)
    try:
        result = make_final_entry(qr_text, temp_path)
        append_row_to_excel(result)
        return {"status": "success", "data": result}
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

# -------------------------------
# 원드라이브 파일 존재 Ping (선택)
# -------------------------------
def _auth():
    return {"Authorization": f"Bearer {ACCESS_TOKEN}"} if ACCESS_TOKEN else {}

@app.get("/onedrive/ping")
def onedrive_ping():
    if not ACCESS_TOKEN:
        return {"error": "ACCESS_TOKEN env not set"}
    try:
        r = requests.get(f"{GRAPH}/me/drive/root:/{FILE_NAME}", headers=_auth())
        return {"status": r.status_code, "json": r.json()}
    except Exception as e:
        return {"error": str(e)}

# -------------------------------
# DEBUG: Azure 환경 확인
# -------------------------------
@app.get("/__debug/azure")
def dbg():
    return {
        "client_id": CLIENT_ID,
        "tenant_id": TENANT_ID,
        "secret_len": len(CLIENT_SECRET) if CLIENT_SECRET else 0,
        "secret_fp": hashlib.sha256((CLIENT_SECRET or '').encode()).hexdigest()[:12],
        "file_name": FILE_NAME,
        "worksheet": WORKSHEET_NAME,
    }



