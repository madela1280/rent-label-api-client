import os
import shutil
import hashlib
import urllib.parse

from dotenv import load_dotenv; load_dotenv()

from fastapi import FastAPI, Request, UploadFile, Form, File
from fastapi.responses import RedirectResponse, JSONResponse
from pydantic import BaseModel

import httpx
import requests

from ocr_utils import make_final_entry
from excel_utils import append_row_to_excel

# -------------------------------
# FastAPI
# -------------------------------
app = FastAPI()

# -------------------------------
# ENV & Constants
# -------------------------------
CLIENT_ID = os.getenv("CLIENT_ID")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# Excel/Graph 관련
FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")
ACCESS_TOKEN = os.getenv("ACCESS_TOKEN")  # /onedrive/ping 용(선택)

# OAuth/Graph
REDIRECT_URI = "https://rent-label-api-client-docker.onrender.com/callback"
SCOPES = "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read"
GRAPH = "https://graph.microsoft.com/v1.0"


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
# 로그인 (Azure OAuth2)
# -------------------------------
@app.get("/login")
def login():
    url = (
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
        f"&response_mode=query"
        f"&scope={SCOPES}"
    )
    return RedirectResponse(url)


# -------------------------------
# 콜백 (인증 코드 → 토큰 교환)
# -------------------------------
@app.get("/callback")
async def callback(request: Request):
    code = request.query_params.get("code")
    if not code:
        return JSONResponse(status_code=400, content={"error": "Authorization code missing"})

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "grant_type": "authorization_code",
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    async with httpx.AsyncClient(timeout=20.0) as client:
        token_response = await client.post(token_url, data=data, headers=headers)

    if token_response.status_code != 200:
        return JSONResponse(
            status_code=500,
            content={"error": "Token exchange failed", "details": token_response.text},
        )

    tok = token_response.json()

    # refresh_token 저장 (서버 로컬 파일)
    try:
        with open("refresh_token.txt", "w", encoding="utf-8") as f:
            f.write(tok.get("refresh_token", ""))
    except Exception as e:
        # 저장 실패해도 토큰은 응답으로 줌
        pass

    return {
        "access_token": tok.get("access_token"),
        "refresh_token": tok.get("refresh_token"),
        "expires_in": tok.get("expires_in"),
        "message": "refresh_token 저장 완료",
    }


# -------------------------------
# 임의로 엑셀 특정 행에 쓰는 API (토큰 직접 전달)
# -------------------------------
class ExcelInput(BaseModel):
    access_token: str
    # 예: ["2025-07-30", "홍길동", "010-1234-5678", "주소", "유축기기종", "기기번호", "송장번호"]
    row: list


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
            # 사용된 범위 가져와서 다음 행 계산
            used_range_url = f"{base_url}/usedRange"
            used_res = await client.get(used_range_url, headers=headers)
            used_data = used_res.json()

            if "address" not in used_data:
                return {"error": "Unable to detect used range", "details": used_data}

            # 예: '유축기출고'!A1:G5 → 마지막 행 5 추출
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


