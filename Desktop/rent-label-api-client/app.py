from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, JSONResponse
from pydantic import BaseModel
import httpx
import urllib.parse

from fastapi import UploadFile, Form, File
import shutil, os

from ocr_utils import make_final_entry
from excel_utils import append_row_to_excel

app = FastAPI()  # ✅ FastAPI 객체 먼저 생성

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
        os.remove(temp_path)

# 인증 설정
CLIENT_ID = "c4c5125d-7475-4eb1-a4ee-f3deb0788280"
CLIENT_SECRET = "Zxr8Q~r-2EK64w4t9yejhU6L4QjJO1IrHLfLdaOa"
TENANT_ID = "405ba8a3-73ff-4423-8925-d9eda360cfa7"
REDIRECT_URI = "https://rent-label-api-client.onrender.com/callback"
SCOPES = "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read"

# 엑셀 파일 위치 (내 파일 기준)
FILE_NAME = "유축기출고.xlsx"
WORKSHEET_NAME = "유축기출고"

@app.get("/")
def root():
    return {"message": "✅ rent-label-api-client is running"}

@app.get("/login")
def login():
    return RedirectResponse(
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/authorize"
        f"?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={REDIRECT_URI}"
        f"&response_mode=query"
        f"&scope={SCOPES}"
    )

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

    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    async with httpx.AsyncClient() as client:
        token_response = await client.post(token_url, data=data, headers=headers)

    if token_response.status_code != 200:
        return JSONResponse(status_code=500, content={"error": "Token exchange failed", "details": token_response.text})

    access_token = token_response.json().get("access_token")
    return {"access_token": access_token}

class ExcelInput(BaseModel):
    access_token: str
    row: list  # 예: ["2025-07-30", "홍길동", "010-1234-5678", "주소", "기기번호", "기종", "송장번호"]

@app.post("/write-excel")
async def write_excel(data: ExcelInput):
    headers = {
        "Authorization": f"Bearer {data.access_token}",
        "Content-Type": "application/json"
    }

    encoded_path = urllib.parse.quote(f"/{FILE_NAME}")
    base_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_path}:/workbook/worksheets('{WORKSHEET_NAME}')"

    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            # 현재 데이터가 입력된 마지막 행 확인
            used_range_url = f"{base_url}/usedRange"
            used_res = await client.get(used_range_url, headers=headers)
            used_data = used_res.json()

            if "address" not in used_data:
                return {"error": "Unable to detect used range", "details": used_data}

            last_row = int(used_data["address"].split("!")[1].split(":")[1][1:])
            next_row = last_row + 1
            target_range = f"A{next_row}:G{next_row}"

            range_url = f"{base_url}/range(address='{target_range}')"
            response = await client.patch(range_url, headers=headers, json={"values": [data.row]})

            if response.status_code != 200:
                return {"error": "Failed to write to Excel", "details": response.text}

    except Exception as e:
        return {"error": "Internal Server Error", "details": str(e)}

    return {
        "status": "success",
        "row": data.row,
        "range": target_range
    }

from fastapi import File

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
        os.remove(temp_path)
