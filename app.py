from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, JSONResponse
import httpx
import os

app = FastAPI()

# 환경 변수
CLIENT_ID = "c4c5125d-7475-4eb1-a4ee-f3deb0788280"
CLIENT_SECRET = "Zxr8Q~r-2EK64w4t9yejhU6L4QjJO1IrHLfLda0a"
TENANT_ID = "405ba8a3-73ff-4423-8925-d9eda360cfa7"
REDIRECT_URI = "https://rent-label-api-client.onrender.com/callback"
SCOPES = "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read"

@app.get("/")
def root():
    return {"message": "rent-label-api-client is running"}

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

    token_data = token_response.json()
    access_token = token_data.get("access_token")

    return {"access_token": access_token}

@app.get("/excel-info")
async def get_excel_info(access_token: str):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    url = "https://graph.microsoft.com/v1.0/me/drive/root:/유축기출고.xlsx:/workbook/worksheets('유축기출고')/range(address='A1:F100')"

    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers)

    if response.status_code != 200:
        return {"error": "Excel read failed", "details": response.text}

    return response.json()

from pydantic import BaseModel

class ExcelInput(BaseModel):
    access_token: str
    row: list  # 예: ["2025-07-30", "홍길동", "010-1234-5678", "서울시...", "SM123456", "기종"]

@app.post("/write-excel")
async def write_excel(data: ExcelInput):
    headers = {
        "Authorization": f"Bearer {data.access_token}",
        "Content-Type": "application/json"
    }

    # 데이터를 추가할 테이블 방식으로 전송 (행 추가)
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/유축기출고.xlsx:/workbook/worksheets('유축기출고')/tables/add"
    
    # 테이블이 없을 경우 생성
    async with httpx.AsyncClient() as client:
        await client.post(url, headers=headers, json={
            "address": "A1:F1",
            "hasHeaders": True
        })

    # 테이블에 행 추가
    row_add_url = "https://graph.microsoft.com/v1.0/me/drive/root:/유축기출고.xlsx:/workbook/worksheets('유축기출고')/tables('Table1')/rows/add"
    async with httpx.AsyncClient() as client:
        response = await client.post(row_add_url, headers=headers, json={
            "values": [data.row]
        })

    if response.status_code != 201:
        return {"error": "Failed to write to Excel", "details": response.text}

    return {"status": "success", "written_row": data.row}



