from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse
import httpx
import os

app = FastAPI()

# 환경 변수 불러오기
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = "https://rent-label-api-client.onrender.com/callback"
SCOPES = "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read"

# 기본 루트 확인용
@app.get("/")
def root():
    return {"message": "rent-label-api-client is running"}

# 🔐 Microsoft 로그인 유도 엔드포인트
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
import json
import urllib.parse

from fastapi.responses import JSONResponse
import httpx

@app.get("/callback")
async def callback(request: Request):
    code = request.query_params.get("code")

    if not code:
        return JSONResponse(status_code=400, content={"error": "Authorization code not found"})

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }

    data = {
        "client_id": CLIENT_ID,
        "scope": SCOPES,
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "client_secret": CLIENT_SECRET
    }

    async with httpx.AsyncClient() as client:
        response = await client.post(token_url, headers=headers, data=data)

    if response.status_code != 200:
        return JSONResponse(status_code=500, content={"error": "Token request failed", "details": response.text})

    token_data = response.json()
    return token_data  # 또는 필요한 항목만 추출해서 반환 가능

@app.get("/excel-info")
async def get_excel_info(request: Request):
    access_token = request.query_params.get("access_token")

    if not access_token:
        return JSONResponse(status_code=400, content={"error": "Access token missing"})

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    # ✅ 여기에 본인의 파일/시트 경로를 정확히 넣으세요 (예시는 임시값)
    excel_api_url = "https://graph.microsoft.com/v1.0/me/drive/root:/유축기출고.xlsx:/workbook/worksheets('Sheet1')/range(address='A1:F10')"

    async with httpx.AsyncClient() as client:
        response = await client.get(excel_api_url, headers=headers)

    if response.status_code != 200:
        return JSONResponse(status_code=500, content={"error": "Excel read failed", "details": response.text})

    return response.json()

from fastapi.responses import JSONResponse
import httpx

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

    # 📌 실제 엑셀 경로와 시트명을 반영해서 수정 필요
    url = "https://graph.microsoft.com/v1.0/me/drive/root:/유축기출고.xlsx:/workbook/worksheets('Sheet1')/range(address='A1:F10')"

    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers)

    if response.status_code != 200:
        return {"error": "Excel read failed", "details": response.text}

    return response.json()
