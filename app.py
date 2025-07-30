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

