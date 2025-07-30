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

