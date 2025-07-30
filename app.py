from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, JSONResponse
import os
import httpx

app = FastAPI()

# 환경 변수 불러오기
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
REDIRECT_URI = "https://rent-label-api-client.onrender.com/callback"
SCOPES = "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read"

# ✅ 기본 확인용 루트
@app.get("/")
def root():
    return {"message": "rent-label-api-client is running"}

# 🔐 Microsoft 로그인 유도
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

# 🔑 인증 코드로 액세스 토큰 교환
@app.get("/callback")
async def callback(request: Request):
    code = request.query_params.get("code")
    if not code:
        return JSONResponse(status_code=400, content={"error": "Authorization code missing"})

    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "code": code,
        "grant_type": "authorization_code",
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPES,
    }

    async with httpx.AsyncClient() as client:
        token_response = await client.post(token_url, headers=headers, data=data)

    if token_response.status_code != 200:
        return JSONResponse(status_code=500, content={"error": "Token exchange failed", "details": token_response.text})

    token_data = token_response.json()
    access_token = token_data.get("access_token")
    return {"access_token": access_token}

# 📊 엑셀 조회 엔드포인트
@app.get("/excel-info")
async def get_excel_info(access_token: str):
    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    url = "https://graph.microsoft.com/v1.0/me/drive/root:/유축기출고.xlsx:/workbook/worksheets('유축기출고')/range(address='A1:F10')"

    async with httpx.AsyncClient() as client:
        response = await client.get(url, headers=headers)

    if response.status_code != 200:
        return {"error": "Excel read failed", "details": response.text}

    return response.json()


