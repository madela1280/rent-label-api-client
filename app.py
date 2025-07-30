from fastapi import FastAPI, Request
from fastapi.responses import RedirectResponse, JSONResponse
import httpx

app = FastAPI()

# 🔐 하드코딩된 인증 정보
CLIENT_ID = "c4c5125d-7475-4eb1-a4ee-f3deb0788280"
CLIENT_SECRET = "Zxr8Q~r-2EK64w4t9yejhU6L4QjJO1IrHLfLda0a"
TENANT_ID = "405ba8a3-73ff-4423-8925-d9eda360cfa7"
REDIRECT_URI = "https://rent-label-api-client.onrender.com/callback"
SCOPES = "https://graph.microsoft.com/.default"

# 🔍 서버 동작 확인용 루트
@app.get("/")
def root():
    return {"message": "rent-label-api-client is running"}

# 🔗 로그인 유도 URL 생성
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

# 🔁 콜백 처리: 인증 코드 → 토큰
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

# 📊 Excel 정보 조회
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


