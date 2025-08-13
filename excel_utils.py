import os
import urllib.parse
import httpx
import msal

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ["offline_access", "Files.ReadWrite.All", "Sites.ReadWrite.All", "User.Read"]

FILE_NAME = os.getenv("FILE_NAME", "유축기출고.xlsx")
WORKSHEET_NAME = os.getenv("WORKSHEET_NAME", "유축기출고")
TABLE_NAME = os.getenv("TABLE_NAME", "출고내역")

REFRESH_FILE = "refresh_token.txt"


def _build_msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )


def _load_refresh_token():
    if not os.path.exists(REFRESH_FILE):
        raise RuntimeError("refresh_token.txt 없음 → /login 후 /callback 먼저 실행하세요.")
    with open(REFRESH_FILE, encoding="utf-8") as f:
        return f.read().strip()


def get_access_token():
    """
    refresh_token으로 새 access_token 발급
    """
    refresh_token = _load_refresh_token()
    app = _build_msal_app()
    result = app.acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)
    if "access_token" not in result:
        raise RuntimeError(f"토큰 갱신 실패: {result}")
    # refresh_token이 새로 내려오면 교체 저장
    if "refresh_token" in result:
        with open(REFRESH_FILE, "w", encoding="utf-8") as f:
            f.write(result["refresh_token"])
    return result["access_token"]


async def append_row_to_excel(row: dict):
    """
    row 예시:
    {
        "출고일": "2025-07-30",
        "대여자명": "홍길동",
        "전화번호": "010-1234-5678",
        "주소": "서울시...",
        "유축기기종": "Symphony",
        "기기번호": "ABC123",
        "송장번호": "123-456"
    }
    """
    try:
        token = get_access_token()
    except Exception as e:
        print("[OneDrive] ACCESS_TOKEN 획득 실패:", e)
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    encoded_path = urllib.parse.quote(f"/{FILE_NAME}")
    table_url = (
        f"{GRAPH_BASE}/me/drive/root:{encoded_path}:/workbook/worksheets('{WORKSHEET_NAME}')"
        f"/tables('{TABLE_NAME}')/rows"
    )

    values = [[
        row.get("출고일", ""),
        row.get("대여자명", ""),
        row.get("전화번호", ""),
        row.get("주소", ""),
        row.get("유축기기종", ""),
        row.get("기기번호", ""),
        row.get("송장번호", ""),
    ]]

    async with httpx.AsyncClient(timeout=20.0) as client:
        res = await client.post(table_url, headers=headers, json={"values": values})

    if res.status_code not in (200, 201):
        print("[OneDrive] 테이블 행 추가 실패:", res.status_code, res.text)
        return

    print(f"[OneDrive] 업로드 성공 → {FILE_NAME} / {WORKSHEET_NAME} / {TABLE_NAME}")
