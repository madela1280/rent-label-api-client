import os, urllib.parse
import httpx

# 환경변수: ACCESS_TOKEN, FILE_NAME, WORKSHEET_NAME, TABLE_NAME
# 기본값: 유축기출고.xlsx / 유축기출고 / 출고내역
def append_row_to_excel(row: dict):
    token = os.getenv("ACCESS_TOKEN") or get_access_token()
    file_name = os.getenv("FILE_NAME", "유축기출고.xlsx")
    sheet_name = os.getenv("WORKSHEET_NAME", "유축기출고")
    table_name = os.getenv("TABLE_NAME", "출고내역")

    if not token:
        print("[OneDrive] ACCESS_TOKEN 없음 → 업로드 생략")
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    try:
        table_url = (
            f"https://graph.microsoft.com/v1.0/me/drive/root:/{file_name}:/workbook/worksheets('{sheet_name}')"
            f"/tables('{table_name}')/rows"
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

        with httpx.Client(timeout=20.0) as client:
            res = client.post(table_url, headers=headers, json={"values": values})
            if res.status_code not in (200, 201):
                print("[OneDrive] 테이블 행 추가 실패:", res.status_code, res.text)
                return

        print(f"[OneDrive] 업로드 성공 → {file_name} / {sheet_name} / {table_name}")

    except Exception as e:
        print("[OneDrive] 예외:", repr(e))

def get_access_token():
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    tenant_id = os.getenv("TENANT_ID")

    if not os.path.exists("refresh_token.txt"):
        raise RuntimeError("refresh_token.txt 없음 → /login 후 /callback 먼저 실행하세요.")

    with open("refresh_token.txt") as f:
        refresh_token = f.read().strip()

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "refresh_token": refresh_token,
        "grant_type": "refresh_token",
        "scope": "offline_access Files.ReadWrite.All Sites.ReadWrite.All User.Read"
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}

    resp = httpx.post(url, data=data, headers=headers)
    if resp.status_code != 200:
        raise RuntimeError(f"토큰 갱신 실패: {resp.text}")

    return resp.json()["access_token"]
