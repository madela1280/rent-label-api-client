import os, urllib.parse
import httpx

# 환경변수 필요: ACCESS_TOKEN, FILE_NAME, WORKSHEET_NAME
# FILE_NAME 기본값: 유축기출고.xlsx, WORKSHEET_NAME 기본값: 유축기출고
def append_row_to_excel(row: dict):
    token = os.getenv("ACCESS_TOKEN")
    file_name = os.getenv("FILE_NAME", "유축기출고.xlsx")
    sheet_name = os.getenv("WORKSHEET_NAME", "유축기출고")
    if not token:
        print("[OneDrive] ACCESS_TOKEN 없음 → 업로드 생략")
        return

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    encoded_path = urllib.parse.quote(f"/{file_name}")
    base_url = f"https://graph.microsoft.com/v1.0/me/drive/root:{encoded_path}:/workbook/worksheets('{sheet_name}')"

    # 업로드 값 준비 (열 순서 고정)
    values = [[
        row.get("출고일",""),
        row.get("대여자명",""),
        row.get("전화번호",""),
        row.get("주소",""),
        row.get("기기번호",""),
        row.get("기종",""),
        row.get("송장번호",""),
    ]]

    try:
        with httpx.Client(timeout=20.0) as client:
            # 마지막 사용된 행 번호 조회
            used_res = client.get(f"{base_url}/usedRange", headers=headers)
            used_data = used_res.json()
            if used_res.status_code != 200 or "address" not in used_data:
                print("[OneDrive] usedRange 조회 실패:", used_res.status_code, used_data)
                return

            # 예: '유축기출고!A1:G23' → 23 추출
            last_row = int(used_data["address"].split("!")[1].split(":")[1][1:])
            next_row = last_row + 1
            target_range = f"A{next_row}:G{next_row}"

            # 범위에 값 적기
            range_url = f"{base_url}/range(address='{target_range}')"
            patch_res = client.patch(range_url, headers=headers, json={"values": values})
            if patch_res.status_code != 200:
                print("[OneDrive] range patch 실패:", patch_res.status_code, patch_res.text)
                return

            print(f"[OneDrive] 업로드 성공 → {file_name} / {sheet_name} / {target_range}")

    except Exception as e:
        print("[OneDrive] 예외:", repr(e))

