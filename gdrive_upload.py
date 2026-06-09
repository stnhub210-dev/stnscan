"""
구글 드라이브 업로드 v2 - OAuth 방식
- 본인 구글 계정으로 로그인 → 내 드라이브에 업로드
- 처음 실행 시 브라우저 로그인 1회 필요
"""

import os, sys, json, subprocess
from datetime import datetime

SCAN_FOLDER       = r"C:\Users\User\Desktop\스캔"
CREDENTIALS       = r"C:\Users\User\Downloads\credentials.json"
TOKEN_FILE        = r"C:\Users\User\Downloads\token.json"
FILE_IDS_JSON     = os.path.join(SCAN_FOLDER, "file_ids.json")
DRIVE_FOLDER_NAME = "스캔관리대장"

SCOPES = ["https://www.googleapis.com/auth/drive.file"]

def install_libs():
    try:
        from googleapiclient.discovery import build
        from google_auth_oauthlib.flow import InstalledAppFlow
    except ImportError:
        print("  라이브러리 설치 중...")
        subprocess.run([sys.executable, "-m", "pip", "install",
            "google-api-python-client",
            "google-auth-httplib2",
            "google-auth-oauthlib", "-q"], check=True)
        print("  설치 완료")

def get_service():
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build

    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # credentials.json 이 서비스 계정이면 OAuth용으로 재생성 필요
            # OAuth 클라이언트 ID 방식으로 진행
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS, SCOPES)
            print("\n  브라우저가 열립니다. 구글 계정으로 로그인하세요.")
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w") as f:
            f.write(creds.to_json())

    return build("drive", "v3", credentials=creds)

def get_or_create_folder(service, name):
    res = service.files().list(
        q=f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        fields="files(id,name)"
    ).execute()
    files = res.get("files", [])
    if files:
        fid = files[0]["id"]
        print(f"  기존 폴더 사용: {name} (ID: {fid})")
        return fid
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder"}
    folder = service.files().create(body=meta, fields="id").execute()
    fid = folder["id"]
    service.permissions().create(
        fileId=fid, body={"type":"anyone","role":"reader"}
    ).execute()
    print(f"  새 폴더 생성: {name} (ID: {fid})")
    return fid

def upload_file(service, local_path, fname, folder_id, existing_ids):
    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(local_path, mimetype="application/pdf", resumable=True)
    if fname in existing_ids:
        fid = existing_ids[fname]
        try:
            service.files().update(fileId=fid, media_body=media).execute()
            return fid, "업데이트"
        except: pass
    meta = {"name": fname, "parents": [folder_id]}
    f = service.files().create(body=meta, media_body=media, fields="id").execute()
    fid = f["id"]
    service.permissions().create(
        fileId=fid, body={"type":"anyone","role":"reader"}
    ).execute()
    return fid, "신규"

def load_existing_ids():
    if os.path.exists(FILE_IDS_JSON):
        try:
            with open(FILE_IDS_JSON, encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return {}

def save_ids(ids):
    with open(FILE_IDS_JSON, "w", encoding="utf-8") as f:
        json.dump(ids, f, ensure_ascii=False, indent=2)

def main():
    print("="*55)
    print("  구글 드라이브 업로드 v2 (OAuth)")
    print("="*55)

    print("\n라이브러리 확인 중...")
    install_libs()

    # credentials.json 이 서비스 계정인 경우 OAuth용으로 변환 불가
    # → OAuth 클라이언트 ID를 새로 만들어야 함
    print("  OAuth 키: credentials_oauth.json 사용")

    oauth_creds = r"C:\Users\User\Downloads\credentials_oauth.json"
    global CREDENTIALS
    CREDENTIALS = oauth_creds

    print("\n구글 드라이브 연결 중...")
    try:
        service = get_service()
        print("  ✅ 로그인 성공")
    except Exception as e:
        print(f"  ❌ 로그인 실패: {e}")
        input("\n아무 키나 누르면 종료...")
        return

    folder_id = get_or_create_folder(service, DRIVE_FOLDER_NAME)
    file_ids = load_existing_ids()

    pdfs = [f for f in os.listdir(SCAN_FOLDER) if f.lower().endswith(".pdf")]
    print(f"\n총 {len(pdfs)}개 PDF 업로드 시작\n{'─'*55}")

    new_cnt = upd_cnt = err_cnt = 0
    for i, fname in enumerate(pdfs, 1):
        fpath = os.path.join(SCAN_FOLDER, fname)
        print(f"[{i}/{len(pdfs)}] {fname[:45]}", end=" ")
        try:
            fid, status = upload_file(service, fpath, fname, folder_id, file_ids)
            file_ids[fname] = fid
            print(f"✅ {status}")
            if status == "신규": new_cnt += 1
            else: upd_cnt += 1
            if i % 10 == 0: save_ids(file_ids)
        except Exception as e:
            print(f"❌ {str(e)[:60]}")
            err_cnt += 1

    save_ids(file_ids)
    print(f"\n{'='*55}")
    print(f"  완료: 신규 {new_cnt} / 업데이트 {upd_cnt} / 오류 {err_cnt}")
    print(f"{'='*55}")

    gen_path = os.path.join(SCAN_FOLDER, "generate_html.py")
    if os.path.exists(gen_path):
        print("\nHTML 재생성 중...")
        subprocess.run([sys.executable, gen_path, "once"])
        print("✅ 스캔관리대장.html 갱신 완료")

    input("\n아무 키나 누르면 종료...")

if __name__ == "__main__":
    main()
