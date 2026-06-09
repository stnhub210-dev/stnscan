"""
구글 드라이브 파일 ID 수집기
- 공유 폴더의 파일 ID를 자동으로 수집
- file_ids.json 저장 → generate_html.py에서 사용
"""
import os, json, re, sys
import urllib.request

FOLDER_ID   = "1hT9xdNEawnkrZ78p26Kcjlm-1wdR5x-X"
SCAN_FOLDER = r"C:\Users\User\Desktop\스캔"
OUTPUT_JSON = os.path.join(SCAN_FOLDER, "file_ids.json")

def fetch_file_ids():
    """구글 드라이브 공개 폴더에서 파일 목록 가져오기"""
    print("구글 드라이브 파일 목록 수집 중...")
    
    # 방법: 구글 드라이브 폴더 HTML 파싱
    url = f"https://drive.google.com/drive/folders/{FOLDER_ID}"
    
    try:
        req = urllib.request.Request(url, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        with urllib.request.urlopen(req, timeout=15) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        
        # 파일 ID 패턴 추출
        # 구글 드라이브 파일 ID는 33자 영숫자
        pattern = r'"([\w-]{28,})".*?"([^"]+\.pdf)"'
        matches = re.findall(pattern, html)
        
        ids = {}
        for fid, fname in matches:
            if len(fid) > 25 and fname.endswith('.pdf'):
                ids[fname] = fid
        
        return ids
    except Exception as e:
        print(f"자동 수집 실패: {e}")
        return {}

def save_ids(ids):
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(ids, f, ensure_ascii=False, indent=2)
    print(f"저장 완료: {len(ids)}개 파일 ID → {OUTPUT_JSON}")

if __name__ == '__main__':
    ids = fetch_file_ids()
    
    if not ids:
        print("\n자동 수집이 안 됩니다.")
        print("수동으로 파일 ID를 입력하거나")
        print("generate_html.py 를 실행하면")
        print("드라이브 폴더 링크로 연결됩니다.\n")
    else:
        save_ids(ids)
    
    # generate_html.py 실행
    import subprocess
    subprocess.run([sys.executable, os.path.join(SCAN_FOLDER, 'generate_html.py'), 'once'])
