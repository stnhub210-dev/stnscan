"""
PDF 파일명 자동 변경 스크립트
- API 키를 api_key.txt 에서 읽음 (CMD 깨짐 문제 해결)
- 텍스트 PDF + 스캔 이미지 PDF 모두 처리 (Claude Vision)
- 파일명을 내용에 맞게 자동 변경
"""

import os, sys, re, json, time, base64, shutil
from pathlib import Path

# ── 설정
SCAN_FOLDER  = os.path.join(os.path.expanduser("~"), "Desktop", "스캔")
KEY_FILE     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "api_key.txt")
DONE_FILE    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "renamed_done.json")
LOG_FILE     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rename_log.txt")

def log(msg):
    ts = time.strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def load_api_key():
    if not os.path.exists(KEY_FILE):
        # 키 파일 없으면 템플릿 생성
        with open(KEY_FILE, "w", encoding="utf-8") as f:
            f.write("여기에_API_키_붙여넣기\n")
        print(f"\n❌ API 키 파일이 없어서 생성했습니다.")
        print(f"   {KEY_FILE}")
        print(f"   파일을 메모장으로 열고 API 키를 붙여넣은 후 다시 실행하세요.\n")
        input("Enter 눌러 종료...")
        sys.exit(1)
    
    with open(KEY_FILE, "r", encoding="utf-8") as f:
        key = f.read().strip()
    
    if not key or key == "여기에_API_키_붙여넣기" or not key.startswith("sk-"):
        print(f"\n❌ api_key.txt 에 올바른 API 키를 입력해주세요.")
        print(f"   현재값: {key[:20]}...")
        input("Enter 눌러 종료...")
        sys.exit(1)
    return key

def load_done():
    if os.path.exists(DONE_FILE):
        with open(DONE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_done(db):
    with open(DONE_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)

def pdf_to_base64_image(pdf_path):
    """PDF 첫 페이지를 이미지로 변환 후 base64"""
    try:
        from pdf2image import convert_from_path
        imgs = convert_from_path(pdf_path, first_page=1, last_page=1, dpi=150)
        if not imgs:
            return None
        import io
        buf = io.BytesIO()
        imgs[0].save(buf, format="JPEG", quality=85)
        return base64.b64encode(buf.getvalue()).decode()
    except Exception as e:
        log(f"  이미지 변환 실패: {e}")
        return None

def extract_text(pdf_path):
    """텍스트 추출 시도"""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages[:3]:
                t = page.extract_text() or ""
                text += t
            if text.strip():
                return text[:3000]
    except:
        pass
    return ""

def ask_claude(api_key, filename, text="", img_b64=None):
    """Claude API 호출 → 새 파일명 반환"""
    import urllib.request
    
    content = []
    
    if img_b64:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64}
        })
    
    prompt = f"""이 PDF 문서의 내용을 보고, 내용에 맞는 한국어 파일명을 제안해주세요.

현재 파일명: {filename}
{"추출된 텍스트: " + text[:500] if text else ""}

규칙:
1. 형식: 날짜_기관명_문서종류 (예: 2026-03-15_인천지법_결정문)
2. 날짜는 문서에 있는 날짜 사용 (없으면 생략)
3. 특수문자 사용 금지 (/, \\, :, *, ?, ", <, >, | 제외)
4. 50자 이내
5. 파일명만 응답 (확장자 제외, 다른 설명 없이)

예시:
- 2026-05-04_인천지법_에스티엔유치권결정문
- 2026-03_에스티엔미디어_송출료소송답변서
- 2026-01_이미정_연봉계약서"""

    content.append({"type": "text", "text": prompt})
    
    body = json.dumps({
        "model": "claude-sonnet-4-5",
        "max_tokens": 200,
        "messages": [{"role": "user", "content": content}]
    }).encode("utf-8")
    
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=body,
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01"
        }
    )
    
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read())
        return data["content"][0]["text"].strip()

def safe_rename(old_path, new_name):
    """안전한 파일명 변경"""
    folder = os.path.dirname(old_path)
    # 특수문자 제거
    safe = re.sub(r'[\\/:*?"<>|]', "_", new_name)
    safe = safe.strip(". ")[:60]
    new_path = os.path.join(folder, safe + ".pdf")
    
    # 충돌 방지
    counter = 1
    while os.path.exists(new_path) and new_path != old_path:
        new_path = os.path.join(folder, f"{safe}_{counter}.pdf")
        counter += 1
    
    if new_path != old_path:
        os.rename(old_path, new_path)
    return new_path

def main():
    print("=" * 55)
    print("  PDF 파일명 자동 변경 (내용 기반)")
    print("=" * 55)
    
    api_key = load_api_key()
    log(f"API 키 로드 완료 ({api_key[:15]}...)")
    
    pdfs = sorted(Path(SCAN_FOLDER).glob("*.pdf"))
    if not pdfs:
        print(f"\n❌ PDF 파일 없음: {SCAN_FOLDER}")
        input("Enter 눌러 종료...")
        return
    
    done_db = load_done()
    
    # 처리 대상 필터
    targets = []
    for p in pdfs:
        fname = p.name
        # 이미 의미있는 이름 (SCAN_, Untitled_ 아닌 것) → 선택적 처리
        is_scan = fname.startswith("SCAN_") or fname.startswith("Untitled_")
        is_done = fname in done_db
        targets.append((p, is_scan, is_done))
    
    scan_count   = sum(1 for _, s, d in targets if s and not d)
    named_count  = sum(1 for _, s, d in targets if not s and not d)
    
    print(f"\n📂 폴더: {SCAN_FOLDER}")
    print(f"   전체 PDF: {len(pdfs)}개")
    print(f"   SCAN_* (이름변경필요): {scan_count}개")
    print(f"   기타 파일: {named_count}개")
    print(f"   처리완료(건너뜀): {sum(1 for _,_,d in targets if d)}개")
    
    print("\n처리할 파일을 선택하세요:")
    print("  1. SCAN_* 파일만 (스캔 파일)")
    print("  2. 전체 파일 모두")
    print("  3. 종료")
    
    choice = input("\n선택 (1/2/3): ").strip()
    if choice == "3":
        return
    
    process_list = []
    for p, is_scan, is_done in targets:
        if is_done:
            continue
        if choice == "1" and not is_scan:
            continue
        process_list.append(p)
    
    if not process_list:
        print("\n처리할 파일이 없습니다.")
        input("Enter 눌러 종료...")
        return
    
    print(f"\n▶ {len(process_list)}개 파일 처리 시작\n")
    success = 0
    
    for i, pdf_path in enumerate(process_list, 1):
        fname = pdf_path.name
        print(f"[{i}/{len(process_list)}] {fname[:45]}")
        
        try:
            # 텍스트 추출 시도
            text = extract_text(str(pdf_path))
            img_b64 = None
            
            if not text.strip():
                print("  → 이미지 PDF, 비전 분석 중...")
                img_b64 = pdf_to_base64_image(str(pdf_path))
            else:
                print("  → 텍스트 추출 성공")
            
            if not text and not img_b64:
                print("  → 건너뜀 (내용 읽기 실패)")
                continue
            
            # Claude 분석
            new_name = ask_claude(api_key, fname, text, img_b64)
            new_name = new_name.replace(".pdf", "").strip()
            
            # 파일명 변경
            new_path = safe_rename(str(pdf_path), new_name)
            new_fname = os.path.basename(new_path)
            
            print(f"  ✅ {fname[:30]} → {new_fname[:40]}")
            done_db[new_fname] = {"original": fname, "renamed_at": time.strftime("%Y-%m-%d %H:%M")}
            save_done(done_db)
            log(f"변경: {fname} → {new_fname}")
            success += 1
            
            time.sleep(0.3)  # API 속도 제한
            
        except Exception as e:
            print(f"  ❌ 오류: {e}")
            log(f"오류 [{fname}]: {e}")
    
    print(f"\n{'='*55}")
    print(f"  완료: {success}/{len(process_list)}개 변경")
    print(f"  로그: {LOG_FILE}")
    print(f"{'='*55}")
    input("\nEnter 눌러 종료...")

if __name__ == "__main__":
    main()
