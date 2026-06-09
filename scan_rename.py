"""
SCAN_ 파일 AI 자동 분석 + 파일명 변경기 v2
- 문서 내용 분석 → 기한 날짜 추출 → 파일명에 포함
- 형식: YYYYMMDD_회사명_문서종류_요약_기한YYYYMMDD.pdf
- 기한 7일 이내 → 🔴 긴급 / 14일 이내 → 🟠 주의 자동 표시
"""

import os, sys, json, base64, re
from datetime import datetime, date

SCAN_FOLDER = r"C:\Users\User\Desktop\스캔"
API_KEY     = "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"
API_URL     = "https://api.anthropic.com/v1/messages"
MODEL       = "claude-opus-4-5"

PROMPT = """이 법률/업무 문서를 분석해서 아래 JSON 형식으로만 답하세요. 다른 말은 절대 하지 마세요.

{{
  "date": "문서 작성일 또는 스캔일 YYYYMMDD (없으면 오늘 {today})",
  "company": "회사/기관명 (에스티엔미디어/에스티엔씨엘오미디어/에스티엔/이강영/플래닛/이창규드림/기타 중 하나)",
  "doctype": "문서종류 (소장/내용증명/회생계획안/회생채권조사확정재판/인사급여/내부보고/공문신청/계약서/입찰/기타 중 하나)",
  "summary": "핵심내용 10자 이내 (예: 대여금청구, 월간보고서, 광고계약)",
  "category": "소송 또는 업무",
  "deadline": "답변기한/제출기한/출석기한 날짜 YYYYMMDD. 없으면 null",
  "deadline_type": "기한 종류 (답변기한/제출기한/출석기한/이행기한/null 중 하나)"
}}

오늘: {today}
주의: deadline은 문서에 명시된 법정기한, 답변기한, 제출기한, 출석일자 등을 찾아주세요.
"""

def install_libs():
    import subprocess
    try:
        import fitz
    except ImportError:
        print("  PyMuPDF 설치 중...")
        subprocess.run([sys.executable, "-m", "pip", "install", "PyMuPDF", "-q"], check=True)

def pdf_to_image_base64(pdf_path):
    import fitz
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
    img_bytes = pix.tobytes("jpeg")
    doc.close()
    return base64.b64encode(img_bytes).decode("utf-8")

def analyze(img_b64):
    import urllib.request
    today = date.today().strftime("%Y%m%d")
    payload = {
        "model": MODEL,
        "max_tokens": 600,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": img_b64}},
                {"type": "text", "text": PROMPT.format(today=today)}
            ]
        }]
    }
    req = urllib.request.Request(
        API_URL,
        data=json.dumps(payload).encode(),
        headers={"Content-Type":"application/json","x-api-key":API_KEY,"anthropic-version":"2023-06-01"},
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode())
    text = data["content"][0]["text"].strip()
    m = re.search(r'\{.*\}', text, re.DOTALL)
    if m: return json.loads(m.group())
    raise ValueError(f"JSON 파싱 실패: {text}")

def urgency_label(deadline_str):
    """기한까지 남은 일수로 긴급도 계산"""
    if not deadline_str: return ""
    try:
        dl = datetime.strptime(str(deadline_str), "%Y%m%d").date()
        days = (dl - date.today()).days
        if days <= 0:  return "🔴 기한초과"
        if days <= 7:  return f"🔴 긴급 (D-{days})"
        if days <= 14: return f"🟠 주의 (D-{days})"
        return f"✅ D-{days}"
    except:
        return ""

def make_name(info):
    d        = re.sub(r'[^\d]','', str(info.get("date","") or ""))[:8] or date.today().strftime("%Y%m%d")
    company  = (info.get("company","기타") or "기타").replace(" ","")
    doctype  = (info.get("doctype","기타") or "기타").replace(" ","")
    summary  = (info.get("summary","") or "").replace(" ","").replace("/","_")[:15]
    deadline = re.sub(r'[^\d]','', str(info.get("deadline","") or ""))[:8]

    name = f"{d}_{company}_{doctype}"
    if summary:  name += f"_{summary}"
    if deadline: name += f"_기한{deadline}"
    name += ".pdf"
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def process():
    files = [f for f in os.listdir(SCAN_FOLDER)
             if f.upper().startswith("SCAN_") and f.lower().endswith(".pdf")]

    if not files:
        print("✅ 변경할 SCAN_ 파일이 없습니다.")
        return

    print(f"\n총 {len(files)}개 SCAN_ 파일 발견\n{'─'*60}")
    results = []

    for i, fname in enumerate(files, 1):
        fpath = os.path.join(SCAN_FOLDER, fname)
        print(f"[{i}/{len(files)}] {fname}")
        try:
            print("  → 이미지 변환 중...")
            img = pdf_to_image_base64(fpath)
            print("  → AI 분석 중...")
            info = analyze(img)

            new_name = make_name(info)
            new_path = os.path.join(SCAN_FOLDER, new_name)

            # 중복 처리
            if os.path.exists(new_path) and new_path != fpath:
                base, ext = os.path.splitext(new_name)
                new_name = f"{base}_2{ext}"
                new_path = os.path.join(SCAN_FOLDER, new_name)

            dl_label = urgency_label(info.get("deadline"))
            print(f"  ✅ → {new_name}")
            if dl_label: print(f"     기한: {info.get('deadline')} {dl_label}")
            print(f"     ({info.get('category')} / {info.get('doctype')} / {info.get('company')})\n")

            results.append({"old":fname,"new":new_name,"path":fpath,"new_path":new_path,"info":info,"dl":dl_label})
        except Exception as e:
            print(f"  ❌ 오류: {e}\n")
            results.append({"old":fname,"new":None,"error":str(e)})

    ok  = [r for r in results if r.get("new")]
    err = [r for r in results if not r.get("new")]

    print(f"\n{'='*60}")
    print(f"  변경 예정: {len(ok)}개 / 오류: {len(err)}개")
    print(f"{'='*60}")
    for r in ok:
        print(f"  {r['old']}")
        print(f"  → {r['new']}  {r.get('dl','')}\n")

    if not ok:
        input("\n아무 키나 누르면 종료..."); return

    ans = input(f"\n{len(ok)}개 파일을 변경하시겠습니까? (Y/N): ").strip().upper()
    if ans == "Y":
        for r in ok:
            try:
                os.rename(r["path"], r["new_path"])
                print(f"  ✅ {r['new']}")
            except Exception as e:
                print(f"  ❌ {r['old']}: {e}")
        print("\n✅ 완료! generate_html.py 실행해서 스캔관리대장을 갱신하세요.")
        print("   → python generate_html.py once")
    else:
        print("취소됐습니다.")

if __name__ == "__main__":
    print("="*60)
    print("  SCAN_ AI 분석 + 파일명 변경기 v2")
    print("  (기한 날짜 자동 추출 → 긴급/주의 자동 표시)")
    print("="*60)
    install_libs()
    process()
    input("\n아무 키나 누르면 종료...")
