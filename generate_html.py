"""
스캔관리대장 HTML 생성기 - 구글 드라이브 API 연동 최종버전
"""
import os, re, json, urllib.request, urllib.parse, time, sys
from datetime import datetime, timedelta

SCAN_FOLDER = r"C:\Users\User\Desktop\스캔"
OUTPUT_HTML = os.path.join(SCAN_FOLDER, "스캔관리대장.html")
FOLDER_ID   = "1hT9xdNEawnkrZ78p26Kcjlm-1wdR5x-X"
FOLDER_URL  = f"https://drive.google.com/drive/folders/{FOLDER_ID}"
API_KEY     = "여기에복사한키입력"  # ← 여기에 API 키 입력
ID_CACHE    = os.path.join(SCAN_FOLDER, "file_ids.json")

def load_lawsuit_excel():
    """소송목록.xlsx 에서 소송 문서 목록 읽기"""
    excel_path = os.path.join(SCAN_FOLDER, "소송목록.xlsx")
    result = {}
    if not os.path.exists(excel_path):
        print("  소송목록.xlsx 없음 → 기본 목록 사용")
        return None
    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        ws = wb["소송목록"]
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row[7]: continue  # 파일명 없으면 건너뜀
            fname = str(row[7]).strip()
            company = str(row[1] or "").strip()
            case_name = str(row[2] or "").strip()
            case_no = str(row[3] or "").strip()
            doc_type = str(row[4] or "").strip()
            urgency = str(row[6] or "일반").strip()
            if fname and case_name:
                result[fname] = (company, case_name, case_no, doc_type, urgency)
        print(f"  소송목록.xlsx 로드 완료: {len(result)}건")
        return result
    except Exception as e:
        print(f"  소송목록.xlsx 읽기 오류: {e}")
        return None

LAWSUIT_FILES_DEFAULT = {
    "(26.05.04)결정_인천지법(25라5671)_에스티엔 유치권.pdf": ("인천지방법원","에스티엔 유치권 결정","25라5671","결정문","긴급"),
    "260406_70거4454.pdf":("법원","70거4454 사건","70거4454","법원서류","주의"),
    "김윤석 소송.pdf":("김윤석","김윤석 소송","","소송관련","주의"),
    "답변서_2026가단104265_(어뮤즈).pdf":("어뮤즈","2026가단104265 답변서","2026가단104265","답변서","주의"),
    "부동산명령.pdf":("법원","부동산 명령","","법원명령","긴급"),
    "이강영 특가중.pdf":("이강영","특정범죄가중처벌 사건","","형사사건","긴급"),
    "이미정 개인회생 이체내역.pdf":("이미정","개인회생 이체내역","","회생관련","주의"),
    "이미정 개인회생 채권자목록.pdf":("이미정","개인회생 채권자목록","","회생관련","주의"),
    "이미정 개인회생 판결문.pdf":("이미정","개인회생 판결문","","판결문","주의"),
    "중부모범 답변서.pdf":("중부모범","중부모범 답변서","","답변서","주의"),
    "중부모범답변.pdf":("중부모범","중부모범 답변","","답변서","주의"),
    "에스티엔미디어 송출료소송.pdf":("에스티엔미디어","송출료 소송","","소송관련","주의"),
    "에스티엔미디어 중부모범.pdf":("에스티엔미디어/중부모범","중부모범 관련 서류","","소송관련","주의"),
    "차주화 진술서 (2).pdf":("차주화","차주화 진술서","","진술서","일반"),
    "사실확인서.pdf":("","사실확인서","","확인서류","일반"),
    "부평테크노타워_관리비.pdf":("부평테크노타워/에스티엔","관리비 미납 최고서 (2,080,246원)","테크노타워관리-26-05-06","최고서/내용증명","긴급"),
}

def get_lawsuit_files():
    loaded = load_lawsuit_excel()
    return loaded if loaded is not None else LAWSUIT_FILES_DEFAULT

COMPANIES = [
    ("에스티엔미디어", ["에스티엔미디어","에스티엔방송","STN","stn"]),
    ("에스티엔뉴스",   ["에스티엔뉴스","stnnews"]),
    ("에스티엔",       ["에스티엔등기","에스티엔 3월","에스티엔 매출","260309_에스티엔","260320_에스티엔","260401","260403_에스티엔","260406_에스티엔","260413_에스티엔","260422_에스티엔","에스티엔"]),
    ("이강영",         ["이강영"]),
    ("플래닛네이처미디어", ["플래닛네이처미디어","플레닛네이처미디어","플래닛네이처","플레닛네이처","플래닛부동산","플레닛"]),
]

def fetch_drive_ids():
    """구글 드라이브 API로 파일 목록 가져오기"""
    print("  구글 드라이브 파일 목록 수집 중...")
    ids = {}
    page_token = None
    
    while True:
        params = {
            "q": f"'{FOLDER_ID}' in parents and trashed=false",
            "fields": "nextPageToken,files(id,name)",
            "pageSize": "1000",
            "key": API_KEY,
        }
        if page_token:
            params["pageToken"] = page_token
        
        url = "https://www.googleapis.com/drive/v3/files?" + urllib.parse.urlencode(params)
        
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read())
            
            for f in data.get("files", []):
                if f["name"].lower().endswith(".pdf"):
                    ids[f["name"]] = f["id"]
            
            page_token = data.get("nextPageToken")
            if not page_token:
                break
                
        except Exception as e:
            print(f"  API 오류: {e}")
            break
    
    if ids:
        with open(ID_CACHE, "w", encoding="utf-8") as f:
            json.dump(ids, f, ensure_ascii=False, indent=2)
        print(f"  파일 ID {len(ids)}개 수집 완료")
    return ids

def load_ids():
    """캐시된 파일 ID 로드 (없으면 API 호출)"""
    if os.path.exists(ID_CACHE):
        mtime = os.path.getmtime(ID_CACHE)
        # 1시간마다 갱신
        if time.time() - mtime < 3600:
            with open(ID_CACHE, "r", encoding="utf-8") as f:
                return json.load(f)
    return fetch_drive_ids()

def get_date(fname):
    m = re.search(r'SCAN_(\d{4})(\d{2})(\d{2})_', fname)
    if m: return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r'\((\d{2})\.(\d{2})\.(\d{2})\)', fname)
    if m: return f"20{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r'^(\d{2})(\d{2})(\d{2})_', fname)
    if m: return f"20{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r'(202\d)[.\-](\d{1,2})[.\-](\d{1,2})', fname)
    if m: return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    return "2026-05-17"

def get_company(fname):
    f = fname.replace(".pdf","")
    for name, keywords in COMPANIES:
        for kw in keywords:
            if kw in f: return name
    for p in ["이미정","김경아","김준영","차주화","정기학","심창호","경민창","김윤석"]:
        if p in f: return p
    return "기타"

def get_company_tags(fname):
    f = fname.replace(".pdf","")
    tags = []
    for name, keywords in COMPANIES:
        for kw in keywords:
            if kw in f: tags.append(name); break
    return tags

def get_doctype(fname):
    f = fname.replace(".pdf","")
    if "내용증명" in f: return "✉️ 내용증명"
    if any(k in f for k in ["계약서","협약서","양도","양수"]): return "📝 계약서"
    if any(k in f for k in ["급여","연봉","용역비","품의","채용","이력서"]): return "👤 인사/급여"
    if any(k in f for k in ["국세","지방세","세무소","과태료","건강보험"]): return "🧾 세금/공과"
    if any(k in f for k in ["등기부","등기권리","사업자등록","법인인감","인감증명"]): return "🏢 법인서류"
    if any(k in f for k in ["보고서","월간","자금수지","매출"]): return "📊 내부보고"
    if any(k in f for k in ["공문","서약서","신청서","확약서","동의서"]): return "📋 공문/신청"
    if any(k in f for k in ["영수증","수령","청구","약정금"]): return "💰 영수/청구"
    if any(k in f for k in ["SCAN","Untitled"]): return "🖨️ 스캔문서"
    if any(k in f for k in ["임대차","전대차"]): return "🏠 임대차"
    if any(k in f for k in ["소송","답변서","판결","결정","명령","회생","진술"]): return "⚖️ 소송관련"
    return "📄 기타"

def is_recent(d, days=7):
    try: return (datetime.now()-datetime.strptime(d,"%Y-%m-%d")).days<=days
    except: return False

def new_b(d): return ' <span class="nb">NEW</span>' if is_recent(d) else ""
def ra(d): return ' data-recent="true"' if is_recent(d) else ""
def ubadge(u): return f'<span class="badge {u}">{u}</span>' if u else ""

def generate(output_path=OUTPUT_HTML):
    LAWSUIT_FILES = get_lawsuit_files()
    print(f"[{datetime.now().strftime('%H:%M:%S')}] HTML 생성 시작...")
    
    # 파일 ID 로드
    ids = load_ids()
    
    # 로컬 파일 목록
    try:
        pdfs = sorted([f for f in os.listdir(SCAN_FOLDER) if f.lower().endswith('.pdf')])
    except:
        pdfs = list(LAWSUIT_FILES.keys())

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    # 전체 문서 구성
    all_docs = []
    for fname in pdfs:
        fid = ids.get(fname, "")
        link = f"https://drive.google.com/file/d/{fid}/view" if fid else FOLDER_URL
        date = get_date(fname)
        all_docs.append({
            "fname": fname, "date": date, "fid": fid, "link": link,
            "company": get_company(fname),
            "doctype": get_doctype(fname),
            "tags": get_company_tags(fname),
            "is_lawsuit": fname in LAWSUIT_FILES,
            "urgency": LAWSUIT_FILES.get(fname, ("","","","",""))[4],
            "info": LAWSUIT_FILES.get(fname),
        })

    all_docs.sort(key=lambda x: x["date"], reverse=True)
    lawsuit_docs = [d for d in all_docs if d["is_lawsuit"]]
    work_docs    = [d for d in all_docs if not d["is_lawsuit"]]
    urgent_n  = sum(1 for d in lawsuit_docs if d["urgency"]=="긴급")
    caution_n = sum(1 for d in lawsuit_docs if d["urgency"]=="주의")
    recent_n  = sum(1 for d in all_docs if is_recent(d["date"]))
    linked_n  = sum(1 for d in all_docs if d["fid"])
    co_counts = {n: sum(1 for d in all_docs if n in d["tags"]) for n,_ in COMPANIES}

    # 소송 행
    l_rows = ""
    for i,d in enumerate(lawsuit_docs,1):
        info = d["info"]
        l_rows += f"""<tr{ra(d['date'])}>
<td class="c">{i}</td><td class="c dc">{d['date']}{new_b(d['date'])}</td>
<td>{info[0]}</td><td>{info[1]}</td><td class="c">{info[2]}</td>
<td class="c">{info[3]}</td><td class="c">{ubadge(d['urgency'])}</td>
<td><a href="{d['link']}" target="_blank" class="fl">📄 {d['fname']}</a></td>
</tr>\n"""

    # 업무 행
    w_rows = ""
    for i,d in enumerate(work_docs,1):
        w_rows += f"""<tr{ra(d['date'])}>
<td class="c">{i}</td><td class="c dc">{d['date']}{new_b(d['date'])}</td>
<td class="c">{d['company']}</td><td class="c">{d['doctype']}</td>
<td><a href="{d['link']}" target="_blank" class="fl">📄 {d['fname']}</a></td>
</tr>\n"""

    # 회사별 탭/패널
    co_tabs = ""
    co_panels = ""
    for cname,_ in COMPANIES:
        cnt = co_counts[cname]
        if cnt == 0: continue
        short = cname.replace("플래닛네이처미디어","플래닛").replace("에스티엔미디어","STN미디어").replace("에스티엔뉴스","STN뉴스").replace("에스티엔","STN")
        co_tabs += f'<button class="tab co-tab" id="tab-co-{cname}" onclick="showTab(\'co-{cname}\',this)">{short} ({cnt})</button>\n'
        docs = [d for d in all_docs if cname in d["tags"]]
        rows = ""
        for i,d in enumerate(docs,1):
            lm = f' {ubadge("주의")}' if d["is_lawsuit"] else ""
            rows += f"""<tr{ra(d['date'])}>
<td class="c">{i}</td><td class="c dc">{d['date']}{new_b(d['date'])}</td>
<td class="c">{d['doctype']}{lm}</td>
<td><a href="{d['link']}" target="_blank" class="fl">📄 {d['fname']}</a></td>
</tr>\n"""
        co_panels += f"""<div id="panel-co-{cname}" class="panel">
<table class="co-tbl"><thead><tr><th>No</th><th>스캔일자</th><th>문서종류</th><th>파일 열기 (구글 드라이브)</th></tr></thead>
<tbody>{rows}</tbody></table></div>\n"""

    link_notice = "" if linked_n == len(pdfs) else f'<div class="notice">⚠️ {len(pdfs)-linked_n}개 파일 링크 없음 → <a href="{FOLDER_URL}" target="_blank">드라이브 폴더에서 직접 찾기</a></div>'

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta http-equiv="refresh" content="300">
<title>스캔관리대장</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"맑은 고딕",sans-serif;background:#f0f2f5;padding:20px;font-size:13px}}
h1{{color:#1F4E79;font-size:22px;margin-bottom:14px}}
.summary{{display:flex;gap:10px;margin-bottom:14px;flex-wrap:wrap}}
.card{{background:white;border-radius:10px;padding:12px 18px;box-shadow:0 2px 6px rgba(0,0,0,.1);text-align:center;min-width:80px}}
.card .num{{font-size:26px;font-weight:bold;color:#1F4E79}}
.card .label{{font-size:11px;color:#666;margin-top:3px}}
.card.red .num{{color:#e53935}}.card.orange .num{{color:#FF9900}}.card.green .num{{color:#2e7d32}}.card.blue .num{{color:#1565c0}}
.toolbar{{display:flex;gap:8px;margin-bottom:10px;align-items:center;flex-wrap:wrap}}
.toolbar input{{padding:7px 12px;border:1px solid #ccc;border-radius:6px;font-size:13px;width:260px;font-family:"맑은 고딕"}}
.toolbar select{{padding:7px 9px;border:1px solid #ccc;border-radius:6px;font-size:13px;font-family:"맑은 고딕"}}
.btn{{padding:7px 14px;border:none;border-radius:6px;cursor:pointer;font-family:"맑은 고딕";font-size:13px;font-weight:bold}}
.btn-reset{{background:#6c757d;color:white}}
.btn-recent{{background:#2e7d32;color:white}}
.btn-recent .cnt{{background:white;color:#2e7d32;border-radius:10px;padding:1px 6px;font-size:11px;margin-left:4px}}
.btn-drive{{background:#4285f4;color:white;text-decoration:none;display:inline-flex;align-items:center}}
.btn-refresh{{background:#1F4E79;color:white;margin-left:auto}}
.btn:hover{{opacity:.85}}
.tabs{{display:flex;gap:3px;flex-wrap:wrap;margin-bottom:0}}
.tab{{padding:9px 16px;background:#ddd;border-radius:8px 8px 0 0;cursor:pointer;font-weight:bold;font-size:12px;border:none;font-family:"맑은 고딕";white-space:nowrap}}
.tab.active{{background:#1F4E79;color:white}}
.tab.green.active{{background:#375623}}
.tab.co-tab.active{{background:#5c3d91;color:white}}
.panel{{display:none}}.panel.active{{display:block}}
table{{width:100%;border-collapse:collapse;background:white;box-shadow:0 2px 6px rgba(0,0,0,.1)}}
th{{padding:9px 11px;text-align:center;border-bottom:2px solid #dee2e6;white-space:nowrap}}
.lawsuit-tbl th{{background:#1F4E79;color:white}}
.work-tbl th{{background:#375623;color:white}}
.co-tbl th{{background:#5c3d91;color:white}}
td{{padding:8px 11px;border-bottom:1px solid #e9ecef;vertical-align:middle}}
tr:nth-child(even) td{{background:#fafafa}}
tr:hover td{{background:#e8f4fd!important}}
tr[data-recent="true"] td{{background:#f0fff4}}
tr[data-recent="true"]:hover td{{background:#d4edda!important}}
.c{{text-align:center}}.dc{{white-space:nowrap;font-weight:500}}
a.fl{{color:#0563C1;text-decoration:none}}
a.fl:hover{{text-decoration:underline}}
.badge{{display:inline-block;padding:2px 9px;border-radius:10px;font-size:11px;font-weight:bold}}
.긴급{{background:#e53935;color:white}}.주의{{background:#FF9900;color:white}}.일반{{background:#6c757d;color:white}}
.nb{{display:inline-block;background:#2e7d32;color:white;font-size:10px;padding:1px 5px;border-radius:7px;margin-left:4px;font-weight:bold}}
.hidden{{display:none!important}}
.rbb{{background:#d4edda;border:1px solid #28a745;border-radius:6px;padding:7px 12px;margin-bottom:8px;color:#155724;font-size:12px;display:none}}
.rbb.show{{display:block}}
.notice{{background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:8px 14px;margin-bottom:10px;font-size:12px}}
.count-info{{color:#666;font-size:12px}}
.updated{{color:#999;font-size:11px;margin-top:10px;text-align:right}}
</style>
</head>
<body>
<h1>📂 스캔관리대장</h1>
<div class="summary">
  <div class="card"><div class="num">{len(lawsuit_docs)}</div><div class="label">소송문서</div></div>
  <div class="card red"><div class="num">{urgent_n}</div><div class="label">🔴 긴급</div></div>
  <div class="card orange"><div class="num">{caution_n}</div><div class="label">🟡 주의</div></div>
  <div class="card"><div class="num">{len(work_docs)}</div><div class="label">업무문서</div></div>
  <div class="card green"><div class="num">{recent_n}</div><div class="label">📅 최근7일</div></div>
  <div class="card blue"><div class="num">{linked_n}</div><div class="label">☁️ 드라이브연결</div></div>
  <div class="card"><div class="num">{len(pdfs)}</div><div class="label">전체</div></div>
</div>
{link_notice}
<div class="toolbar">
  <input type="text" id="si" placeholder="🔍 파일명, 회사명, 사건명..." oninput="doSearch()">
  <select id="sc" onchange="doSearch()">
    <option value="all">전체탭</option>
    <option value="lawsuit">소송만</option>
    <option value="work">업무만</option>
  </select>
  <button class="btn btn-reset" onclick="clearAll()">초기화</button>
  <button class="btn btn-recent" id="rb" onclick="toggleRecent()">📅 최근 7일 <span class="cnt">{recent_n}</span></button>
  <a href="{FOLDER_URL}" target="_blank" class="btn btn-drive">☁️ 드라이브 폴더</a>
  <span id="ci" class="count-info"></span>
  <button class="btn btn-refresh" onclick="location.reload()">🔄 새로고침</button>
</div>
<div id="rbb" class="rbb">📅 최근 7일 이내 문서만 표시 중 &nbsp;<a href="#" onclick="clearAll();return false;">전체 보기</a></div>
<div class="tabs">
  <button class="tab active" id="tab-lawsuit" onclick="showTab('lawsuit',this)">⚖️ 소송관리 ({len(lawsuit_docs)})</button>
  <button class="tab green" id="tab-work" onclick="showTab('work',this)">📋 업무문서 ({len(work_docs)})</button>
  {co_tabs}
</div>
<div id="panel-lawsuit" class="panel active">
<table class="lawsuit-tbl"><thead><tr><th>No</th><th>스캔일자</th><th>회사(기관)명</th><th>사건명</th><th>사건번호</th><th>문서종류</th><th>긴급도</th><th>파일 열기</th></tr></thead>
<tbody id="lb">{l_rows}</tbody></table></div>
<div id="panel-work" class="panel">
<table class="work-tbl"><thead><tr><th>No</th><th>스캔일자</th><th>회사(기관)명</th><th>문서종류</th><th>파일 열기</th></tr></thead>
<tbody id="wb">{w_rows}</tbody></table></div>
{co_panels}
<p class="updated">※ 갱신: {now_str} | ☁️ 구글 드라이브 연동 | 드라이브 링크: {linked_n}/{len(pdfs)}개</p>
<script>
let rm=false;
function showTab(n,btn){{
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('panel-'+n).classList.add('active');
  btn.classList.add('active');
  applyF();
}}
function toggleRecent(){{
  rm=!rm;
  document.getElementById('rb').style.background=rm?'#155724':'#2e7d32';
  document.getElementById('rbb').classList.toggle('show',rm);
  applyF();
}}
function doSearch(){{rm=false;document.getElementById('rb').style.background='#2e7d32';document.getElementById('rbb').classList.remove('show');applyF();}}
function applyF(){{
  const q=document.getElementById('si').value.toLowerCase().trim();
  let v=0;
  document.querySelectorAll('tbody tr').forEach(tr=>{{
    const ok=(!q||tr.textContent.toLowerCase().includes(q))&&(!rm||tr.dataset.recent==='true');
    tr.classList.toggle('hidden',!ok);
    if(ok)v++;
  }});
  document.getElementById('ci').textContent=(q||rm)?`표시: ${{v}}건`:'';
}}
function clearAll(){{
  document.getElementById('si').value='';document.getElementById('sc').value='all';
  rm=false;document.getElementById('rb').style.background='#2e7d32';
  document.getElementById('rbb').classList.remove('show');
  document.getElementById('ci').textContent='';
  document.querySelectorAll('tr').forEach(tr=>tr.classList.remove('hidden'));
}}
</script>
</body></html>"""

    with open(output_path,'w',encoding='utf-8') as f:
        f.write(html)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 완료! ({len(pdfs)}개 파일, 드라이브 링크: {linked_n}개)")

if __name__=='__main__':
    if len(sys.argv)>1 and sys.argv[1]=='once':
        generate(); sys.exit(0)
    print("="*50)
    print("  스캔관리대장 자동 갱신 감시")
    print(f"  폴더: {SCAN_FOLDER}")
    print("  종료: Ctrl+C\n")
    prev=0
    while True:
        try:
            files=[f for f in os.listdir(SCAN_FOLDER) if f.lower().endswith('.pdf')]
            if len(files)!=prev:
                generate(); prev=len(files)
        except Exception as e: print(f"오류: {e}")
        time.sleep(5)
