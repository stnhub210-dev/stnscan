"""
스캔관리대장 자동 생성기 v3
- SCAN_ 파일명에서 날짜 정확 추출
- 회사명/문서종류 정확 감지
- 로컬 API 서버 (포트 8765) - 새로고침 버튼용
"""

import os, json, re, sys, time, subprocess, threading
from datetime import datetime, date
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.request

# ─────────── 설정 ───────────────────────────
SCAN_FOLDER = r"C:\Users\User\Desktop\스캔"
OUTPUT_HTML = os.path.join(SCAN_FOLDER, "스캔관리대장.html")
XLSX_PATH   = os.path.join(SCAN_FOLDER, "소송목록.xlsx")
FOLDER_ID   = "1hT9xdNEawnkrZ78p26Kcjlm-1wdR5x-X"
API_PORT    = 8765
# ────────────────────────────────────────────

def extract_date(fname, fpath=None):
    """파일명에서 유효한 날짜 추출. 없으면 파일 수정일 사용."""
    # SCAN_YYYYMMDD_ 형식 우선 처리
    m = re.match(r'SCAN_(\d{8})_', fname, re.IGNORECASE)
    if m:
        d = m.group(1)
        y, mo, day = int(d[:4]), int(d[4:6]), int(d[6:8])
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}"

    # 파일명 앞부분 YYYYMMDD_ 형식
    m = re.match(r'(\d{8})[_\-]', fname)
    if m:
        d = m.group(1)
        y, mo, day = int(d[:4]), int(d[4:6]), int(d[6:8])
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}"

    # 파일명 어디서든 YYYY.MM.DD 형식
    m = re.search(r'(\d{4})[.\-](\d{2})[.\-](\d{2})', fname)
    if m:
        y, mo, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}"

    # 파일명 어디서든 8자리 숫자
    for m in re.finditer(r'(\d{8})', fname):
        d = m.group(1)
        y, mo, day = int(d[:4]), int(d[4:6]), int(d[6:8])
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}"

    # 파일 수정일
    if fpath and os.path.exists(fpath):
        mtime = os.path.getmtime(fpath)
        return datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")

    return date.today().strftime("%Y-%m-%d")

def detect_company(fname):
    f = fname
    if any(k in f for k in ["에스티엔미디어","STN미디어","stnmedia","STNmedia"]): return "에스티엔미디어"
    if any(k in f for k in ["씨엘오미디어","CLO미디어","clomedia"]):              return "에스티엔씨엘오미디어"
    if any(k in f for k in ["이강영"]):                                           return "이강영"
    if any(k in f for k in ["플래닛","planet","Planet"]):                         return "플래닛"
    if any(k in f for k in ["이창규","드림"]):                                    return "이창규드림"
    if any(k in f for k in ["에스티엔","STN","stn"]):                             return "에스티엔"
    return "기타"

def detect_category(fname):
    lawsuit_kw = ["소장","소송","내용증명","회생","채권","법원","재판","확정","가소","가합","간회합","간이회생"]
    return "소송" if any(k in fname for k in lawsuit_kw) else "업무"

def detect_doctype(fname):
    if "내용증명"   in fname: return "내용증명"
    if "소장"       in fname: return "소장"
    if "회생계획"   in fname: return "회생계획안"
    if "확정재판"   in fname: return "회생채권조사확정재판"
    if "채권"       in fname: return "채권신고"
    if "회생"       in fname: return "회생관련"
    if any(k in fname for k in ["급여","임금","인사"]): return "인사/급여"
    if "보고서"     in fname or "보고"  in fname: return "내부보고"
    if any(k in fname for k in ["공문","신청"]): return "공문/신청"
    if "계약"       in fname: return "계약서"
    if "입찰"       in fname: return "입찰"
    if fname.upper().startswith("SCAN_"): return "스캔문서"
    return "기타"

def detect_urgency(fname):
    # A안: 파일명에 기한YYYYMMDD 포함된 경우 날짜 계산
    m = re.search(r'기한(\d{8})', fname)
    if m:
        try:
            deadline = datetime.strptime(m.group(1), "%Y%m%d").date()
            days_left = (deadline - date.today()).days
            if days_left <= 0:  return "긴급"   # 이미 지남
            if days_left <= 7:  return "긴급"   # 7일 이내
            if days_left <= 14: return "주의"   # 14일 이내
        except: pass

    # 키워드 긴급
    if any(k in fname for k in ["긴급","urgent","즉시","당일","가압류","집행"]):
        return "긴급"

    # 키워드 주의
    if any(k in fname for k in ["주의","중요"]):
        return "주의"

    # B안: 소송문서는 기본 주의
    if detect_category(fname) == "소송":
        return "주의"

    return ""

def load_lawsuit_data():
    try:
        import openpyxl
        wb = openpyxl.load_workbook(XLSX_PATH)
        ws = wb.active
        rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if r[0]]
        print(f"  소송목록.xlsx 로드 완료: {len(rows)}건")
        return rows
    except Exception as e:
        print(f"  소송목록.xlsx 로드 실패: {e}")
        return []

CLASSIFICATIONS = os.path.join(SCAN_FOLDER, "classifications.json")

def load_classifications(scan_folder=SCAN_FOLDER):
    path = os.path.join(scan_folder, "classifications.json")
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return {}

def save_classifications(data):
    with open(CLASSIFICATIONS, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def fetch_drive_ids(scan_folder=SCAN_FOLDER):
    try:
        jp = os.path.join(scan_folder, "file_ids.json")
        if os.path.exists(jp):
            with open(jp, encoding="utf-8") as f:
                return json.load(f)
    except: pass
    return {}

def make_link(fname, drive_ids):
    fid = drive_ids.get(fname)
    if fid: return f"https://drive.google.com/file/d/{fid}/view"
    # 로컬 파일 직접 열기
    local = os.path.join(SCAN_FOLDER, fname).replace("\\", "/")
    return f"file:///{local}"

# ── HTML 생성 ────────────────────────────────
def generate(scan_folder=SCAN_FOLDER, output_path=OUTPUT_HTML):
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] HTML 생성 시작...")
    drive_ids = fetch_drive_ids(scan_folder)
    load_lawsuit_data()

    pdfs = [f for f in os.listdir(scan_folder) if f.lower().endswith('.pdf')]

    today_str = date.today().strftime("%Y-%m-%d")

    classifications = load_classifications(scan_folder)
    docs = []
    for fname in pdfs:
        fpath = os.path.join(scan_folder, fname)
        override = classifications.get(fname, {})
        company  = override.get("company")  or detect_company(fname)
        category = override.get("category") or detect_category(fname)
        doctype  = override.get("doctype")  or detect_doctype(fname)
        docs.append({
            "fname":    fname,
            "date":     extract_date(fname, fpath),
            "company":  company,
            "category": category,
            "doctype":  doctype,
            "urgency":  detect_urgency(fname),
            "link":     make_link(fname, drive_ids),
        })

    # 날짜 내림차순 정렬
    docs.sort(key=lambda d: d["date"], reverse=True)

    def new_badge(d):
        return ' <span class="nbadge">NEW</span>' if d["date"] == today_str else ""

    def ub(u):
        if u == "긴급": return '<span class="b-red">긴급</span>'
        if u == "주의": return '<span class="b-org">주의</span>'
        return ""

    def rc(d):
        if d["date"] == today_str: return ' class="r-new"'
        if d["urgency"] == "긴급": return ' class="r-urg"'
        return ""

    def rows(lst):
        out = ""
        for i, d in enumerate(lst):
            icon = "⚖️" if d["category"] == "소송" else "📋"
            safe     = d["fname"].replace("'", "\\'")
            company  = d["company"].replace("'", "\\'")
            category = d["category"].replace("'", "\\'")
            doctype  = d["doctype"].replace("'", "\\'")
            fname_short = d["fname"][:30] + ("…" if len(d["fname"])>30 else "")
            out += (
                f'<tr{rc(d)}>'
                f'<td class="c">{i+1}</td>'
                f'<td class="c dc">{d["date"]}{new_badge(d)}</td>'
                f'<td class="c">{icon} {d["category"]}</td>'
                f'<td>{d["company"]}</td>'
                f'<td>{d["doctype"]}</td>'
                f'<td class="c">{ub(d["urgency"])}</td>'
                f'<td><a href="{d["link"]}" target="_blank" class="fl">📄 {fname_short}</a></td>'
                f'<td class="c"><button class="cbtn" onclick="openModal(&quot;{safe}&quot;,&quot;{company}&quot;,&quot;{category}&quot;,&quot;{doctype}&quot;)">✏️</button></td>'
                f'</tr>\n'
            )
        return out

    def filt(key, val): return [d for d in docs if d[key] == val]

    ld = filt("category","소송"); wd = filt("category","업무")
    sm = filt("company","에스티엔미디어"); sn = filt("company","에스티엔")
    lk = filt("company","이강영");         pl = filt("company","플래닛")
    urg_cnt  = sum(1 for d in docs if d["urgency"]=="긴급")
    note_cnt = sum(1 for d in docs if d["urgency"]=="주의")
    from datetime import timedelta
    week_ago = (date.today() - timedelta(days=7)).strftime("%Y-%m-%d")
    today_cnt= sum(1 for d in docs if d["date"] >= week_ago)
    drv_cnt  = len(drive_ids)
    updated  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def panel(pid, lst, active=""):
        return f'<div id="p-{pid}" class="panel {active}"><table><thead><tr><th>No</th><th>스캔일자</th><th>구분</th><th>회사(기관)명</th><th>문서종류</th><th>긴급도</th><th>파일 열기</th><th>분류</th></tr></thead><tbody>{rows(lst)}</tbody></table></div>'

    PW_HASH = "9998d4912395a5bd5dbd699f33725cdea181a3215f744318c03168c548f9b77b"

    html = f"""<!DOCTYPE html>
<html lang="ko"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>스캔관리대장</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Malgun Gothic',sans-serif;background:#f0f2f5;color:#333;font-size:14px}}

/* ── 잠금 화면 ── */
#lockScreen{{
  position:fixed;top:0;left:0;width:100%;height:100%;
  background:linear-gradient(135deg,#0d1b6e 0%,#1a237e 50%,#283593 100%);
  z-index:9999;display:flex;align-items:center;justify-content:center;
}}
#lockScreen.hide{{display:none}}
.lock-box{{
  background:rgba(255,255,255,0.06);backdrop-filter:blur(12px);
  border:1px solid rgba(255,255,255,0.15);border-radius:20px;
  padding:48px 40px;min-width:320px;max-width:380px;width:90%;
  text-align:center;box-shadow:0 24px 64px rgba(0,0,0,0.5);
}}
.lock-logo{{font-size:48px;margin-bottom:12px}}
.lock-title{{color:#fff;font-size:22px;font-weight:700;margin-bottom:4px;letter-spacing:-0.5px}}
.lock-sub{{color:rgba(255,255,255,0.5);font-size:12px;margin-bottom:32px}}
.lock-label{{color:rgba(255,255,255,0.7);font-size:12px;text-align:left;margin-bottom:6px;display:block}}
.lock-input{{
  width:100%;padding:13px 16px;border-radius:10px;border:1.5px solid rgba(255,255,255,0.2);
  background:rgba(255,255,255,0.08);color:#fff;font-size:16px;
  outline:none;letter-spacing:2px;text-align:center;
  transition:border .2s;
}}
.lock-input::placeholder{{color:rgba(255,255,255,0.3);letter-spacing:0}}
.lock-input:focus{{border-color:rgba(255,255,255,0.6);background:rgba(255,255,255,0.12)}}
.lock-btn{{
  width:100%;margin-top:16px;padding:14px;border:none;border-radius:10px;
  background:#e8f0fe;color:#1a237e;font-size:15px;font-weight:700;
  cursor:pointer;transition:all .2s;letter-spacing:0.5px;
}}
.lock-btn:hover{{background:#fff;transform:translateY(-1px);box-shadow:0 4px 16px rgba(0,0,0,0.2)}}
.lock-err{{color:#ff6b6b;font-size:12px;margin-top:10px;min-height:18px}}
.lock-notice{{color:rgba(255,255,255,0.3);font-size:11px;margin-top:24px;line-height:1.6}}
</style>
.hdr{{background:#1a237e;color:#fff;padding:14px 24px;display:flex;align-items:center;justify-content:space-between}}
.hdr h1{{font-size:19px}}
.hdr span{{font-size:11px;opacity:.7}}
.stats{{display:flex;gap:10px;padding:14px 24px;flex-wrap:wrap}}
.stat{{background:#fff;border-radius:10px;padding:12px 18px;min-width:100px;text-align:center;
       box-shadow:0 2px 6px rgba(0,0,0,.08);cursor:pointer;transition:all .15s}}
.stat:hover{{transform:translateY(-3px);box-shadow:0 6px 16px rgba(0,0,0,.15)}}
.stat .n{{font-size:26px;font-weight:700;color:#1a237e}}
.stat .l{{font-size:11px;color:#666;margin-top:3px}}
.stat.red .n{{color:#e53935}}.stat.org .n{{color:#ef6c00}}.stat.grn .n{{color:#2e7d32}}
.bar{{display:flex;gap:8px;padding:0 24px 12px;flex-wrap:wrap;align-items:center}}
.bar input{{flex:1;min-width:180px;padding:8px 13px;border:1px solid #ddd;border-radius:8px;font-size:13px}}
.btn{{padding:8px 15px;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:600}}
.b-gray{{background:#78909c;color:#fff}}.b-blue{{background:#1565c0;color:#fff}}
.b-teal{{background:#00897b;color:#fff;font-size:14px;padding:9px 18px}}
.btn:hover{{opacity:.85}}
#rstat{{font-size:12px;color:#666;margin-left:6px}}
.tabs{{display:flex;gap:4px;padding:0 24px;flex-wrap:wrap;border-bottom:2px solid #ddd}}
.tab{{padding:8px 14px;border:none;background:#e0e0e0;cursor:pointer;font-size:13px;color:#555;
      border-bottom:3px solid transparent;margin-bottom:-2px;border-radius:6px 6px 0 0;font-weight:600;transition:all .15s}}
.tab:hover{{background:#bdbdbd;color:#333}}
.tab.on{{color:#fff;font-weight:700;border-bottom:none}}
#t-all.on{{background:#1a237e}}
#t-lawsuit.on{{background:#c62828}}
#t-work.on{{background:#1565c0}}
#t-sm.on{{background:#6a1b9a}}
#t-sn.on{{background:#00695c}}
#t-lk.on{{background:#e65100}}
#t-pl.on{{background:#2e7d32}}
.content{{padding:14px 24px}}
.panel{{display:none}}.panel.on{{display:block}}
table{{width:100%;border-collapse:collapse;background:#fff;border-radius:10px;
       overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,.07)}}
th{{background:#1a237e;color:#fff;padding:10px 9px;font-size:13px;text-align:left}}
td{{padding:8px 9px;font-size:13px;border-bottom:1px solid #f0f0f0}}
tr:hover td{{background:#f5f5f5}}
.c{{text-align:center}}.dc{{font-size:12px;color:#555}}
.fl{{color:#1565c0;text-decoration:none;font-weight:500;font-size:12px}}
.fl:hover{{text-decoration:underline}}
.nbadge{{background:#2e7d32;color:#fff;font-size:10px;padding:2px 5px;border-radius:10px;margin-left:3px}}
.b-red{{background:#e53935;color:#fff;font-size:11px;padding:2px 7px;border-radius:10px}}
.b-org{{background:#ef6c00;color:#fff;font-size:11px;padding:2px 7px;border-radius:10px}}
.r-new td{{background:#e8f5e9!important}}.r-urg td{{background:#fff3e0!important}}
.foot{{color:#999;font-size:11px;margin-top:8px;text-align:right;padding-bottom:20px}}
.cbtn{{background:#e8eaf6;border:none;border-radius:6px;cursor:pointer;padding:4px 8px;font-size:13px}}
.cbtn:hover{{background:#c5cae9}}
.modal-bg{{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:1000;align-items:center;justify-content:center}}
.modal-bg.show{{display:flex}}
.modal{{background:#fff;border-radius:12px;padding:24px;min-width:320px;box-shadow:0 8px 32px rgba(0,0,0,.2)}}
.modal h3{{font-size:16px;margin-bottom:16px;color:#1a237e}}
.modal label{{font-size:13px;color:#555;display:block;margin-bottom:4px}}
.modal select{{width:100%;padding:8px;border:1px solid #ddd;border-radius:6px;font-size:13px;margin-bottom:12px}}
.modal-btns{{display:flex;gap:8px;justify-content:flex-end;margin-top:8px}}
</style></head><body>

<!-- ── 잠금 화면 ── -->
<div id="lockScreen">
  <div class="lock-box">
    <div class="lock-logo">📁</div>
    <div class="lock-title">스캔관리대장</div>
    <div class="lock-sub">STN 문서관리시스템 · stnscan.co.kr</div>
    <label class="lock-label">비밀번호</label>
    <input class="lock-input" id="lockPw" type="password" placeholder="비밀번호를 입력하세요"
           onkeydown="if(event.key==='Enter')unlock()">
    <button class="lock-btn" onclick="unlock()">🔓 열기</button>
    <div class="lock-err" id="lockErr"></div>
    <div class="lock-notice">관계자 외 접근 금지<br>© STN Media Group</div>
  </div>
</div>
<!-- ── /잠금 화면 ── -->

<div class="hdr"><h1>📁 스캔관리대장</h1><span>마지막 생성: {updated}</span></div>
<div class="stats">
  <div class="stat"     onclick="goTab('all')">  <div class="n">{len(docs)}</div><div class="l">전체</div></div>
  <div class="stat red" onclick="goTab('lawsuit')"><div class="n">{len(ld)}</div><div class="l">소송문서</div></div>
  <div class="stat red" onclick="filterKw('긴급')"><div class="n">{urg_cnt}</div><div class="l">● 긴급</div></div>
  <div class="stat org" onclick="filterKw('주의')"><div class="n">{note_cnt}</div><div class="l">● 주의</div></div>
  <div class="stat"     onclick="goTab('work')"> <div class="n">{len(wd)}</div><div class="l">업무문서</div></div>
  <div class="stat grn" onclick="filterToday()"> <div class="n">{today_cnt}</div><div class="l">📅 최근 7일</div></div>
  <div class="stat" style="cursor:default">      <div class="n">{drv_cnt}</div><div class="l">☁ 드라이브</div></div>
</div>
<div class="bar">
  <input id="si" placeholder="🔍 파일명, 회사명, 사건명..." oninput="search()">
  <button class="btn b-gray" onclick="clear_()">🔍 검색</button>
  <button class="btn b-blue" onclick="filterToday()">📅 최근 7일</button>
  <button class="btn b-teal" onclick="refresh_()" id="rb">🔄 새로고침</button>
  <span id="rstat"></span>
</div>
<div class="tabs">
  <button class="tab on"  id="t-all"     onclick="goTab('all',this)">📂 전체 ({len(docs)})</button>
  <button class="tab"     id="t-lawsuit" onclick="goTab('lawsuit',this)">⚖️ 소송관리 ({len(ld)})</button>
  <button class="tab"     id="t-work"    onclick="goTab('work',this)">📋 업무문서 ({len(wd)})</button>
  <button class="tab"     id="t-sm"      onclick="goTab('sm',this)">STN미디어 ({len(sm)})</button>
  <button class="tab"     id="t-sn"      onclick="goTab('sn',this)">STN뉴스 ({len(sn)})</button>
  <button class="tab"     id="t-lk"      onclick="goTab('lk',this)">이강영 ({len(lk)})</button>
  <button class="tab"     id="t-pl"      onclick="goTab('pl',this)">플래닛 ({len(pl)})</button>
</div>
<!-- 분류 변경 모달 -->
<div class="modal-bg" id="modalBg">
  <div class="modal">
    <h3>✏️ 분류 변경</h3>
    <div id="modalFname" style="font-size:12px;color:#666;margin-bottom:12px;word-break:break-all"></div>
    <label>회사(기관)명</label>
    <select id="mCompany">
      <option>에스티엔미디어</option>
      <option>에스티엔</option>
      <option>에스티엔씨엘오미디어</option>
      <option>이강영</option>
      <option>플래닛</option>
      <option>이창규드림</option>
      <option>기타</option>
    </select>
    <label>구분</label>
    <select id="mCategory">
      <option>업무</option>
      <option>소송</option>
    </select>
    <label>문서종류</label>
    <select id="mDoctype">
      <option>기타</option>
      <option>소장</option>
      <option>내용증명</option>
      <option>회생계획안</option>
      <option>회생채권조사확정재판</option>
      <option>인사/급여</option>
      <option>내부보고</option>
      <option>공문/신청</option>
      <option>계약서</option>
      <option>입찰</option>
      <option>스캔문서</option>
    </select>
    <div class="modal-btns">
      <button class="btn b-gray" onclick="closeModal()">취소</button>
      <button class="btn b-blue" onclick="saveClassification()">저장</button>
    </div>
  </div>
</div>

<div class="content">
  {panel('all',    docs, 'on')}
  {panel('lawsuit',ld)}
  {panel('work',   wd)}
  {panel('sm',     sm)}
  {panel('sn',     sn)}
  {panel('lk',     lk)}
  {panel('pl',     pl)}
  <div class="foot">🕐 {updated} 기준 · 총 {len(docs)}개 파일</div>
</div>
<script>
let cur='all';
const tabMap={{'all':0,'lawsuit':1,'work':2,'sm':3,'sn':4,'lk':5,'pl':6}};

function goTab(name, btn){{
  cur=name;
  // 패널 전환
  document.querySelectorAll('.panel').forEach(p=>{{p.classList.remove('on');p.style.display='none';}});
  const panel=document.getElementById('p-'+name);
  if(panel){{panel.classList.add('on');panel.style.display='block';}}
  // 탭 활성화
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  const b=btn||document.getElementById('t-'+name);
  if(b) b.classList.add('on');
  search();
}}
function search(){{
  const q=document.getElementById('si').value.toLowerCase();
  const rows=document.getElementById('p-'+cur).querySelectorAll('tbody tr');
  rows.forEach(r=>r.style.display=(!q||r.textContent.toLowerCase().includes(q))?'':'none');
}}
function clear_(){{document.getElementById('si').value='';search();}}
function filterToday(){{
  // 최근 3일 날짜 계산
  const today = new Date();
  const rows = document.getElementById('p-all').querySelectorAll('tbody tr');
  document.getElementById('si').value = '';
  goTab('all', document.getElementById('t-all'));
  rows.forEach(r => {{
    const dateCell = r.querySelector('.dc');
    if(!dateCell) return;
    const dateStr = dateCell.textContent.trim().substring(0,10);
    const rowDate = new Date(dateStr);
    const diff = (today - rowDate) / (1000*60*60*24);
    r.style.display = diff <= 7 ? '' : 'none';
  }});
}}
function filterKw(kw){{
  document.getElementById('si').value=kw;
  goTab('all');search();
}}
let _curFname = '';
function openModal(fname, company, category, doctype){{
  _curFname = fname;
  document.getElementById('modalFname').textContent = fname;
  document.getElementById('mCompany').value  = company  || '기타';
  document.getElementById('mCategory').value = category || '업무';
  document.getElementById('mDoctype').value  = doctype  || '기타';
  document.getElementById('modalBg').classList.add('show');
}}
function closeModal(){{
  document.getElementById('modalBg').classList.remove('show');
}}
async function saveClassification(){{
  const data = {{
    fname:    _curFname,
    company:  document.getElementById('mCompany').value,
    category: document.getElementById('mCategory').value,
    doctype:  document.getElementById('mDoctype').value
  }};
  try{{
    const r = await fetch('http://localhost:{API_PORT}/reclassify', {{
      method:'POST',
      headers:{{'Content-Type':'application/json'}},
      body: JSON.stringify(data)
    }});
    const res = await r.json();
    if(res.ok){{
      closeModal();
      // 새로고침 버튼과 동일하게 재생성
      document.getElementById('rb').click();
    }} else {{
      alert('저장 실패: ' + res.msg);
    }}
  }} catch(e){{
    alert('서버 연결 실패 — 스캔관리대장_실행.bat 확인');
  }}
}}
document.getElementById('modalBg').addEventListener('click', function(e){{
  if(e.target === this) closeModal();
}});

async function refresh_(){{
  const btn=document.getElementById('rb');
  const st=document.getElementById('rstat');
  btn.disabled=true;btn.textContent='⏳ 갱신 중...';st.textContent='';
  try{{
    const r=await fetch('http://localhost:{API_PORT}/refresh',{{method:'POST',signal:AbortSignal.timeout(60000)}});
    const d=await r.json();
    if(d.ok){{st.style.color='#2e7d32';st.textContent='✅ '+d.msg;setTimeout(()=>location.reload(),1500);}}
    else{{st.style.color='#e53935';st.textContent='❌ '+d.msg;}}
  }}catch(e){{
    st.style.color='#e53935';
    st.textContent='❌ 서버 연결 실패 — 스캔관리대장_실행.bat 확인';
  }}finally{{btn.disabled=false;btn.textContent='🔄 새로고침';}}
}}

/* ── 잠금 화면 ── */
const PW_HASH = "{PW_HASH}";
async function sha256(str){{
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(str));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join('');
}}
async function unlock(){{
  const pw = document.getElementById('lockPw').value;
  const err = document.getElementById('lockErr');
  if(!pw){{ err.textContent='비밀번호를 입력하세요.'; return; }}
  const hash = await sha256(pw);
  if(hash === PW_HASH){{
    sessionStorage.setItem('stn_auth','1');
    document.getElementById('lockScreen').classList.add('hide');
    err.textContent='';
  }} else {{
    err.textContent='❌ 비밀번호가 올바르지 않습니다.';
    document.getElementById('lockPw').value='';
    document.getElementById('lockPw').focus();
  }}
}}
// 페이지 로드 시 인증 여부 확인
(function(){{
  if(sessionStorage.getItem('stn_auth')==='1'){{
    document.getElementById('lockScreen').classList.add('hide');
  }} else {{
    document.getElementById('lockPw').focus();
  }}
}})();
/* ── /잠금 화면 ── */
</script></body></html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] ✅ HTML 생성 완료 ({len(pdfs)}개 파일)")
    return len(pdfs)

def drive_upload_new(scan_folder=SCAN_FOLDER):
    """새 파일만 드라이브에 업로드"""
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload

        TOKEN       = r"C:\Users\User\Downloads\token.json"
        OAUTH_CREDS = r"C:\Users\User\Downloads\credentials_oauth.json"
        SCOPES      = ["https://www.googleapis.com/auth/drive.file"]
        FOLDER_NAME = "스캔관리대장"

        creds = None
        if os.path.exists(TOKEN):
            creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                with open(TOKEN, "w") as f:
                    f.write(creds.to_json())
            else:
                print("  드라이브 로그인 필요 -> python gdrive_upload.py 실행")
                return 0

        service = build("drive", "v3", credentials=creds)
        res = service.files().list(
            q=f"name='{FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id)"
        ).execute()
        files = res.get("files", [])
        if not files:
            print("  드라이브 폴더 없음")
            return 0
        folder_id = files[0]["id"]

        fid_path = os.path.join(scan_folder, "file_ids.json")
        file_ids = {}
        if os.path.exists(fid_path):
            with open(fid_path, encoding="utf-8") as f:
                file_ids = json.load(f)

        pdfs = [f for f in os.listdir(scan_folder) if f.lower().endswith(".pdf")]
        new_files = [f for f in pdfs if f not in file_ids]
        if not new_files:
            return 0

        print(f"  드라이브 업로드: {len(new_files)}개 신규 파일")
        cnt = 0
        for fname in new_files:
            fpath = os.path.join(scan_folder, fname)
            try:
                media = MediaFileUpload(fpath, mimetype="application/pdf", resumable=True)
                meta  = {"name": fname, "parents": [folder_id]}
                f = service.files().create(body=meta, media_body=media, fields="id").execute()
                fid = f["id"]
                service.permissions().create(
                    fileId=fid, body={"type":"anyone","role":"reader"}
                ).execute()
                file_ids[fname] = fid
                cnt += 1
                print(f"    [{cnt}] {fname[:50]}")
            except Exception as e:
                print(f"    ERR {fname[:40]}: {e}")

        with open(fid_path, "w", encoding="utf-8") as f:
            json.dump(file_ids, f, ensure_ascii=False, indent=2)
        print(f"  드라이브 업로드 완료: {cnt}개")
        return cnt

    except ImportError:
        print("  google 라이브러리 없음 -> python gdrive_upload.py 먼저 실행")
        return 0
    except Exception as e:
        print(f"  드라이브 업로드 오류: {e}")
        return 0

def git_push():
    """
    스캔관리대장.html 을 GitHub main 브랜치에 푸시.
    - 한국어 파일명 깨짐 방지: core.quotepath false 설정
    - 커밋 메시지 영문 사용 (cp949 인코딩 충돌 방지)
    - nothing to commit 상태는 정상 처리 (오류 아님)
    - stderr/stdout 모두 utf-8 + errors='ignore' 디코딩
    """
    def run(cmd):
        return subprocess.run(
            cmd, cwd=SCAN_FOLDER,
            capture_output=True
        )

    try:
        # 한국어 파일명 인코딩 설정
        run(["git", "config", "core.quotepath", "false"])
        run(["git", "config", "core.autocrlf", "false"])

        # HTML 파일 스테이징
        r = run(["git", "add", "스캔관리대장.html"])
        if r.returncode != 0:
            err = r.stderr.decode("utf-8", "ignore").strip()
            print(f"  git add 실패: {err}")
            return False, f"git add 실패: {err}"

        # 변경 여부 확인
        status = run(["git", "status", "--porcelain"])
        status_out = status.stdout.decode("utf-8", "ignore").strip()
        if not status_out:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] 변경 없음 — push 생략")
            return True, "변경 없음"

        # 커밋 (메시지 영문 — cp949 충돌 방지)
        msg = f"update {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        r = run(["git", "commit", "-m", msg])
        if r.returncode != 0:
            out = r.stdout.decode("utf-8", "ignore").strip()
            err = r.stderr.decode("utf-8", "ignore").strip()
            if "nothing to commit" in out or "nothing to commit" in err:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] nothing to commit — push 생략")
                return True, "변경 없음"
            print(f"  git commit 실패: {err or out}")
            return False, f"git commit 실패: {err or out}"

        # 푸시
        r = run(["git", "push", "origin", "main"])
        if r.returncode != 0:
            err = r.stderr.decode("utf-8", "ignore").strip()
            print(f"  git push 실패: {err}")
            return False, f"git push 실패: {err}"

        print(f"[{datetime.now().strftime('%H:%M:%S')}] GitHub push 완료")
        return True, "GitHub push 완료"

    except FileNotFoundError:
        print("  git 명령어를 찾을 수 없음 — Git 설치 확인")
        return False, "git 명령어 없음"
    except Exception as e:
        print(f"  git_push 예외: {e}")
        return False, str(e)



class Handler(BaseHTTPRequestHandler):
    def log_message(self,*a): pass
    def send_json(self, data, code=200):
        body=json.dumps(data,ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type","application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin","*")
        self.send_header("Access-Control-Allow-Methods","POST,GET,OPTIONS")
        self.end_headers(); self.wfile.write(body)
    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin","*")
        self.send_header("Access-Control-Allow-Methods","POST,GET,OPTIONS")
        self.end_headers()
    def do_POST(self):
        if self.path=="/refresh":
            try:
                cnt=generate(SCAN_FOLDER,OUTPUT_HTML)
                ok,msg=git_push()
                self.send_json({"ok":True,"msg":f"갱신완료({cnt}개) — {msg}"})
            except Exception as e:
                self.send_json({"ok":False,"msg":str(e)},500)
        elif self.path=="/reclassify":
            try:
                length = int(self.headers.get("Content-Length",0))
                body = json.loads(self.rfile.read(length).decode("utf-8"))
                fname    = body.get("fname","")
                company  = body.get("company","")
                category = body.get("category","")
                doctype  = body.get("doctype","")
                if not fname:
                    self.send_json({"ok":False,"msg":"파일명 없음"},400)
                    return
                data = load_classifications()
                data[fname] = {"company":company,"category":category,"doctype":doctype}
                save_classifications(data)
                cnt = generate(SCAN_FOLDER,OUTPUT_HTML)
                self.send_json({"ok":True,"msg":f"분류 변경 완료 ({fname})"})
            except Exception as e:
                self.send_json({"ok":False,"msg":str(e)},500)
        else:
            self.send_json({"ok":False,"msg":"not found"},404)
    def do_GET(self):
        if self.path=="/ping": self.send_json({"ok":True})

if __name__=="__main__":
    if len(sys.argv)>1 and sys.argv[1]=="once":
        # GitHub Actions 환경 감지: SCAN_FOLDER가 없으면 현재 디렉토리 사용
        ci_folder = SCAN_FOLDER if os.path.isdir(SCAN_FOLDER) else os.path.abspath(".")
        ci_output = os.path.join(ci_folder, "스캔관리대장.html")
        generate(ci_folder, ci_output)
        sys.exit(0)

    print("="*50)
    print("  스캔관리대장 v3  |  포트:",API_PORT)
    print("="*50)
    t=threading.Thread(target=lambda:HTTPServer(("127.0.0.1",API_PORT),Handler).serve_forever(),daemon=True)
    t.start()
    generate(SCAN_FOLDER,OUTPUT_HTML)
    prev=0
    while True:
        try:
            n=len([f for f in os.listdir(SCAN_FOLDER) if f.lower().endswith('.pdf')])
            if n!=prev:
                generate(SCAN_FOLDER,OUTPUT_HTML)
                prev=n
                uploaded=drive_upload_new(SCAN_FOLDER)
                if uploaded>0:
                    generate(SCAN_FOLDER,OUTPUT_HTML)
                git_push()
        except Exception as e: print(f"오류:{e}")
        time.sleep(5)
