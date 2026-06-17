"""
스캔관리대장 자동 생성기 v4
- 한글 파일명: 파일 생성일(ctime) 기준 날짜 적용
- 날짜+시간 기준 최신순 정렬
- 분류 모달에 날짜 수동 수정 기능
- 새 파일 감지 알림 배지 (30초 폴링)
- /filecount API 추가
"""

import os, json, re, sys, time, subprocess, threading
from datetime import datetime, date, timedelta
from http.server import HTTPServer, BaseHTTPRequestHandler

SCAN_FOLDER = r"C:\Users\User\Desktop\스캔"
OUTPUT_HTML = os.path.join(SCAN_FOLDER, "스캔관리대장.html")
XLSX_PATH   = os.path.join(SCAN_FOLDER, "소송목록.xlsx")
FOLDER_ID   = "1hT9xdNEawnkrZ78p26Kcjlm-1wdR5x-X"
API_PORT    = 8765


def _file_ts(fpath):
    try:
        return min(os.path.getctime(fpath), os.path.getmtime(fpath))
    except:
        return time.time()


def extract_date_and_ts(fname, fpath=None):
    ts = _file_ts(fpath) if (fpath and os.path.exists(fpath)) else time.time()

    m = re.match(r'SCAN_(\d{8})_', fname, re.IGNORECASE)
    if m:
        d = m.group(1)
        y, mo, day = int(d[:4]), int(d[4:6]), int(d[6:8])
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}", ts

    m = re.match(r'(\d{8})[_\-]', fname)
    if m:
        d = m.group(1)
        y, mo, day = int(d[:4]), int(d[4:6]), int(d[6:8])
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}", ts

    m = re.search(r'(\d{4})[.\-](\d{2})[.\-](\d{2})', fname)
    if m:
        y, mo, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}", ts

    for m in re.finditer(r'(\d{8})', fname):
        d = m.group(1)
        y, mo, day = int(d[:4]), int(d[4:6]), int(d[6:8])
        if 2020 <= y <= 2035 and 1 <= mo <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mo:02d}-{day:02d}", ts

    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d"), ts


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
    kw = ["소장","소송","내용증명","회생","채권","법원","재판","확정","가소","가합","간회합","간이회생"]
    return "소송" if any(k in fname for k in kw) else "업무"

def detect_doctype(fname):
    if "내용증명" in fname: return "내용증명"
    if "소장"     in fname: return "소장"
    if "회생계획" in fname: return "회생계획안"
    if "확정재판" in fname: return "회생채권조사확정재판"
    if "채권"     in fname: return "채권신고"
    if "회생"     in fname: return "회생관련"
    if any(k in fname for k in ["급여","임금","인사"]): return "인사/급여"
    if "보고서" in fname or "보고" in fname: return "내부보고"
    if any(k in fname for k in ["공문","신청"]): return "공문/신청"
    if "계약"   in fname: return "계약서"
    if "입찰"   in fname: return "입찰"
    if fname.upper().startswith("SCAN_"): return "스캔문서"
    return "기타"

def detect_urgency(fname):
    m = re.search(r'기한(\d{8})', fname)
    if m:
        try:
            dl = datetime.strptime(m.group(1), "%Y%m%d").date()
            left = (dl - date.today()).days
            if left <= 7:  return "긴급"
            if left <= 14: return "주의"
        except: pass
    if any(k in fname for k in ["긴급","urgent","즉시","당일","가압류","집행"]): return "긴급"
    if any(k in fname for k in ["주의","중요"]): return "주의"
    if detect_category(fname) == "소송": return "주의"
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

def load_classifications():
    if os.path.exists(CLASSIFICATIONS):
        try:
            with open(CLASSIFICATIONS, encoding="utf-8") as f:
                return json.load(f)
        except: pass
    return {}

def save_classifications(data):
    with open(CLASSIFICATIONS, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def fetch_drive_ids():
    try:
        jp = os.path.join(SCAN_FOLDER, "file_ids.json")
        if os.path.exists(jp):
            with open(jp, encoding="utf-8") as f:
                return json.load(f)
    except: pass
    return {}

def make_link(fname, drive_ids):
    fid = drive_ids.get(fname)
    if fid: return f"https://drive.google.com/file/d/{fid}/view"
    local = os.path.join(SCAN_FOLDER, fname).replace("\\", "/")
    return f"file:///{local}"


def generate(scan_folder=SCAN_FOLDER, output_path=OUTPUT_HTML):
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] HTML 생성 시작...")
    drive_ids = fetch_drive_ids()
    load_lawsuit_data()

    pdfs      = [f for f in os.listdir(scan_folder) if f.lower().endswith('.pdf')]
    today_str = date.today().strftime("%Y-%m-%d")
    clsf      = load_classifications()

    docs = []
    for fname in pdfs:
        fpath    = os.path.join(scan_folder, fname)
        ov       = clsf.get(fname, {})
        company  = ov.get("company")  or detect_company(fname)
        category = ov.get("category") or detect_category(fname)
        doctype  = ov.get("doctype")  or detect_doctype(fname)
        date_str, sort_ts = extract_date_and_ts(fname, fpath)
        if ov.get("date"):
            date_str = ov["date"]
            try: sort_ts = datetime.strptime(ov["date"], "%Y-%m-%d").timestamp()
            except: pass
        docs.append({
            "fname": fname, "date": date_str, "sort_ts": sort_ts,
            "company": company, "category": category, "doctype": doctype,
            "urgency": detect_urgency(fname), "link": make_link(fname, drive_ids),
        })

    docs.sort(key=lambda d: d["sort_ts"], reverse=True)

    def nbadge(d):
        return '<span class="nbadge">NEW</span>' if d["date"] == today_str else ""

    def ubadge(u):
        if u == "긴급": return '<span class="b-red">긴급</span>'
        if u == "주의": return '<span class="b-org">주의</span>'
        return ""

    def rowcls(d):
        if d["date"] == today_str: return ' class="r-new"'
        if d["urgency"] == "긴급": return ' class="r-urg"'
        return ""

    def mkrows(lst):
        out = ""
        for i, d in enumerate(lst):
            icon  = "&#9878;" if d["category"] == "소송" else "&#128203;"
            safe  = d["fname"].replace("'", "\\'")
            co    = d["company"].replace("'", "\\'")
            ca    = d["category"].replace("'", "\\'")
            dt    = d["doctype"].replace("'", "\\'")
            sd    = d["date"]
            short = d["fname"][:28] + ("..." if len(d["fname"]) > 28 else "")
            out += (
                f'<tr{rowcls(d)}>'
                f'<td class="c">{i+1}</td>'
                f'<td class="c dc">{d["date"]} {nbadge(d)}</td>'
                f'<td class="c">{icon} {d["category"]}</td>'
                f'<td>{d["company"]}</td>'
                f'<td>{d["doctype"]}</td>'
                f'<td class="c">{ubadge(d["urgency"])}</td>'
                f'<td><a href="{d["link"]}" target="_blank" class="fl">&#128196; {short}</a></td>'
                f'<td class="c"><button class="cbtn" '
                f'onclick="openModal(\'{safe}\',\'{co}\',\'{ca}\',\'{dt}\',\'{sd}\')">&#9998;</button></td>'
                f'</tr>\n'
            )
        return out

    def filt(k, v): return [d for d in docs if d[k] == v]

    ld = filt("category","소송"); wd = filt("category","업무")
    sm = filt("company","에스티엔미디어"); sn = filt("company","에스티엔")
    lk = filt("company","이강영");         pl = filt("company","플래닛")
    urg  = sum(1 for d in docs if d["urgency"] == "긴급")
    warn = sum(1 for d in docs if d["urgency"] == "주의")
    wago = (date.today() - timedelta(days=7)).strftime("%Y-%m-%d")
    tcnt = sum(1 for d in docs if d["date"] >= wago)
    dcnt = len(drive_ids)
    upd  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    N    = len(docs)

    def panel(pid, lst, on=""):
        hd = ("<thead><tr>"
              "<th>No</th><th>스캔일자</th><th>구분</th>"
              "<th>회사(기관)명</th><th>문서종류</th><th>긴급도</th>"
              "<th>파일 열기</th><th>분류</th>"
              "</tr></thead>")
        return (f'<div id="p-{pid}" class="panel{" on" if on else ""}">'
                f'<table>{hd}<tbody>{mkrows(lst)}</tbody></table></div>')

    # ── HTML (f-string, CSS/JS 중괄호는 {{ }} 로 이스케이프) ──────────
    html = f"""<!DOCTYPE html>
<html lang="ko"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>스캔관리대장</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Malgun Gothic',sans-serif;background:#f0f2f5;color:#333;font-size:14px}}
.hdr{{background:#1a237e;color:#fff;padding:14px 24px;display:flex;align-items:center;justify-content:space-between}}
.hdr h1{{font-size:19px}}.hdr span{{font-size:11px;opacity:.7}}
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
      border-bottom:3px solid transparent;margin-bottom:-2px;border-radius:6px 6px 0 0;
      font-weight:600;transition:all .15s}}
.tab:hover{{background:#bdbdbd;color:#333}}
.tab.on{{color:#fff;font-weight:700;border-bottom:none}}
#t-all.on{{background:#1a237e}}#t-lawsuit.on{{background:#c62828}}
#t-work.on{{background:#1565c0}}#t-sm.on{{background:#6a1b9a}}
#t-sn.on{{background:#00695c}}#t-lk.on{{background:#e65100}}#t-pl.on{{background:#2e7d32}}
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
.nbadge{{background:#2e7d32;color:#fff;font-size:10px;padding:2px 5px;
         border-radius:10px;margin-left:3px;vertical-align:middle}}
.b-red{{background:#e53935;color:#fff;font-size:11px;padding:2px 7px;border-radius:10px}}
.b-org{{background:#ef6c00;color:#fff;font-size:11px;padding:2px 7px;border-radius:10px}}
.r-new td{{background:#e8f5e9!important}}.r-urg td{{background:#fff3e0!important}}
.foot{{color:#999;font-size:11px;margin-top:8px;text-align:right;padding-bottom:20px}}
.cbtn{{background:#e8eaf6;border:none;border-radius:6px;cursor:pointer;padding:4px 8px;font-size:13px}}
.cbtn:hover{{background:#c5cae9}}
.modal-bg{{display:none;position:fixed;top:0;left:0;width:100%;height:100%;
           background:rgba(0,0,0,.5);z-index:1000;align-items:center;justify-content:center}}
.modal-bg.show{{display:flex}}
.modal{{background:#fff;border-radius:12px;padding:24px;min-width:320px;
        box-shadow:0 8px 32px rgba(0,0,0,.2)}}
.modal h3{{font-size:16px;margin-bottom:12px;color:#1a237e}}
.mfname{{font-size:11px;color:#888;margin-bottom:14px;word-break:break-all;
         padding:6px 8px;background:#f5f5f5;border-radius:6px}}
.modal label{{font-size:12px;color:#555;display:block;margin-bottom:4px;font-weight:600}}
.modal select,
.modal input[type="date"]{{width:100%;padding:8px;border:1px solid #ddd;
                           border-radius:6px;font-size:13px;margin-bottom:12px;
                           background:#fff}}
.modal-btns{{display:flex;gap:8px;justify-content:flex-end;margin-top:4px}}
</style>
</head><body>
<div class="hdr">
  <h1>&#128193; 스캔관리대장</h1>
  <span>마지막 생성: {upd}</span>
</div>
<div class="stats">
  <div class="stat"     onclick="goTab('all')">    <div class="n">{N}</div>   <div class="l">전체</div></div>
  <div class="stat red" onclick="goTab('lawsuit')"> <div class="n">{len(ld)}</div><div class="l">소송문서</div></div>
  <div class="stat red" onclick="filterKw('긴급')"> <div class="n">{urg}</div> <div class="l">&#9679; 긴급</div></div>
  <div class="stat org" onclick="filterKw('주의')"> <div class="n">{warn}</div><div class="l">&#9679; 주의</div></div>
  <div class="stat"     onclick="goTab('work')">    <div class="n">{len(wd)}</div><div class="l">업무문서</div></div>
  <div class="stat grn" onclick="filterToday()">    <div class="n">{tcnt}</div><div class="l">&#128197; 최근 7일</div></div>
  <div class="stat" style="cursor:default">         <div class="n">{dcnt}</div><div class="l">&#9729; 드라이브</div></div>
</div>
<div class="bar">
  <input id="si" placeholder="&#128269; 파일명, 회사명, 사건명..." oninput="doSearch()">
  <button class="btn b-gray" onclick="clearSearch()">초기화</button>
  <button class="btn b-blue" onclick="filterToday()">&#128197; 최근 7일</button>
  <button class="btn b-teal" id="rb" onclick="doRefresh()">&#128260; 새로고침</button>
  <span id="rstat"></span>
</div>
<div class="tabs">
  <button class="tab on" id="t-all"     onclick="goTab('all',this)">&#128194; 전체 ({N})</button>
  <button class="tab"    id="t-lawsuit" onclick="goTab('lawsuit',this)">&#9878; 소송관리 ({len(ld)})</button>
  <button class="tab"    id="t-work"    onclick="goTab('work',this)">&#128203; 업무문서 ({len(wd)})</button>
  <button class="tab"    id="t-sm"      onclick="goTab('sm',this)">STN미디어 ({len(sm)})</button>
  <button class="tab"    id="t-sn"      onclick="goTab('sn',this)">STN뉴스 ({len(sn)})</button>
  <button class="tab"    id="t-lk"      onclick="goTab('lk',this)">이강영 ({len(lk)})</button>
  <button class="tab"    id="t-pl"      onclick="goTab('pl',this)">플래닛 ({len(pl)})</button>
</div>

<!-- ===== 분류 변경 모달 ===== -->
<div class="modal-bg" id="modalBg">
  <div class="modal">
    <h3>&#9998; 분류 변경</h3>
    <div class="mfname" id="mFname"></div>

    <label>&#128197; 스캔 날짜</label>
    <input type="date" id="mDate">

    <label>&#127970; 회사(기관)명</label>
    <select id="mCompany">
      <option>에스티엔미디어</option>
      <option>에스티엔</option>
      <option>에스티엔씨엘오미디어</option>
      <option>이강영</option>
      <option>플래닛</option>
      <option>이창규드림</option>
      <option>기타</option>
    </select>

    <label>&#128221; 구분</label>
    <select id="mCategory">
      <option>업무</option>
      <option>소송</option>
    </select>

    <label>&#128196; 문서종류</label>
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
      <button class="btn b-blue" onclick="saveModal()">저장</button>
    </div>
  </div>
</div>
<!-- ========================== -->

<div class="content">
  {panel('all',    docs, True)}
  {panel('lawsuit',ld)}
  {panel('work',   wd)}
  {panel('sm',     sm)}
  {panel('sn',     sn)}
  {panel('lk',     lk)}
  {panel('pl',     pl)}
  <div class="foot">&#128336; {upd} 기준 &middot; 총 {N}개 파일</div>
</div>

<script>
var _cur = 'all';
var _fname = '';

function goTab(name, btn) {{
  _cur = name;
  document.querySelectorAll('.panel').forEach(function(p) {{
    p.classList.remove('on'); p.style.display = 'none';
  }});
  var p = document.getElementById('p-' + name);
  if (p) {{ p.classList.add('on'); p.style.display = 'block'; }}
  document.querySelectorAll('.tab').forEach(function(t) {{ t.classList.remove('on'); }});
  var b = btn || document.getElementById('t-' + name);
  if (b) b.classList.add('on');
  doSearch();
}}

function doSearch() {{
  var q = document.getElementById('si').value.toLowerCase();
  var rows = document.getElementById('p-' + _cur).querySelectorAll('tbody tr');
  rows.forEach(function(r) {{
    r.style.display = (!q || r.textContent.toLowerCase().indexOf(q) >= 0) ? '' : 'none';
  }});
}}

function clearSearch() {{
  document.getElementById('si').value = '';
  doSearch();
}}

function filterToday() {{
  var today = new Date();
  goTab('all', document.getElementById('t-all'));
  document.getElementById('si').value = '';
  document.getElementById('p-all').querySelectorAll('tbody tr').forEach(function(r) {{
    var dc = r.querySelector('.dc');
    if (!dc) return;
    var diff = (today - new Date(dc.textContent.trim().substring(0, 10))) / 86400000;
    r.style.display = diff <= 7 ? '' : 'none';
  }});
}}

function filterKw(kw) {{
  document.getElementById('si').value = kw;
  goTab('all'); doSearch();
}}

/* ── 모달 ── */
function openModal(fname, company, category, doctype, scanDate) {{
  _fname = fname;
  document.getElementById('mFname').textContent = fname;
  document.getElementById('mDate').value     = scanDate  || '';
  document.getElementById('mCompany').value  = company   || '기타';
  document.getElementById('mCategory').value = category  || '업무';
  document.getElementById('mDoctype').value  = doctype   || '기타';
  document.getElementById('modalBg').classList.add('show');
}}

function closeModal() {{
  document.getElementById('modalBg').classList.remove('show');
}}

document.getElementById('modalBg').addEventListener('click', function(e) {{
  if (e.target === this) closeModal();
}});

function saveModal() {{
  var payload = {{
    fname    : _fname,
    date     : document.getElementById('mDate').value,
    company  : document.getElementById('mCompany').value,
    category : document.getElementById('mCategory').value,
    doctype  : document.getElementById('mDoctype').value
  }};
  fetch('http://localhost:{API_PORT}/reclassify', {{
    method  : 'POST',
    headers : {{'Content-Type': 'application/json'}},
    body    : JSON.stringify(payload)
  }}).then(function(r) {{ return r.json(); }})
    .then(function(res) {{
      if (res.ok) {{ closeModal(); doRefresh(); }}
      else {{ alert('저장 실패: ' + res.msg); }}
    }}).catch(function() {{
      alert('서버 연결 실패 - 스캔관리대장_실행.bat 확인');
    }});
}}

/* ── 새로고침 ── */
function doRefresh() {{
  var btn = document.getElementById('rb');
  var st  = document.getElementById('rstat');
  btn.disabled = true; btn.textContent = '...갱신중...'; st.textContent = '';
  fetch('http://localhost:{API_PORT}/refresh', {{ method: 'POST' }})
    .then(function(r) {{ return r.json(); }})
    .then(function(d) {{
      if (d.ok) {{
        st.style.color = '#2e7d32'; st.textContent = 'OK ' + d.msg;
        setTimeout(function() {{ location.reload(); }}, 1500);
      }} else {{
        st.style.color = '#e53935'; st.textContent = 'ERR ' + d.msg;
      }}
    }}).catch(function() {{
      st.style.color = '#e53935';
      st.textContent = 'ERR 서버 연결 실패';
    }}).finally(function() {{
      btn.disabled = false; btn.textContent = '&#128260; 새로고침';
    }});
}}

/* ── 새 파일 감지 (30초 폴링) ── */
(function() {{
  var known = {N};
  function poll() {{
    fetch('http://localhost:{API_PORT}/filecount')
      .then(function(r) {{ return r.json(); }})
      .then(function(d) {{
        if (d.count > known) {{ showToast(d.count - known); known = d.count; }}
      }}).catch(function() {{}});
  }}
  function showToast(n) {{
    document.title = '[' + n + '개 신규] 스캔관리대장';
    var el = document.getElementById('_nb');
    if (!el) {{
      el = document.createElement('div');
      el.id = '_nb';
      el.style.cssText = [
        'position:fixed;bottom:24px;right:24px;z-index:9999;',
        'background:#2e7d32;color:#fff;padding:14px 20px;',
        'border-radius:12px;font-size:14px;font-weight:700;',
        'box-shadow:0 4px 16px rgba(0,0,0,.3);cursor:pointer;'
      ].join('');
      el.onclick = function() {{
        doRefresh(); el.remove(); document.title = '스캔관리대장';
      }};
      document.body.appendChild(el);
    }}
    el.textContent = '[신규] ' + n + '개 파일 - 클릭하여 새로고침';
  }}
  setInterval(poll, 30000);
  setTimeout(poll, 5000);
}})();
</script>
</body></html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[{datetime.now().strftime('%H:%M:%S')}] HTML 생성 완료 ({len(pdfs)}개 파일)")
    return len(pdfs)


def drive_upload_new(scan_folder=SCAN_FOLDER):
    try:
        from google.oauth2.credentials import Credentials
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        TOKEN  = r"C:\Users\User\Downloads\token.json"
        SCOPES = ["https://www.googleapis.com/auth/drive.file"]
        FNAME  = "스캔관리대장"
        creds  = None
        if os.path.exists(TOKEN):
            creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
                with open(TOKEN, "w") as f: f.write(creds.to_json())
            else:
                print("  드라이브 로그인 필요 -> python gdrive_upload.py 실행"); return 0
        svc = build("drive", "v3", credentials=creds)
        res = svc.files().list(
            q=f"name='{FNAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id)").execute()
        fls = res.get("files", [])
        if not fls: print("  드라이브 폴더 없음"); return 0
        fid_path = os.path.join(scan_folder, "file_ids.json")
        file_ids = json.load(open(fid_path, encoding="utf-8")) if os.path.exists(fid_path) else {}
        pdfs     = [f for f in os.listdir(scan_folder) if f.lower().endswith(".pdf")]
        new_files = [f for f in pdfs if f not in file_ids]
        if not new_files: return 0
        cnt = 0
        for fname in new_files:
            try:
                media = MediaFileUpload(os.path.join(scan_folder, fname), mimetype="application/pdf", resumable=True)
                fobj  = svc.files().create(body={"name": fname, "parents": [fls[0]["id"]]},
                                           media_body=media, fields="id").execute()
                svc.permissions().create(fileId=fobj["id"], body={"type":"anyone","role":"reader"}).execute()
                file_ids[fname] = fobj["id"]; cnt += 1
            except Exception as e: print(f"  ERR {fname[:40]}: {e}")
        with open(fid_path, "w", encoding="utf-8") as f:
            json.dump(file_ids, f, ensure_ascii=False, indent=2)
        return cnt
    except ImportError: print("  google 라이브러리 없음"); return 0
    except Exception as e: print(f"  드라이브 오류: {e}"); return 0


def git_push():
    try:
        os.chdir(SCAN_FOLDER)
        subprocess.run(["git","add","스캔관리대장.html"], check=True, capture_output=True)
        subprocess.run(["git","commit","-m",f"update {datetime.now().strftime('%Y-%m-%d %H:%M')}"],
                       check=True, capture_output=True)
        subprocess.run(["git","push"], check=True, capture_output=True)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] GitHub 업로드 완료")
        return True, "GitHub 업로드 완료"
    except subprocess.CalledProcessError as e:
        return False, (e.stderr.decode("utf-8","ignore").strip() if e.stderr else "오류")
    except Exception as e:
        return False, str(e)


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def send_json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False).encode()
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
    def do_GET(self):
        if self.path == "/ping":
            self.send_json({"ok": True})
        elif self.path == "/filecount":
            try:
                cnt = len([f for f in os.listdir(SCAN_FOLDER) if f.lower().endswith('.pdf')])
                self.send_json({"ok": True, "count": cnt})
            except Exception as e:
                self.send_json({"ok": False, "count": 0, "msg": str(e)})
    def do_POST(self):
        if self.path == "/refresh":
            try:
                cnt = generate(SCAN_FOLDER, OUTPUT_HTML)
                ok, msg = git_push()
                self.send_json({"ok": True, "msg": f"갱신완료({cnt}개) - {msg}"})
            except Exception as e:
                self.send_json({"ok": False, "msg": str(e)}, 500)
        elif self.path == "/reclassify":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body   = json.loads(self.rfile.read(length).decode("utf-8"))
                fname  = body.get("fname","")
                if not fname:
                    self.send_json({"ok": False, "msg": "파일명 없음"}, 400); return
                data = load_classifications()
                data[fname] = {
                    "company" : body.get("company",""),
                    "category": body.get("category",""),
                    "doctype" : body.get("doctype",""),
                }
                if body.get("date"):
                    data[fname]["date"] = body["date"]
                save_classifications(data)
                generate(SCAN_FOLDER, OUTPUT_HTML)
                self.send_json({"ok": True, "msg": f"분류 변경 완료 ({fname})"})
            except Exception as e:
                self.send_json({"ok": False, "msg": str(e)}, 500)
        else:
            self.send_json({"ok": False, "msg": "not found"}, 404)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "once":
        generate(SCAN_FOLDER, OUTPUT_HTML); sys.exit(0)

    print("=" * 50)
    print(f"  스캔관리대장 v4  |  포트: {API_PORT}")
    print("=" * 50)
    t = threading.Thread(
        target=lambda: HTTPServer(("127.0.0.1", API_PORT), Handler).serve_forever(),
        daemon=True)
    t.start()
    generate(SCAN_FOLDER, OUTPUT_HTML)
    prev = 0
    while True:
        try:
            n = len([f for f in os.listdir(SCAN_FOLDER) if f.lower().endswith('.pdf')])
            if n != prev:
                generate(SCAN_FOLDER, OUTPUT_HTML)
                prev = n
                if drive_upload_new(SCAN_FOLDER) > 0:
                    generate(SCAN_FOLDER, OUTPUT_HTML)
                git_push()
        except Exception as e:
            print(f"오류: {e}")
        time.sleep(5)
