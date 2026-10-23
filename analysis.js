/* ══════════════════════════════════════════════════════════════════
   STAY BUT 어드민 — 매출 분석
   [2026-10-23] index.html 안에 있던 매출분석·시트 코드를 그대로 옮겼습니다.
   내용은 한 글자도 바꾸지 않았습니다. (원래 index.html 6244~7780번째 줄)

   index.html 에서 원래 있던 자리에 <script src="/analysis.js"> 한 줄로
   불러옵니다. 실행 순서가 예전과 똑같아야 해서 꼭 그 자리여야 합니다.
   (앞쪽 코드 → 이 파일 → 뒤쪽 코드)

   들어 있는 것: 매출분석 시트(renderSheet), 인사이트 카드(analysisRender),
   엑셀·이미지 내보내기, 목표 설정, 객단가 하한선(RATE_FLOOR) 등
   ══════════════════════════════════════════════════════════════════ */

let __analysisRows = null;
let analysisView = 'all';
//  📊 매출 분석 시트 (전체 화면 뷰 · 시안 디자인 + 라이브 데이터)
let __sheetRows = null, __sheetView = null, __sheetWeek = 'all', __sheetGoal = {}, __sheetCharts = [], __sheetStyled = false;
function sheetStyle_() {
  if (__sheetStyled) return; __sheetStyled = true;
  const css = `
  /* [2026-09-17] 매출분석 — 시안 구조 적용. · 여기 로컬 변수들이 밝은 색으로 고정돼 있어서 다크에서 흰 박스가 남았습니다. · 전부 전역 토큰에 연결합니다. 계산·데이터는 건드리지 않았습니다. */
  #analysis-sheet{
    --sl:var(--text);--sl2:var(--text-soft);--mut:var(--text-muted);--ln:var(--border);
    --grn:var(--success-600);--red:var(--danger-500);
    --bk:#2e5891;--air:#b85c3f;--ago:#6b3fa0;--exp:#a37e2b;
    color:var(--text);font-family:Pretendard,-apple-system,'Malgun Gothic','Noto Sans KR',sans-serif;}
  html[data-theme="dark"] #analysis-sheet{--bk:#8fbde8;--air:#e08a6a;--ago:#c48ad8;--exp:#f5c94a;}
  #analysis-sheet *{box-sizing:border-box;}
  /* [2026-09-17] color 를 안 주고 있어서 다크에서 제목이 배경에 묻혔습니다. */
  #analysis-sheet .as-title{font-size:22px;font-weight:800;margin:0;letter-spacing:-.02em;color:var(--text);}
  #analysis-sheet .as-sub{font-size:12.5px;color:var(--mut);margin-top:3px;}
  #analysis-sheet .as-tabs{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin:16px 0 12px;}
  #analysis-sheet .as-tab{font-size:12.5px;font-weight:700;padding:6px 13px;border-radius:8px;background:transparent;color:var(--mut);border:1px solid transparent;cursor:pointer;}
  #analysis-sheet .as-tab:hover{border-color:var(--border);}
  #analysis-sheet .as-tab.on{background:var(--surface);border-color:var(--border-strong);color:var(--text);font-weight:800;}
  /* [2026-09-17] 기간 + 주차를 담는 한 줄짜리 컨트롤 바 */
  #analysis-sheet .as-ctrl{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin:16px 0 20px;
    padding:8px 12px;background:var(--surface-soft);border:1px solid var(--border);border-radius:11px;}
  #analysis-sheet .as-ctrl-sp{flex:1;min-width:0;}
  #analysis-sheet .as-hbtn{border:1px solid var(--border-strong);background:var(--surface);color:var(--text-soft);
    font-size:11.5px;font-weight:700;padding:7px 13px;border-radius:9px;cursor:pointer;}
  #analysis-sheet .as-hbtn:hover{color:var(--text);}
  #analysis-sheet .as-card{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px 18px;}
  #analysis-sheet .as-sec{font-size:15px;font-weight:800;margin:22px 2px 10px;}
  /* 시안: 카드 사이 간격 11 → 13 */
  #analysis-sheet .as-grid{display:grid;gap:13px;}
  /* 시안: 큰 그림자와 20px 라운드를 뺐습니다. 카드가 아니라 숫자가 보이게. */
  #analysis-sheet .as-hero{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:16px 18px;color:var(--text);}
  #analysis-sheet .as-hero .lab{font-size:12px;color:var(--text-muted);font-weight:600;} #analysis-sheet .as-hero .big{font-size:36px;font-weight:800;margin-top:4px;color:var(--text);letter-spacing:-.035em;line-height:1.05;} #analysis-sheet .as-hero .subx{font-size:11.5px;color:var(--text-muted);margin-top:6px;}
  #analysis-sheet .as-delta{display:inline-flex;align-items:center;gap:3px;font-size:11px;font-weight:800;border-radius:7px;padding:2px 7px;}
  #analysis-sheet .as-delta.up{color:var(--success-600);background:var(--success-50);} #analysis-sheet .as-delta.down{color:var(--danger-600);background:var(--danger-50);} #analysis-sheet .as-delta.flat{color:var(--text-muted);background:var(--surface-soft);}
  #analysis-sheet .as-banner{display:flex;align-items:center;gap:9px;background:var(--surface-soft);border:1px solid var(--border);border-radius:9px;padding:9px 11px;margin-top:9px;font-size:12px;line-height:1.5;color:var(--text-soft);}
  #analysis-sheet .as-banner .ic{color:var(--sl);font-weight:800;}
  #analysis-sheet .as-banner b{color:var(--sl);}
  #analysis-sheet .as-banner .ic{width:34px;height:34px;border-radius:10px;background:#3b6db5;display:grid;place-items:center;flex:none;color:#fff;font-size:16px;}
  #analysis-sheet .as-prog{height:10px;background:rgba(255,255,255,.18);border-radius:6px;overflow:hidden;margin:8px 0;}
  #analysis-sheet .as-prog>div{height:100%;background:linear-gradient(90deg,#4a8a3d,#c9e0c1);border-radius:6px;transition:width .4s;}
  #analysis-sheet .as-statrow{display:flex;gap:26px;flex-wrap:wrap;margin-top:16px;padding-top:15px;border-top:1px solid var(--border);}
  #analysis-sheet .as-statrow .v{font-size:18px;font-weight:800;} #analysis-sheet .as-statrow .k{font-size:11px;color:var(--mut);}
  #analysis-sheet .as-gin{width:120px;padding:6px 9px;border-radius:8px;border:1px solid rgba(255,255,255,.3);background:rgba(255,255,255,.12);color:#fff;font-size:13px;}
  #analysis-sheet .as-gbtn{font-size:12px;font-weight:700;padding:6px 11px;border-radius:8px;border:0;cursor:pointer;background:#c9e0c1;color:#123;}
  #analysis-sheet .as-gpre{background:rgba(255,255,255,.16);color:#fff;}
  /* 시안: KPI 는 흰 카드가 아니라 한 톤 눌린 면 — 큰 숫자와 층이 생깁니다 */
  #analysis-sheet .as-kpi{background:var(--surface-soft);border:1px solid var(--border);border-radius:10px;padding:12px 13px;}
  #analysis-sheet .as-kpi .k{font-size:11px;color:var(--mut);font-weight:700;} #analysis-sheet .as-kpi .v{font-size:20px;font-weight:800;margin-top:4px;letter-spacing:-.025em;color:var(--text);} #analysis-sheet .as-kpi .d{font-size:10.5px;color:var(--text-muted);margin-top:3px;}
  #analysis-sheet .as-mini{height:6px;background:var(--border);border-radius:4px;margin-top:8px;overflow:hidden;} #analysis-sheet .as-mini>div{height:100%;border-radius:4px;}
  #analysis-sheet table.as-tbl{border-collapse:collapse;width:100%;font-size:12.5px;}
  /* [2026-09-18] 표 글씨 12 → 13px, 여백도 조금 넉넉하게 */
  #analysis-sheet .as-tbl{font-size:13px;}
  #analysis-sheet .as-tbl th,#analysis-sheet .as-tbl td{padding:10px 10px;text-align:right;border-bottom:1px solid var(--ln);}
  #analysis-sheet .as-tbl th:first-child,#analysis-sheet .as-tbl td:first-child{text-align:left;}
  #analysis-sheet .as-tbl th{font-size:12px;color:var(--mut);font-weight:700;background:var(--surface-soft);}
  #analysis-sheet .as-tbl tr.tot td{font-weight:800;border-top:1.5px solid var(--border-strong);background:var(--surface-soft);color:var(--text);}
  #analysis-sheet .as-pos{color:var(--grn);font-weight:700;}
  #analysis-sheet .as-pill{display:inline-block;font-size:12px;font-weight:700;padding:3px 9px;border-radius:6px;}
  #analysis-sheet .as-chartbox{position:relative;height:220px;}
  #analysis-sheet .as-subtabs{display:flex;gap:6px;flex-wrap:wrap;margin:-6px 0 14px;}
  #analysis-sheet .as-subtab{font-size:12px;font-weight:700;padding:5px 12px;border-radius:8px;background:var(--surface-soft);color:var(--text-muted);border:1px solid var(--border);cursor:pointer;}
  #analysis-sheet .as-subtab.on{background:var(--accent);color:#fff;border-color:var(--accent);}
  #analysis-sheet .as-basis{display:inline-block;font-size:10.5px;font-weight:700;color:#2e5891;background:#e8f0fa;border-radius:6px;padding:2px 8px;margin-left:6px;vertical-align:middle;}
  #analysis-sheet .as-kpigrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-top:12px;}
  #analysis-sheet .as-note{font-size:12.5px;color:var(--mut);margin-top:8px;line-height:1.7;}
  #analysis-sheet .as-lever{background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:14px 16px;margin-bottom:10px;border-left:3px solid var(--sl);}
  #analysis-sheet .as-lever .top{display:flex;align-items:center;gap:8px;flex-wrap:wrap;}
  #analysis-sheet .as-lever .rank{font-size:11px;font-weight:800;background:var(--sl);color:#fff;padding:3px 9px;border-radius:6px;}
  #analysis-sheet .as-lever .nm{font-size:14.5px;font-weight:800;} #analysis-sheet .as-lever .eff{margin-left:auto;font-size:12px;color:var(--grn);font-weight:800;}
  #analysis-sheet .as-lever p{margin:7px 0 0;font-size:12.5px;} #analysis-sheet .as-as{color:#a37e2b;font-weight:700;} #analysis-sheet .as-to{color:#3a7030;font-weight:700;}
  /* ── 카드 격자 레이아웃 ── */
  #analysis-sheet .as-shell{display:block;}
  #analysis-sheet .as-side{display:none;}
  #analysis-sheet .as-side .sl{font-size:10.5px;font-weight:800;color:var(--mut);letter-spacing:.04em;padding:3px 6px 7px;}
  #analysis-sheet .as-side i{display:block;font-style:normal;font-size:12px;padding:7px 8px;border-radius:7px;color:var(--text-soft);cursor:pointer;user-select:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
  #analysis-sheet .as-side i:hover{background:var(--surface-soft);}
  #analysis-sheet .as-side i.on{background:var(--surface-soft);color:var(--sl);font-weight:700;}
  #analysis-sheet .as-side i b{font-weight:400;color:var(--border-strong);margin-right:5px;}
  #analysis-sheet .as-side i.on b{color:var(--text-muted);}
  #analysis-sheet .as-cards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;min-width:0;}
  /* 시안: 카드 안쪽 여백 10/12 → 16/18, 라운드 11 → 12 */
  #analysis-sheet .as-c{border:1px solid var(--border);border-radius:12px;background:var(--surface);padding:16px 18px;min-width:0;overflow:hidden;}
  #analysis-sheet .as-c.w2{grid-column:span 2;}
  #analysis-sheet .as-c.w3{grid-column:span 3;}
  #analysis-sheet .as-ch{display:flex;align-items:baseline;gap:8px;margin-bottom:13px;}
  #analysis-sheet .as-ct{font-size:13.5px;font-weight:800;letter-spacing:-.005em;color:var(--text);}
  #analysis-sheet .as-csub{font-size:12px;color:var(--mut);font-weight:600;}
  #analysis-sheet .as-cx{margin-left:auto;color:var(--border-strong);font-size:12px;cursor:pointer;padding:0 2px;}
  #analysis-sheet .as-cx:hover{color:var(--red);}
  #analysis-sheet .as-c .as-kpigrid{margin-top:0;}
  #analysis-sheet .as-c .as-note{margin-top:7px;}
  @media(max-width:1100px){
    #analysis-sheet .as-cards{grid-template-columns:repeat(2,minmax(0,1fr));}
    #analysis-sheet .as-c.w3{grid-column:span 2;}
  }
  @media(max-width:760px){
    #analysis-sheet .as-shell{grid-template-columns:minmax(0,1fr);}
    #analysis-sheet .as-side{position:static;}
    #analysis-sheet .as-side i{display:inline-block;margin:2px 3px 2px 0;}
    #analysis-sheet .as-cards{grid-template-columns:minmax(0,1fr);}
    #analysis-sheet .as-c.w2,#analysis-sheet .as-c.w3{grid-column:span 1;}
  }
  @media(max-width:640px){#analysis-sheet .as-g2,#analysis-sheet .as-g3{grid-template-columns:1fr!important;}}`;
  const st = document.createElement('style'); st.id = 'as-style'; st.textContent = css; document.head.appendChild(st);
}
async function loadAnalysisSheet() {
  const host = document.getElementById('analysis-sheet');
  if (!host) return;
  sheetStyle_();
  host.innerHTML = '<div style="padding:44px;color:#78736b;">데이터를 불러오는 중…</div>';
  const { data, error } = await db.from('bookings')
    .select('created_at, booked_at, checkin_date, checkout_date, room_id, channel_id, gross_amount, channel_commission, net_amount, nights, num_guests, guest_country, booking_status, channels(name, code), rooms(room_number)')
    .eq('tenant_id', '00000000-0000-0000-0000-000000000001')
    .order('checkin_date', { ascending: true });
  if (error) { host.innerHTML = `<div style="padding:44px;color:#b85c3f;">로드 실패: ${error.message}</div>`; return; }
  __sheetRows = data || [];
  const months = sheetMonths_();
  if (__sheetView === null || (__sheetView !== 'all' && !months.includes(__sheetView))) {
    const cur = new Date().toISOString().slice(0, 7);
    __sheetView = months.includes(cur) ? cur : (months.length ? months[months.length - 1] : 'all');
  }
  renderSheet();
}
function sheetMonths_() {
  return Array.from(new Set((__sheetRows || []).map(b => revMonth_(b)).filter(Boolean))).sort();
}
function sheetDaysIn_(ym) { const [y, m] = ym.split('-').map(Number); return new Date(y, m, 0).getDate(); }
window.sheetSetView = function (v) { __sheetView = v; __sheetWeek = 'all'; renderSheet(); };
/* [2026-09-01] 객단가 하한선 — 「이 아래로는 안 판다」고 정해둔 선. · 7~8월 실제 판매 데이터로 잡았습니다 (7월 1만원 · 8월 5천원 단위로 조정). · 여기까지 내려오면 99% 이상 팔립니다. 평상시 가격이 아니라 빈 날을 안 남기려는 마지막 선입니다. · 채널에 건 값이 그대로 들어오지 않습니다 — 에어비앤비 막바지 할인 10%, · 부킹닷컴 회원 혜택 15~20%가 더 빠집니다. */
const RATE_FLOOR = {
  '101': [130000, 140000],
  '102': [110000, 120000],
  '201': [140000, 140000],
  '202': [140000, 160000],
  '301': [150000, 150000],
};
const FLOOR_DISCOUNT = { air: [0.10, 0.10], bk: [0.20, 0.15] };   // [최대, 최소] 할인율
// ── 매출 인사이트 카드 목록 (사이드바에서 켜고 끄기 · 선택은 브라우저에 저장) ──
const AS_CARDS = [
  // [2026-08-24] 배치 개편 · '채널별 거래액'은 '채널별 실적'에 통합(거래액 중복 제거)
  { k: 'kpi',   t: '핵심 지표',        w: 3 },
  { k: 'fee',   t: '채널별 실적',      w: 3 },
  { k: 'room',  t: '객실별 ADR·판매율', w: 3 },
  { k: 'cxl',   t: '취소 현황',        w: 1 },
  { k: 'trend', t: '월별 추이',        w: 1 },
  { k: 'book',  t: '예약 건수 추이',   w: 1 },
  { k: 'nts',   t: '판매 숙박 추이',   w: 1 },
  { k: 'band',  t: '가격대별 판매 분포', w: 1 },
  { k: 'wk',    t: '평일 / 주말',      w: 1 },
  { k: 'dow',   t: '요일별 체크인',    w: 1 },
  { k: 'lead',  t: '예약 리드타임',    w: 1 },
  { k: 'nat',   t: '국적 분포',        w: 1 },
  /* [2026-10-15] 「판매 일보」 다시 켬 — 아래로 길어져도 수치를 다 보고 싶다고 하셨습니다. · 숨기려면 아래 한 줄을 다시 주석 처리하면 됩니다. */
  { k: 'daily', t: '판매 일보',        w: 3 },
];
window.sheetSetWeek = function (w) { __sheetWeek = w; renderSheet(); };
window.sheetSetGoal = function (v) {
  let g = v;
  if (g == null) { const el = document.getElementById('as-goal-input'); g = el ? parseInt(el.value.replace(/[^0-9]/g, '')) : 0; }
  if (g > 0 && __sheetView !== 'all') { __sheetGoal[__sheetView] = g; renderSheet(); }
};
// 매출 분석 다운로드 (이미지 PNG / 엑셀 XLSX)
window.sheetExportImage = function (btn) {
  const el = document.getElementById('analysis-sheet'); if (!el) return;
  const orig = btn ? btn.textContent : ''; if (btn) { btn.textContent = '⏳'; btn.disabled = true; }
  ensureHtml2Canvas_()
    .then(() => html2canvas(el, { backgroundColor: '#f5f5f4', scale: 2, useCORS: true, ignoreElements: (n) => n.classList && n.classList.contains('as-noexport') }))
    .then(canvas => { const a = document.createElement('a'); a.href = canvas.toDataURL('image/png'); a.download = '매출분석_' + ((window.__sheetExport && window.__sheetExport.scope) || '').replace(/[^0-9가-힣]/g, '') + '.png'; a.click(); })
    .catch(e => alert('이미지 생성 실패: ' + (e.message || e)))
    .finally(() => { if (btn) { btn.textContent = orig; btn.disabled = false; } });
};
window.sheetExportExcel = function () {
  if (typeof XLSX === 'undefined') { alert('엑셀 라이브러리 로딩 실패. 새로고침 후 다시 시도해주세요.'); return; }
  const d = window.__sheetExport; if (!d) { alert('데이터가 없습니다.'); return; }
  const wb = XLSX.utils.book_new();
  const add = (name, header, rows) => { const ws = XLSX.utils.aoa_to_sheet([header, ...(rows || [])]); XLSX.utils.book_append_sheet(wb, ws, name); };
  add('요약', ['지표', '값'], d.kpi);
  add('채널별', ['채널', '거래액', '건수', '비중(%)'], d.channels);
  add('객실별ADR', ['객실', '숙박일', '점유율(%)', 'ADR', '매출', '+5%월', '+10%월'], d.rooms);
  add('요일', ['구분', 'ADR', '박수', '매출'], d.weekday);
  add('국적', ['국적', '건수'], d.nationality);
  add('판매일보', ['날짜', '판매객실', '가동률(%)', '매출', 'ADR', 'IN', 'OUT'], d.daily);
  XLSX.writeFile(wb, '매출분석_' + ((d.scope) || '').replace(/[^0-9가-힣]/g, '') + '.xlsx');
};
function renderSheet() {
  const host = document.getElementById('analysis-sheet');
  if (!host) return;
  __sheetCharts.forEach(c => { try { c.destroy(); } catch (e) {} });
  __sheetCharts = [];
  const all = __sheetRows || [];
  const months = sheetMonths_();
  const KRW = (n) => (typeof fmtKRW === 'function') ? fmtKRW(n || 0) : ('₩' + Math.round(n || 0).toLocaleString('ko-KR'));
  const man = (n) => '약 ' + Math.round((n || 0) / 10000).toLocaleString('ko-KR') + '만원';
  // 스코프
  const wkOf = (d) => Math.min(5, Math.max(1, Math.ceil(parseInt((d || '').slice(8, 10), 10) / 7)));
  const dOf = s => new Date(s + 'T00:00:00');
  const pad2 = n => String(n).padStart(2, '0');
  const overlapN = (b, s, e) => { if (!b.checkin_date || !b.checkout_date) return 0; const ci = dOf(b.checkin_date), co = dOf(b.checkout_date); const a = ci > s ? ci : s, z = co < e ? co : e; const n = Math.round((z - a) / 86400000); return n > 0 ? n : 0; };
  let scope = all, scopeLabel = '전체 누적', spanDays, winStart, winEnd;
  if (__sheetView !== 'all') {
    scope = all.filter(b => revMonth_(b) === __sheetView);
    const mo = parseInt(__sheetView.split('-')[1], 10);
    const dim = sheetDaysIn_(__sheetView);
    if (__sheetWeek !== 'all') {
      scope = scope.filter(b => wkOf(revDate_(b)) === __sheetWeek);
      scopeLabel = `${mo}월 ${__sheetWeek}주차`;
      const lo = (__sheetWeek - 1) * 7 + 1, hi = Math.min(dim, __sheetWeek * 7);
      spanDays = Math.max(1, hi - lo + 1);
      winStart = dOf(`${__sheetView}-${pad2(lo)}`); winEnd = new Date(winStart); winEnd.setDate(winStart.getDate() + spanDays);
    } else {
      scopeLabel = `${mo}월`;
      spanDays = dim;
      winStart = dOf(`${__sheetView}-01`); winEnd = new Date(winStart); winEnd.setMonth(winEnd.getMonth() + 1);
    }
  } else {
    spanDays = months.reduce((s, m) => s + sheetDaysIn_(m), 0) || 30;
    winStart = months.length ? dOf(months[0] + '-01') : new Date();
    winEnd = new Date(winStart); if (months.length) { winEnd = dOf(months[months.length - 1] + '-01'); winEnd.setMonth(winEnd.getMonth() + 1); }
  }
  const active = scope.filter(b => b.booking_status !== 'cancelled');
  // 점유(숙박일/OCC)는 체크아웃 월과 무관하게 '해당 기간을 실제 점유한 모든 예약' 기준
  const occPool = all.filter(b => b.booking_status !== 'cancelled');
  // [2026-08-24] cachedRooms 가 아직 안 실려 있으면 객실별 표가 '합계'만 나오던 문제.
  //              캐시가 비어 있을 땐 예약 데이터(rooms(room_number))에서 직접 객실 목록을 만듭니다.
  const _rmCache = (typeof cachedRooms !== 'undefined' && cachedRooms) ? cachedRooms.filter(r => r.is_active).map(r => r.room_number) : [];
  const _rmNums = Array.from(new Set((_rmCache.length ? _rmCache : all.map(b => b.rooms && b.rooms.room_number)).filter(Boolean)))
    .sort((a, b) => String(a).localeCompare(String(b)));
  const rooms = _rmNums.map(n => ({ room_number: n }));
  const roomCount = rooms.length || 5;
  const gross = active.reduce((s, b) => s + (b.gross_amount || 0), 0);
  const net = active.reduce((s, b) => s + (b.net_amount || 0), 0);
  const nights = active.reduce((s, b) => s + (b.nights || 0), 0);
  const bookingsN = active.length;
  const adr = nights > 0 ? gross / nights : 0;
  const saleable = Math.max(1, spanDays * roomCount);   // 판매가능 객실박
  // 아래 셋은 '예약 기준' 참고값입니다. 화면에 쓰는 지표는 그 밑의 실제 숙박 기준입니다.
  const occ = Math.min(100, nights / saleable * 100);
  const revpar = gross / saleable;
  const occNights = occPool.reduce((s, b) => s + overlapN(b, winStart, winEnd), 0);
  const occReal = Math.min(100, occNights / saleable * 100);
  const _actOverlap = active.reduce((s, b) => s + overlapN(b, winStart, winEnd), 0);
  const spillN  = Math.max(0, nights - _actOverlap);      // 이 기간 예약인데 다음 달로 넘어가는 밤
  const carryN  = Math.max(0, occNights - _actOverlap);   // 지난 달 예약이 이 기간에 묵은 밤
  const vacantN = Math.max(0, saleable - occNights);      // 공실
  // ── [2026-08-25] 숙박업 셈법으로 통일 ─────────────────────────────
  // 예전에는 '기준일이 이 기간인 예약의 총 박수(nights)'를 분모로 썼습니다.
  // 그 박수에는 다음 달로 넘어가는 밤까지 들어 있어 판매가능 박수를 넘길 수 있었고,
  // 그때 점유율만 100%로 잘려 ADR × 점유율 ≠ RevPAR 이 되었습니다 (그 경고문의 원인).
  // 이제 '이 기간에 실제로 묵은 밤(occNights)' 하나만 분모로 씁니다.
  //  · 실제 숙박은 기간을 넘을 수 없으므로 점유율이 잘릴 일이 없고
  //  · 객실매출도 같은 밤에만 배분하므로 ADR × 점유율 = RevPAR 이 항상 성립합니다.
  const roomRev = occPool.reduce((sum, b) => {
    const n = b.nights || 0; if (n <= 0) return sum;
    return sum + (b.gross_amount || 0) / n * overlapN(b, winStart, winEnd);
  }, 0);
  const adrR    = occNights > 0 ? roomRev / occNights : 0;   // ADR = 객실매출 / 실제 숙박
  const occR    = saleable  > 0 ? occNights / saleable * 100 : 0;
  const revparR = saleable  > 0 ? roomRev / saleable : 0;    // = adrR × occR / 100
  const losR    = bookingsN > 0 ? occNights / bookingsN : 0; // 평균 숙박일
  const cancels = scope.filter(b => b.booking_status === 'cancelled');
  const cancelLost = cancels.reduce((s, b) => s + (b.gross_amount || 0), 0);
  const cancelRate = scope.length > 0 ? cancels.length / scope.length * 100 : 0;
  // 채널
  const chMap = {};
  active.forEach(b => { const n = b.channels?.name || '직접'; const o = chMap[n] || (chMap[n] = { rev: 0, cnt: 0, nights: 0, fee: 0, leadSum: 0, leadN: 0 }); o.rev += (b.gross_amount || 0); o.fee += (b.channel_commission || 0); o.cnt++; o.nights += (b.nights || 0); if (b.created_at && b.checkin_date) { let ld = Math.floor((dOf(b.checkin_date) - new Date(b.created_at)) / 86400000); if (ld < 0) ld = 0; o.leadSum += ld; o.leadN++; } });
  const chCancel = {};
  cancels.forEach(b => { const n = b.channels?.name || '직접'; chCancel[n] = (chCancel[n] || 0) + 1; });
  const chArr = Object.entries(chMap).map(([n, o]) => { const cc = chCancel[n] || 0; return { n, rev: o.rev, cnt: o.cnt, nights: o.nights, fee: o.fee, net: o.rev - o.fee, feeRate: o.rev > 0 ? o.fee / o.rev * 100 : 0, netPerNight: o.nights > 0 ? (o.rev - o.fee) / o.nights : 0, share: gross > 0 ? o.rev / gross * 100 : 0, color: opsChColor(n), adr: o.nights > 0 ? o.rev / o.nights : 0, los: o.cnt > 0 ? o.nights / o.cnt : 0, cancelRate: (o.cnt + cc) > 0 ? cc / (o.cnt + cc) * 100 : 0, avgLead: o.leadN > 0 ? o.leadSum / o.leadN : 0 }; }).sort((a, b) => b.rev - a.rev);
  // 객실별
  const rAgg = {}; rooms.forEach(r => rAgg[r.room_number] = { sold: 0, rev: 0, occN: 0 });
  active.forEach(b => { const rn = b.rooms?.room_number; if (!rn || !rAgg[rn]) return; rAgg[rn].sold += (b.nights || 0); rAgg[rn].rev += (b.gross_amount || 0); });
  occPool.forEach(b => { const rn = b.rooms?.room_number; if (!rn || !rAgg[rn]) return; rAgg[rn].occN += overlapN(b, winStart, winEnd); });
  // 객실별도 전체 KPI와 같은 분모(기준일 예약의 박수)를 사용 → ADR × 점유율 = 객실 RevPAR 성립
  const rStats = rooms.map(r => {
    const a = rAgg[r.room_number];
    return { room: r.room_number, sold: a.sold, occN: a.occN, rev: a.rev,
             adr: a.sold > 0 ? a.rev / a.sold : 0,
             occ: spanDays > 0 ? Math.min(100, a.sold / spanDays * 100) : 0,      // 예약 기준
             occReal: spanDays > 0 ? Math.min(100, a.occN / spanDays * 100) : 0 }; // 참고(실제 점유)
  }).filter(s => s.sold > 0 || s.rev > 0 || s.occN > 0);
  const adrVals = rStats.filter(s => s.adr > 0).map(s => s.adr);
  const minAdr = adrVals.length ? Math.min(...adrVals) : 0, maxAdr = adrVals.length ? Math.max(...adrVals) : 0;
  const maxRoomRev = Math.max(...rStats.map(s => s.rev), 1);
  /* [2026-10-02] 객실 순서를 매출순 → 객실번호순(101·102·201·202·301)으로. · 매출순으로 줄이 매달 바뀌니까 「101호가 어디 있지」를 매번 찾아야 했습니다. · 객실 위치는 고정이라 번호순이 눈에 익습니다. 숫자 계산은 그대로입니다. */
  const roomBySales = [...rStats].sort((a, b) =>
    String(a.room || a.name || '').localeCompare(String(b.room || b.name || ''), 'en', { numeric: true }));
  // 요일
  const wk = { wd: { rev: 0, nts: 0 }, we: { rev: 0, nts: 0 } };
  active.forEach(b => { if (!b.checkin_date) return; const dow = new Date(b.checkin_date + 'T00:00:00').getDay(); const t = (dow === 5 || dow === 6) ? wk.we : wk.wd; t.rev += (b.gross_amount || 0); t.nts += (b.nights || 0); });
  const wdAdr = wk.wd.nts > 0 ? wk.wd.rev / wk.wd.nts : 0, weAdr = wk.we.nts > 0 ? wk.we.rev / wk.we.nts : 0;
  const wePrem = wdAdr > 0 ? (weAdr - wdAdr) / wdAdr * 100 : 0;
  const wkMax = Math.max(wdAdr, weAdr, 1);
  // 요일별 체크인 & 예약 리드타임
  const DOWK = ['일', '월', '화', '수', '목', '금', '토'];
  const dowCnt = [0, 0, 0, 0, 0, 0, 0];
  active.forEach(b => { if (!b.checkin_date) return; dowCnt[dOf(b.checkin_date).getDay()]++; });
  const dowMax = Math.max(...dowCnt, 1);
  const busyDow = dowCnt.indexOf(Math.max(...dowCnt));
  const leadB = [['당일', 0, 0], ['1~3일', 1, 3], ['4~7일', 4, 7], ['8~14일', 8, 14], ['15~30일', 15, 30], ['31일+', 31, 99999]];
  const leadCnt = leadB.map(() => 0); let leadSum = 0, leadN = 0;
  active.forEach(b => { if (!b.created_at || !b.checkin_date) return; let ld = Math.floor((dOf(b.checkin_date) - new Date(b.created_at)) / 86400000); if (ld < 0) ld = 0; leadSum += ld; leadN++; const i = leadB.findIndex(([, lo, hi]) => ld >= lo && ld <= hi); if (i >= 0) leadCnt[i]++; });
  const leadMax = Math.max(...leadCnt, 1); const avgLead = leadN > 0 ? leadSum / leadN : 0;
  // 국적
  const natMap = {};
  let unknownN = 0;
  active.forEach(b => { const c = (b.guest_country || '').toUpperCase(); if (!c) { const chn = b.channels?.name || ''; if (chn.indexOf('에어') >= 0) { unknownN++; return; } natMap['__'] = natMap['__'] || { code: null, n: 0 }; natMap['__'].n++; return; } natMap[c] = natMap[c] || { code: c, n: 0 }; natMap[c].n++; });
  const natFlag = (cc) => cc && /^[A-Z]{2}$/.test(cc) ? `<img src="https://flagcdn.com/w40/${cc.toLowerCase()}.png" alt="${cc}" loading="lazy" style="width:20px;height:14px;border-radius:2px;border:.5px solid #e8e6e1;vertical-align:-2px;margin-right:7px;flex:none;" onerror="this.style.display='none'">` : '<span style="margin-right:7px;">🌐</span>';
  const natName = (cc) => cc ? ((typeof COUNTRY_KO !== 'undefined' && COUNTRY_KO[cc]) ? COUNTRY_KO[cc] : cc) : '미상';
  const natArr = Object.entries(natMap).map(([k, o]) => ({ code: o.code, name: natName(o.code), v: o.n })).sort((a, b) => b.v - a.v).slice(0, 8);
  const natMax = Math.max(...natArr.map(x => x.v), 1);
  const foreignN = active.filter(b => { const c = (b.guest_country || '').toUpperCase(); return c && c !== 'KR'; }).length;
  const knownN = active.filter(b => (b.guest_country || '')).length;
  // 월별 추세 (전체)
  const monthly = months.map(m => {
    const ma = all.filter(b => revMonth_(b) === m && b.booking_status !== 'cancelled');
    return { label: parseInt(m.split('-')[1], 10) + '월', gross: ma.reduce((s, b) => s + (b.gross_amount || 0), 0), cnt: ma.length, nights: ma.reduce((s, b) => s + (b.nights || 0), 0) };
  });
  // 취소 월별 추이 (7·8·9월 비교용) — [2026-08-24] 추가
  const cxlMonthly = months.map(m => {
    const mAll = all.filter(b => revMonth_(b) === m);
    const mCxl = mAll.filter(b => b.booking_status === 'cancelled');
    return {
      key: m,
      label: parseInt(m.split('-')[1], 10) + '월',
      total: mAll.length,
      cnt: mCxl.length,
      rate: mAll.length > 0 ? mCxl.length / mAll.length * 100 : 0,
      lost: mCxl.reduce((s, b) => s + (b.gross_amount || 0), 0),
    };
  }).filter(x => x.total > 0);
  // 목표
  let goal = null;
  if (__sheetView !== 'all') { const mo = parseInt(__sheetView.split('-')[1], 10); goal = (__sheetView in __sheetGoal) ? __sheetGoal[__sheetView] : (mo >= 8 ? 24000000 : null); }
  // ── 전략 ──
  const lowAdrRoom = [...rStats].filter(s => s.adr > 0).sort((a, b) => a.adr - b.adr)[0];
  const allEff5 = Math.round(gross * 0.05);
  const lowAdrEff = lowAdrRoom ? Math.round(lowAdrRoom.rev * 0.05) : 0;
  const cancelRecover = Math.round(cancelLost * 0.3);
  const lowOccRoom = [...rStats].filter(s => s.sold > 0).sort((a, b) => a.occ - b.occ)[0];
  const emptyN = lowOccRoom ? Math.max(0, Math.round(spanDays - lowOccRoom.sold)) : 0;
  const fillN = Math.floor(emptyN * 0.5);
  const fillEff = (lowOccRoom && lowOccRoom.adr > 0) ? Math.round(fillN * lowOccRoom.adr) : 0;
  const topCh = chArr[0];
  const lever = (rank, tone, nm, eff, as, to, ef) => `
    <div class="as-lever" style="border-left-color:${tone};">
      <div class="top"><span class="rank" style="background:${tone};">${rank}</span><span class="nm">${nm}</span>${eff ? `<span class="eff">${eff}</span>` : ''}</div>
      <p><span class="as-as">현황</span> · ${as}</p>
      <p><span class="as-to">실행</span> · ${to}</p>
      <p><b>효과</b> · ${ef}</p>
    </div>`;
  let levers = '';
  if (lowAdrRoom) levers += lever('1순위 · 즉효', '#1f1f23', `객단가 인상 (${RH(lowAdrRoom.room)}·평일)`, `+${KRW(allEff5)} / 월`,
    `${RH(lowAdrRoom.room)} ADR ${KRW(lowAdrRoom.adr)}(최저) · 평일 ADR ${KRW(wdAdr)} (주말 대비 ${wePrem >= 0 ? '−' : '+'}${Math.abs(wePrem).toFixed(0)}%)`,
    `${RH(lowAdrRoom.room)} +10%, 평일가 +5% → 전 객실 ADR +5%`,
    `${RH(lowAdrRoom.room)} 월 <b>+${KRW(lowAdrEff)}</b> · 전 객실 +5% = <b>+${KRW(allEff5)}</b> → 거래액 <b>${KRW(gross + allEff5)}</b>`);
  if (cancelLost > 0) levers += lever('2순위 · 즉효', '#1f1f23', '취소 방어', `손실 ${KRW(cancelLost)} 회수`,
    `취소율 ${cancelRate.toFixed(0)}% · 손실 ${KRW(cancelLost)}`, `비환불(선결제) 요금제↑ · 취소 즉시 재판매 알림`,
    `취소분의 30%만 재판매해도 <b>+${KRW(cancelRecover)}</b>`);
  if (lowOccRoom && fillEff > 0) levers += lever('3순위 · 즉효', '#1f1f23', `${RH(lowOccRoom.room)} 공실 채우기`, `+${KRW(fillEff)}`,
    `${RH(lowOccRoom.room)} 점유 ${lowOccRoom.occ.toFixed(0)}%(최저) · 빈 ${emptyN}박 · ADR ${KRW(lowOccRoom.adr)}`,
    `막판가·최소박 완화 · 평일 프로모션 (인상보다 채우기 우선)`, `빈 ${emptyN}박 중 ${fillN}박만 채워도 <b>+${KRW(fillEff)}</b>`);
  if (topCh && topCh.share > 55) levers += lever('4순위 · 구조', '#78736b', '채널 다변화', '',
    `${topCh.n} ${topCh.share.toFixed(0)}% 편중 · 채널 정책 리스크`, `타 채널·직접예약 노출↑ → 의존도 60%대로`,
    `리스크 분산 + 직접예약↑ 시 <b>정산액 개선</b>`);
  const monthOpts = `<option value="all"${__sheetView === 'all' ? ' selected' : ''}>전체 누적</option>` +
    months.map(m => `<option value="${m}"${__sheetView === m ? ' selected' : ''}>${parseInt(m.split('-')[1], 10)}월</option>`).join('');
  /* [2026-09-18] 다크에서 기간·기준 드롭다운의 글씨가 안 보였습니다. · color 가 #1f1f23 (거의 검정) 으로 박혀 있어서 어두운 배경에 검은 글씨였습니다. · 테두리도 밝은 회색 고정이라 상자만 덩그러니 보였습니다. */
  const _selCss = 'border:1px solid var(--border-strong);border-radius:9px;padding:8px 30px 8px 13px;'
    + 'font-size:14px;font-weight:700;background:var(--surface);color:var(--text);cursor:pointer;';
  const basisOpts = `<option value="checkin_date"${REV_BASIS === 'checkin_date' ? ' selected' : ''}>체크인 기준</option>` +
                    `<option value="checkout_date"${REV_BASIS === 'checkout_date' ? ' selected' : ''}>체크아웃 기준</option>`;
  const tabsHtml = `<span style="font-size:13px;color:var(--mut);font-weight:600;">기간</span>` +
    `<select onchange="sheetSetView(this.value)" style="${_selCss}">${monthOpts}</select>` +
    `<span style="font-size:13px;color:var(--mut);font-weight:600;margin-left:6px;">기준</span>` +
    `<select onchange="setRevBasis(this.value)" title="매출을 체크인 날짜로 잡을지, 체크아웃 날짜로 잡을지" style="${_selCss}">${basisOpts}</select>` +
    `<span style="margin-left:auto;font-size:11.5px;color:#a8a39a;">2026 · ${REV_BASIS_LABEL}</span>`;
  let weekTabsHtml = '';
  if (__sheetView !== 'all') {
    const wset = Array.from(new Set(all.filter(b => revMonth_(b) === __sheetView && b.booking_status !== 'cancelled').map(b => wkOf(revDate_(b))))).sort((a, b) => a - b);
    if (wset.length) {
      weekTabsHtml = `<div class="as-subtabs"><button class="as-subtab ${__sheetWeek === 'all' ? 'on' : ''}" onclick="sheetSetWeek('all')">월 전체</button>` +
        wset.map(w => `<button class="as-subtab ${__sheetWeek === w ? 'on' : ''}" onclick="sheetSetWeek(${w})">${w}주차</button>`).join('') + `</div>`;
    }
  }
  const goalPct = goal ? Math.min(100, gross / goal * 100) : 0;
  const goalHtml = goal ? `
      <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:10px;">
        <div>
          <div style="font-size:23px;font-weight:800;letter-spacing:-.02em;">${goalPct.toFixed(1)}<span style="font-size:14px;color:var(--mut);">%</span></div>
          <div style="font-size:10.5px;color:var(--mut);font-weight:600;">취소 제외</div>
        </div>
        <div style="text-align:right;">
          <div style="font-size:10.5px;color:var(--mut);margin-bottom:4px;">목표 금액</div>
          <div style="display:flex;gap:6px;justify-content:flex-end;flex-wrap:wrap;align-items:center;">
            <input id="as-goal-input" value="${goal.toLocaleString()}" style="width:112px;padding:6px 9px;border:1px solid #d6d3cd;border-radius:8px;font-size:13px;text-align:right;color:#1f2937;">
            <button onclick="sheetSetGoal()" style="background:var(--accent);color:#fff;border:0;border-radius:8px;padding:6px 13px;font-size:12px;font-weight:700;cursor:pointer;">설정</button>
          </div>
          <div style="display:flex;gap:5px;justify-content:flex-end;margin-top:6px;">
            <span onclick="sheetSetGoal(20000000)" style="cursor:pointer;font-size:11px;padding:4px 9px;border-radius:7px;background:${goal===20000000?'#1f1f23':'#f5f5f4'};color:${goal===20000000?'#fff':'#57534c'};font-weight:${goal===20000000?'700':'400'};">2,000만</span>
            <span onclick="sheetSetGoal(24000000)" style="cursor:pointer;font-size:11px;padding:4px 9px;border-radius:7px;background:${goal===24000000?'#1f1f23':'#f5f5f4'};color:${goal===24000000?'#fff':'#57534c'};font-weight:${goal===24000000?'700':'400'};">2,400만</span>
            <span onclick="sheetSetGoal(30000000)" style="cursor:pointer;font-size:11px;padding:4px 9px;border-radius:7px;background:${goal===30000000?'#1f1f23':'#f5f5f4'};color:${goal===30000000?'#fff':'#57534c'};font-weight:${goal===30000000?'700':'400'};">3,000만</span>
          </div>
        </div>
      </div>
      <div style="height:9px;background:#e8e6e1;border-radius:6px;overflow:hidden;margin:11px 0 7px;">
        <div style="width:${Math.min(100, goalPct)}%;height:100%;background:var(--accent);border-radius:6px;"></div>
      </div>
      <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;">
        <div><div style="font-size:10px;color:var(--mut);font-weight:700;">현재 거래액</div><div style="font-size:13px;font-weight:800;">${KRW(gross)}</div></div>
        <div style="text-align:right;"><div style="font-size:10px;color:var(--mut);font-weight:700;">목표까지</div><div style="font-size:13px;font-weight:800;color:${(goal - gross) > 0 ? '#57534c' : '#4a8a3d'};">${(goal - gross) > 0 ? (KRW(goal - gross) + ' 남음') : '달성'}</div></div>
      </div>` : '';
  // ✦ AI 인사이트 (실데이터 규칙 기반 자동 문장)
  const hiAdrRoom = [...rStats].filter(s => s.adr > 0).sort((a, b) => b.adr - a.adr)[0];
  const ins = [];
  if (roomBySales[0] && roomBySales[0].rev > 0) ins.push(['🏆', `<b>${RH(roomBySales[0].room)}가 매출 1위</b> — 전체 거래액의 <b>${gross > 0 ? Math.round(roomBySales[0].rev / gross * 100) : 0}%</b>(${KRW(roomBySales[0].rev)}) · 점유 ${roomBySales[0].occ.toFixed(0)}%.`]);
  if (hiAdrRoom) ins.push(['💎', `<b>${RH(hiAdrRoom.room)} ADR 최고</b> ${KRW(hiAdrRoom.adr)}/박${hiAdrRoom.occ < occR ? ` — 점유 ${hiAdrRoom.occ.toFixed(0)}%로 평균 이하라 가격 방어 vs 점유 회복 판단 필요.` : '.'}`]);
  if (lowOccRoom) ins.push(['⚠️', `<b>${RH(lowOccRoom.room)} 점유 ${lowOccRoom.occ.toFixed(0)}%로 최저</b> — 막판가·최소박 완화로 공실 채우기 대상.`]);
  if (topCh) ins.push(['🔗', `<b>${topCh.n}</b>이 거래액의 <b>${topCh.share.toFixed(0)}%</b> — ${topCh.share > 55 ? '채널 편중, 타 채널·직접예약 다변화 필요.' : '채널 분산 양호.'}`]);
  if (wdAdr > 0 && weAdr > 0) ins.push(['📈', `주말 ADR이 평일 대비 <b>${wePrem >= 0 ? '+' : ''}${wePrem.toFixed(0)}%</b> — ${wePrem < 8 ? '주말 프리미엄이 약함, 주말가 상향 여력.' : '주말 프리미엄 정상 작동.'}`]);
  ins.push(['📅', `체크인은 <b>${DOWK[busyDow]}요일</b>에 최다(${dowCnt[busyDow]}건) — 가격·청소 인력 배치 참고.`]);
  if (avgLead > 0) ins.push(['⏱️', `평균 <b>${avgLead.toFixed(0)}일 전</b> 예약 — ${avgLead < 7 ? '막판 예약 비중이 커 당일가 관리가 중요.' : '조기예약 비중, 얼리버드 프로모션 여력.'}`]);
  if (cancelLost > 0) ins.push(['🛑', `취소율 <b>${cancelRate.toFixed(0)}%</b> · 손실 <b>${KRW(cancelLost)}</b> — 비환불 요금제·즉시 재판매로 방어.`]);
  const aiHtml = `
    <div class="as-sec" style="margin-top:18px;">✦ AI 인사이트 <span style="color:var(--mut);font-weight:600;font-size:12px;">· ${scopeLabel} 자동 분석</span></div>
    <div class="as-card" style="background:linear-gradient(180deg,#fbfcfe,#fff);">
      ${ins.map(([ic, t]) => `<div style="display:flex;gap:10px;align-items:flex-start;padding:7px 0;font-size:13px;line-height:1.55;"><span style="font-size:15px;flex:none;">${ic}</span><span>${t}</span></div>`).join('')}
    </div>`;
  const roomRows = rStats.map(s => {
    const tag = s.adr === minAdr && s.adr > 0 ? ' <span class="as-pill" style="background:#fbeae3;color:var(--red);">최저가</span>' : (s.adr === maxAdr && s.adr > 0 ? ' <span class="as-pill" style="background:#eae7fb;color:var(--ago);">최고가</span>' : '');
    return `<tr><td>${RH(s.room)}${tag}</td><td>${s.sold}</td><td>${s.occ.toFixed(0)}%</td><td>${KRW(s.adr)}</td><td>${KRW(s.rev)}</td><td class="as-pos">+${KRW(s.rev * 0.05)}</td><td class="as-pos">+${KRW(s.rev * 0.10)}</td></tr>`;
  }).join('');
  /* [2026-09-01] 객단가 하한선 표 — 그 달 실제 ADR 과 나란히 놓아 · 하한선을 지켰는지, 올릴 여지가 있는지 한눈에 보이게 합니다. */
  const _fw = (a, b) => a === b ? KRW(a) : KRW(a) + ' ~ ' + KRW(b);
  const _rmAdr = {}; rStats.forEach(x => { _rmAdr[String(x.room)] = x.adr; });
  const floorRows = Object.keys(RATE_FLOOR).sort().map(rm => {
    const [lo, hi] = RATE_FLOOR[rm];
    const air = [Math.round(lo * (1 - FLOOR_DISCOUNT.air[0]) / 1000) * 1000,
                 Math.round(hi * (1 - FLOOR_DISCOUNT.air[1]) / 1000) * 1000];
    const bk  = [Math.round(lo * (1 - FLOOR_DISCOUNT.bk[0]) / 1000) * 1000,
                 Math.round(hi * (1 - FLOOR_DISCOUNT.bk[1]) / 1000) * 1000];
    const cur = _rmAdr[rm] || 0;
    const gap = cur ? cur - lo : null;
    const gapHtml = gap === null ? '<span style="color:var(--mut);">—</span>'
      : gap >= 0 ? `<span class="as-pos">+${KRW(gap)}</span>`
                 : `<span style="color:var(--red);">하한선 아래</span>`;
    return `<tr><td style="font-weight:700;">${RH(rm)}</td><td>${_fw(lo, hi)}</td>`
         + `<td>${_fw(air[0], air[1])}</td><td>${_fw(bk[0], bk[1])}</td>`
         + `<td>${cur ? KRW(cur) : '—'}</td><td>${gapHtml}</td></tr>`;
  }).join('');
  const floorTable = `
      <div style="height:14px;"></div>
      <div class="as-sec" style="margin:0 0 3px;">객단가 하한선 <span style="color:var(--mut);font-weight:600;font-size:11.5px;">· 리드타임 3일 기준</span></div>
      <div style="font-size:11px;color:var(--mut);margin-bottom:7px;">7~8월 실제 판매 데이터로 잡은 선 — 7월 1만원, 8월 5천원 단위로 조정</div>
      <div style="overflow-x:auto;">
      <table class="as-tbl"><thead><tr><th>객실</th><th>CMS 하한선</th><th>에어비앤비 −10%</th><th>부킹닷컴 −15~20%</th><th>이 기간 ADR</th><th>여유</th></tr></thead>
      <tbody>${floorRows}</tbody></table>
      </div>
      <div class="as-note">· <b>하한선</b>은 「이 아래로는 안 판다」고 정해둔 선입니다. 여기까지 내려오면 99% 이상 팔립니다.<br>
      · <b>여유</b>가 「하한선 아래」면 그 달 실제 ADR 이 하한선보다 낮았다는 뜻 — 하한선이 안 지켜졌거나 너무 높게 잡힌 것입니다.<br>
      · 채널에 건 값이 그대로 들어오지 않습니다. 에어비앤비 막바지 할인 10%, 부킹닷컴 회원 혜택 15~20%가 더 빠집니다.</div>`;
  // 전월 대비 델타 (특정 월 · 월 전체일 때만)
  let mom = null;
  if (__sheetView !== 'all' && __sheetWeek === 'all') {
    const idx = months.indexOf(__sheetView);
    if (idx > 0) {
      const pm = months[idx - 1];
      const pa = all.filter(b => revMonth_(b) === pm && b.booking_status !== 'cancelled');
      const pGross = pa.reduce((s, b) => s + (b.gross_amount || 0), 0);
      const pNights = pa.reduce((s, b) => s + (b.nights || 0), 0);
      const pDays = sheetDaysIn_(pm);
      const pWs = dOf(pm + '-01'); const pWe = new Date(pWs); pWe.setMonth(pWe.getMonth() + 1);
      const pSaleable = Math.max(1, pDays * roomCount);
      // 전월도 같은 셈법(실제 숙박 기준)으로 맞춥니다
      const pOccN = occPool.reduce((t, b) => t + overlapN(b, pWs, pWe), 0);
      const pRoomRev = occPool.reduce((t, b) => {
        const n = b.nights || 0; if (n <= 0) return t;
        return t + (b.gross_amount || 0) / n * overlapN(b, pWs, pWe);
      }, 0);
      mom = { label: parseInt(pm.split('-')[1], 10) + '월', gross: pGross, bookings: pa.length,
              nights: pOccN,
              adr: pOccN > 0 ? pRoomRev / pOccN : 0,
              occ: pOccN / pSaleable * 100,
              revpar: pRoomRev / pSaleable,
              los: pa.length > 0 ? pOccN / pa.length : 0 };
    }
  }
  // 이번 달·미래 달은 아직 파는 중 → 완료된 달과 증감률로 비교하지 않음
  const _nowYm = (() => { const t = new Date(); return t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0'); })();
  const _inProgress = (__sheetView !== 'all') && (__sheetView >= _nowYm);
  const dPct = (cur, prev) => { if (_inProgress || !mom || !prev) return ''; const p = (cur - prev) / Math.abs(prev) * 100; const cls = p > 0.5 ? 'up' : p < -0.5 ? 'down' : 'flat'; const ar = p > 0.5 ? '▲' : p < -0.5 ? '▼' : '—'; return `<span class="as-delta ${cls}">${ar} ${Math.abs(p).toFixed(0)}% <span style="opacity:.55;font-weight:600;">vs ${mom.label}</span></span>`; };
  const dPt = (cur, prev) => { if (_inProgress || !mom || prev == null) return ''; const d = cur - prev; const cls = d > 0.5 ? 'up' : d < -0.5 ? 'down' : 'flat'; const ar = d > 0.5 ? '▲' : d < -0.5 ? '▼' : '—'; return `<span class="as-delta ${cls}">${ar} ${Math.abs(d).toFixed(0)}p <span style="opacity:.55;font-weight:600;">vs ${mom.label}</span></span>`; };
  // 스파크라인 (월별 거래액)
  let sparkSvg = '';
  { const vals = monthly.map(x => x.gross); if (vals.length >= 2) { const mx = Math.max(...vals, 1), mn = Math.min(...vals, 0), W = 150, H = 46; const pts = vals.map((v, i) => { const x = i / (vals.length - 1) * W; const y = H - ((v - mn) / ((mx - mn) || 1)) * (H - 8) - 4; return x.toFixed(0) + ',' + y.toFixed(0); }).join(' '); const ci = __sheetView !== 'all' ? months.indexOf(__sheetView) : vals.length - 1; const cx = ci >= 0 ? ci / (vals.length - 1) * W : 0; const cy = ci >= 0 ? H - ((vals[ci] - mn) / ((mx - mn) || 1)) * (H - 8) - 4 : 0; sparkSvg = `<svg width="150" height="46" viewBox="0 0 150 46"><polyline points="${pts}" fill="none" stroke="#3b6db5" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/><circle cx="${cx.toFixed(0)}" cy="${cy.toFixed(0)}" r="3.5" fill="#3b6db5"/></svg>`; } }
  // 핵심 한 줄 인사이트
  // [2026-08-24] 자동 코멘트('OO호가 매출을 이끌고…', '+10% 인상 시 월 …') 제거 — 표에 숫자만 둡니다.
  // 📋 판매 일보 — 일자별 판매객실·가동률·매출(일할 배분)·ADR·IN/OUT
  const DOW_K = ['일', '월', '화', '수', '목', '금', '토'];
  // 매출은 쪼개지 않고 기준일(체크인 또는 체크아웃)에 예약 전액을 계상 → 합계가 상단 거래액과 정확히 일치
  const revByDate = (() => {
    const m = Object.create(null);
    for (const b of active) {
      const k = revDate_(b); if (!k) continue;
      const nn = b.nights || Math.max(1, Math.round((dOf(b.checkout_date) - dOf(b.checkin_date)) / 86400000));
      if (!m[k]) m[k] = { rev: 0, nts: 0, cnt: 0 };
      m[k].rev += (b.gross_amount || 0); m[k].nts += nn; m[k].cnt++;
    }
    return m;
  })();
  const dailyRows = (() => {
    const out = [], lim = new Date(winEnd);
    for (let d = new Date(winStart); d < lim; d.setDate(d.getDate() + 1)) {
      const cur = new Date(d);
      const ymd = cur.getFullYear() + '-' + pad2(cur.getMonth() + 1) + '-' + pad2(cur.getDate());
      let sold = 0, cin = 0, cout = 0;
      for (const b of occPool) {
        if (!b.checkin_date || !b.checkout_date) continue;
        if (b.checkin_date <= ymd && ymd < b.checkout_date) sold++;
        if (b.checkin_date === ymd) cin++;
        if (b.checkout_date === ymd) cout++;
      }
      const rv = revByDate[ymd] || { rev: 0, nts: 0, cnt: 0 };
      out.push({ ymd, dow: cur.getDay(), sold, cin, cout, rev: rv.rev, revNts: rv.nts, revCnt: rv.cnt });
    }
    return out;
  })();
  const dlTot = dailyRows.reduce((a, r) => ({ sold: a.sold + r.sold, rev: a.rev + r.rev, revNts: a.revNts + r.revNts, revCnt: a.revCnt + r.revCnt, cin: a.cin + r.cin, cout: a.cout + r.cout }), { sold: 0, rev: 0, revNts: 0, revCnt: 0, cin: 0, cout: 0 });
  const dlCap = dailyRows.length * roomCount;
  const _todayYmd = (() => { const t = new Date(); return t.getFullYear() + '-' + pad2(t.getMonth() + 1) + '-' + pad2(t.getDate()); })();
  // ── 수수료 ──
  const feeTot = active.reduce((s, b) => s + Number(b.channel_commission || 0), 0);
  const netTot = gross - feeTot;
  const feeRate = gross > 0 ? feeTot / gross * 100 : 0;
  const netPerNight = nights > 0 ? netTot / nights : 0;
  // ── 가격대별 판매 분포 (박당 요금) ──
  const BANDS = [[0, 100000, '10만 미만'], [100000, 130000, '10~13만'], [130000, 160000, '13~16만'],
                 [160000, 200000, '16~20만'], [200000, 250000, '20~25만'], [250000, 1e12, '25만 이상']];
  const bandRows = BANDS.map(([lo, hi, label]) => ({ label, lo, hi, cnt: 0, nts: 0, rev: 0 }));
  active.forEach(b => {
    const nn = Number(b.nights || 0); if (nn <= 0) return;
    const per = Number(b.gross_amount || 0) / nn;
    const r = bandRows.find(x => per >= x.lo && per < x.hi);
    if (r) { r.cnt++; r.nts += nn; r.rev += Number(b.gross_amount || 0); }
  });
  const bandMaxN = Math.max(1, ...bandRows.map(r => r.nts));
  const bandTop = bandRows.slice().sort((a, b) => b.nts - a.nts)[0];
  // ── KPI 미니 스파크라인 (월별 6개) ──
  const spark_ = (vals, cur) => {
    const v = vals.filter(x => x != null);
    if (v.length < 2) return '';
    const mx = Math.max(...v), mn = Math.min(...v, 0), W = 52, H = 15;
    const pts = vals.map((x, i) => {
      const y = H - ((x - mn) / ((mx - mn) || 1)) * (H - 3) - 1.5;
      return (i / (vals.length - 1) * W).toFixed(1) + ',' + y.toFixed(1);
    }).join(' ');
    const ci = (cur != null && cur >= 0 && cur < vals.length) ? cur : vals.length - 1;
    const cx = ci / (vals.length - 1) * W;
    const cy = H - ((vals[ci] - mn) / ((mx - mn) || 1)) * (H - 3) - 1.5;
    return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" style="vertical-align:-2px;margin-left:5px;"><polyline points="${pts}" fill="none" stroke="#a8a39a" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/><circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="2" fill="#1f1f23"/></svg>`;
  };
  const _mIdx = (__sheetView !== 'all') ? months.indexOf(__sheetView) : (monthly.length - 1);
  const spCnt = spark_(monthly.map(x => x.cnt), _mIdx);
  const spAdr = spark_(monthly.map(x => x.nights > 0 ? Math.round(x.gross / x.nights) : 0), _mIdx);
  const spNts = spark_(monthly.map(x => x.nights), _mIdx);
  // ── 한 줄 분석 문장 ──
  // [2026-08-24] 자동 서술('거래액이 지난달보다 …', 'OO월은 아직 예약을 받는 중이라 …') 제거.
  window.__sheetExport = {
    scope: scopeLabel,
    kpi: [['거래액(GMV)', gross], ['정산액', net], ['객실매출', Math.round(roomRev)], ['예약건수', bookingsN], ['숙박일수', occNights], ['평균숙박일', +losR.toFixed(2)], ['판매가능박수', saleable], ['ADR', Math.round(adrR)], ['OCC(%)', Math.round(occR)], ['RevPAR', Math.round(revparR)], ['취소건', cancels.length], ['취소손실', cancelLost]],
    channels: chArr.map(c => [c.n, c.rev, c.cnt, Number(c.share.toFixed(1))]),
    rooms: rStats.map(s => [RH(s.room), s.occN, Math.round(s.occ), Math.round(s.adr), s.rev, Math.round(s.rev * 0.05), Math.round(s.rev * 0.10)]),
    weekday: [['평일', Math.round(wdAdr), wk.wd.nts, wk.wd.rev], ['주말(금·토)', Math.round(weAdr), wk.we.nts, wk.we.rev]],
    nationality: natArr.map(x => [x.name, x.v]),
    daily: dailyRows.map(r => [r.ymd, r.sold, roomCount > 0 ? Math.round(r.sold / roomCount * 100) : 0, Math.round(r.rev), r.revNts > 0 ? Math.round(r.rev / r.revNts) : 0, r.cin, r.cout]),
  };
  // ── 카드 내용 ──
  const CARD = {};
  CARD.kpi = `
      <div style="display:flex;align-items:flex-end;gap:16px;flex-wrap:wrap;margin-bottom:10px;">
        <div style="min-width:0;">
          <div style="font-size:11px;font-weight:700;color:var(--mut);">거래액 (GMV) · ${scopeLabel}<span class="as-basis">${REV_BASIS_LABEL}</span></div>
          <div style="font-size:29px;font-weight:800;letter-spacing:-.03em;line-height:1.1;margin-top:3px;">${KRW(gross)}</div>
          <div style="display:flex;align-items:center;gap:8px;margin-top:5px;flex-wrap:wrap;font-size:11px;color:var(--mut);">${dPct(gross, mom && mom.gross)}<span>${man(gross)}</span></div>
          <div style="font-size:11px;color:var(--mut);margin-top:4px;">객실매출 <b style="color:var(--ink);font-weight:700;">${KRW(roomRev)}</b> — 이 기간에 묵은 밤만 배분 (ADR·RevPAR 기준)</div>
        </div>
        ${sparkSvg ? `<div style="margin-left:auto;text-align:right;"><div style="font-size:10px;color:#78736b;font-weight:700;margin-bottom:2px;">월별 추이</div>${sparkSvg}</div>` : ''}
      </div>
      <div class="as-kpigrid">
        <div class="as-kpi"><div class="k">예약 건수</div><div class="v" style="font-size:19px;">${bookingsN}건${spCnt}</div><div class="d">${REV_BASIS_LABEL}</div>${dPct(bookingsN, mom && mom.bookings) ? `<div style="margin-top:5px;">${dPct(bookingsN, mom.bookings)}</div>` : ''}</div>
        <div class="as-kpi"><div class="k">숙박일수</div><div class="v" style="font-size:19px;">${occNights}박${spNts}</div><div class="d">이 기간에 실제로 묵은 밤 · 공실 <b style="color:${vacantN > 0 ? '#a37e2b' : '#4a8a3d'};">${vacantN}박</b></div>${dPct(occNights, mom && mom.nights) ? `<div style="margin-top:5px;">${dPct(occNights, mom.nights)}</div>` : ''}</div>
        <div class="as-kpi"><div class="k">평균 숙박</div><div class="v" style="font-size:19px;">${losR.toFixed(2)}박</div><div class="d">숙박일수 ÷ 예약 건수</div>${dPct(losR, mom && mom.los) ? `<div style="margin-top:5px;">${dPct(losR, mom.los)}</div>` : ''}</div>
        <div class="as-kpi"><div class="k">ADR (박당)</div><div class="v" style="font-size:19px;">${KRW(adrR)}${spAdr}</div><div class="d">객실매출 ÷ ${occNights}박</div>${dPct(adrR, mom && mom.adr) ? `<div style="margin-top:5px;">${dPct(adrR, mom.adr)}</div>` : ''}</div>
        <div class="as-kpi"><div class="k">점유율 (OCC)</div><div class="v" style="font-size:19px;">${occR.toFixed(0)}%</div><div class="d">${occNights} / ${saleable}박</div>${dPt(occR, mom && mom.occ) ? `<div style="margin-top:5px;">${dPt(occR, mom.occ)}</div>` : ''}</div>
        <div class="as-kpi"><div class="k">RevPAR</div><div class="v" style="font-size:19px;">${KRW(revparR)}</div><div class="d">ADR × 점유율</div>${dPct(revparR, mom && mom.revpar) ? `<div style="margin-top:5px;">${dPct(revparR, mom.revpar)}</div>` : ''}</div>
      </div>
      <!-- [2026-08-25] RevPAR 계산 불일치 경고 삭제 — 분모를 실제 숙박으로 통일해 어긋날 일이 없습니다 -->`;
  // 채널이 2개 이상이고 가장 작은 채널도 5% 넘을 때만 도넛 (97:2 같은 경우엔 막대가 낫다)
  const _showDonut = chArr.length >= 2 && Math.min.apply(null, chArr.map(c => c.share)) >= 5;
  CARD.ch = `
      ${_showDonut ? `<div style="display:flex;align-items:center;gap:12px;margin-bottom:11px;padding-bottom:10px;border-bottom:1px solid #f5f5f4;">
        <div style="width:94px;height:94px;flex:none;position:relative;"><canvas id="asChDonut"></canvas></div>
        <div style="flex:1;min-width:0;">
          ${chArr.map(c => `<div style="display:flex;align-items:center;gap:6px;font-size:11.5px;margin:4px 0;min-width:0;">
            <span style="width:9px;height:9px;border-radius:2px;background:${c.color};flex:none;"></span>
            <span style="font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${c.n}</span>
            <span style="margin-left:auto;font-weight:800;white-space:nowrap;">${KRW(c.rev)}</span>
            <span style="width:42px;flex:none;text-align:right;color:var(--mut);font-weight:700;">${c.share.toFixed(1)}%</span>
          </div>`).join('')}
        </div>
      </div>` : ''}
      ${chArr.map((c, i) => `
        <div style="padding:${i ? '9px' : '0'} 0 9px;${i < chArr.length - 1 ? 'border-bottom:1px solid #f5f5f4;' : ''}">
          <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:5px;min-width:0;">
            <div style="display:flex;align-items:center;gap:6px;min-width:0;">${channelLogo_(c.n)}<span style="font-size:10.5px;color:#a8a39a;white-space:nowrap;">${c.cnt}건</span></div>
            ${_showDonut ? '' : `<div style="text-align:right;white-space:nowrap;"><span style="font-size:14px;font-weight:800;">${KRW(c.rev)}</span> <span style="font-size:11px;color:var(--mut);font-weight:700;">${c.share.toFixed(1)}%</span></div>`}
          </div>
          ${_showDonut ? '' : `<div style="height:7px;background:var(--surface-soft);border-radius:5px;overflow:hidden;margin-bottom:6px;"><div style="width:${Math.max(2, c.share)}%;height:100%;background:${c.color};border-radius:5px;"></div></div>`}
          <div style="display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:4px;">
            <div style="background:#fafaf9;border-radius:6px;padding:4px 6px;min-width:0;"><div style="font-size:9px;color:#78736b;font-weight:700;">ADR</div><div style="font-size:11px;font-weight:800;">${KRW(c.adr)}</div></div>
            <div style="background:#fafaf9;border-radius:6px;padding:4px 6px;min-width:0;"><div style="font-size:9px;color:#78736b;font-weight:700;">평균숙박</div><div style="font-size:11px;font-weight:800;">${c.los.toFixed(1)}박</div></div>
            <div style="background:#fafaf9;border-radius:6px;padding:4px 6px;min-width:0;"><div style="font-size:9px;color:#78736b;font-weight:700;">취소율</div><div style="font-size:11px;font-weight:800;${c.cancelRate > 0 ? 'color:var(--red);' : ''}">${c.cancelRate.toFixed(0)}%</div></div>
            <div style="background:#fafaf9;border-radius:6px;padding:4px 6px;min-width:0;"><div style="font-size:9px;color:#78736b;font-weight:700;">리드</div><div style="font-size:11px;font-weight:800;">${c.avgLead.toFixed(0)}일</div></div>
          </div>
        </div>`).join('')}
      ${topCh && topCh.share > 55 ? `<div class="as-note">${topCh.n} 의존도 ${topCh.share.toFixed(1)}%</div>` : ''}`;
  const _feeSorted = chArr.slice().sort((a, b) => b.netPerNight - a.netPerNight);
  const _best = _feeSorted[0], _worst = _feeSorted[_feeSorted.length - 1];
  CARD.fee = `
      <div class="as-kpigrid" style="grid-template-columns:repeat(auto-fit,minmax(112px,1fr));margin-bottom:10px;">
        <div class="as-kpi"><div class="k">거래액</div><div class="v" style="font-size:17px;">${KRW(gross)}</div></div>
        <div class="as-kpi"><div class="k">수수료</div><div class="v" style="font-size:17px;color:var(--red);">−${KRW(feeTot)}</div><div class="d">${feeRate.toFixed(1)}%</div></div>
        <div class="as-kpi"><div class="k">실수령</div><div class="v" style="font-size:17px;">${KRW(netTot)}</div><div class="d">${(100 - feeRate).toFixed(1)}%</div></div>
        <div class="as-kpi"><div class="k">박당 실수령</div><div class="v" style="font-size:17px;">${KRW(netPerNight)}</div><div class="d">ADR ${KRW(adr)} 중</div></div>
      </div>
      <div style="overflow-x:auto;">
      <table class="as-tbl"><thead><tr><th>채널</th><th>건수</th><th>박수</th><th>비중</th><th>거래액</th><th>ADR</th><th>평균숙박</th><th>수수료</th><th>수수료율</th><th>실수령</th><th>박당 실수령</th><th>취소율</th><th>리드</th></tr></thead>
      <tbody>${chArr.map(c => `<tr>
        <td>${channelLogo_(c.n)}</td>
        <td>${c.cnt}</td><td>${c.nights}</td>
        <td>${c.share.toFixed(1)}%</td>
        <td>${KRW(c.rev)}</td>
        <td>${KRW(c.adr)}</td><td>${c.los.toFixed(1)}박</td>
        <td style="color:var(--red);">−${KRW(c.fee)}</td>
        <td style="font-weight:800;color:${c.feeRate >= 20 ? 'var(--red)' : (c.feeRate >= 18 ? '#a37e2b' : 'var(--grn)')};">${c.feeRate.toFixed(2)}%</td>
        <td><b>${KRW(c.net)}</b></td>
        <td style="font-weight:800;${_best && c.n === _best.n ? 'color:var(--grn);' : (_worst && c.n === _worst.n && chArr.length > 1 ? 'color:var(--red);' : '')}">${KRW(c.netPerNight)}${_best && c.n === _best.n ? ' ↑' : (_worst && c.n === _worst.n && chArr.length > 1 ? ' ↓' : '')}</td>
        <td style="${c.cancelRate > 0 ? 'color:var(--red);' : ''}">${c.cancelRate.toFixed(0)}%</td>
        <td>${c.avgLead.toFixed(0)}일</td>
      </tr>`).join('')}
      <tr class="tot"><td>합계</td><td>${bookingsN}</td><td>${nights}</td><td>100%</td><td>${KRW(gross)}</td><td>${KRW(adr)}</td><td>${bookingsN > 0 ? (nights / bookingsN).toFixed(1) : '0.0'}박</td><td>−${KRW(feeTot)}</td><td>${feeRate.toFixed(2)}%</td><td>${KRW(netTot)}</td><td>${KRW(netPerNight)}</td><td>${cancelRate.toFixed(0)}%</td><td>${avgLead.toFixed(0)}일</td></tr>
      </tbody></table>
      </div>
      ${(_best && _worst && chArr.length > 1 && _best.netPerNight > _worst.netPerNight) ? `<div class="as-note" style="margin-top:8px;">
        박당 실수령 최고 <b>${_best.n} ${KRW(_best.netPerNight)}</b> · 최저 <b>${_worst.n} ${KRW(_worst.netPerNight)}</b> · 차이 ${KRW(_best.netPerNight - _worst.netPerNight)}/박
      </div>` : ''}`;
  CARD.band = `
      ${bandRows.map(r => `<div style="display:flex;align-items:center;gap:8px;margin:5px 0;font-size:11.5px;min-width:0;">
        <span style="width:62px;flex:none;font-weight:700;white-space:nowrap;">${r.label}</span>
        <span style="flex:1;height:11px;background:var(--surface-soft);border-radius:5px;overflow:hidden;min-width:0;"><span style="display:block;width:${r.nts / bandMaxN * 100}%;height:100%;background:${bandTop && r.label === bandTop.label ? 'var(--sl)' : '#a8a39a'};border-radius:5px;"></span></span>
        <span style="width:38px;flex:none;text-align:right;font-weight:800;">${r.nts}박</span>
        <span style="width:36px;flex:none;text-align:right;color:var(--mut);font-size:10.5px;">${nights > 0 ? Math.round(r.nts / nights * 100) : 0}%</span>
      </div>`).join('')}
      <div class="as-note">${bandTop && bandTop.nts ? `가장 많이 팔린 가격대는 <b>${bandTop.label}</b>(${bandTop.nts}박, 전체의 ${nights > 0 ? Math.round(bandTop.nts / nights * 100) : 0}%)입니다.` : '판매 기록이 부족합니다.'}<br>박당 요금(거래액 ÷ 박수) 기준 · 취소 제외</div>`;
  CARD.trend = `<div class="as-chartbox" style="height:172px;"><canvas id="asGross"></canvas></div>
      <div class="as-note" style="display:flex;gap:11px;flex-wrap:wrap;">
        <span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:13px;border-top:2px solid var(--sl);"></span>확정</span>
        <span style="display:inline-flex;align-items:center;gap:4px;"><span style="width:13px;border-top:2px dashed #a8a39a;"></span>판매 진행 중</span>
      </div>`;
  CARD.book = `<div class="as-chartbox" style="height:172px;"><canvas id="asBook"></canvas></div>`;
  CARD.nts  = `<div class="as-chartbox" style="height:172px;"><canvas id="asNights"></canvas></div>`;
  CARD.dow  = `<div class="as-chartbox" style="height:158px;"><canvas id="asDow"></canvas></div>
      <div class="as-note">최다 <b>${DOWK[busyDow]}요일</b> · ${dowCnt[busyDow]}건</div>`;
  CARD.lead = `<div class="as-chartbox" style="height:158px;"><canvas id="asLead"></canvas></div>
      <div class="as-note">평균 <b>${avgLead.toFixed(1)}일 전</b> 예약 · 수집일 기준 근사치.</div>`;
  /* [2026-09-17] 객실별 — 7열짜리 표를 없애고 막대 하나로 합쳤습니다. · 표에 있던 판매박수·판매율·ADR·매출을 전부 막대 오른쪽에 붙여서 · 한 줄에 다 보입니다. +5% / +10% 열은 뺐습니다 (아래 하한선 표와 겹칩니다). · 계산값(rStats)은 그대로 씁니다 — 숫자는 하나도 안 바뀝니다. */
  CARD.room = `
      ${roomBySales.map((s) => {
        const tag = (s.adr === minAdr && s.adr > 0)
          ? ' <span class="as-pill" style="background:var(--danger-50);color:var(--red);">최저가</span>'
          : ((s.adr === maxAdr && s.adr > 0)
            ? ' <span class="as-pill" style="background:var(--info-50);color:var(--bk);">최고가</span>' : '');
        const col = s.occ >= 90 ? 'var(--grn)' : (s.occ < 75 ? 'var(--text-muted)' : 'var(--bk)');
        return `
        <div style="display:flex;align-items:center;gap:11px;padding:9px 0;border-top:1px solid var(--border);min-width:0;">
          <div style="width:96px;flex:none;font-size:12.5px;font-weight:800;">${RH(s.room)}${tag}</div>
          <div style="flex:1;height:22px;background:var(--surface-soft);border-radius:6px;position:relative;overflow:hidden;min-width:60px;">
            <div style="position:absolute;inset:0 auto 0 0;width:${Math.min(100, s.occ).toFixed(0)}%;background:${col};border-radius:6px;opacity:.9;"></div>
            <div style="position:absolute;left:9px;top:0;bottom:0;display:flex;align-items:center;font-size:11px;font-weight:800;color:#fff;z-index:2;">${s.occ.toFixed(0)}%</div>
          </div>
          <div style="width:96px;flex:none;text-align:right;font-size:12.5px;font-weight:800;font-variant-numeric:tabular-nums;">${KRW(s.rev)}</div>
          <div style="width:88px;flex:none;text-align:right;font-size:11.5px;font-weight:700;color:var(--mut);font-variant-numeric:tabular-nums;">${KRW(s.adr)}</div>
          <div style="width:62px;flex:none;text-align:right;font-size:11px;font-weight:700;color:var(--mut);">${s.sold}/${spanDays}박</div>
        </div>`; }).join('')}
      <div style="display:flex;align-items:center;gap:11px;padding:10px 0 0;border-top:1.5px solid var(--border-strong);font-weight:800;">
        <div style="width:96px;flex:none;font-size:12.5px;">합계</div>
        <div style="flex:1;font-size:11.5px;color:var(--mut);">점유율 ${occ.toFixed(0)}%</div>
        <div style="width:96px;flex:none;text-align:right;font-size:12.5px;font-variant-numeric:tabular-nums;">${KRW(gross)}</div>
        <div style="width:88px;flex:none;text-align:right;font-size:11.5px;font-variant-numeric:tabular-nums;">${KRW(adr)}</div>
        <div style="width:62px;flex:none;text-align:right;font-size:11px;">${nights}박</div>
      </div>
      ${floorTable}
      <div class="as-note">· 막대 = <b>판매율</b> (그 방의 판매 박수 ÷ ${spanDays}일) · 오른쪽은 매출 / ADR / 판매박수.<br>· 달을 걸쳐 묵는 예약 때문에 100%를 넘을 수 있어요 (예: 32박/31일) · 박수 × ADR = 매출이 행마다 맞습니다.</div>`;
  CARD.wk = `
      <div class="as-grid as-g2" style="grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;">
        <div class="as-kpi"><div class="k">평일 ADR</div><div class="v" style="font-size:17px;">${KRW(wdAdr)}</div><div class="as-mini"><div style="width:${wdAdr / wkMax * 100}%;background:var(--sl);"></div></div><div class="d">${wk.wd.nts}박</div></div>
        <div class="as-kpi"><div class="k">주말(금·토) ADR</div><div class="v" style="font-size:17px;">${KRW(weAdr)}</div><div class="as-mini"><div style="width:${weAdr / wkMax * 100}%;background:var(--sl);"></div></div><div class="d">${wk.we.nts}박</div></div>
      </div>
      <div class="as-note">주말이 평일보다 <b>${wePrem >= 0 ? '+' : ''}${wePrem.toFixed(0)}%</b> · 평일 판매 비중 ${nights > 0 ? Math.round(wk.wd.nts / nights * 100) : 0}%.</div>`;
  CARD.nat = `
      ${natArr.length ? natArr.map(x => `<div style="display:flex;align-items:center;gap:8px;margin:5px 0;font-size:11.5px;min-width:0;"><div style="width:104px;flex:none;display:flex;align-items:center;overflow:hidden;white-space:nowrap;">${natFlag(x.code)}${x.name}</div><div style="flex:1;height:11px;background:var(--surface-soft);border-radius:5px;overflow:hidden;min-width:0;"><div style="width:${100 * x.v / natMax}%;height:100%;background:var(--sl);border-radius:5px;"></div></div><div style="width:32px;flex:none;text-align:right;font-weight:700;">${x.v}건</div></div>`).join('') : '<div style="color:var(--mut);font-size:12px;">국적이 입력된 예약이 없어요.</div>'}
      <div class="as-note">에어비앤비 국가 미제공 ${unknownN}건 제외</div>`;
  CARD.cxl = `
      <div class="as-kpigrid" style="grid-template-columns:repeat(3,minmax(0,1fr));">
        <div class="as-kpi"><div class="k">취소 건수</div><div class="v" style="font-size:17px;">${cancels.length}건</div><div class="d">${scope.length}건 중</div></div>
        <div class="as-kpi"><div class="k">취소율</div><div class="v" style="font-size:17px;color:var(--red);">${cancelRate.toFixed(1)}%</div><div class="d">확정 대비</div></div>
        <div class="as-kpi"><div class="k">취소 손실</div><div class="v" style="font-size:17px;color:var(--red);">${KRW(cancelLost)}</div><div class="d">취소분 합계</div></div>
      </div>
      ${cxlMonthly.length >= 2 ? `<div style="margin-top:14px;padding-top:13px;border-top:1px solid var(--border);">
        <div style="font-size:11px;font-weight:700;color:var(--mut);margin-bottom:11px;">월별 취소율</div>
        <div style="display:flex;align-items:flex-end;gap:14px;height:104px;padding:0 2px;">
          ${cxlMonthly.map(x => { const _mx = Math.max.apply(null, cxlMonthly.map(y => y.rate)) || 1; const _cur = x.key === __sheetView; return `<div style="flex:1;min-width:0;display:flex;flex-direction:column;align-items:center;justify-content:flex-end;height:100%;">
            <div style="font-size:11.5px;font-weight:800;margin-bottom:5px;color:${_cur ? 'var(--red)' : '#57534c'};white-space:nowrap;">${x.rate.toFixed(1)}%</div>
            <div style="width:100%;height:${Math.max(3, x.rate / _mx * 72).toFixed(0)}px;border-radius:4px 4px 0 0;background:${_cur ? '#c0392b' : '#d6d3cd'};"></div>
          </div>`; }).join('')}
        </div>
        <div style="display:flex;gap:14px;padding-top:7px;border-top:1px solid var(--border);">
          ${cxlMonthly.map(x => { const _cur = x.key === __sheetView; return `<div style="flex:1;min-width:0;text-align:center;">
            <div style="font-size:12px;font-weight:${_cur ? '800' : '700'};color:${_cur ? 'var(--ink)' : '#57534c'};">${x.label}</div>
            <div style="font-size:10.5px;color:var(--mut);margin-top:2px;white-space:nowrap;">${x.cnt}/${x.total}건</div>
            <div style="font-size:10.5px;color:var(--mut);margin-top:1px;white-space:nowrap;">${man(x.lost)}</div>
          </div>`; }).join('')}
        </div>
      </div>` : ''}
      <div class="as-note">${scopeLabel} 체크인 예정이었던 예약 중 취소분 · 수수료는 부과되지 않음.</div>`;
  CARD.daily = `
      <div style="overflow:auto;max-height:420px;border:1px solid #e8e6e1;border-radius:8px;">
      <table class="as-tbl" style="margin:0;">
        <thead><tr style="position:sticky;top:0;z-index:2;background:var(--surface);box-shadow:0 1px 0 #e8e6e1;">
          <th style="text-align:left;">날짜</th><th>판매</th><th>가동률</th><th>매출</th><th>ADR</th><th>IN</th><th>OUT</th>
        </tr></thead>
        <tbody>${dailyRows.map(r => {
          const pc = roomCount > 0 ? r.sold / roomCount * 100 : 0;
          const we = (r.dow === 5 || r.dow === 6);
          const full = roomCount > 0 && r.sold >= roomCount;
          const zero = r.sold === 0;
          const isT = r.ymd === _todayYmd;
          return `<tr style="${we ? 'background:#fafaf9;' : ''}${isT ? 'box-shadow:inset 3px 0 0 #3b6db5;' : ''}">
            <td style="text-align:left;white-space:nowrap;font-weight:${isT ? '800' : '600'};">${r.ymd.slice(5).replace('-', '/')}<span style="color:${we ? '#b85c3f' : '#78736b'};font-weight:700;margin-left:5px;">${DOW_K[r.dow]}</span>${isT ? '<span class="as-pill" style="background:#e8f0fa;color:#3b6db5;margin-left:6px;">오늘</span>' : ''}</td>
            <td style="${full ? 'font-weight:800;color:#1f1f23;' : (zero ? 'font-weight:700;color:#d6d3cd;' : 'font-weight:700;color:#57534c;')}">${r.sold}/${roomCount}</td>
            <td><div style="display:flex;align-items:center;gap:6px;justify-content:flex-end;">
              <span style="font-weight:700;min-width:30px;">${pc.toFixed(0)}%</span>
              <span style="width:46px;height:6px;background:#f5f5f4;border-radius:4px;overflow:hidden;display:inline-block;"><span style="display:block;width:${Math.min(100, pc)}%;height:100%;background:${full ? '#1f1f23' : (pc >= 60 ? '#78736b' : '#d6d3cd')};"></span></span>
            </div></td>
            <td style="font-weight:700;${r.rev > 0 ? '' : 'color:#d6d3cd;'}">${r.rev > 0 ? KRW(r.rev) : '—'}</td>
            <td style="color:var(--mut);">${r.revNts > 0 ? KRW(r.rev / r.revNts) : '—'}</td>
            <td style="font-weight:700;color:${r.cin ? '#1f1f23' : '#d6d3cd'};">${r.cin || '·'}</td>
            <td style="font-weight:700;color:${r.cout ? '#78736b' : '#d6d3cd'};">${r.cout || '·'}</td>
          </tr>`;
        }).join('')}
        <tr class="tot"><td style="text-align:left;">합계 · ${dailyRows.length}일</td><td>${dlTot.sold}박</td><td>${dlCap > 0 ? (dlTot.sold / dlCap * 100).toFixed(0) : 0}%</td><td>${KRW(dlTot.rev)}</td><td>${dlTot.revNts > 0 ? KRW(dlTot.rev / dlTot.revNts) : '—'}</td><td>${dlTot.cin}</td><td>${dlTot.cout}</td></tr>
        </tbody>
      </table>
      </div>
      <div class="as-note">· <b>매출</b>은 예약 전액을 <b>${REV_BASIS_LABEL}</b> 날짜에 계상 → <b>합계 ${KRW(dlTot.rev)} = 상단 거래액</b>${Math.abs(dlTot.rev - gross) < 1 ? ' ✓' : ` (차이 ${KRW(Math.abs(dlTot.rev - gross))} — 확인 필요)`} · 연박은 기준일에 전액 잡히므로 사이 날짜는 비어 있을 수 있어요.<br>· <b>판매·가동률</b>은 그날 실제로 묵은 객실 수 · 금·토 음영, 만실 초록, 오늘 파란 줄.</div>`;
  const cardsHtml = AS_CARDS.filter(c => CARD[c.k]).map(c =>
    `<div class="as-c ${c.w === 3 ? 'w3' : (c.w === 2 ? 'w2' : '')}">
       <div class="as-ch"><span class="as-ct">${c.t}</span><span class="as-csub">${scopeLabel}</span></div>
       ${CARD[c.k]}
     </div>`).join('');
  host.innerHTML = `
    <div class="as-noexport" style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;">
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;min-width:0;">
        <h1 class="as-title" style="margin:0;font-size:16px;">매출 인사이트</h1>
        <span class="as-pill" style="background:var(--surface-soft);color:var(--text-soft);">성수점 · ${roomCount}실</span>
        <span class="as-pill" style="background:var(--info-50);color:var(--info-600);">${REV_BASIS_LABEL}</span>
        <span class="as-pill" style="background:var(--surface-soft);color:var(--text-muted);">취소 제외</span>
      </div>
      <div style="display:flex;gap:6px;">
        <button onclick="sheetExportExcel()" title="엑셀로 저장" class="as-hbtn">엑셀</button>
        <button onclick="sheetExportImage(this)" title="이미지로 저장" class="as-hbtn">이미지</button>
      </div>
    </div>
    <!-- [2026-09-17] 기간 선택과 주차 선택이 두 줄이었습니다. 한 줄로 합쳤습니다. -->
    <div class="as-ctrl as-noexport">
      <div class="as-tabs" style="margin:0;">${tabsHtml}</div>
      <div class="as-ctrl-sp"></div>
      ${weekTabsHtml}
    </div>
    <div class="as-shell">
      <div class="as-cards">${cardsHtml}</div>
    </div>
    <div style="height:34px;"></div>
  `;
  // ── 차트 ──
  /* [2026-09-17] 차트 색이 밝은 값으로 고정돼 있어 다크에서 선·눈금이 안 보였습니다. · 현재 테마를 읽어서 정합니다. */
  const _dk = document.documentElement.getAttribute('data-theme') === 'dark';
  const GRID = _dk ? '#2a2a30' : '#f5f5f4';
  const SL   = _dk ? '#c9c9d0' : '#1f1f23';
  const LBL  = _dk ? '#9a9aa2' : '#57534c';
  const yMan = { grid: { color: GRID }, ticks: { callback: v => Math.round(v / 10000) + '만' } };
  const asDL = { id: 'asDL', afterDatasetsDraw(ch) { const fmt = ch.$fmt; if (!fmt) return; const ctx = ch.ctx; ch.data.datasets.forEach((ds, di) => { const meta = ch.getDatasetMeta(di); if (meta.hidden) return; meta.data.forEach((el, i) => { const t = fmt(ds.data[i]); if (!t) return; ctx.save(); ctx.fillStyle = LBL; ctx.font = "700 10.5px Pretendard,-apple-system,'Malgun Gothic',sans-serif"; ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(t, el.x, el.y - 6); ctx.restore(); }); }); } };
  const FMT = { asGross: v => v ? Math.round(v / 10000).toLocaleString('ko-KR') + '만' : '', asBook: v => v ? v + '건' : '', asNights: v => v ? v + '박' : '', asDow: v => v ? v + '건' : '', asLead: v => v ? v + '건' : '' };
  const mk = (id, cfg) => { const el = document.getElementById(id); if (!el) return; cfg.plugins = (cfg.plugins || []).concat([asDL]); cfg.options = cfg.options || {}; cfg.options.layout = { padding: { top: 22 } }; const ch = new Chart(el, cfg); ch.$fmt = FMT[id]; __sheetCharts.push(ch); };
  const mLabels = monthly.map(x => x.label);
  const _futIdx = months.findIndex(m => m >= _nowYm);   // 이번 달부터가 '판매 진행 중'
  const _isFut = k => _futIdx >= 0 && k >= _futIdx;
  mk('asGross', {
    type: 'line',
    data: { labels: mLabels, datasets: [{
      data: monthly.map(x => x.gross), borderColor: SL, backgroundColor: 'rgba(38,50,74,.08)',
      fill: true, tension: .3, pointRadius: 3.5,
      pointBackgroundColor: monthly.map((_, k) => _isFut(k) ? '#fff' : SL),
      pointBorderColor: monthly.map((_, k) => _isFut(k) ? '#a8a39a' : SL),
      pointBorderWidth: 1.6,
      segment: {
        borderDash: c => _isFut(c.p1DataIndex) ? [5, 4] : undefined,
        borderColor: c => _isFut(c.p1DataIndex) ? '#a8a39a' : SL
      }
    }] },
    options: { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: {
        label: c => KRW(c.parsed.y) + (_isFut(c.dataIndex) ? ' (판매 진행 중)' : '') } } },
      scales: { x: { grid: { display: false } }, y: yMan } } });
  mk('asBook', { type: 'bar', data: { labels: mLabels, datasets: [{ data: monthly.map(x => x.cnt), backgroundColor: '#2e5891', borderRadius: 6 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => c.parsed.y + '건' } } }, scales: { x: { grid: { display: false } }, y: { grid: { color: GRID } } } } });
  mk('asNights', { type: 'line', data: { labels: mLabels, datasets: [{ data: monthly.map(x => x.nights), borderColor: '#7c3aed', backgroundColor: 'rgba(124,58,237,.10)', fill: true, tension: .35, pointRadius: 4, pointBackgroundColor: '#7c3aed' }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => c.parsed.y + '박' } } }, scales: { x: { grid: { display: false } }, y: { grid: { color: GRID } } } } });
  mk('asChDonut', { type: 'doughnut',
    data: { labels: chArr.map(c => c.n), datasets: [{ data: chArr.map(c => c.rev), backgroundColor: chArr.map(c => c.color), borderWidth: 0 }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: '58%',
      plugins: { legend: { display: false }, tooltip: { callbacks: {
        label: c => c.label + ' ' + KRW(c.parsed) + ' (' + (chArr[c.dataIndex] ? chArr[c.dataIndex].share.toFixed(1) : '0') + '%)' } } } } });
  mk('asDow', { type: 'bar', data: { labels: DOWK, datasets: [{ data: dowCnt, backgroundColor: DOWK.map((d, i) => i === busyDow ? '#2e5891' : '#d6d3cd'), borderRadius: 6 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => c.parsed.y + '건' } } }, scales: { x: { grid: { display: false } }, y: { grid: { color: GRID }, ticks: { precision: 0 } } } } });
  mk('asLead', { type: 'bar', data: { labels: leadB.map(b => b[0]), datasets: [{ data: leadCnt, backgroundColor: '#7c3aed', borderRadius: 6 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => c.parsed.y + '건' } } }, scales: { x: { grid: { display: false } }, y: { grid: { color: GRID }, ticks: { precision: 0 } } } } });
}
async function openAnalysis() {
  const modal = document.getElementById('analysis-modal');
  if (!modal) return;
  modal.classList.add('open');
  const body = document.getElementById('analysis-body');
  body.innerHTML = '<div class="empty-state" style="padding:40px;">데이터를 분석하는 중…</div>';
  const { data, error } = await db.from('bookings')
    .select('created_at, checkin_date, checkout_date, room_id, channel_id, gross_amount, nights, num_guests, guest_country, booking_status, channels(name, code), rooms(room_number)')
    .order('checkin_date', { ascending: true });
  if (error) { body.innerHTML = `<div class="empty-state" style="padding:40px;">분석 데이터 로드 실패: ${error.message}</div>`; return; }
  __analysisRows = data || [];
  analysisBuildToggle();
  analysisRender();
}
function analysisMonths_() {
  const set = new Set();
  (__analysisRows || []).forEach(b => { const k = revMonth_(b); if (k) set.add(k); });
  return Array.from(set).sort();
}
function analysisBuildToggle() {
  const el = document.getElementById('analysis-view-toggle');
  if (!el) return;
  let html = `<button data-view="all" class="${analysisView === 'all' ? 'active' : ''}">전체 누적</button>`;
  analysisMonths_().forEach(m => {
    const mo = parseInt(m.split('-')[1], 10);
    html += `<button data-view="${m}" class="${analysisView === m ? 'active' : ''}">${mo}월</button>`;
  });
  el.innerHTML = html;
}
function analysisRender() {
  const body = document.getElementById('analysis-body');
  if (!body) return;
  const all = __analysisRows || [];
  let scope = all, scopeLabel = '전체 누적';
  if (analysisView !== 'all') {
    scope = all.filter(b => revMonth_(b) === analysisView);
    const [y, mo] = analysisView.split('-');
    scopeLabel = `${y}년 ${parseInt(mo, 10)}월 · ${REV_BASIS_LABEL}`;
  }
  const scopeEl = document.getElementById('analysis-scope-label');
  if (scopeEl) scopeEl.textContent = scopeLabel + ' 기준';
  const active = scope.filter(b => b.booking_status !== 'cancelled');
  if (active.length === 0) {
    body.innerHTML = `<div class="empty-state" style="padding:40px;">${scopeLabel}에 분석할 예약 데이터가 없습니다.</div>`;
    return;
  }
  const perNight_ = (b) => (b.nights > 0 ? (b.gross_amount || 0) / b.nights : 0);
  const leadDays_ = (b) => {
    if (!b.created_at || !b.checkin_date) return null;
    const d = Math.floor((new Date(b.checkin_date + 'T00:00:00') - new Date(b.created_at)) / 86400000);
    return d < 0 ? 0 : d;
  };
  const man_ = (v) => Math.round(v / 10000).toLocaleString();
  // ---------- 1) 채널별 비교 ----------
  const ch = {};
  scope.forEach(b => {
    const k = b.channels?.name || '직접';
    const o = ch[k] || (ch[k] = { cnt: 0, rev: 0, nights: 0, lead: [], total: 0, cancel: 0 });
    o.total++;
    if (b.booking_status === 'cancelled') { o.cancel++; return; }
    o.cnt++; o.rev += (b.gross_amount || 0); o.nights += (b.nights || 0);
    const ld = leadDays_(b); if (ld != null) o.lead.push(ld);
  });
  const chRows = Object.entries(ch).sort((a, b) => b[1].rev - a[1].rev).map(([name, o]) => {
    const adr = o.nights > 0 ? o.rev / o.nights : 0;
    const avgLead = o.lead.length ? Math.round(o.lead.reduce((s, v) => s + v, 0) / o.lead.length) : 0;
    const cancelRate = o.total > 0 ? (o.cancel / o.total * 100) : 0;
    return { name, cnt: o.cnt, adr, avgLead, cancelRate };
  });
  const chTable = `
    <table class="an-table">
      <thead><tr><th>채널</th><th>확정 건수</th><th>ADR(박당)</th><th>평균 리드타임</th><th>취소율</th></tr></thead>
      <tbody>${chRows.map(r => `
        <tr><td><span style="display:inline-block;width:9px;height:9px;border-radius:2px;background:${opsChColor(r.name)};margin-right:7px;vertical-align:middle;"></span>${r.name}</td>
          <td class="num">${r.cnt}건</td>
          <td class="num">${fmtKRW(r.adr)}</td>
          <td class="num">${r.avgLead}일 전</td>
          <td class="num">${r.cancelRate.toFixed(0)}%</td></tr>`).join('')}
      </tbody>
    </table>`;
  // ---------- 2) 평일 / 주말 요금 ----------
  const wk = { weekday: { rev: 0, nights: 0, cnt: 0 }, weekend: { rev: 0, nights: 0, cnt: 0 } };
  active.forEach(b => {
    if (!b.checkin_date) return;
    const dow = new Date(b.checkin_date + 'T00:00:00').getDay(); // 0일~6토
    const bucket = (dow === 5 || dow === 6) ? wk.weekend : wk.weekday; // 금·토 = 주말
    bucket.rev += (b.gross_amount || 0); bucket.nights += (b.nights || 0); bucket.cnt++;
  });
  const wkAdr = (o) => o.nights > 0 ? o.rev / o.nights : 0;
  const wkMax = Math.max(wkAdr(wk.weekday), wkAdr(wk.weekend), 1);
  const wkTable = `
    <table class="an-table">
      <thead><tr><th>구분</th><th>건수</th><th>ADR(박당)</th><th class="an-bar-cell">상대</th></tr></thead>
      <tbody>
        <tr><td>평일 (일~목)</td><td class="num">${wk.weekday.cnt}건</td><td class="num">${fmtKRW(wkAdr(wk.weekday))}</td>
          <td><div class="an-mini"><div style="width:${(wkAdr(wk.weekday) / wkMax * 100).toFixed(0)}%"></div></div></td></tr>
        <tr><td>주말 (금·토)</td><td class="num">${wk.weekend.cnt}건</td><td class="num">${fmtKRW(wkAdr(wk.weekend))}</td>
          <td><div class="an-mini"><div style="width:${(wkAdr(wk.weekend) / wkMax * 100).toFixed(0)}%; background:#c0451f;"></div></div></td></tr>
      </tbody>
    </table>`;
  // ---------- 3) 리드타임 분포 ----------
  const buckets = [['당일', 0, 0], ['1–3일', 1, 3], ['4–7일', 4, 7], ['8–14일', 8, 14], ['15–30일', 15, 30], ['31일+', 31, 9999]];
  const ldCounts = buckets.map(() => 0);
  active.forEach(b => { const ld = leadDays_(b); if (ld == null) return; const i = buckets.findIndex(([, lo, hi]) => ld >= lo && ld <= hi); if (i >= 0) ldCounts[i]++; });
  const ldMax = Math.max(...ldCounts, 1);
  const ldTable = `
    <table class="an-table">
      <thead><tr><th>예약 시점</th><th>건수</th><th class="an-bar-cell">분포</th></tr></thead>
      <tbody>${buckets.map((bk, i) => `
        <tr><td>${bk[0]}</td><td class="num">${ldCounts[i]}건</td>
          <td><div class="an-mini"><div style="width:${(ldCounts[i] / ldMax * 100).toFixed(0)}%"></div></div></td></tr>`).join('')}
      </tbody>
    </table>`;
  // ---------- 4) 잘 팔린 가격대 (박당요금) ----------
  const bands = [['10–15만', 100000, 149999], ['15–20만', 150000, 199999], ['20–25만', 200000, 249999], ['25–30만', 250000, 299999], ['30만+', 300000, 1e12]];
  const bandCounts = bands.map(() => 0);
  active.forEach(b => { const pn = perNight_(b); if (pn <= 0) return; const i = bands.findIndex(([, lo, hi]) => pn >= lo && pn <= hi); if (i >= 0) bandCounts[i]++; });
  const bandMax = Math.max(...bandCounts, 1);
  const topBand = bands[bandCounts.indexOf(Math.max(...bandCounts))];
  const bandTable = `
    <table class="an-table">
      <thead><tr><th>박당 요금대</th><th>건수</th><th class="an-bar-cell">분포</th></tr></thead>
      <tbody>${bands.map((bd, i) => `
        <tr><td>${bd[0]}${bd === topBand ? ' <span class="an-chip" style="background:#EAF3DE;color:#3B6D11;">최다</span>' : ''}</td>
          <td class="num">${bandCounts[i]}건</td>
          <td><div class="an-mini"><div style="width:${(bandCounts[i] / bandMax * 100).toFixed(0)}%; background:#639922;"></div></div></td></tr>`).join('')}
      </tbody>
    </table>`;
  // ---------- 5~7) 객실별 실판매 기준가 · 점유율 · RevPAR · 공실 · 프라이싱 제안 ----------
  const median_ = (arr) => { if (!arr.length) return 0; const s = [...arr].sort((a, b) => a - b); const m = Math.floor(s.length / 2); return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
  const pad2_ = (n) => String(n).padStart(2, '0');
  const ymd_ = (d) => `${d.getFullYear()}-${pad2_(d.getMonth() + 1)}-${pad2_(d.getDate())}`;
  let spanDays;
  if (analysisView === 'all') {
    const cis = active.map(b => b.checkin_date).filter(Boolean).sort();
    const cos = active.map(b => revDate_(b)).filter(Boolean).sort();
    const first = cis[0], last = cos[cos.length - 1] || cis[cis.length - 1];
    spanDays = (first && last) ? Math.max(1, Math.round((new Date(last) - new Date(first)) / 86400000)) : 30;
  } else {
    const [yy, mm2] = analysisView.split('-').map(Number);
    spanDays = new Date(yy, mm2, 0).getDate();
  }
  const rooms = (cachedRooms || []).filter(r => r.is_active).sort((a, b) => (a.room_number || '').localeCompare(b.room_number || ''));
  const roomAgg = {};
  rooms.forEach(r => { roomAgg[r.room_number] = { soldNights: 0, rev: 0, wd: [], we: [], lead: [], total: 0, cancel: 0 }; });
  scope.forEach(b => {
    const rn = b.rooms?.room_number; if (!rn || !roomAgg[rn]) return;
    roomAgg[rn].total++;
    if (b.booking_status === 'cancelled') { roomAgg[rn].cancel++; return; }
    roomAgg[rn].rev += (b.gross_amount || 0);
    roomAgg[rn].soldNights += (b.nights || 0);
    const pn = b.nights > 0 ? (b.gross_amount || 0) / b.nights : 0;
    const dow = b.checkin_date ? new Date(b.checkin_date + 'T00:00:00').getDay() : 0;
    if (pn > 0) ((dow === 5 || dow === 6) ? roomAgg[rn].we : roomAgg[rn].wd).push(pn);
    const ld = leadDays_(b); if (ld != null) roomAgg[rn].lead.push(ld);
  });
  const roomStats = rooms.map(r => {
    const a = roomAgg[r.room_number];
    const occ = spanDays > 0 ? Math.min(100, a.soldNights / spanDays * 100) : 0;
    const revpar = spanDays > 0 ? a.rev / spanDays : 0;
    const baseWd = median_(a.wd), baseWe = median_(a.we);
    const cancelRate = a.total > 0 ? a.cancel / a.total * 100 : 0;
    const avgLead = a.lead.length ? Math.round(a.lead.reduce((s, v) => s + v, 0) / a.lead.length) : 0;
    return { room: r.room_number, occ, revpar, baseWd, baseWe, soldNights: a.soldNights, cancelRate, avgLead };
  });
  const roomStatTable = `
    <table class="an-table">
      <thead><tr><th>객실</th><th>점유율</th><th>RevPAR</th><th>평일 기준가</th><th>주말 기준가</th></tr></thead>
      <tbody>${roomStats.map(s => `
        <tr><td>${RH(s.room)}</td>
          <td class="num">${s.occ.toFixed(0)}%</td>
          <td class="num">${fmtKRW(s.revpar)}</td>
          <td class="num">${s.baseWd > 0 ? fmtKRW(s.baseWd) : '—'}</td>
          <td class="num">${s.baseWe > 0 ? fmtKRW(s.baseWe) : '—'}</td></tr>`).join('')}
      </tbody>
    </table>
    <div style="font-size:11px; color:var(--text-muted); margin-top:8px;">· 기준가 = 실제 판매 박당요금의 중앙값(이상치 방지) · 점유율·RevPAR은 기간 ${spanDays}일 기준</div>`;
  // 공실 (다가오는 4주) — scope와 무관하게 미래 기준
  const today0 = new Date(); today0.setHours(0, 0, 0, 0);
  const HORIZON = 28;
  const vac = {}; rooms.forEach(r => vac[r.room_number] = 0);
  for (let i = 0; i < HORIZON; i++) {
    const day = new Date(today0); day.setDate(day.getDate() + i);
    const dstr = ymd_(day);
    rooms.forEach(r => {
      const occupied = (__analysisRows || []).some(b => b.booking_status !== 'cancelled' && b.rooms?.room_number === r.room_number && b.checkin_date && b.checkout_date && b.checkin_date <= dstr && dstr < b.checkout_date);
      if (!occupied) vac[r.room_number]++;
    });
  }
  const vacTotal = Object.values(vac).reduce((s, v) => s + v, 0);
  const vacMax = Math.max(...Object.values(vac), 1);
  const vacTable = `
    <table class="an-table">
      <thead><tr><th>객실</th><th>빈 날(4주)</th><th class="an-bar-cell">공실</th></tr></thead>
      <tbody>${rooms.map(r => `
        <tr><td>${RH(r.room_number)}</td><td class="num">${vac[r.room_number]}일</td>
          <td><div class="an-mini"><div style="width:${(vac[r.room_number] / vacMax * 100).toFixed(0)}%; background:#c0451f;"></div></div></td></tr>`).join('')}
      </tbody>
    </table>
    <div style="font-size:11px; color:var(--text-muted); margin-top:8px;">· 오늘부터 28일 중 예약 없는 객실-일 합계 ${vacTotal}일 — 리드타임을 넘겼는데 빈 날은 인하 타이밍</div>`;
  // 프라이싱 제안 (실판매 기준가 + 점유율·리드타임·취소율)
  const suggCards = roomStats.map(s => {
    let tag, bg, fg, advice;
    if (s.soldNights === 0) {
      tag = '데이터 부족'; bg = 'var(--surface-soft)'; fg = 'var(--text-muted)';
      advice = '판매 기록이 적어 제안을 보류합니다. 노출·최소박 완화로 초기 예약을 모으세요.';
    } else if (s.occ >= 70 && s.cancelRate < 12) {
      tag = '인상 여력'; bg = '#EAF3DE'; fg = '#3B6D11';
      advice = `수요 강함(점유 ${s.occ.toFixed(0)}%${s.avgLead >= 8 ? ' · 조기예약 많음' : ''}, 취소 ${s.cancelRate.toFixed(0)}%) → 주말 기준가 <b>+5~10%</b> 시도 여력.`;
    } else if (s.occ <= 35) {
      tag = '인하 / 완화'; bg = '#FAECE7'; fg = '#993C1D';
      advice = `공실 많음(점유 ${s.occ.toFixed(0)}%) → 평일·당일가 <b>−5%</b> 또는 <b>최소박 완화</b>로 점유 회복.`;
    } else {
      tag = '유지'; bg = 'var(--surface-soft)'; fg = 'var(--text-soft)';
      advice = `적정 수준(점유 ${s.occ.toFixed(0)}%) → 현 기준가 유지, 주말만 소폭(+3~5%) 테스트.`;
    }
    return `
      <div style="border:1px solid var(--border); border-radius:10px; padding:13px 15px;">
        <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;">
          <span style="font-weight:800; font-size:14px; color:var(--text);">${RH(s.room)}</span>
          <span class="an-chip" style="background:${bg}; color:${fg};">${tag}</span>
        </div>
        <div style="font-size:12px; color:var(--text-soft); margin-bottom:8px;">평일 ${s.baseWd > 0 ? fmtKRW(s.baseWd) : '—'} · 주말 ${s.baseWe > 0 ? fmtKRW(s.baseWe) : '—'} · 점유 ${s.occ.toFixed(0)}% · 리드 ${s.avgLead}일</div>
        <div style="font-size:12.5px; color:var(--text); line-height:1.6;">${advice}</div>
      </div>`;
  }).join('');
  //  ✦ 데이터팀 리포트 상단 4종 (① 요약 KPI · ② 핵심 인사이트 · ③ 월별 추세 · ④ 매출 기여도)
  const roomCount = rooms.length || 5;
  // ---- ① Executive Summary KPI (현재 스코프) ----
  const revTotal    = active.reduce((s, b) => s + (b.gross_amount || 0), 0);
  const nightsTotal = active.reduce((s, b) => s + (b.nights || 0), 0);
  const adrAll      = nightsTotal > 0 ? revTotal / nightsTotal : 0;
  const bookingsN   = active.length;
  const occAll      = (spanDays > 0 && roomCount > 0) ? Math.min(100, nightsTotal / (spanDays * roomCount) * 100) : 0;
  const revparAll   = (spanDays > 0 && roomCount > 0) ? revTotal / (spanDays * roomCount) : 0;
  const totalScope  = scope.length;
  const cancelScope = scope.filter(b => b.booking_status === 'cancelled').length;
  const cancelRateAll = totalScope > 0 ? cancelScope / totalScope * 100 : 0;
  const cancelledRevLost = scope.filter(b => b.booking_status === 'cancelled').reduce((s, b) => s + (b.gross_amount || 0), 0);
  // 전월 비교 (특정 월 스코프일 때만)
  let prevKpi = null;
  if (analysisView !== 'all') {
    const months = analysisMonths_();
    const idx = months.indexOf(analysisView);
    if (idx > 0) {
      const pm = months[idx - 1];
      const pActive = all.filter(b => revMonth_(b) === pm && b.booking_status !== 'cancelled');
      const [py, pmo] = pm.split('-').map(Number);
      const pSpan = new Date(py, pmo, 0).getDate();
      const pRev = pActive.reduce((s, b) => s + (b.gross_amount || 0), 0);
      const pNights = pActive.reduce((s, b) => s + (b.nights || 0), 0);
      prevKpi = {
        label: `${pmo}월`,
        rev: pRev, bookings: pActive.length,
        adr: pNights > 0 ? pRev / pNights : 0,
        occ: (pSpan > 0 && roomCount > 0) ? Math.min(100, pNights / (pSpan * roomCount) * 100) : 0,
        revpar: (pSpan > 0 && roomCount > 0) ? pRev / (pSpan * roomCount) : 0,
      };
    }
  }
  const deltaPct = (cur, prev) => {
    if (prevKpi == null || prev == null || prev === 0) return '';
    const pct = (cur - prev) / Math.abs(prev) * 100;
    const cls = pct > 0.5 ? 'up' : pct < -0.5 ? 'down' : 'flat';
    const arrow = pct > 0.5 ? '▲' : pct < -0.5 ? '▼' : '—';
    return `<div class="an-kpi-delta ${cls}">${arrow} ${Math.abs(pct).toFixed(0)}% <span style="opacity:0.6;font-weight:600;">vs ${prevKpi.label}</span></div>`;
  };
  const deltaPt = (cur, prev) => {
    if (prevKpi == null || prev == null) return '';
    const d = cur - prev;
    const cls = d > 0.5 ? 'up' : d < -0.5 ? 'down' : 'flat';
    const arrow = d > 0.5 ? '▲' : d < -0.5 ? '▼' : '—';
    return `<div class="an-kpi-delta ${cls}">${arrow} ${Math.abs(d).toFixed(0)}p <span style="opacity:0.6;font-weight:600;">vs ${prevKpi.label}</span></div>`;
  };
  const kpiCard = (label, val, delta) => `<div class="an-kpi"><div class="an-kpi-label">${label}</div><div class="an-kpi-val">${val}</div>${delta || ''}</div>`;
  const execSummaryHtml = `
    <div class="an-exec-band">
      <div class="an-exec-title">✦ Executive Summary · ${scopeLabel}</div>
      <div class="an-kpi-grid">
        ${kpiCard('총매출', fmtKRW(revTotal), deltaPct(revTotal, prevKpi?.rev))}
        ${kpiCard('예약 건수', bookingsN + '건', deltaPct(bookingsN, prevKpi?.bookings))}
        ${kpiCard('ADR (박당)', fmtKRW(adrAll), deltaPct(adrAll, prevKpi?.adr))}
        ${kpiCard('점유율 (OCC)', occAll.toFixed(0) + '%', deltaPt(occAll, prevKpi?.occ))}
        ${kpiCard('RevPAR', fmtKRW(revparAll), deltaPct(revparAll, prevKpi?.revpar))}
        ${kpiCard('평균 예약가치', bookingsN > 0 ? fmtKRW(revTotal / bookingsN) : '—', '')}
        ${kpiCard('총 판매 박수', nightsTotal.toLocaleString() + '박', '')}
        ${kpiCard('취소율', cancelRateAll.toFixed(0) + '%', cancelledRevLost > 0 ? `<div class="an-kpi-delta down">−${fmtKRW(cancelledRevLost)} 손실</div>` : '')}
      </div>
    </div>`;
  // ---- ④용 객실/채널 매출 집계 (인사이트에서도 재사용) ----
  const roomRevArr = rooms.map(r => {
    const a = roomAgg[r.room_number] || { rev: 0, soldNights: 0 };
    const st = roomStats.find(s => s.room === r.room_number) || { occ: 0 };
    return { room: r.room_number, rev: a.rev, adr: a.soldNights > 0 ? a.rev / a.soldNights : 0, occ: st.occ };
  }).sort((x, y) => y.rev - x.rev);
  const totalRoomRev = roomRevArr.reduce((s, r) => s + r.rev, 0) || 1;
  const chRevArr = Object.entries(ch).map(([name, o]) => ({ name, rev: o.rev, cancelRate: o.total > 0 ? o.cancel / o.total * 100 : 0 })).sort((a, b) => b.rev - a.rev);
  const totalChRev = chRevArr.reduce((s, c) => s + c.rev, 0) || 1;
  // ---- ② 핵심 인사이트 (데이터 기반 자동 문장) ----
  const insights = [];
  if (roomRevArr[0] && roomRevArr[0].rev > 0) {
    const t = roomRevArr[0];
    insights.push({ ic: '🏆', txt: `<b>${RH(t.room)}가 매출 1위</b> — 전체 매출의 <b>${(t.rev / totalRoomRev * 100).toFixed(0)}%</b>(${fmtKRW(t.rev)})를 책임지는 핵심 객실. 점유율 ${t.occ.toFixed(0)}%.` });
  }
  const byAdr = [...roomRevArr].filter(r => r.adr > 0).sort((a, b) => b.adr - a.adr);
  if (byAdr[0]) {
    const t = byAdr[0];
    const note = t.occ < occAll ? ` 다만 점유율(${t.occ.toFixed(0)}%)이 평균 이하라 <b>가격 방어 vs 점유 회복</b> 판단 필요.` : ` 점유율도 ${t.occ.toFixed(0)}%로 견조.`;
    insights.push({ ic: '💎', txt: `<b>${RH(t.room)}가 ADR 최고</b> (${fmtKRW(t.adr)}/박).${note}` });
  }
  const byOcc = [...roomRevArr].sort((a, b) => a.occ - b.occ);
  if (byOcc[0] && byOcc[0].occ < 45) {
    insights.push({ ic: '⚠️', txt: `<b>${RH(byOcc[0].room)} 점유율이 ${byOcc[0].occ.toFixed(0)}%로 최저</b> — 노출 강화·최소박 완화 또는 평일·당일가 인하 검토 대상.` });
  }
  const earlyN = active.filter(b => { const ld = leadDays_(b); return ld != null && ld >= 8; }).length;
  if (bookingsN > 0) {
    insights.push({ ic: '📅', txt: `예약의 <b>${(earlyN / bookingsN * 100).toFixed(0)}%</b>가 <b>8일 이상 조기예약</b> — 조기예약 할인 여력과 얼리버드 프로모션 타이밍 판단에 활용.` });
  }
  if (chRevArr[0] && chRevArr[0].rev > 0) {
    const t = chRevArr[0];
    insights.push({ ic: '🔗', txt: `<b>${t.name}</b>이(가) 매출의 <b>${(t.rev / totalChRev * 100).toFixed(0)}%</b>를 차지 (취소율 ${t.cancelRate.toFixed(0)}%). 채널 의존도와 수수료 구조를 함께 볼 지점.` });
  }
  const wdAdr2 = wkAdr(wk.weekday), weAdr2 = wkAdr(wk.weekend);
  if (wdAdr2 > 0 && weAdr2 > 0) {
    const prem = (weAdr2 - wdAdr2) / wdAdr2 * 100;
    insights.push({ ic: '📈', txt: `주말(금·토) ADR이 평일 대비 <b>${prem >= 0 ? '+' : ''}${prem.toFixed(0)}%</b> — ${prem < 8 ? '주말 프리미엄이 약함, 주말가 상향 여력.' : '주말 프리미엄 정상 작동 중.'}` });
  }
  if (typeof vacTotal === 'number' && vacTotal > roomCount * 7) {
    insights.push({ ic: '🛏️', txt: `다가오는 4주간 <b>공실 ${vacTotal}일</b> — 리드타임을 넘겼는데 빈 날은 막판 인하/프로모션으로 채울 타이밍.` });
  }
  const insightsHtml = `
    <div class="an-sec">
      <div class="an-sec-title">✦ 핵심 인사이트</div>
      <div class="an-sec-sub">데이터에서 자동 추출한 요약 — ${scopeLabel} 기준</div>
      ${insights.length ? insights.slice(0, 6).map(o => `<div class="an-insight"><span class="ic">${o.ic}</span><span>${o.txt}</span></div>`).join('') : '<div style="color:var(--text-muted);font-size:13px;">인사이트를 낼 만한 데이터가 아직 부족합니다.</div>'}
    </div>`;
  // ---- ③ 월별 추세 (전체 데이터 기준, 스코프 무관) ----
  const monthly = analysisMonths_().map(m => {
    const mActive = all.filter(b => revMonth_(b) === m && b.booking_status !== 'cancelled');
    const [my, mm] = m.split('-').map(Number);
    const mSpan = new Date(my, mm, 0).getDate();
    const mRev = mActive.reduce((s, b) => s + (b.gross_amount || 0), 0);
    const mNights = mActive.reduce((s, b) => s + (b.nights || 0), 0);
    return { m, label: `${mm}월`, rev: mRev, cnt: mActive.length, adr: mNights > 0 ? mRev / mNights : 0, occ: (mSpan > 0 && roomCount > 0) ? Math.min(100, mNights / (mSpan * roomCount) * 100) : 0 };
  });
  const monthRevMax = Math.max(...monthly.map(x => x.rev), 1);
  // 전월 대비 성장률
  monthly.forEach((x, i) => { x.growth = (i > 0 && monthly[i - 1].rev > 0) ? (x.rev - monthly[i - 1].rev) / monthly[i - 1].rev * 100 : null; });
  const bestMonth = monthly.length ? monthly.reduce((a, b) => b.rev > a.rev ? b : a) : null;
  const growths = monthly.map(x => x.growth).filter(v => v != null);
  const avgGrowth = growths.length ? growths.reduce((s, v) => s + v, 0) / growths.length : null;
  const growthCell = (g) => {
    if (g == null) return '<span style="color:var(--text-muted);">—</span>';
    const cls = g > 0.5 ? 'up' : g < -0.5 ? 'down' : 'flat';
    const color = cls === 'up' ? '#3B6D11' : cls === 'down' ? '#C0392B' : 'var(--text-muted)';
    const arrow = g > 0.5 ? '▲' : g < -0.5 ? '▼' : '—';
    return `<span style="color:${color}; font-weight:700;">${arrow} ${Math.abs(g).toFixed(0)}%</span>`;
  };
  const monthlyHtml = monthly.length > 1 ? `
    <div class="an-sec">
      <div class="an-sec-title">📊 월별 추세</div>
      <div class="an-sec-sub">체크인 월 기준 매출·건수·ADR·점유율·전월대비 성장 (전체 데이터)</div>
      <table class="an-table">
        <thead><tr><th>월</th><th>매출</th><th class="an-bar-cell">매출 추이</th><th>전월대비</th><th>건수</th><th>ADR</th><th>점유율</th></tr></thead>
        <tbody>${monthly.map(x => `
          <tr${x.m === analysisView ? ' style="background:var(--accent-soft);"' : ''}><td>${x.label}${x.m === analysisView ? ' <span class="an-chip" style="background:var(--accent);color:#fff;">보는 중</span>' : ''}${bestMonth && x.m === bestMonth.m ? ' <span class="an-chip" style="background:#EAF3DE;color:#3B6D11;">최고매출</span>' : ''}</td>
            <td class="num">${fmtKRW(x.rev)}</td>
            <td><div class="an-mini"><div style="width:${(x.rev / monthRevMax * 100).toFixed(0)}%; background:var(--accent);"></div></div></td>
            <td class="num">${growthCell(x.growth)}</td>
            <td class="num">${x.cnt}건</td>
            <td class="num">${fmtKRW(x.adr)}</td>
            <td class="num">${x.occ.toFixed(0)}%</td></tr>`).join('')}
        </tbody>
      </table>
      <div style="font-size:11.5px; color:var(--text-muted); margin-top:8px;">${bestMonth ? `· 최고 매출: <b>${bestMonth.label}</b> (${fmtKRW(bestMonth.rev)})` : ''}${avgGrowth != null ? ` · 월평균 성장률 ${avgGrowth >= 0 ? '+' : ''}${avgGrowth.toFixed(0)}%` : ''}</div>
    </div>` : '';
  // ---- ④ 매출 기여도 (객실·채널 파레토) ----
  let cum = 0;
  const roomParetoRows = roomRevArr.map(r => {
    const share = r.rev / totalRoomRev * 100; cum += share;
    return `<tr><td>${RH(r.room)}</td><td class="num">${fmtKRW(r.rev)}</td><td class="num">${share.toFixed(0)}%</td><td class="num">${cum.toFixed(0)}%</td>
      <td><div class="an-mini"><div style="width:${share.toFixed(0)}%; background:#639922;"></div></div></td></tr>`;
  }).join('');
  const chParetoRows = chRevArr.map(c => {
    const share = c.rev / totalChRev * 100;
    return `<tr><td>${c.name}</td><td class="num">${fmtKRW(c.rev)}</td><td class="num">${share.toFixed(0)}%</td>
      <td><div class="an-mini"><div style="width:${share.toFixed(0)}%; background:#185FA5;"></div></div></td></tr>`;
  }).join('');
  const contributionHtml = `
    <div class="an-grid2">
      <div class="an-sec">
        <div class="an-sec-title">💰 객실별 매출 기여도</div>
        <div class="an-sec-sub">매출 비중 내림차순 · 누적 — 상위 몇 개 방이 매출을 끄는지(집중도)</div>
        <table class="an-table">
          <thead><tr><th>객실</th><th>매출</th><th>비중</th><th>누적</th><th class="an-bar-cell">비중</th></tr></thead>
          <tbody>${roomParetoRows}</tbody>
        </table>
      </div>
      <div class="an-sec">
        <div class="an-sec-title">💰 채널별 매출 기여도</div>
        <div class="an-sec-sub">채널 매출 비중 — 의존도 파악</div>
        <table class="an-table">
          <thead><tr><th>채널</th><th>매출</th><th>비중</th><th class="an-bar-cell">비중</th></tr></thead>
          <tbody>${chParetoRows}</tbody>
        </table>
      </div>
    </div>`;
  //  ✦ 상세 지표 확장 (요일 · 투숙기간 · 인원 · 리드타임요금 · 객실×채널 · 취소 · 국적)
  // 요일별 수요·ADR (체크인 요일)
  const DOWK = ['일', '월', '화', '수', '목', '금', '토'];
  const dowAgg = DOWK.map(() => ({ cnt: 0, rev: 0, nights: 0 }));
  active.forEach(b => { if (!b.checkin_date) return; const d = new Date(b.checkin_date + 'T00:00:00').getDay(); dowAgg[d].cnt++; dowAgg[d].rev += (b.gross_amount || 0); dowAgg[d].nights += (b.nights || 0); });
  const dowCntMax = Math.max(...dowAgg.map(o => o.cnt), 1);
  const dowHtml = `
    <div class="an-sec">
      <div class="an-sec-title">8 · 요일별 수요 · 요금</div>
      <div class="an-sec-sub">체크인 요일별 건수와 ADR — 어느 요일이 잘 나가고 비싸게 받는지 (금·토 = 주말)</div>
      <table class="an-table">
        <thead><tr><th>요일</th><th>건수</th><th class="an-bar-cell">수요</th><th>ADR</th></tr></thead>
        <tbody>${dowAgg.map((o, i) => `
          <tr${(i === 5 || i === 6) ? ' style="background:var(--surface-soft);"' : ''}><td>${DOWK[i]}${(i === 5 || i === 6) ? ' <span class="an-chip" style="background:#FAECE7;color:#993C1D;">주말</span>' : ''}</td>
            <td class="num">${o.cnt}건</td>
            <td><div class="an-mini"><div style="width:${(o.cnt / dowCntMax * 100).toFixed(0)}%"></div></div></td>
            <td class="num">${o.nights > 0 ? fmtKRW(o.rev / o.nights) : '—'}</td></tr>`).join('')}
        </tbody>
      </table>
    </div>`;
  // 투숙 기간(LOS) 분포
  const losB = [['1박', 1, 1], ['2박', 2, 2], ['3박', 3, 3], ['4–6박', 4, 6], ['7박+', 7, 9999]];
  const losCnt = losB.map(() => 0), losRev = losB.map(() => 0);
  let losSum = 0, losN = 0;
  active.forEach(b => { const n = b.nights || 0; if (n <= 0) return; losSum += n; losN++; const i = losB.findIndex(([, lo, hi]) => n >= lo && n <= hi); if (i >= 0) { losCnt[i]++; losRev[i] += (b.gross_amount || 0); } });
  const avgLos = losN > 0 ? losSum / losN : 0;
  const losMax = Math.max(...losCnt, 1);
  const losHtml = `
    <div class="an-sec">
      <div class="an-sec-title">9 · 투숙 기간 (LOS) 분포</div>
      <div class="an-sec-sub">예약별 숙박일수 분포 · 평균 투숙 <b>${avgLos.toFixed(1)}박</b> — 최소박 설정·연박 할인 판단</div>
      <table class="an-table">
        <thead><tr><th>숙박일</th><th>건수</th><th class="an-bar-cell">분포</th><th>매출</th></tr></thead>
        <tbody>${losB.map((bk, i) => `
          <tr><td>${bk[0]}</td><td class="num">${losCnt[i]}건</td>
            <td><div class="an-mini"><div style="width:${(losCnt[i] / losMax * 100).toFixed(0)}%; background:#639922;"></div></div></td>
            <td class="num">${fmtKRW(losRev[i])}</td></tr>`).join('')}
        </tbody>
      </table>
    </div>`;
  // 인원 구성 분포
  const partyB = [['1명', 1, 1], ['2명', 2, 2], ['3명', 3, 3], ['4명+', 4, 99]];
  const partyCnt = partyB.map(() => 0);
  let gSum = 0, gN = 0;
  active.forEach(b => { const g = b.num_guests || 0; if (g <= 0) return; gSum += g; gN++; const i = partyB.findIndex(([, lo, hi]) => g >= lo && g <= hi); if (i >= 0) partyCnt[i]++; });
  const avgGuests = gN > 0 ? gSum / gN : 0;
  const partyMax = Math.max(...partyCnt, 1);
  const partyHtml = `
    <div class="an-sec">
      <div class="an-sec-title">10 · 인원 구성</div>
      <div class="an-sec-sub">예약 인원 분포 · 평균 <b>${avgGuests.toFixed(1)}명</b> — 객실 타입·침구 구성 판단</div>
      <table class="an-table">
        <thead><tr><th>인원</th><th>건수</th><th class="an-bar-cell">분포</th></tr></thead>
        <tbody>${partyB.map((bk, i) => `
          <tr><td>${bk[0]}</td><td class="num">${partyCnt[i]}건</td>
            <td><div class="an-mini"><div style="width:${(partyCnt[i] / partyMax * 100).toFixed(0)}%; background:#185FA5;"></div></div></td></tr>`).join('')}
        </tbody>
      </table>
    </div>`;
  // 리드타임대별 ADR (조기예약 vs 임박 요금 차이)
  const leadAdr = buckets.map(() => ({ rev: 0, nights: 0, cnt: 0 }));
  active.forEach(b => { const ld = leadDays_(b); if (ld == null) return; const i = buckets.findIndex(([, lo, hi]) => ld >= lo && ld <= hi); if (i >= 0) { leadAdr[i].rev += (b.gross_amount || 0); leadAdr[i].nights += (b.nights || 0); leadAdr[i].cnt++; } });
  const leadAdrVals = leadAdr.map(o => o.nights > 0 ? o.rev / o.nights : 0);
  const leadAdrMax = Math.max(...leadAdrVals, 1);
  const leadAdrHtml = `
    <div class="an-sec">
      <div class="an-sec-title">11 · 예약 시점별 요금 (조기 vs 임박)</div>
      <div class="an-sec-sub">리드타임 구간별 ADR — 일찍 예약한 손님이 더/덜 냈는지. 조기할인·막판할인 설계 근거</div>
      <table class="an-table">
        <thead><tr><th>예약 시점</th><th>건수</th><th>ADR</th><th class="an-bar-cell">상대</th></tr></thead>
        <tbody>${buckets.map((bk, i) => `
          <tr><td>${bk[0]}</td><td class="num">${leadAdr[i].cnt}건</td>
            <td class="num">${leadAdrVals[i] > 0 ? fmtKRW(leadAdrVals[i]) : '—'}</td>
            <td><div class="an-mini"><div style="width:${(leadAdrVals[i] / leadAdrMax * 100).toFixed(0)}%; background:#8a6d1f;"></div></div></td></tr>`).join('')}
        </tbody>
      </table>
    </div>`;
  // 객실 × 채널 매출 매트릭스
  const mChNames = chRevArr.map(c => c.name);
  const matrix = {}; rooms.forEach(r => { matrix[r.room_number] = {}; });
  active.forEach(b => { const rn = b.rooms?.room_number; const cn = b.channels?.name || '직접'; if (!rn || !matrix[rn]) return; matrix[rn][cn] = (matrix[rn][cn] || 0) + (b.gross_amount || 0); });
  const matrixHtml = `
    <div class="an-sec">
      <div class="an-sec-title">12 · 객실 × 채널 매출 매트릭스</div>
      <div class="an-sec-sub">어느 방이 어느 채널에서 팔리는지 — 채널별 객실 편중 파악</div>
      <table class="an-table">
        <thead><tr><th>객실</th>${mChNames.map(n => `<th style="text-align:right;">${n}</th>`).join('')}<th style="text-align:right;">합계</th></tr></thead>
        <tbody>${rooms.map(r => {
          const row = matrix[r.room_number] || {};
          const tot = Object.values(row).reduce((s, v) => s + v, 0);
          return `<tr><td>${RH(r.room_number)}</td>${mChNames.map(n => `<td class="num">${row[n] ? fmtKRW(row[n]) : '—'}</td>`).join('')}<td class="num" style="font-weight:700;">${fmtKRW(tot)}</td></tr>`;
        }).join('')}
        </tbody>
      </table>
    </div>`;
  // 취소 상세 (채널별 취소·손실매출)
  const cancelByCh = {};
  scope.filter(b => b.booking_status === 'cancelled').forEach(b => { const cn = b.channels?.name || '직접'; const o = cancelByCh[cn] || (cancelByCh[cn] = { cnt: 0, rev: 0 }); o.cnt++; o.rev += (b.gross_amount || 0); });
  const cancelRows = Object.entries(cancelByCh).sort((a, b) => b[1].rev - a[1].rev);
  const cancelHtml = `
    <div class="an-sec">
      <div class="an-sec-title">13 · 취소 분석</div>
      <div class="an-sec-sub">전체 취소율 <b>${cancelRateAll.toFixed(0)}%</b> · 취소로 인한 잠재 매출 손실 <b>${fmtKRW(cancelledRevLost)}</b></div>
      ${cancelRows.length ? `<table class="an-table">
        <thead><tr><th>채널</th><th>취소 건수</th><th>손실 매출</th></tr></thead>
        <tbody>${cancelRows.map(([n, o]) => `<tr><td>${n}</td><td class="num">${o.cnt}건</td><td class="num" style="color:#C0392B;">${fmtKRW(o.rev)}</td></tr>`).join('')}</tbody>
      </table>` : '<div style="color:var(--text-muted);font-size:13px;">이 기간 취소 건이 없습니다. 👍</div>'}
    </div>`;
  // 국적 믹스
  const countryAgg = {};
  active.forEach(b => { const c = b.guest_country || null; const key = c || '미상'; const o = countryAgg[key] || (countryAgg[key] = { cnt: 0, rev: 0, code: c }); o.cnt++; o.rev += (b.gross_amount || 0); });
  const countryRows = Object.entries(countryAgg).sort((a, b) => b[1].cnt - a[1].cnt).slice(0, 10);
  const countryMax = Math.max(...countryRows.map(([, o]) => o.cnt), 1);
  const countryHtml = countryRows.length ? `
    <div class="an-sec">
      <div class="an-sec-title">14 · 게스트 국적 믹스</div>
      <div class="an-sec-sub">국적별 예약 건수·매출 (상위 10) — 언어 안내·타겟 마케팅 참고</div>
      <table class="an-table">
        <thead><tr><th>국적</th><th>건수</th><th class="an-bar-cell">비중</th><th>매출</th></tr></thead>
        <tbody>${countryRows.map(([key, o]) => `
          <tr><td>${o.code ? fmtCountry(o.code) : '미상'}</td><td class="num">${o.cnt}건</td>
            <td><div class="an-mini"><div style="width:${(o.cnt / countryMax * 100).toFixed(0)}%; background:#5a4f8a;"></div></div></td>
            <td class="num">${fmtKRW(o.rev)}</td></tr>`).join('')}
        </tbody>
      </table>
    </div>` : '';
  // ── 🔮 예상 매출 (선택 월이 현재/미래 & 여유 있을 때만) ──
  let forecastHtml = '';
  if (analysisView !== 'all') {
    const nowMonth = new Date().toISOString().slice(0, 7);
    const cap = spanDays * roomCount;
    const remaining = Math.max(0, cap - nightsTotal);
    if (analysisView >= nowMonth && remaining >= 5 && adrAll > 0) {
      const scen = [['보수적', 0.4], ['중립', 0.6], ['낙관', 0.8]];
      const rows = scen.map(([label, rate]) => {
        const addNights = Math.round(remaining * rate);
        const rev = Math.round(revTotal + addNights * adrAll);
        const occ = Math.min(100, (nightsTotal + addNights) / cap * 100);
        return `<tr><td>${label}</td><td>${occ.toFixed(0)}%</td><td style="text-align:right;font-weight:700;">${fmtKRW(rev)}</td></tr>`;
      }).join('');
      const fmo = parseInt(analysisView.split('-')[1], 10);
      forecastHtml = `
      <div class="an-sec" style="border:1px solid #16161b; border-radius:12px; padding:14px 16px; margin-bottom:14px; background:#FBFAF5;">
        <div class="an-sec-title">🔮 ${fmo}월 예상 매출 (Forecast)</div>
        <div class="an-sec-sub">현재 확정 <b>${fmtKRW(revTotal)}</b> · OCC ${occAll.toFixed(0)}% · 남은 가용 ${remaining}박. 남은 날짜가 현재 ADR(${fmtKRW(adrAll)}/박)로 채워질 때 예상.</div>
        <table class="an-table" style="margin-top:8px;"><thead><tr><th>시나리오</th><th>예상 OCC</th><th style="text-align:right;">예상 거래액</th></tr></thead><tbody>${rows}</tbody></table>
      </div>`;
    }
  }
  // ── 🎯 목표 달성률 밴드 (특정 월 · 8월 이후 기본 2,400만, 편집형) ──
  window.__goalOverride = window.__goalOverride || {};
  let goalBandHtml = '';
  if (analysisView !== 'all') {
    const _gmo = parseInt(analysisView.split('-')[1], 10);
    const goal = (analysisView in window.__goalOverride) ? window.__goalOverride[analysisView] : (_gmo >= 8 ? 24000000 : null);
    if (goal && goal > 0) {
      const gpct = Math.min(100, revTotal / goal * 100);
      const gleft = goal - revTotal;
      const _pbtn = (v,l) => `<button onclick="anSetGoal(${v})" style="font-size:12px;font-weight:700;padding:5px 10px;border-radius:8px;border:0;cursor:pointer;background:rgba(255,255,255,.16);color:#fff;">${l}</button>`;
      goalBandHtml = `
      <div style="background:linear-gradient(135deg,#1f1f23,#3f3c36);border-radius:14px;padding:16px 18px;color:#fff;margin-bottom:14px;box-shadow:0 4px 14px rgba(38,50,74,.20);">
        <div style="display:flex;justify-content:space-between;align-items:baseline;flex-wrap:wrap;gap:8px;">
          <div style="font-size:13px;opacity:.82;">${scopeLabel} · 목표 달성률</div>
          <div style="font-size:23px;font-weight:800;">${gpct.toFixed(1)}%</div>
        </div>
        <div style="height:10px;background:rgba(255,255,255,.18);border-radius:6px;overflow:hidden;margin:9px 0;">
          <div style="width:${gpct}%;height:100%;background:linear-gradient(90deg,#4a8a3d,#c9e0c1);border-radius:6px;transition:width .4s;"></div>
        </div>
        <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;font-size:12px;">
          <span style="opacity:.85;">현재 <b>${fmtKRW(revTotal)}</b> · 목표</span>
          <input id="an-goal-input" type="text" value="${goal.toLocaleString()}" style="width:118px;padding:5px 8px;border-radius:8px;border:1px solid rgba(255,255,255,.3);background:rgba(255,255,255,.12);color:#fff;font-size:13px;">
          <button onclick="anSetGoal()" style="font-size:12px;font-weight:700;padding:5px 11px;border-radius:8px;border:0;cursor:pointer;background:#c9e0c1;color:#123;">설정</button>
          ${_pbtn(20000000,'2,000만')}${_pbtn(24000000,'2,400만')}${_pbtn(30000000,'3,000만')}
          <span style="margin-left:auto;opacity:.9;font-weight:700;">${gleft>0?('목표까지 '+fmtKRW(gleft)):'목표 달성 🎉'}</span>
        </div>
      </div>`;
    }
  }
  // ── 🎯 전략 권고안 (실데이터 기반, 우선순위·기대효과 원 단위) ──
  const _lowAdrRoom = [...roomRevArr].filter(r => r.adr > 0).sort((a,b) => a.adr - b.adr)[0];
  const _lowAdrEffect = _lowAdrRoom ? Math.round(_lowAdrRoom.rev * 0.05) : 0;
  const _allEffect5 = Math.round(revTotal * 0.05);
  const _cancelRecover = Math.round(cancelledRevLost * 0.3);
  const _lowOccRoom = [...roomStats].filter(s => s.soldNights > 0).sort((a,b) => a.occ - b.occ)[0];
  const _lowOccInfo = _lowOccRoom ? roomRevArr.find(r => r.room === _lowOccRoom.room) : null;
  const _emptyNights = _lowOccRoom ? Math.max(0, Math.round(spanDays - _lowOccRoom.soldNights)) : 0;
  const _fillN = Math.floor(_emptyNights * 0.5);
  const _fillEffect = (_lowOccInfo && _lowOccInfo.adr > 0) ? Math.round(_fillN * _lowOccInfo.adr) : 0;
  const _topCh = chRevArr[0];
  const _topChShare = _topCh ? (_topCh.rev / totalChRev * 100) : 0;
  const _wdA = wkAdr(wk.weekday), _weA = wkAdr(wk.weekend);
  const _wePrem = _wdA > 0 ? (_weA - _wdA) / _wdA * 100 : 0;
  const _lever = (rank, tone, name, eff, asis, todo, effect) => `
    <div style="background:var(--surface);border:1px solid var(--border);border-radius:12px;padding:13px 15px;margin-bottom:10px;border-left:4px solid ${tone};">
      <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
        <span style="font-size:11px;font-weight:800;background:${tone};color:#fff;padding:3px 9px;border-radius:6px;">${rank}</span>
        <span style="font-size:14px;font-weight:800;color:var(--text);">${name}</span>
        ${eff ? `<span style="margin-left:auto;font-size:12px;color:#4a8a3d;font-weight:800;">${eff}</span>` : ''}
      </div>
      <div style="font-size:12.5px;margin-top:7px;color:var(--text-soft);"><b style="color:#a37e2b;">현황</b> · ${asis}</div>
      <div style="font-size:12.5px;margin-top:4px;color:var(--text);"><b style="color:#3a7030;">실행</b> · ${todo}</div>
      <div style="font-size:12.5px;margin-top:4px;color:var(--text);"><b>효과</b> · ${effect}</div>
    </div>`;
  let _levers = '';
  if (_lowAdrRoom) _levers += _lever('1순위 · 즉효', '#1f1f23', `객단가 인상 (${RH(_lowAdrRoom.room)}·평일)`,
    `+${fmtKRW(_allEffect5)} / 월`,
    `${RH(_lowAdrRoom.room)} ADR ${fmtKRW(_lowAdrRoom.adr)}(최저) · 평일 ADR ${fmtKRW(_wdA)} (주말 대비 ${_wePrem>=0?'−':'+'}${Math.abs(_wePrem).toFixed(0)}%)`,
    `${RH(_lowAdrRoom.room)} +10%, 평일가 +5% → 전 객실 ADR +5%`,
    `${RH(_lowAdrRoom.room)} 월 <b>+${fmtKRW(_lowAdrEffect)}</b> · 전 객실 +5% = <b>+${fmtKRW(_allEffect5)}</b> → 거래액 <b>${fmtKRW(revTotal + _allEffect5)}</b>`);
  if (cancelledRevLost > 0) _levers += _lever('2순위 · 즉효', '#1f1f23', '취소 방어',
    `손실 ${fmtKRW(cancelledRevLost)} 회수`,
    `취소율 ${cancelRateAll.toFixed(0)}% · 손실 ${fmtKRW(cancelledRevLost)}`,
    `비환불(선결제) 요금제↑ · 취소 즉시 재판매 알림`,
    `취소분의 30%만 재판매해도 <b>+${fmtKRW(_cancelRecover)}</b>`);
  if (_lowOccRoom && _fillEffect > 0) _levers += _lever('3순위 · 즉효', '#1f1f23', `${RH(_lowOccRoom.room)} 공실 채우기`,
    `+${fmtKRW(_fillEffect)}`,
    `${RH(_lowOccRoom.room)} 점유 ${_lowOccRoom.occ.toFixed(0)}%(최저) · 빈 ${_emptyNights}박 · ADR ${_lowOccInfo?fmtKRW(_lowOccInfo.adr):'—'}`,
    `막판가·최소박 완화 · 평일 프로모션 (인상보다 채우기 우선)`,
    `빈 ${_emptyNights}박 중 ${_fillN}박만 채워도 <b>+${fmtKRW(_fillEffect)}</b>`);
  if (_topCh && _topChShare > 55) _levers += _lever('4순위 · 구조', '#78736b', '채널 다변화',
    '',
    `${_topCh.name} ${_topChShare.toFixed(0)}% 편중 · 채널 정책 리스크`,
    `타 채널·직접예약 노출↑ → 의존도 60%대로`,
    `리스크 분산 + 직접예약↑ 시 <b>정산액 개선</b>`);
  const strategyHtml = `
    <div class="an-sec" style="border:2px solid #1f1f23;border-radius:14px;padding:16px;margin-bottom:14px;">
      <div class="an-sec-title">🎯 전략 권고안 — 우선순위별 실행안</div>
      <div class="an-sec-sub">${scopeLabel} 실데이터 기반 · 기대효과 원 단위 · 규칙 기반 참고안</div>
      ${_levers || '<div style="color:var(--text-muted);font-size:13px;">전략을 낼 데이터가 부족합니다.</div>'}
      <div style="margin-top:6px;background:var(--accent);color:#fff;border-radius:10px;padding:12px 14px;font-size:12.5px;">
        <b>결론</b> — 1순위(전 객실 +5%)만으로 거래액 <b>${fmtKRW(revTotal + _allEffect5)}</b> 도달. 2·3순위까지 병행 시 추가 <b>+${fmtKRW(_cancelRecover + _fillEffect)}</b> 여력.
      </div>
    </div>`;
  body.innerHTML = `
    ${goalBandHtml}
    ${execSummaryHtml}
    ${forecastHtml}
    ${insightsHtml}
    ${monthlyHtml}
    ${contributionHtml}
    <div class="an-divider-head">📋 상세 분석</div>
    <div class="an-sec">
      <div class="an-sec-title">1 · 채널별 비교</div>
      <div class="an-sec-sub">채널별 확정 건수 · ADR · 평균 리드타임 · 취소율 — 어느 채널에서 언제 얼마 받는지 파악</div>
      ${chTable}
    </div>
    <div class="an-grid2">
      <div class="an-sec">
        <div class="an-sec-title">2 · 평일 / 주말 요금</div>
        <div class="an-sec-sub">체크인 요일 기준 (주말 = 금·토) 박당 평균요금·건수</div>
        ${wkTable}
      </div>
      <div class="an-sec">
        <div class="an-sec-title">3 · 리드타임 분포</div>
        <div class="an-sec-sub">예약 시점(체크인까지 남은 일수)별 건수</div>
        ${ldTable}
      </div>
    </div>
    <div class="an-sec">
      <div class="an-sec-title">4 · 잘 팔린 가격대</div>
      <div class="an-sec-sub">박당 요금 구간별 확정 건수 — 실제 가장 많이 팔린 가격 확인</div>
      ${bandTable}
    </div>
    <div class="an-sec">
      <div class="an-sec-title">5 · 객실별 점유율 · RevPAR · 실판매 기준가</div>
      <div class="an-sec-sub">점유율×가격을 한 표에서 — 점유율 높은데 기준가 낮은 방 = 인상 여력, 반대는 인하 신호</div>
      ${roomStatTable}
    </div>
    <div class="an-sec">
      <div class="an-sec-title">6 · 공실 캘린더 (다가오는 4주)</div>
      <div class="an-sec-sub">객실별 빈 날 — 임박했는데 안 팔린 방은 인하/프로모션 타이밍</div>
      ${vacTable}
    </div>
    ${dowHtml}
    <div class="an-grid2">${losHtml}${partyHtml}</div>
    ${leadAdrHtml}
    ${matrixHtml}
    <div class="an-grid2">${cancelHtml}${countryHtml}</div>
    <div class="an-sec" style="border:2px solid #16161b;">
      <div class="an-sec-title">✦ 프라이싱 제안</div>
      <div class="an-sec-sub">위 분석(실판매 기준가 · 점유율 · 리드타임 · 취소율)을 종합한 객실별 조정 방향 — 규칙 기반, 참고용</div>
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px,1fr)); gap:12px;">${suggCards}</div>
      <div style="font-size:11px; color:var(--text-muted); margin-top:12px;">· 자동 규칙에 따른 참고 제안입니다. 실제 요금은 시장·경쟁·이벤트를 함께 보고 결정하세요.</div>
    </div>
  `;
}
// 목표 금액 설정/편집 (매출 분석 목표 달성률 밴드)
window.anSetGoal = function (v) {
  window.__goalOverride = window.__goalOverride || {};
  let g = v;
  if (g == null) { const el = document.getElementById('an-goal-input'); g = el ? parseInt(el.value.replace(/[^0-9]/g, '')) : 0; }
  if (g > 0 && typeof analysisView !== 'undefined' && analysisView !== 'all') { window.__goalOverride[analysisView] = g; analysisRender(); }
};
