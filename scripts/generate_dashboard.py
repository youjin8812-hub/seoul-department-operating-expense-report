"""
인터랙티브 대시보드 생성 (rename+정리 <- generate_dashboard_enhanced.py)
- output/dashboard.html  (유일한 최종 대시보드. 구버전 output/dashboard.html은 archive/dashboard_v1.html로 이동됨)

임베드 데이터는 해당월/비목/전체부서명/집행금액 4개 필드만 포함한다.
전화번호/작성자/집행대상/제목/문서url 등은 애초에 데이터에 포함하지 않아
구조적으로 개인정보 노출이 불가능하도록 설계했다(화면 미표시가 아니라 데이터 자체 미포함).

필터·집계·해석문 생성은 전량 브라우저 내 자바스크립트에서 수행되며,
페이지 로드시 원본 총액과의 일치 여부를 자체 검증해 배지로 표시한다.
검증 기준값은 이 스크립트가 매 실행마다 원본 CSV에서 재계산해 주입하므로
데이터셋이 바뀌어도 하드코딩된 값에 의존하지 않는다.
"""
import argparse
import json
import os
import pandas as pd
import plotly.offline as pyo

OUT_PATH = "output/dashboard.html"
REQUIRED_COLUMNS = ["해당년도", "해당월", "전체부서명", "비목", "집행금액"]


def load_raw_rows(csv_path):
    df = pd.read_csv(csv_path, encoding="utf-8-sig")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼 누락: {missing}")
    df["집행금액"] = df["집행금액"].astype("int64")
    df["해당월"] = df["해당월"].astype(int)
    rows = [
        {"m": int(r["해당월"]), "c": r["비목"], "d": r["전체부서명"], "a": int(r["집행금액"])}
        for _, r in df[["해당월", "비목", "전체부서명", "집행금액"]].iterrows()
    ]
    return rows, df


def compute_verification_constants(rows, df):
    total = sum(r["a"] for r in rows)
    assert total == int(df["집행금액"].sum()), f"총액 불일치: {total}"
    n_months = df["해당월"].nunique()
    coverage = df.groupby("전체부서명")["해당월"].nunique()
    same_period = coverage[coverage == n_months]
    keys = set()
    for r in rows[:5]:
        keys.update(r.keys())
    assert keys == {"m", "c", "d", "a"}, f"예상치 못한 필드 포함(개인정보 컬럼 유입 의심): {keys}"

    amt = df["집행금액"]
    q1, q3 = amt.quantile(0.25), amt.quantile(0.75)
    iqr_fence = float(q3 + 1.5 * (q3 - q1))
    p95 = float(amt.quantile(0.95))

    years = sorted(df["해당년도"].unique().tolist())
    months = sorted(df["해당월"].unique().tolist())
    period_label = f"{years[0]}년 {months[0]}~{months[-1]}월" if len(years) == 1 else f"{years[0]}~{years[-1]}년"

    print(f"[검증 통과] 총액: {total:,} | 행수: {len(rows)} | 동일기간부서: {len(same_period)}/{df['전체부서명'].nunique()} "
          f"| 전체월수: {n_months} | 임베드 필드: {sorted(keys)}")

    return {
        "total": total, "row_count": len(rows), "n_months": n_months,
        "iqr_fence": iqr_fence, "p95": p95, "period_label": period_label,
    }


HTML_TEMPLATE = r"""<html><head><meta charset="utf-8">
<title>서울시 본청 부서운영업무추진비 분석·보고서 생성기 — 인터랙티브 대시보드</title>
<meta name="description" content="Seoul Department Operating Expense Report Generator — 서울시 본청 부서운영업무추진비 인터랙티브 대시보드 (__PERIOD_LABEL__)">
<meta property="og:title" content="서울시 본청 부서운영업무추진비 분석·보고서 생성기">
<meta property="og:description" content="Seoul Department Operating Expense Report Generator — 필터·KPI·자동 해석문을 갖춘 인터랙티브 HTML 대시보드">
<meta property="og:type" content="website">
<style>
:root{--navy:#1f4e8c;--navy2:#3f72af;--gray:#888;--bg:#f7f8fa;--border:#e2e5ea;}
*{box-sizing:border-box;}
body{font-family:'Malgun Gothic',sans-serif;margin:0;background:var(--bg);color:#222;}
header{background:var(--navy);color:#fff;padding:14px 20px;}
header h1{margin:0;font-size:20px;}
header p{margin:4px 0 0;font-size:12px;opacity:0.85;}
#verifyBadge{display:inline-block;margin-left:12px;padding:2px 10px;border-radius:10px;font-size:11px;font-weight:bold;}
#verifyBadge.ok{background:#2e7d32;}
#verifyBadge.fail{background:#c62828;}
.kpi-row{display:flex;gap:10px;padding:14px 20px 0;flex-wrap:wrap;}
.kpi-card{flex:1;min-width:140px;background:#fff;border:1px solid var(--border);border-radius:8px;padding:10px 12px;}
.kpi-card .label{font-size:11px;color:var(--gray);}
.kpi-card .value{font-size:18px;font-weight:bold;color:var(--navy);margin-top:2px;}
.summary-strip{margin:10px 20px 0;padding:8px 12px;background:#eef2f8;border-left:4px solid var(--navy);font-size:12.5px;border-radius:4px;}
.layout{display:flex;gap:14px;padding:14px 20px 30px;align-items:flex-start;}
#filterPanel{width:250px;flex-shrink:0;background:#fff;border:1px solid var(--border);border-radius:8px;padding:14px;position:sticky;top:14px;}
#filterPanel h3{font-size:13px;margin:0 0 6px;color:var(--navy);}
.filter-block{margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid #eee;}
.filter-block label{display:block;font-size:12px;margin:2px 0;cursor:pointer;}
.filter-block input[type=text]{width:100%;padding:4px;font-size:12px;margin-bottom:4px;box-sizing:border-box;}
#deptList{max-height:150px;overflow-y:auto;border:1px solid #eee;padding:4px;font-size:11px;}
.small-btn{font-size:10.5px;padding:2px 6px;margin-right:4px;cursor:pointer;background:#eee;border:1px solid #ccc;border-radius:3px;}
.metric-radio label{display:inline-block;margin-right:8px;font-size:12px;}
.charts{flex:1;display:grid;grid-template-columns:1fr 1fr;gap:14px;}
.chart-card{background:#fff;border:1px solid var(--border);border-radius:8px;padding:10px;}
.chart-card.full{grid-column:1 / span 2;}
.chart-card h4{margin:0 0 4px;font-size:13px;color:#222;}
.interp{font-size:11.5px;color:#444;background:#f4f6f9;padding:6px 8px;border-radius:4px;margin-top:6px;min-height:2em;}
.footnote{font-size:10.5px;color:#999;padding:0 20px 20px;}
input[type=range]{width:120px;}
#topNVal{font-size:12px;font-weight:bold;}
</style></head>
<body>
<header>
  <h1>서울시 본청 부서운영업무추진비 분석·보고서 생성기 — 인터랙티브 대시보드</h1>
  <p>분석기간 __PERIOD_LABEL__ | 집계 기준: 해당년도→해당월→전체부서명→비목→집행금액 합계 | 개인정보(전화번호/작성자/집행대상) 원문은 이 화면과 데이터에 포함되지 않음
    <span id="verifyBadge">검증 중...</span>
  </p>
</header>

<div class="kpi-row" id="kpiRow"></div>
<div class="summary-strip" id="summaryStrip"></div>

<div class="layout">
  <div id="filterPanel">
    <h3>월 선택</h3>
    <div class="filter-block" id="monthFilter"></div>

    <h3>비목 선택</h3>
    <div class="filter-block" id="catFilter"></div>

    <h3>부서 검색·선택</h3>
    <div class="filter-block">
      <input type="text" id="deptSearch" placeholder="부서명 검색...">
      <span class="small-btn" id="deptAll">전체선택</span><span class="small-btn" id="deptNone">전체해제</span>
      <div id="deptList"></div>
    </div>

    <h3>동일기간 비교</h3>
    <div class="filter-block">
      <label><input type="checkbox" id="samePeriodOnly"> 동일기간(전체 기간) 보유 부서만 보기</label>
    </div>

    <h3>상위 N</h3>
    <div class="filter-block">
      <input type="range" id="topN" min="3" max="20" value="10">
      <span id="topNVal">10</span>개
    </div>

    <h3>지표 전환</h3>
    <div class="filter-block metric-radio" id="metricFilter">
      <label><input type="radio" name="metric" value="sum" checked> 집행액</label>
      <label><input type="radio" name="metric" value="count"> 건수</label>
      <label><input type="radio" name="metric" value="avg"> 평균금액</label>
    </div>
  </div>

  <div class="charts">
    <div class="chart-card"><h4>1. 월별 집행 현황</h4><div id="c1"></div><div class="interp" id="i1"></div></div>
    <div class="chart-card"><h4>2. 월별 비목 구성 변화 (100%)</h4><div id="c2"></div><div class="interp" id="i2"></div></div>
    <div class="chart-card"><h4>3. 부서 비교 (상위N)</h4><div id="c3"></div><div class="interp" id="i3"></div></div>
    <div class="chart-card"><h4>4. 부서×비목 히트맵 (구성비 %)</h4><div id="c4"></div><div class="interp" id="i4"></div></div>
    <div class="chart-card"><h4>5. 부서별 집행액-건수 산점도</h4><div id="c5"></div><div class="interp" id="i5"></div></div>
    <div class="chart-card"><h4>6. 집중도 파레토</h4><div id="c6"></div><div class="interp" id="i6"></div></div>
    <div class="chart-card"><h4>7. 금액 분포 (고액 집행 후보 기준선)</h4><div id="c7"></div><div class="interp" id="i7"></div></div>
    <div class="chart-card full"><h4>8. 반복 집행 패턴 (상위N)</h4><div id="c8"></div><div class="interp" id="i8"></div></div>
  </div>
</div>

<p class="footnote">
집계 기준: 필터 선택 조건 내 해당월×전체부서명×비목×집행금액 | 지표 전환 시 집행액=합계, 건수=행수, 평균금액=합계/행수로 재계산 |
고액 집행 후보는 오류·부적정 집행으로 단정하지 않음 | 반복 패턴의 정례성·사유는 데이터로 확인되지 않음 |
동일기간 비교는 전체 기간 데이터를 보유한 부서만을 대상으로 함 | 이 대시보드는 DOCX/HWPX 보고서의 상세 부가 자료입니다.
</p>

<script>__PLOTLY_JS__</script>
<script>
const RAW = __RAW_DATA__;
const TOTAL_CHECK = __TOTAL_CHECK__;
const ROW_COUNT_CHECK = __ROW_COUNT_CHECK__;
const CATS = Array.from(new Set(RAW.map(r=>r.c))).sort();
const MONTHS = Array.from(new Set(RAW.map(r=>r.m))).sort((a,b)=>a-b);
const CAT_COLOR = {"기관":"#1f4e8c","부서":"#c98a1f","시책":"#3f9142","정원":"#a13f6b"};

const ALL_DEPTS = Array.from(new Set(RAW.map(r=>r.d))).sort();

// 동일기간(전체 월) 보유 부서 계산 (구조적 사실, 필터와 무관하게 원본 데이터 기준)
const deptMonthCoverage = {};
RAW.forEach(r=>{
  if(!deptMonthCoverage[r.d]) deptMonthCoverage[r.d] = new Set();
  deptMonthCoverage[r.d].add(r.m);
});
const SAME_PERIOD_DEPTS = new Set(Object.keys(deptMonthCoverage).filter(d => deptMonthCoverage[d].size === MONTHS.length));

function fmtWon(n){ return Math.round(n).toLocaleString('ko-KR') + '원'; }
function fmtPct(n){ return n.toFixed(1) + '%'; }

// ---------------- 상태 ----------------
const state = {
  months: new Set(MONTHS),
  cats: new Set(CATS),
  deptSelected: new Set(ALL_DEPTS),
  samePeriodOnly: false,
  topN: 10,
  metric: 'sum',
};

// ---------------- 집계 유틸 ----------------
function metricAgg(rows, metric){
  if(metric === 'count') return rows.length;
  const sum = rows.reduce((s,r)=>s+r.a,0);
  if(metric === 'sum') return sum;
  return rows.length ? sum/rows.length : 0;
}
function metricLabel(metric){
  return metric === 'sum' ? '집행액(원)' : (metric === 'count' ? '건수' : '평균금액(원)');
}

function getFilteredData(){
  const activeDepts = state.samePeriodOnly
    ? new Set([...state.deptSelected].filter(d => SAME_PERIOD_DEPTS.has(d)))
    : state.deptSelected;
  return RAW.filter(r => state.months.has(r.m) && state.cats.has(r.c) && activeDepts.has(r.d));
}

// ---------------- 필터 UI 초기화 ----------------
function buildMonthFilter(){
  const el = document.getElementById('monthFilter');
  el.innerHTML = MONTHS.map(m=>`<label><input type="checkbox" class="monthCk" value="${m}" checked> ${m}월</label>`).join('');
  el.querySelectorAll('.monthCk').forEach(cb=>cb.addEventListener('change', ()=>{
    state.months = new Set([...el.querySelectorAll('.monthCk:checked')].map(c=>parseInt(c.value)));
    render();
  }));
}
function buildCatFilter(){
  const el = document.getElementById('catFilter');
  el.innerHTML = CATS.map(c=>`<label><input type="checkbox" class="catCk" value="${c}" checked> ${c}</label>`).join('');
  el.querySelectorAll('.catCk').forEach(cb=>cb.addEventListener('change', ()=>{
    state.cats = new Set([...el.querySelectorAll('.catCk:checked')].map(c=>c.value));
    render();
  }));
}
function buildDeptList(filterText){
  const el = document.getElementById('deptList');
  const text = (filterText||'').trim();
  const shown = text ? ALL_DEPTS.filter(d=>d.includes(text)) : ALL_DEPTS;
  el.innerHTML = shown.map(d=>{
    const checked = state.deptSelected.has(d) ? 'checked' : '';
    const tag = SAME_PERIOD_DEPTS.has(d) ? ' [동일기간]' : '';
    return `<label><input type="checkbox" class="deptCk" value="${d}" ${checked}> ${d}${tag}</label>`;
  }).join('');
  el.querySelectorAll('.deptCk').forEach(cb=>cb.addEventListener('change', (e)=>{
    if(e.target.checked) state.deptSelected.add(e.target.value);
    else state.deptSelected.delete(e.target.value);
    render();
  }));
}
function initFilters(){
  buildMonthFilter();
  buildCatFilter();
  buildDeptList('');
  document.getElementById('deptSearch').addEventListener('input', (e)=> buildDeptList(e.target.value));
  document.getElementById('deptAll').addEventListener('click', ()=>{ state.deptSelected = new Set(ALL_DEPTS); buildDeptList(document.getElementById('deptSearch').value); render(); });
  document.getElementById('deptNone').addEventListener('click', ()=>{ state.deptSelected = new Set(); buildDeptList(document.getElementById('deptSearch').value); render(); });
  document.getElementById('samePeriodOnly').addEventListener('change', (e)=>{ state.samePeriodOnly = e.target.checked; render(); });
  const topNInput = document.getElementById('topN');
  topNInput.addEventListener('input', (e)=>{ state.topN = parseInt(e.target.value); document.getElementById('topNVal').textContent = state.topN; render(); });
  document.querySelectorAll('input[name=metric]').forEach(r=>r.addEventListener('change', (e)=>{ state.metric = e.target.value; render(); }));
}

// ---------------- KPI & 요약 ----------------
function updateKPIs(filtered){
  const total = filtered.reduce((s,r)=>s+r.a,0);
  const count = filtered.length;
  const deptCount = new Set(filtered.map(r=>r.d)).size;
  const byMonth = {};
  filtered.forEach(r=>{ byMonth[r.m] = (byMonth[r.m]||0) + r.a; });
  let topMonth = '-', topMonthVal = -1;
  Object.entries(byMonth).forEach(([m,v])=>{ if(v>topMonthVal){ topMonthVal=v; topMonth=m; } });
  const samePeriodInSelection = new Set([...filtered.map(r=>r.d)].filter(d=>SAME_PERIOD_DEPTS.has(d))).size;

  const cards = [
    ['총 집행액', fmtWon(total)],
    ['집행 건수', count.toLocaleString('ko-KR')+'건'],
    ['부서 수', deptCount+'개'],
    ['최고 집행월', topMonth==='-' ? '-' : topMonth+'월 ('+fmtWon(topMonthVal)+')'],
    ['동일기간 비교가능 부서', samePeriodInSelection+'개 (전체 '+SAME_PERIOD_DEPTS.size+'개)'],
  ];
  document.getElementById('kpiRow').innerHTML = cards.map(([l,v])=>
    `<div class="kpi-card"><div class="label">${l}</div><div class="value">${v}</div></div>`).join('');

  const byCat = {};
  filtered.forEach(r=>{ byCat[r.c] = (byCat[r.c]||0) + r.a; });
  let domCat='-', domCatVal=-1;
  Object.entries(byCat).forEach(([c,v])=>{ if(v>domCatVal){domCatVal=v; domCat=c;} });
  const byDept = {};
  filtered.forEach(r=>{ byDept[r.d] = (byDept[r.d]||0) + r.a; });
  let topDept='-', topDeptVal=-1;
  Object.entries(byDept).forEach(([d,v])=>{ if(v>topDeptVal){topDeptVal=v; topDept=d;} });

  const monthList = [...state.months].sort((a,b)=>a-b).join(',')+'월';
  document.getElementById('summaryStrip').innerHTML =
    `<b>현재 선택</b>: ${monthList} · 비목 ${[...state.cats].join('/')} · 부서 ${deptCount}개 선택`
    + (state.samePeriodOnly ? ' · 동일기간 보유 부서만' : '')
    + ` &nbsp;|&nbsp; <b>주요 비목</b>: ${domCat==='-'?'집계 대상 없음':domCat+' ('+fmtWon(domCatVal)+')'}`
    + ` &nbsp;|&nbsp; <b>최고 집행 부서</b>: ${topDept==='-'?'집계 대상 없음':topDept+' ('+fmtWon(topDeptVal)+')'}`;
}

function updateVerifyBadge(){
  const total = RAW.reduce((s,r)=>s+r.a,0);
  const monthlySum = MONTHS.reduce((s,m)=> s + RAW.filter(r=>r.m===m).reduce((s2,r)=>s2+r.a,0), 0);
  const catSum = CATS.reduce((s,c)=> s + RAW.filter(r=>r.c===c).reduce((s2,r)=>s2+r.a,0), 0);
  const ok = total === TOTAL_CHECK && monthlySum === TOTAL_CHECK && catSum === TOTAL_CHECK && RAW.length === ROW_COUNT_CHECK;
  const badge = document.getElementById('verifyBadge');
  badge.textContent = ok ? `검증 통과 (총액 ${total.toLocaleString('ko-KR')}원 = 기준값 일치, 동일기간부서 ${SAME_PERIOD_DEPTS.size}개 확인)` : '검증 실패 - 데이터 확인 필요';
  badge.className = ok ? 'ok' : 'fail';
}

// ---------------- 차트 1: 월별 집행 현황 ----------------
function renderC1(filtered){
  const vals = MONTHS.map(m => metricAgg(filtered.filter(r=>r.m===m), state.metric));
  const maxIdx = vals.indexOf(Math.max(...vals));
  const colors = vals.map((v,i)=> i===maxIdx ? '#1f4e8c' : '#c9d6e3');
  Plotly.react('c1', [{x: MONTHS.map(m=>m+'월'), y: vals, type:'bar', marker:{color:colors},
    hovertemplate:'%{x}: %{y:,.0f}<extra></extra>'}],
    {margin:{t:10,l:50,r:10,b:30}, height:260, yaxis:{title:metricLabel(state.metric)}});
  const lowIdx = vals.indexOf(Math.min(...vals.filter(v=>true)));
  document.getElementById('i1').textContent = vals.every(v=>v===0) ? '선택된 조건에 해당하는 데이터가 없습니다.' :
    `선택된 조건에서 최고 ${metricLabel(state.metric)} 월은 ${MONTHS[maxIdx]}월(${vals[maxIdx].toLocaleString('ko-KR')}), 최저는 ${MONTHS[lowIdx]}월(${vals[lowIdx].toLocaleString('ko-KR')})입니다.`;
}

// ---------------- 차트 2: 월별 비목 구성 변화 ----------------
function renderC2(filtered){
  const activeCats = [...state.cats];
  const perMonth = MONTHS.map(m=>{
    const rowsM = filtered.filter(r=>r.m===m);
    const raw = {}; activeCats.forEach(c=> raw[c] = metricAgg(rowsM.filter(r=>r.c===c), state.metric));
    const sum = Object.values(raw).reduce((a,b)=>a+b,0);
    const share = {}; activeCats.forEach(c=> share[c] = sum>0 ? raw[c]/sum*100 : 0);
    return share;
  });
  const traces = activeCats.map(c=>({
    x: MONTHS.map(m=>m+'월'), y: perMonth.map(s=>s[c]), name:c, type:'bar',
    marker:{color:CAT_COLOR[c]||'#999'}, hovertemplate:'%{x} '+c+': %{y:.1f}%<extra></extra>'
  }));
  Plotly.react('c2', traces, {barmode:'stack', margin:{t:10,l:40,r:10,b:30}, height:260, yaxis:{title:'비중(%)', range:[0,100]}, legend:{orientation:'h', y:-0.2}});
  let maxDiffCat='-', maxDiffVal=0, maxDiffMonth='-';
  activeCats.forEach(c=>{
    const s = perMonth.map(p=>p[c]);
    const diff = Math.max(...s) - Math.min(...s);
    if(diff > maxDiffVal){ maxDiffVal = diff; maxDiffCat = c; maxDiffMonth = MONTHS[s.indexOf(Math.max(...s))]; }
  });
  document.getElementById('i2').textContent = filtered.length===0 ? '선택된 조건에 해당하는 데이터가 없습니다.' :
    `선택된 조건에서 ${maxDiffCat} 비중의 월별 변동폭이 가장 크며(${maxDiffVal.toFixed(1)}%p), ${maxDiffMonth}월에 가장 높은 비중을 보입니다.`;
}

// ---------------- 차트 3: 부서 비교 ----------------
function renderC3(filtered){
  const depts = Array.from(new Set(filtered.map(r=>r.d)));
  const vals = depts.map(d=> metricAgg(filtered.filter(r=>r.d===d), state.metric));
  let pairs = depts.map((d,i)=>({d, v: vals[i]})).sort((a,b)=>b.v-a.v).slice(0, state.topN);
  pairs = pairs.reverse();
  Plotly.react('c3', [{x: pairs.map(p=>p.v), y: pairs.map(p=>p.d), type:'bar', orientation:'h',
    marker:{color:'#1f4e8c'}, hovertemplate:'%{y}: %{x:,.0f}<extra></extra>'}],
    {margin:{t:10,l:220,r:20,b:30}, height:Math.max(260, pairs.length*22), xaxis:{title:metricLabel(state.metric)}});
  document.getElementById('i3').textContent = pairs.length===0 ? '선택된 조건에 해당하는 부서가 없습니다.' :
    `현재 조건에서 ${metricLabel(state.metric)} 기준 1위는 ${pairs[pairs.length-1].d}(${pairs[pairs.length-1].v.toLocaleString('ko-KR')})이며, 표시된 부서는 ${pairs.length}개입니다.`;
}

// ---------------- 차트 4: 히트맵 ----------------
function renderC4(filtered){
  const depts = Array.from(new Set(filtered.map(r=>r.d)));
  const totals = depts.map(d => metricAgg(filtered.filter(r=>r.d===d), state.metric));
  let order = depts.map((d,i)=>({d, v: totals[i]})).sort((a,b)=>b.v-a.v).slice(0, state.topN).map(o=>o.d);
  const activeCats = [...state.cats];
  const z = order.map(d=>{
    const rowsD = filtered.filter(r=>r.d===d);
    const raw = activeCats.map(c=> metricAgg(rowsD.filter(r=>r.c===c), state.metric));
    const sum = raw.reduce((a,b)=>a+b,0);
    return raw.map(v => sum>0 ? v/sum*100 : 0);
  });
  Plotly.react('c4', [{z, x:activeCats, y:order, type:'heatmap', colorscale:'Blues',
    hovertemplate:'%{y} / %{x}: %{z:.1f}%<extra></extra>'}],
    {margin:{t:10,l:220,r:20,b:30}, height:Math.max(260, order.length*22)});
  document.getElementById('i4').textContent = order.length===0 ? '선택된 조건에 해당하는 부서가 없습니다.' :
    `상위 ${order.length}개 부서의 비목 구성비를 표시합니다. 값은 각 부서 내 비목별 비중(%)입니다.`;
}

// ---------------- 차트 5: 산점도 ----------------
function renderC5(filtered){
  const depts = Array.from(new Set(filtered.map(r=>r.d)));
  const points = depts.map(d=>{
    const rowsD = filtered.filter(r=>r.d===d);
    const sum = rowsD.reduce((s,r)=>s+r.a,0);
    const cnt = rowsD.length;
    const avg = cnt ? sum/cnt : 0;
    const byCat = {}; rowsD.forEach(r=>{ byCat[r.c]=(byCat[r.c]||0)+r.a; });
    let domCat='-', domVal=-1;
    Object.entries(byCat).forEach(([c,v])=>{ if(v>domVal){domVal=v;domCat=c;} });
    return {d, sum, cnt, avg, domCat};
  }).sort((a,b)=>b.sum-a.sum).slice(0, state.topN);
  Plotly.react('c5', [{
    x: points.map(p=>p.cnt), y: points.map(p=>p.sum), text: points.map(p=>p.d+'<br>평균 '+Math.round(p.avg).toLocaleString('ko-KR')+'원<br>주비목: '+p.domCat),
    mode:'markers', type:'scatter',
    marker:{size: points.map(p=>Math.max(8, Math.sqrt(p.avg)/25)), color: points.map(p=>CAT_COLOR[p.domCat]||'#999')},
    hovertemplate:'%{text}<br>건수 %{x} / 총액 %{y:,.0f}<extra></extra>'
  }], {margin:{t:10,l:60,r:10,b:40}, height:280, xaxis:{title:'집행 건수'}, yaxis:{title:'집행액(원)'}});
  document.getElementById('i5').textContent = points.length===0 ? '선택된 조건에 해당하는 부서가 없습니다.' :
    `점 크기는 건당 평균 금액, 색상은 주 비목을 나타냅니다(동일기간 비교 가능 부서 위주로 해석 권장). 표시 부서 ${points.length}개.`;
}

// ---------------- 차트 6: 파레토 ----------------
function renderC6(filtered){
  const depts = Array.from(new Set(filtered.map(r=>r.d)));
  const totals = depts.map(d=>({d, v: metricAgg(filtered.filter(r=>r.d===d), state.metric)})).sort((a,b)=>b.v-a.v);
  const grandTotal = totals.reduce((s,t)=>s+t.v,0);
  const top = totals.slice(0, state.topN);
  let cum = 0;
  const cumShare = top.map(t=>{ cum += t.v; return grandTotal>0 ? cum/grandTotal*100 : 0; });
  Plotly.react('c6', [
    {x: top.map((t,i)=>i+1), y: top.map(t=>t.v), type:'bar', name:metricLabel(state.metric), marker:{color:'#1f4e8c'}, yaxis:'y1'},
    {x: top.map((t,i)=>i+1), y: cumShare, type:'scatter', mode:'lines+markers', name:'누적 점유율(%)', line:{color:'#b23b3b'}, yaxis:'y2'}
  ], {margin:{t:10,l:60,r:60,b:30}, height:260,
      yaxis:{title:metricLabel(state.metric)}, yaxis2:{title:'누적 점유율(%)', overlaying:'y', side:'right', range:[0,105]},
      legend:{orientation:'h', y:-0.2}});
  document.getElementById('i6').textContent = top.length===0 ? '선택된 조건에 해당하는 부서가 없습니다.' :
    `상위 ${top.length}개 부서 누적 점유율은 ${cumShare[cumShare.length-1].toFixed(1)}%입니다. 집중도가 높다는 사실 자체를 부정적으로 해석하지 않습니다.`;
}

// ---------------- 차트 7: 금액 분포 ----------------
const FULL_IQR_FENCE = __IQR_FENCE__;
const FULL_P95 = __P95__;
function renderC7(filtered){
  const amounts = filtered.map(r=>r.a);
  Plotly.react('c7', [{x: amounts, type:'histogram', marker:{color:'#3f72af'}, nbinsx:40}],
    {margin:{t:10,l:50,r:10,b:30}, height:260, xaxis:{title:'집행금액(원)'}, yaxis:{title:'건수'},
     shapes:[
       {type:'line', x0:FULL_IQR_FENCE, x1:FULL_IQR_FENCE, y0:0, y1:1, yref:'paper', line:{color:'#e08a00', dash:'dash'}},
       {type:'line', x0:FULL_P95, x1:FULL_P95, y0:0, y1:1, yref:'paper', line:{color:'#b23b3b', dash:'dash'}},
     ]});
  const overIqr = amounts.filter(a=>a>FULL_IQR_FENCE).length;
  const overP95 = amounts.filter(a=>a>=FULL_P95).length;
  document.getElementById('i7').textContent = amounts.length===0 ? '선택된 조건에 해당하는 데이터가 없습니다.' :
    `점선은 전체 데이터 기준 IQR 상단경계(${FULL_IQR_FENCE.toLocaleString('ko-KR')}원, 주황)와 상위5% 기준선(${FULL_P95.toLocaleString('ko-KR')}원, 빨강)입니다. `
    + `현재 선택 범위에서 IQR 기준 고액 집행 후보 ${overIqr}건, 상위5% 기준 ${overP95}건입니다. 오류·부적정 집행을 의미하지 않습니다.`;
}

// ---------------- 차트 8: 반복 패턴 ----------------
function renderC8(filtered){
  const groups = {};
  filtered.forEach(r=>{
    const key = r.d+'|'+r.c+'|'+r.a;
    if(!groups[key]) groups[key] = {d:r.d, c:r.c, a:r.a, months:new Set(), count:0};
    groups[key].months.add(r.m); groups[key].count++;
  });
  let list = Object.values(groups).filter(g=>g.count>=2).sort((a,b)=>b.count-a.count).slice(0, state.topN);
  list = list.reverse();
  Plotly.react('c8', [{
    x: list.map(g=>g.count), y: list.map(g=>g.d),
    text: list.map(g=>g.c+' · '+g.a.toLocaleString('ko-KR')+'원 · '+g.months.size+'개월 분산'),
    type:'bar', orientation:'h', marker:{color: list.map(g=>CAT_COLOR[g.c]||'#999')},
    hovertemplate:'%{y}<br>%{text}<br>반복 %{x}회<extra></extra>',
  }], {margin:{t:10,l:220,r:10,b:30}, height:Math.max(220, list.length*24)});
  document.getElementById('i8').textContent = list.length===0 ? '현재 조건에서 2회 이상 반복되는 동일 금액 조합이 없습니다.' :
    `현재 조건에서 반복 조합 ${list.length}개를 표시합니다(부서·비목·금액 동일 조합 기준). 반복 사유나 정례성은 데이터로 확인되지 않습니다.`;
}

// ---------------- 렌더 진입점 ----------------
function render(){
  const filtered = getFilteredData();
  updateKPIs(filtered);
  renderC1(filtered);
  renderC2(filtered);
  renderC3(filtered);
  renderC4(filtered);
  renderC5(filtered);
  renderC6(filtered);
  renderC7(filtered);
  renderC8(filtered);
}

initFilters();
updateVerifyBadge();
render();
</script>
</body></html>
"""


def run(csv_path):
    os.makedirs("output", exist_ok=True)
    rows, df = load_raw_rows(csv_path)
    consts = compute_verification_constants(rows, df)

    html = HTML_TEMPLATE.replace("__RAW_DATA__", json.dumps(rows, ensure_ascii=False))
    html = html.replace("__PLOTLY_JS__", pyo.get_plotlyjs())
    html = html.replace("__TOTAL_CHECK__", str(consts["total"]))
    html = html.replace("__ROW_COUNT_CHECK__", str(consts["row_count"]))
    html = html.replace("__IQR_FENCE__", str(consts["iqr_fence"]))
    html = html.replace("__P95__", str(consts["p95"]))
    html = html.replace("__PERIOD_LABEL__", consts["period_label"])

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[generate_dashboard] 저장 완료: {OUT_PATH} | 크기(bytes): {len(html.encode('utf-8')):,}")
    return OUT_PATH


def main():
    parser = argparse.ArgumentParser(description="인터랙티브 대시보드 HTML 생성")
    parser.add_argument("--input", default="input/seoul_expenses.csv", help="원본 CSV 경로")
    args = parser.parse_args()
    run(args.input)


if __name__ == "__main__":
    main()
