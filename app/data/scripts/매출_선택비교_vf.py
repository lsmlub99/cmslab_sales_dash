#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
매출 선택비교 생성기 v1
- 3단(1단/2단/3단) 비교 가능한 HTML 대시보드
- 팀 / 브랜드 / 채널 / 구분(연도) / 월 다중 선택 지원
- 단별 합계 + 단간 증감액 / 증감율 자동 계산
- 그룹화 기준에 따른 상세 표 제공

데이터 추출 로직은 `매출 Dashboard_vf.py`의 extract_data / read_base_date 를 재사용.

Usage:
  python "매출_선택비교_vf.py"
  python "매출_선택비교_vf.py" --no-browser
  python "매출_선택비교_vf.py" --data "경로\파일.xlsx" --out 매출_선택비교.html
"""
import os, sys, json, argparse, webbrowser, subprocess
import importlib.util

# ─────────────────────────────────────────────────────────────────────
# 1. 매출 Dashboard_vf.py 의 데이터 추출 로직 로드 (코드 중복 방지)
# ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
DASHBOARD_MOD_PATH = os.path.join(SCRIPT_DIR, '매출 Dashboard_vf.py')


def load_dashboard_module():
    spec = importlib.util.spec_from_file_location('dashboard_main', DASHBOARD_MOD_PATH)
    mod = importlib.util.module_from_spec(spec)
    sys.modules['dashboard_main'] = mod
    spec.loader.exec_module(mod)  # main() 은 __name__=='__main__' 일 때만 실행되므로 안전
    return mod


# ─────────────────────────────────────────────────────────────────────
# 2. HTML 생성
# ─────────────────────────────────────────────────────────────────────
CSS = """
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --blue:#1a56a0; --blue-dark:#154080;
  --bg:#f5f6fa; --card:#ffffff;
  --border:#e5e7eb; --text:#1a1a2e; --muted:#6c757d;
  --pos:#166534; --neg:#b91c1c;
  --pos-bg:#e8f5e2; --neg-bg:#fdf0e8;
  --tA:#1F4E78; --tB:#375623; --tC:#833C0C;
  --d-team:#3b82f6; --d-brand:#8b5cf6; --d-channel:#0ea5e9; --d-metric:#f59e0b; --d-month:#10b981;
  --shadow:0 1px 4px rgba(0,0,0,.05);
  --radius:10px;
  --ff:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;
}
body{font-family:var(--ff);background:var(--bg);color:var(--text);font-size:13px;line-height:1.5;min-height:100vh}

/* HEADER */
.header{
  background:#fff;height:44px;
  display:flex;align-items:center;padding:0 clamp(12px,2vw,24px);gap:14px;
  box-shadow:0 2px 8px rgba(0,0,0,.10);border-bottom:1px solid var(--border);
  position:sticky;top:0;z-index:200;
}
.hdr-company{color:#833C0C;font-size:15px;font-weight:800;white-space:nowrap}
.hdr-divider{width:1px;height:20px;background:#e5e7eb}
.hdr-badge{background:var(--blue);color:#fff;font-size:12.5px;font-weight:700;padding:4px 14px;border-radius:20px;white-space:nowrap}
.hdr-right{margin-left:auto;color:var(--muted);font-size:11px;text-align:right;white-space:nowrap}
.hdr-right strong{color:var(--text);font-size:12px}

/* MODE BAR */
.mode-bar{
  background:var(--blue);padding:6px clamp(12px,2vw,24px);
  display:flex;align-items:center;gap:8px;flex-wrap:wrap;
  position:sticky;top:44px;z-index:190;
  color:#fff;font-size:11px;font-weight:700;
  box-shadow:0 3px 8px rgba(0,0,0,.15);
}
.mb-label{
  background:#DDEBF7;color:#1a56a0;
  padding:2px 8px;border-radius:5px;border:1px solid #9cc2e8;
  font-size:11px;font-weight:800;
}
.mb-sep{color:rgba(255,255,255,.3);padding:0 4px;font-size:13px}
.mb-btn,.sync-btn{
  padding:3px 12px;border-radius:14px;cursor:pointer;
  font-size:11px;font-weight:700;font-family:var(--ff);
  background:rgba(255,255,255,.12);
  border:1px solid rgba(255,255,255,.3);color:rgba(255,255,255,.85);
  transition:all .14s;
}
.mb-btn:hover,.sync-btn:hover{background:rgba(255,255,255,.22);color:#fff}
.mb-btn.active{background:#fff;color:var(--blue);border-color:#fff}
.sync-btn.active{background:#f59e0b;border-color:#d97706;color:#fff}

/* LAYOUT */
.layout{display:flex;gap:12px;padding:12px clamp(12px,2vw,24px) 40px;align-items:flex-start}
.palette{
  flex:0 0 220px;
  background:var(--card);border:1px solid var(--border);border-radius:var(--radius);
  box-shadow:var(--shadow);padding:8px 10px;
  position:sticky;top:96px;max-height:calc(100vh - 110px);overflow-y:auto;
}
.palette::-webkit-scrollbar{width:5px}
.palette::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px}
.pal-hdr{font-size:10.5px;font-weight:700;color:var(--muted);padding:2px 4px 6px;border-bottom:1px solid var(--border);margin-bottom:6px;line-height:1.4}
.pal-dim{margin-bottom:4px}
.pal-dim summary{
  list-style:none;cursor:pointer;
  padding:5px 6px;border-radius:6px;
  font-size:12px;font-weight:800;color:#374151;
  display:flex;align-items:center;gap:6px;
  user-select:none;
}
.pal-dim summary::-webkit-details-marker{display:none}
.pal-dim summary::before{content:"▶";font-size:8px;color:#9ca3af;transition:transform .15s}
.pal-dim[open] summary::before{transform:rotate(90deg)}
.pal-dim summary:hover{background:#f3f4f6}
.pal-tag{display:inline-block;width:8px;height:8px;border-radius:2px;flex-shrink:0}
.pal-tag.team{background:var(--d-team)}
.pal-tag.brand{background:var(--d-brand)}
.pal-tag.channel{background:var(--d-channel)}
.pal-tag.metric{background:var(--d-metric)}
.pal-tag.month{background:var(--d-month)}
.pal-count{margin-left:auto;color:#9ca3af;font-weight:600;font-size:10.5px}
.pal-chips{display:flex;flex-direction:column;gap:3px;padding:4px 2px 4px 12px}
.pal-chip{
  display:flex;align-items:center;gap:6px;
  padding:4px 8px;border-radius:6px;cursor:grab;
  font-size:11.5px;font-weight:600;color:#374151;
  background:#f9fafb;border:1px solid #e5e7eb;
  transition:all .12s;user-select:none;
}
.pal-chip:hover{background:#fff;border-color:#9ca3af;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.pal-chip:active{cursor:grabbing}
.pal-chip.dragging{opacity:.4}
.pal-chip.all{background:#fef9c3;border-color:#fde047;color:#92400e;font-weight:800}
.pal-chip.all:hover{background:#fef08a;border-color:#facc15}
.pal-chip-label{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pal-presence{display:flex;gap:2px;flex-shrink:0}
.pal-presence span{
  display:inline-block;width:14px;height:14px;border-radius:50%;
  font-size:9px;font-weight:800;color:#fff;text-align:center;line-height:14px;
  background:#e5e7eb;
}
.pal-presence span.on{}
.pal-presence span.on.A{background:var(--tA)}
.pal-presence span.on.B{background:var(--tB)}
.pal-presence span.on.C{background:var(--tC)}

/* RIGHT AREA */
.right{flex:1;display:flex;flex-direction:column;gap:12px;min-width:0}

/* TIER GRID */
.tier-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
.tier{
  background:var(--card);border:1px solid var(--border);border-radius:var(--radius);
  box-shadow:var(--shadow);padding:10px 12px 12px;
  display:flex;flex-direction:column;gap:6px;
  min-height:200px;transition:background .15s,border-color .15s,box-shadow .15s;
}
.tier[data-tier="A"]{border-top:4px solid var(--tA)}
.tier[data-tier="B"]{border-top:4px solid var(--tB)}
.tier[data-tier="C"]{border-top:4px solid var(--tC)}
.tier.drag-over{background:#fef9c3;border-color:#f59e0b;box-shadow:0 0 0 3px rgba(245,158,11,.3)}
.tier.none-mode{background:#f3f4f6}
.tier.none-mode .tier-body{opacity:.45}
.tier-head{display:flex;align-items:center;justify-content:space-between;padding-bottom:6px;border-bottom:1px solid var(--border)}
.tier-title{font-size:22px;font-weight:900;line-height:1}
.tier[data-tier="A"] .tier-title{color:var(--tA)}
.tier[data-tier="B"] .tier-title{color:var(--tB)}
.tier[data-tier="C"] .tier-title{color:var(--tC)}
.tier-actions{display:flex;gap:4px}
.tier-btn{
  background:#f3f4f6;border:1px solid var(--border);
  padding:3px 9px;border-radius:6px;font-size:10.5px;font-weight:700;color:var(--muted);
  cursor:pointer;font-family:var(--ff);
}
.tier-btn:hover{background:#e5e7eb;color:var(--text)}
.tier-btn.none{background:#fee2e2;border-color:#fca5a5;color:#b91c1c}
.tier-btn.none:hover{background:#fecaca}
.tier-body{display:flex;flex-direction:column;gap:5px;flex:1}
.tier-dim{display:flex;flex-direction:row;align-items:flex-start;gap:6px;font-size:11px}
.td-label{
  flex:0 0 38px;font-weight:800;color:#fff;text-align:center;
  padding:2px 4px;border-radius:4px;line-height:1.3;font-size:10.5px;
}
.td-label.team{background:var(--d-team)}
.td-label.brand{background:var(--d-brand)}
.td-label.channel{background:var(--d-channel)}
.td-label.metric{background:var(--d-metric)}
.td-label.month{background:var(--d-month)}
.td-chips{flex:1;display:flex;flex-wrap:wrap;gap:3px;min-height:18px;align-items:center}
.td-chip{
  display:inline-flex;align-items:center;gap:2px;
  padding:1px 3px 1px 7px;border-radius:10px;
  font-size:11px;font-weight:600;max-width:170px;
  background:#e5e7eb;color:#374151;
}
.td-chip-label{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:140px}
.td-chip-x{cursor:pointer;font-size:13px;line-height:1;color:#6b7280;font-weight:700;padding:0 4px}
.td-chip-x:hover{color:#dc2626}
.td-all{color:var(--muted);font-style:italic;font-size:10.5px}
.tier-summary{
  margin-top:6px;padding-top:8px;border-top:1px dashed var(--border);
  display:flex;justify-content:space-between;align-items:baseline;font-size:11px;color:var(--muted);
}
.tier-sum-amt{font-size:18px;font-weight:800;color:var(--text)}
.tier-sum-amt.zero{color:var(--muted)}
.kpi-unit{font-size:11px;font-weight:500;color:var(--muted);margin-left:3px}

/* KPI */
.kpi-strip{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}
.kpi-card{
  background:var(--card);border:1px solid var(--border);border-radius:var(--radius);
  box-shadow:var(--shadow);padding:10px 14px;
  display:flex;flex-direction:column;gap:6px;min-width:0;
}
.kpi-card.tierA{border-left:4px solid var(--tA)}
.kpi-card.tierB{border-left:4px solid var(--tB)}
.kpi-card.tierC{border-left:4px solid var(--tC)}
.kpi-card.delta{background:#fafafa}
.kpi-tag{font-size:11px;font-weight:800;color:var(--muted);letter-spacing:.3px}
.kpi-sub{font-size:10px;color:#9ca3af;line-height:1.3;min-height:24px;overflow:hidden;text-overflow:ellipsis}
.kpi-amt{font-size:20px;font-weight:800;color:var(--text);line-height:1.1}
.kpi-delta-row{display:flex;align-items:baseline;gap:8px;flex-wrap:wrap}
.kpi-delta-amt{font-size:17px;font-weight:800}
.kpi-delta-pct{font-size:12px;font-weight:700}
.pos{color:var(--pos)}
.neg{color:var(--neg)}
.zero,.neu{color:var(--muted)}

/* SECTION */
.section-card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden}
.section-head{display:flex;align-items:center;gap:14px;padding:10px 14px;border-bottom:1px solid var(--border);flex-wrap:wrap}
.section-title{font-size:13px;font-weight:800}
.section-sub{font-size:11px;color:var(--muted)}
.gb-area{display:flex;align-items:center;gap:6px;margin-left:auto;flex-wrap:wrap}
.gb-label{font-size:11px;font-weight:800;color:#374151;background:#DDEBF7;padding:3px 8px;border-radius:5px;border:1px solid #9cc2e8}
.gb-chip{padding:3px 10px;border-radius:14px;font-size:11px;font-weight:700;cursor:pointer;background:#f3f4f6;color:#6b7280;border:1px solid #e5e7eb;font-family:var(--ff)}
.gb-chip:hover{background:#e5e7eb}
.gb-chip.active{background:var(--blue);color:#fff;border-color:var(--blue)}
.gb-chip.auto{background:#fef08a;color:#92400e;border-color:#fbbf24}
.gb-chip.auto.active{background:#f59e0b;color:#fff;border-color:#d97706}

.tbl-wrap{overflow:auto;max-height:580px}
.tbl-wrap::-webkit-scrollbar{width:6px;height:6px}
.tbl-wrap::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px}
table.cmp{border-collapse:collapse;width:100%;font-size:12.5px;white-space:nowrap}
table.cmp thead{position:sticky;top:0;z-index:5}
table.cmp thead th{padding:8px 10px;text-align:center;border:1px solid #3d5a80;font-size:12px;font-weight:700;color:#fff}
.h-grp{background:#374151}
.h-tA{background:var(--tA)}
.h-tB{background:var(--tB)}
.h-tC{background:var(--tC)}
.h-delta{background:#5C2508}
table.cmp tbody td{padding:6px 10px;border-bottom:1px solid #f3f4f6;border-right:1px solid #f3f4f6;text-align:right;font-weight:500}
table.cmp tbody td.grp{text-align:left;font-weight:700;color:#1e3a5f;background:#fff}
table.cmp tbody tr:nth-child(even) td{background:#fafafa}
table.cmp tbody tr:hover td{background:#f0f4ff}
table.cmp tfoot td{padding:8px 10px;text-align:right;font-weight:800;font-size:12.5px;background:#374151;color:#fff;border-right:1px solid #475569}
table.cmp tfoot td.grp{text-align:left;background:#1f2937;color:#fff}
td.pos{color:var(--pos);background:var(--pos-bg) !important}
td.neg{color:var(--neg);background:var(--neg-bg) !important}

/* responsive */
@media (max-width:1280px){.kpi-strip{grid-template-columns:repeat(3,1fr)}}
@media (max-width:1024px){
  .layout{flex-direction:column}
  .palette{flex:0 0 auto;position:static;max-height:280px;width:100%}
  .tier-grid{grid-template-columns:1fr}
  .kpi-strip{grid-template-columns:repeat(2,1fr)}
}
"""


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>매출 선택비교 | CMSLAB</title>
<style>__CSS__</style>
</head>
<body>
<header class="header">
  <div class="hdr-company">CMSLAB</div>
  <div class="hdr-divider"></div>
  <div class="hdr-badge">매출 선택비교</div>
  <div class="hdr-right">
    기준일자 <strong>__BASE_DATE__</strong><br>
    레코드 <strong id="hdrRecCount">0</strong>건
  </div>
</header>

<div class="mode-bar">
  <span class="mb-label">선택모드</span>
  <button class="mb-btn active" data-mode="multi">복수</button>
  <button class="mb-btn" data-mode="single">단수</button>
  <span class="mb-sep">|</span>
  <span class="mb-label">일괄적용</span>
  <button class="sync-btn" data-sync="team">팀</button>
  <button class="sync-btn" data-sync="brand">브랜드</button>
  <button class="sync-btn" data-sync="channel">채널</button>
  <button class="sync-btn" data-sync="metric">구분</button>
  <button class="sync-btn" data-sync="month">월</button>
  <span class="mb-sep">|</span>
  <button class="sync-btn" id="syncAll">전체</button>
</div>

<div class="layout">
  <aside class="palette" id="palette">
    <div class="pal-hdr">팔레트<br>드래그 → A·B·C 칸에 놓기 / 클릭은 비활성</div>
  </aside>

  <div class="right">
    <section class="tier-grid" id="tierGrid"></section>
    <section class="kpi-strip" id="kpiStrip"></section>
    <section class="section-card">
      <div class="section-head">
        <div>
          <div class="section-title">상세 비교</div>
          <div class="section-sub">A·B·C 매출과 단간 증감 · 그룹화 기준 다중 선택 가능</div>
        </div>
        <div class="gb-area">
          <span class="gb-label">그룹화</span>
          <button class="gb-chip auto active" data-gb="__auto__">자동 분해</button>
          <button class="gb-chip" data-gb="team">팀</button>
          <button class="gb-chip" data-gb="brand">브랜드</button>
          <button class="gb-chip" data-gb="channel">채널</button>
          <button class="gb-chip" data-gb="month">월</button>
        </div>
      </div>
      <div class="tbl-wrap">
        <table class="cmp" id="cmpTable"></table>
      </div>
    </section>
  </div>
</div>

<script>
const RAW = __DATA__;
__JSAPP__
</script>
</body>
</html>
"""


JS_APP = r"""
// ─── 메타 ───
const METRICS = [
  { key:'y2024',  label:'2024 실적' },
  { key:'y2025',  label:'2025 실적' },
  { key:'actual', label:'2026 실적' },
  { key:'plan',   label:'2026 계획' },
];
const MONTHS = [1,2,3,4,5,6,7,8,9,10,11,12];
const TEAM_ORDER  = ['RBD1팀','RBD2팀','동북아MC팀','Global사업팀','GEC팀','일본사업팀','중국사업팀','메디컬팀'];
const BRAND_ORDER = ['CFC','CEX','DMB','SUS','기타'];

function uniqOrdered(records, key, order){
  const set = new Set(records.map(r=>r[key]));
  const known = order.filter(v=>set.has(v));
  const extras = [...set].filter(v=>!order.includes(v)).sort();
  return [...known, ...extras];
}
const TEAMS    = uniqOrdered(RAW,'team',TEAM_ORDER);
const BRANDS   = uniqOrdered(RAW,'brand',BRAND_ORDER);
const CHANNELS = [...new Set(RAW.map(r=>r.channel))].sort((a,b)=>a.localeCompare(b,'ko'));

const DIMS = [
  { key:'team',    label:'팀',    values:TEAMS,    fmt:v=>v },
  { key:'brand',   label:'브랜드', values:BRANDS,   fmt:v=>v },
  { key:'channel', label:'채널',   values:CHANNELS, fmt:v=>v },
  { key:'metric',  label:'구분',   values:METRICS.map(m=>m.key), fmt:k=>(METRICS.find(m=>m.key===k)||{}).label||k },
  { key:'month',   label:'월',     values:MONTHS,   fmt:v=>v+'월' },
];
function dimByKey(k){ return DIMS.find(d=>d.key===k); }

const TIERS = ['A','B','C'];
const DEFAULT_METRIC = { A:'y2025', B:'plan', C:'actual' };
function makeDefaultTier(t){
  return { team:new Set(), brand:new Set(), channel:new Set(), metric:new Set([DEFAULT_METRIC[t]]), month:new Set() };
}
function makeEmptyTier(){
  return { team:new Set(), brand:new Set(), channel:new Set(), metric:new Set(), month:new Set() };
}
const SEL = { A:makeDefaultTier('A'), B:makeDefaultTier('B'), C:makeDefaultTier('C') };

let MODE = 'multi';   // 'single' | 'multi'
const SYNC = { team:false, brand:false, channel:false, metric:false, month:false };
const GROUP_BY = new Set();
let AUTO_SPLIT = true;
const PRESENCE_DIRTY = { changed:true };

// ─── 포매팅 ───
function fmt(v, digits=1){
  if (v == null || isNaN(v)) return '-';
  const sign = v < 0 ? '-' : '';
  return sign + Math.abs(v).toLocaleString('ko-KR', { minimumFractionDigits:digits, maximumFractionDigits:digits });
}
function fmtSigned(v, digits=1){
  if (v == null || isNaN(v)) return '-';
  if (v === 0) return '0';
  return (v > 0 ? '+' : '-') + Math.abs(v).toLocaleString('ko-KR', { minimumFractionDigits:digits, maximumFractionDigits:digits });
}
function fmtPct(v){
  if (v == null || !isFinite(v)) return '-';
  if (v === 0) return '0.0%';
  return (v > 0 ? '+' : '-') + Math.abs(v).toFixed(1) + '%';
}
function deltaClass(v){ return v > 0 ? 'pos' : v < 0 ? 'neg' : 'zero'; }
function parseChipVal(dim, val){
  if (val === '__ALL__') return '__ALL__';
  return (dim === 'month') ? parseInt(val,10) : val;
}

// ─── 필터/합계 ───
function recordMatches(r, sel){
  if (sel.team.size    && !sel.team.has(r.team))       return false;
  if (sel.brand.size   && !sel.brand.has(r.brand))     return false;
  if (sel.channel.size && !sel.channel.has(r.channel)) return false;
  if (sel.month.size   && !sel.month.has(r.month))     return false;
  return true;
}
function recordValue(r, sel){
  let v = 0;
  for (const m of sel.metric) v += (r[m] || 0);
  return v;
}
function tierSum(tier){
  const sel = SEL[tier];
  if (sel.metric.size === 0) return 0;
  let sum = 0;
  for (const r of RAW){
    if (!recordMatches(r, sel)) continue;
    sum += recordValue(r, sel);
  }
  return sum;
}
function tierIsNone(tier){
  const s = SEL[tier];
  return s.metric.size === 0 && !s.team.size && !s.brand.size && !s.channel.size && !s.month.size;
}

// ─── 상태 변경 ───
function addValue(tier, dim, val){
  const targets = SYNC[dim] ? TIERS : [tier];
  if (val === '__ALL__'){
    // 전체 = 해당 차원의 필터 비우기 (모든 값 포함)
    for (const t of targets) SEL[t][dim].clear();
  } else {
    for (const t of targets){
      if (MODE === 'single') SEL[t][dim].clear();
      SEL[t][dim].add(val);
    }
  }
  renderAll();
}
function removeValue(tier, dim, val){
  const targets = SYNC[dim] ? TIERS : [tier];
  for (const t of targets){
    SEL[t][dim].delete(val);
  }
  renderAll();
}
function resetTier(tier){
  SEL[tier] = makeDefaultTier(tier);
  renderAll();
}
function noneTier(tier){
  SEL[tier] = makeEmptyTier();
  renderAll();
}

// ─── 팔레트 렌더 ───
function tiersContaining(dim, val){
  return TIERS.filter(t => SEL[t][dim].has(val));
}
function renderPalette(){
  const pal = document.getElementById('palette');
  // 헤더 외 dim 섹션 비우기
  pal.querySelectorAll('.pal-dim').forEach(n=>n.remove());
  for (const d of DIMS){
    const det = document.createElement('details');
    det.className = 'pal-dim';
    det.dataset.dim = d.key;
    det.open = (d.key !== 'channel');   // 채널은 길어 기본 접힘
    const summary = document.createElement('summary');
    summary.innerHTML = `<span class="pal-tag ${d.key}"></span>${d.label}<span class="pal-count">${d.values.length}</span>`;
    det.appendChild(summary);
    const chips = document.createElement('div');
    chips.className = 'pal-chips';
    // 전체 칩 — 드롭 시 해당 차원 필터 비움 (= 전체)
    {
      const all = document.createElement('div');
      all.className = 'pal-chip all';
      all.draggable = true;
      all.dataset.dim = d.key;
      all.dataset.val = '__ALL__';
      const allPresent = TIERS.filter(t => SEL[t][d.key].size === 0);
      const allPres = TIERS.map(t => `<span class="${allPresent.includes(t)?'on '+t:''}">${t}</span>`).join('');
      all.innerHTML = `<span class="pal-chip-label">전체</span><span class="pal-presence">${allPres}</span>`;
      chips.appendChild(all);
    }
    for (const v of d.values){
      const ch = document.createElement('div');
      ch.className = 'pal-chip';
      ch.draggable = true;
      ch.dataset.dim = d.key;
      ch.dataset.val = String(v);
      const present = tiersContaining(d.key, v);
      const pres = TIERS.map(t => `<span class="${present.includes(t)?'on '+t:''}">${t}</span>`).join('');
      ch.innerHTML = `<span class="pal-chip-label">${d.fmt(v)}</span><span class="pal-presence">${pres}</span>`;
      chips.appendChild(ch);
    }
    det.appendChild(chips);
    pal.appendChild(det);
  }
}

// ─── A/B/C 패널 렌더 ───
function renderTierGrid(){
  const grid = document.getElementById('tierGrid');
  grid.innerHTML = TIERS.map(renderTier).join('');
  document.getElementById('hdrRecCount').textContent = RAW.length.toLocaleString();
}
function renderTier(t){
  const sel = SEL[t];
  const isNone = tierIsNone(t);
  let dimRows = '';
  for (const d of DIMS){
    const chips = [...sel[d.key]];
    const inner = chips.length === 0
      ? `<span class="td-all">— 전체 —</span>`
      : chips.map(v=>{
          const lbl = d.fmt(v);
          return `<span class="td-chip" data-tier="${t}" data-dim="${d.key}" data-val="${v}"><span class="td-chip-label">${lbl}</span><span class="td-chip-x">×</span></span>`;
        }).join('');
    dimRows += `
      <div class="tier-dim">
        <span class="td-label ${d.key}">${d.label}</span>
        <div class="td-chips">${inner}</div>
      </div>`;
  }
  const total = tierSum(t);
  return `
    <div class="tier ${isNone?'none-mode':''}" data-tier="${t}">
      <div class="tier-head">
        <div class="tier-title">${t}</div>
        <div class="tier-actions">
          <button class="tier-btn" data-action="reset" data-tier="${t}">초기화</button>
          <button class="tier-btn none" data-action="none" data-tier="${t}">없음</button>
        </div>
      </div>
      <div class="tier-body">${dimRows}</div>
      <div class="tier-summary">
        <span>합계</span>
        <span class="tier-sum-amt ${total===0?'zero':''}">${fmt(total)}<span class="kpi-unit">백만</span></span>
      </div>
    </div>`;
}

// ─── KPI ───
function renderKPI(){
  const totals = { A: tierSum('A'), B: tierSum('B'), C: tierSum('C') };
  function deltaCard(tag, sub, a, b){
    const diff = a - b;
    const pct = b === 0 ? null : (diff / Math.abs(b)) * 100;
    const dCls = deltaClass(diff);
    return `
      <div class="kpi-card delta">
        <div class="kpi-tag">${tag}</div>
        <div class="kpi-sub">${sub}</div>
        <div class="kpi-delta-row">
          <span class="kpi-delta-amt ${dCls}">${fmtSigned(diff)}<span class="kpi-unit">백만</span></span>
          <span class="kpi-delta-pct ${dCls}">${fmtPct(pct)}</span>
        </div>
      </div>`;
  }
  const tierCard = (t) => `
    <div class="kpi-card tier${t}">
      <div class="kpi-tag" style="color:var(--t${t})">${t} 합계</div>
      <div class="kpi-sub">${shortSummary(SEL[t])}</div>
      <div><span class="kpi-amt">${fmt(totals[t])}</span><span class="kpi-unit">백만</span></div>
    </div>`;
  document.getElementById('kpiStrip').innerHTML =
    tierCard('A') + tierCard('B') + tierCard('C')
    + deltaCard('Δ B − A', 'B가 A 대비', totals.B, totals.A)
    + deltaCard('Δ C − B', 'C가 B 대비', totals.C, totals.B)
    + deltaCard('Δ C − A', 'C가 A 대비', totals.C, totals.A);
}
function shortSummary(sel){
  if (sel.metric.size === 0 && !sel.team.size && !sel.brand.size && !sel.channel.size && !sel.month.size)
    return '없음 (0)';
  const parts = [];
  if (sel.metric.size) parts.push([...sel.metric].map(k=>(METRICS.find(m=>m.key===k)||{}).label||k).join('+'));
  if (sel.month.size)  parts.push([...sel.month].sort((a,b)=>a-b).map(m=>m+'월').join(','));
  const nm = { team:'팀', brand:'브랜드', channel:'채널' };
  for (const d of ['team','brand','channel']){
    if (sel[d].size === 1) parts.push([...sel[d]][0]);
    else if (sel[d].size > 1) parts.push(`${nm[d]} ${sel[d].size}개`);
  }
  return parts.join(' · ') || '전체';
}

// ─── 자동 분해 ───
function autoSplitDims(){
  if (!AUTO_SPLIT) return [];
  const out = [];
  for (const d of ['team','brand','channel','month']){
    const u = new Set();
    for (const t of TIERS) for (const v of SEL[t][d]) u.add(v);
    if (u.size > 1) out.push(d);
  }
  return out;
}
function activeGroupDims(){
  const set = new Set([...GROUP_BY, ...autoSplitDims()]);
  return DIMS.map(d=>d.key).filter(k=>set.has(k));
}

// ─── 비교 표 ───
function renderTable(){
  const dims = activeGroupDims();
  const tbl = document.getElementById('cmpTable');

  let rowKeys;
  if (dims.length === 0){
    rowKeys = [{}];
  } else {
    const seen = new Map();
    for (const r of RAW){
      const key = dims.map(d=>r[d]).join('§');
      if (!seen.has(key)){
        const gv = {};
        for (const d of dims) gv[d] = r[d];
        seen.set(key, gv);
      }
    }
    const orderMap = {
      team:    v => { const i = TEAMS.indexOf(v);    return i < 0 ? 999 : i; },
      brand:   v => { const i = BRANDS.indexOf(v);   return i < 0 ? 999 : i; },
      channel: v => { const i = CHANNELS.indexOf(v); return i < 0 ? 999 : i; },
      month:   v => v - 1,
    };
    rowKeys = [...seen.values()].sort((a,b)=>{
      for (const d of dims){
        const af = orderMap[d](a[d]), bf = orderMap[d](b[d]);
        if (af !== bf) return af - bf;
      }
      return 0;
    });
  }

  function tierGroupSum(t, gv){
    const sel = SEL[t];
    if (sel.metric.size === 0) return 0;
    let sum = 0;
    for (const r of RAW){
      let ok = true;
      for (const d of dims){ if (r[d] !== gv[d]) { ok=false; break; } }
      if (!ok) continue;
      if (!recordMatches(r, sel)) continue;
      sum += recordValue(r, sel);
    }
    return sum;
  }

  let totA=0, totB=0, totC=0;
  const rows = [];
  for (const gv of rowKeys){
    const vA = tierGroupSum('A', gv);
    const vB = tierGroupSum('B', gv);
    const vC = tierGroupSum('C', gv);
    if (dims.length > 0 && vA === 0 && vB === 0 && vC === 0) continue;
    totA += vA; totB += vB; totC += vC;
    const dBA = vB - vA, pBA = vA === 0 ? null : (dBA/Math.abs(vA))*100;
    const dCB = vC - vB, pCB = vB === 0 ? null : (dCB/Math.abs(vB))*100;
    const dCA = vC - vA, pCA = vA === 0 ? null : (dCA/Math.abs(vA))*100;
    const grpCells = dims.length === 0
      ? `<td class="grp">전체</td>`
      : dims.map(d=>{
          const v = gv[d];
          const lbl = (d === 'month') ? (v+'월') : v;
          return `<td class="grp">${lbl}</td>`;
        }).join('');
    rows.push(`<tr>${grpCells}
      <td>${fmt(vA)}</td>
      <td>${fmt(vB)}</td>
      <td>${fmt(vC)}</td>
      <td class="${deltaClass(dBA)}">${fmtSigned(dBA)}</td>
      <td class="${deltaClass(dBA)}">${fmtPct(pBA)}</td>
      <td class="${deltaClass(dCB)}">${fmtSigned(dCB)}</td>
      <td class="${deltaClass(dCB)}">${fmtPct(pCB)}</td>
      <td class="${deltaClass(dCA)}">${fmtSigned(dCA)}</td>
      <td class="${deltaClass(dCA)}">${fmtPct(pCA)}</td>
    </tr>`);
  }
  const fBA = totB-totA, fPBA = totA===0?null:(fBA/Math.abs(totA))*100;
  const fCB = totC-totB, fPCB = totB===0?null:(fCB/Math.abs(totB))*100;
  const fCA = totC-totA, fPCA = totA===0?null:(fCA/Math.abs(totA))*100;

  const dimLabels = dims.length === 0 ? ['합계'] : dims.map(d=>dimByKey(d).label);
  const hdrGrp = dimLabels.map(l=>`<th rowspan="2" class="h-grp">${l}</th>`).join('');
  const colspanFoot = dimLabels.length;
  const colsTotal = dimLabels.length + 9;

  tbl.innerHTML = `
    <thead>
      <tr>
        ${hdrGrp}
        <th colspan="3" class="h-grp">매출 (백만원)</th>
        <th colspan="2" class="h-delta">Δ B vs A</th>
        <th colspan="2" class="h-delta">Δ C vs B</th>
        <th colspan="2" class="h-delta">Δ C vs A</th>
      </tr>
      <tr>
        <th class="h-tA">A</th>
        <th class="h-tB">B</th>
        <th class="h-tC">C</th>
        <th class="h-delta">증감액</th>
        <th class="h-delta">증감율</th>
        <th class="h-delta">증감액</th>
        <th class="h-delta">증감율</th>
        <th class="h-delta">증감액</th>
        <th class="h-delta">증감율</th>
      </tr>
    </thead>
    <tbody>
      ${rows.length ? rows.join('') : `<tr><td colspan="${colsTotal}" style="text-align:center;padding:20px;color:var(--muted)">선택 조건에 해당하는 데이터가 없습니다.</td></tr>`}
    </tbody>
    ${ dims.length === 0 ? '' : `
    <tfoot>
      <tr>
        <td class="grp" colspan="${colspanFoot}">합계</td>
        <td>${fmt(totA)}</td>
        <td>${fmt(totB)}</td>
        <td>${fmt(totC)}</td>
        <td>${fmtSigned(fBA)}</td>
        <td>${fmtPct(fPBA)}</td>
        <td>${fmtSigned(fCB)}</td>
        <td>${fmtPct(fPCB)}</td>
        <td>${fmtSigned(fCA)}</td>
        <td>${fmtPct(fPCA)}</td>
      </tr>
    </tfoot>`}
  `;
}

// ─── 모드/그룹화 UI 갱신 ───
function updateModeBar(){
  document.querySelectorAll('.mb-btn[data-mode]').forEach(b=>{
    b.classList.toggle('active', b.dataset.mode === MODE);
  });
  document.querySelectorAll('.sync-btn[data-sync]').forEach(b=>{
    b.classList.toggle('active', SYNC[b.dataset.sync] === true);
  });
  const allOn = Object.values(SYNC).every(v=>v);
  document.getElementById('syncAll').classList.toggle('active', allOn);
}
function updateGroupByUI(){
  document.querySelectorAll('.gb-chip').forEach(b=>{
    const gb = b.dataset.gb;
    if (gb === '__auto__') b.classList.toggle('active', AUTO_SPLIT);
    else                   b.classList.toggle('active', GROUP_BY.has(gb));
  });
}

function renderAll(){
  renderPalette();
  renderTierGrid();
  renderKPI();
  renderTable();
  updateModeBar();
  updateGroupByUI();
}

// ─── 클릭 이벤트 ───
document.addEventListener('click', e=>{
  // A/B/C 내 칩 × 제거
  const rmx = e.target.closest('.td-chip-x');
  if (rmx){
    const chip = rmx.closest('.td-chip');
    removeValue(chip.dataset.tier, chip.dataset.dim, parseChipVal(chip.dataset.dim, chip.dataset.val));
    return;
  }
  // 초기화 / 없음
  const tb = e.target.closest('.tier-btn[data-action]');
  if (tb){
    if (tb.dataset.action === 'reset') resetTier(tb.dataset.tier);
    else                                noneTier(tb.dataset.tier);
    return;
  }
  // 단수/복수
  const mb = e.target.closest('.mb-btn[data-mode]');
  if (mb){ MODE = mb.dataset.mode; updateModeBar(); return; }
  // 일괄적용 (개별)
  const sb = e.target.closest('.sync-btn[data-sync]');
  if (sb){ SYNC[sb.dataset.sync] = !SYNC[sb.dataset.sync]; updateModeBar(); return; }
  // 일괄적용 (전체)
  if (e.target.closest('#syncAll')){
    const allOn = Object.values(SYNC).every(v=>v);
    for (const k of Object.keys(SYNC)) SYNC[k] = !allOn;
    updateModeBar();
    return;
  }
  // 그룹화
  const gb = e.target.closest('.gb-chip');
  if (gb){
    const k = gb.dataset.gb;
    if (k === '__auto__') AUTO_SPLIT = !AUTO_SPLIT;
    else if (GROUP_BY.has(k)) GROUP_BY.delete(k);
    else GROUP_BY.add(k);
    updateGroupByUI();
    renderTable();
    return;
  }
});

// ─── 드래그앤드롭 ───
let DRAG_DATA = null;
document.addEventListener('dragstart', e=>{
  const chip = e.target.closest('.pal-chip');
  if (!chip) return;
  DRAG_DATA = { dim: chip.dataset.dim, val: chip.dataset.val };
  chip.classList.add('dragging');
  if (e.dataTransfer){
    e.dataTransfer.effectAllowed = 'copy';
    try { e.dataTransfer.setData('text/plain', JSON.stringify(DRAG_DATA)); } catch(_){}
  }
});
document.addEventListener('dragend', e=>{
  const chip = e.target.closest('.pal-chip');
  if (chip) chip.classList.remove('dragging');
  document.querySelectorAll('.tier.drag-over').forEach(n=>n.classList.remove('drag-over'));
  DRAG_DATA = null;
});
document.addEventListener('dragover', e=>{
  const tier = e.target.closest('.tier');
  if (!tier || !DRAG_DATA) return;
  e.preventDefault();
  if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
  tier.classList.add('drag-over');
});
document.addEventListener('dragleave', e=>{
  const tier = e.target.closest('.tier');
  if (!tier) return;
  if (e.relatedTarget && tier.contains(e.relatedTarget)) return;
  tier.classList.remove('drag-over');
});
document.addEventListener('drop', e=>{
  const tier = e.target.closest('.tier');
  if (!tier || !DRAG_DATA) return;
  e.preventDefault();
  tier.classList.remove('drag-over');
  addValue(tier.dataset.tier, DRAG_DATA.dim, parseChipVal(DRAG_DATA.dim, DRAG_DATA.val));
  DRAG_DATA = null;
});

renderAll();
"""


def make_html(records, base_date: str) -> str:
    data_json = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
    html = HTML_TEMPLATE
    html = html.replace('__CSS__', CSS)
    html = html.replace('__BASE_DATE__', base_date or '미상')
    html = html.replace('__DATA__', data_json)
    html = html.replace('__JSAPP__', JS_APP)
    return html


# ─────────────────────────────────────────────────────────────────────
# 3. MAIN
# ─────────────────────────────────────────────────────────────────────
def open_browser(html_path: str):
    abs_path = os.path.abspath(html_path)
    url = 'file:///' + abs_path.replace('\\', '/')
    print(f"\n🌐 브라우저 오픈 중: {url}")
    try:
        if sys.platform == 'win32':
            os.startfile(abs_path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', abs_path])
        else:
            subprocess.Popen(['xdg-open', abs_path])
        print("✅ 브라우저가 열렸습니다!")
    except Exception:
        try:
            webbrowser.open(url)
            print("✅ 브라우저가 열렸습니다!")
        except Exception:
            print(f"⚠️  자동 오픈 실패. 직접 파일을 여세요: {abs_path}")


def main():
    ap = argparse.ArgumentParser(description='매출 선택비교 HTML 생성기 v1')
    ap.add_argument('--data', default=None, help='Excel 파일 경로 (기본: 자동탐색)')
    ap.add_argument('--out',  default=None, help='출력 HTML 파일 경로')
    ap.add_argument('--no-browser', action='store_true', help='브라우저 자동 오픈 비활성화')
    args = ap.parse_args()

    if not os.path.exists(DASHBOARD_MOD_PATH):
        print(f"❌ '매출 Dashboard_vf.py' 를 찾을 수 없습니다: {DASHBOARD_MOD_PATH}")
        sys.exit(1)
    dash = load_dashboard_module()

    xlsx = dash.find_excel(args.data)
    if not xlsx:
        print("❌ Excel 파일을 찾을 수 없습니다.")
        print(f"   기본 파일명: {dash.EXCEL_FILENAME}")
        sys.exit(1)
    print(f"📂 Excel 읽는 중: {xlsx}")

    records = dash.extract_data(xlsx)
    print(f"✅ 추출 완료: {len(records):,}건")

    base_date = dash.read_base_date(xlsx)
    if base_date:
        print(f"📅 기준일자: {base_date}")

    # 비교 페이지는 fw* 주차 컬럼 불필요 → 사이즈 절감
    slim = []
    for r in records:
        slim.append({
            'team': r['team'], 'channel': r['channel'], 'brand': r['brand'],
            'month': r['month'],
            'y2024': r['y2024'], 'y2025': r['y2025'],
            'plan': r['plan'], 'actual': r['actual'],
        })

    out_path = args.out or os.path.join(SCRIPT_DIR, '매출_선택비교.html')
    html = make_html(slim, base_date)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    kb = os.path.getsize(out_path) / 1024
    print(f"✅ 생성 완료: {out_path}  ({kb:.0f} KB)")

    if not args.no_browser:
        open_browser(out_path)
    else:
        print(f"\n▶  직접 파일을 열어주세요: {os.path.abspath(out_path)}")


if __name__ == '__main__':
    main()
