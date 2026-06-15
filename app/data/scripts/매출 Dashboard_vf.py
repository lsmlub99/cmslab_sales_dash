#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
매출 대시보드 생성기 v3
- 스타일 가이드 적용 (헤더/필터바/탭바/KPI/테이블/차트)
- HTML 생성 후 브라우저 자동 오픈
- Excel 파일 경로 자동 탐색

Usage:
  python generate_dashboard_v3.py
  python generate_dashboard_v3.py --data "경로\파일.xlsx" --out dashboard.html
"""
import os, sys, json, argparse, webbrowser, subprocess
import datetime
import pandas as pd

# ═══════════════════════════════════════════════════════════════════════════════
# 1. DATA EXTRACTION
# ═══════════════════════════════════════════════════════════════════════════════
TARGET_TEAMS = [
    'RBD1팀', 'RBD2팀', '동북아MC팀', 'Global사업팀',
    'GEC팀',  '일본사업팀', '중국사업팀', '메디컬팀'
]
MONTHS_EN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

# 엑셀 col1의 브랜드 섹션 헤더 → 대시보드 브랜드 코드
BRAND_MAP = {
    '셀퓨전씨':       'CFC',
    '셀퓨전씨 엑스퍼트': 'CEX',
    '더마블록':        'DMB',
    '수이스킨':        'SUS',
}
# col1 헤더 중 브랜드 섹션이 아닌 메타/레이아웃 행 (무시)
_BRAND_IGNORE = {'▣ 씨엠에스랩 매출', '(단위 : 백만원)', '영업그룹 Code', 'Code'}

COL_Y2024  = 5   # Jan~Dec cols 5-16,  Total=17
COL_Y2025B = 18  # Jan~Dec cols 18-29, Total=30  (Budget)
COL_Y2025A = 31  # Jan~Dec cols 31-42, Total=43  (Actual)
COL_Y2026P = 44  # Jan~Dec cols 44-55, Total=56  (Plan/Budget)
COL_Y2026A = 57  # Jan~Dec cols 57-68, Total=69  (Actual/실행계획)

# Weekly snapshots per month: {month: [(col, label), ...]}
# col 레이블 기준 (row5): Jan_1=100, Jan_2=105, Jan_3=110, Jan_4=115
#                          Feb_1=120, Feb_2=125, Feb_3=130
#                          Mar_1=136, Mar_2=141, Mar_3=146, Mar_4=151
#                          Apr_1=157, Apr_2=162, Apr_3=167, Apr_4=172, Apr_5=177
#                          May_1=178, May_2=183, May_3=188, May_4=192, May_5=197
#                          Jun_1=198, Jun_2=202
# 주의: 이 하드코딩은 폴백용. 실제 실행 시에는 extract_data() 안의 detect_weekly_cols()가
#       Excel row 5의 'Jan_1' ~ 'Dec_N' 라벨을 자동 스캔해 이 값을 덮어쓴다.
WEEKLY_COLS = {
    1: [(100,'1주차'),(105,'2주차'),(110,'3주차'),(115,'4주차')],
    2: [(120,'1주차'),(125,'2주차'),(130,'3주차')],
    3: [(136,'1주차'),(141,'2주차'),(146,'3주차'),(151,'4주차')],
    4: [(157,'1주차'),(162,'2주차'),(167,'3주차'),(172,'4주차'),(177,'5주차')],
    5: [(178,'1주차'),(183,'2주차'),(188,'3주차'),(192,'4주차'),(197,'5주차')],
    6: [(198,'1주차'),(202,'2주차')],
}

# 런타임에 detect_weekly_cols()가 채워주는 자동감지 결과. make_html에서 WEEKLY_META 주입에 사용.
_DETECTED_WEEKLY_COLS: dict = {}

_MONTH_ABBR_TO_NUM = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
    'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12,
}

def detect_weekly_cols(df, scan_start: int = 100) -> dict:
    """Excel row 5에서 'Jan_1'~'Dec_N' 패턴의 주차 라벨을 스캔.
    반환 형태: {month_int: [(col, '1주차'), (col, '2주차'), ...]}
    - 라벨의 _N 숫자에 빈자리가 있어도 출현한 컬럼들을 1주차부터 순번 매김.
    - row 5 가 없거나 매치가 0개면 빈 dict 반환 → 호출측에서 폴백 사용.
    """
    import re
    pat = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)_(\d+)$')
    by_month: dict = {}
    if df.shape[0] <= 5:
        return by_month
    for c in range(scan_start, df.shape[1]):
        v = df.iloc[5, c]
        if not isinstance(v, str):
            continue
        m = pat.match(v.strip())
        if not m:
            continue
        mn = _MONTH_ABBR_TO_NUM[m.group(1)]
        by_month.setdefault(mn, []).append((c, int(m.group(2))))
    result: dict = {}
    for mn, lst in by_month.items():
        lst.sort(key=lambda t: (t[1], t[0]))  # 라벨 _N → 칼럼 순
        result[mn] = [(c, f'{i}주차') for i, (c, _) in enumerate(lst, 1)]
    return result

def extract_data(xlsx_path: str) -> list:
    import re
    global _DETECTED_WEEKLY_COLS
    df = pd.read_excel(xlsx_path, sheet_name='영업그룹별 매출액', header=None)

    # ── 주차 칼럼 자동감지 (실패 시 하드코딩 폴백) ──
    detected = detect_weekly_cols(df)
    if detected:
        _DETECTED_WEEKLY_COLS = detected
        wkcols = detected
    else:
        _DETECTED_WEEKLY_COLS = {}
        wkcols = WEEKLY_COLS

    # ── col1을 순회하며 행 → 브랜드 맵 구축 ──
    row_brand = {}
    current_brand = '기타'
    for idx in df.index:
        v = df.iloc[idx, 1]
        if pd.notna(v) and isinstance(v, str):
            s = v.strip()
            if re.match(r'^A\d+$', s):
                pass  # 데이터 행 (code) → 현재 brand 유지
            elif s in BRAND_MAP:
                current_brand = BRAND_MAP[s]
            elif s in _BRAND_IGNORE:
                pass  # 메타/헤더 행 (brand 변경 없음)
            else:
                current_brand = '기타'  # 디바이스/키즈토즈 등
        row_brand[idx] = current_brand

    mask = (
        df.iloc[:, 1].astype(str).str.match(r'^A\d+$') &
        df.iloc[:, 3].isin(TARGET_TEAMS)
    )
    clean = df[mask].copy()

    records = []
    for idx, row in clean.iterrows():
        team    = str(row.iloc[3])
        channel = str(row.iloc[2])
        code    = str(row.iloc[1])
        gubun   = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ''
        brand   = row_brand.get(idx, '기타')

        # RBD1팀 자사몰-국내 → 자사몰-국내(SUS)
        if code == 'A109' and team == 'RBD1팀':
            channel = '자사몰-국내(SUS)'

        # 중국법인매출 → 구분Code로 CFC/CEX 분리
        if channel == '중국법인매출' and team == '중국사업팀':
            channel = '상해신이-CFC' if gubun == 'Z05' else '상해신이-CEX'

        def val(col):
            v = row.iloc[col]
            return round(float(v) / 1_000_000, 3) if pd.notna(v) else 0.0

        # 현재 진행 월 = wkcols 중 가장 최신(가장 큰 col 번호)을 가진 월.
        # 이 월에 한해 연간 BJ-equiv 칼럼은 보통 미동기화 상태이므로
        # 최신 weekly snapshot 값으로 actual 을 덮어쓴다.
        # (Apr 등 이미 확정된 월은 BJ-equiv == latest snapshot 이라 차이 없음.
        #  Jan-Mar 등 과거 월의 snapshot 은 의미가 다를 수 있어 건드리지 않음.)
        _current_mn = max(wkcols, key=lambda m: max(c for c,_ in wkcols[m])) if wkcols else None

        for mi in range(12):
            mn = mi + 1
            weekly = {}
            if mn in wkcols:
                for wk_idx, (wk_col, wk_lbl) in enumerate(wkcols[mn], 1):
                    v = row.iloc[wk_col]
                    wv = round(float(v) / 1_000_000, 3) if pd.notna(v) and float(v) != 0 else None
                    weekly[f'fw{wk_idx}'] = wv

            actual_val = val(COL_Y2026A + mi)
            # 현재 진행 월은 무조건 최신 weekly snapshot 칼럼을 actual 로 사용한다.
            # 칸이 비어있는 행은 0 으로 간주 (BJ 잔존값으로 인한 합계 오류 방지)
            if mn == _current_mn and mn in wkcols:
                latest_col = max(c for c, _ in wkcols[mn])
                actual_val = val(latest_col)

            rec = {
                'team':    team,
                'channel': channel,
                'brand':   brand,
                'code':    code,
                'month':   mn,
                'y2024':   val(COL_Y2024  + mi),
                'y2025b':  val(COL_Y2025B + mi),
                'y2025':   val(COL_Y2025A + mi),
                'plan':    val(COL_Y2026P + mi),
                'actual':  actual_val,
            }
            rec.update(weekly)
            records.append(rec)
    return records


# ═══════════════════════════════════════════════════════════════════════════════
# 2. CHART.JS – embed from local file (offline support)
# ═══════════════════════════════════════════════════════════════════════════════
def load_chartjs() -> str:
    """chart.umd.js를 스크립트 폴더 또는 임시 경로에서 탐색 후 반환."""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(script_dir, 'chart.umd.js'),
        os.path.join(script_dir, 'chart.js'),
        '/tmp/chartjs_embed.js',
        '/tmp/package/dist/chart.umd.js',
    ]
    for p in candidates:
        if os.path.exists(p):
            with open(p, encoding='utf-8') as f:
                return f.read()
    return None  # CDN fallback


# ═══════════════════════════════════════════════════════════════════════════════
# 3. HTML GENERATION
# ═══════════════════════════════════════════════════════════════════════════════
CSS = """
/* ─── RESET ─── */
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --blue:#1a56a0;
  --blue-dark:#154080;
  --bg:#f5f6fa;
  --card:#ffffff;
  --border:#e5e7eb;
  --text:#1a1a2e;
  --muted:#6c757d;
  --label:#374151;
  --pos-dark:#375623;
  --neg-dark:#C55A11;
  --pos-light:#e8f5e2;
  --neg-light:#fdf0e8;
  --col-actual:#808080;
  --col-plan:#375623;
  --col-latest:#1F4E78;
  --col-cmp1:#5C2508;
  --col-cmp2:#833C0C;
  --shadow:0 1px 4px rgba(0,0,0,.05);
  --radius:10px;
  border-radius:10px;
  --ff:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;
}
body{font-family:var(--ff);background:var(--bg);color:var(--text);font-size:13px;line-height:1.5;min-height:100vh;overflow-x:hidden}

/* ─── STICKY STACK ─── */
.sticky-header   {position:sticky;top:0;z-index:200}
.sticky-filterbar{position:sticky;top:44px;z-index:190}
.sticky-kpi      {position:sticky;top:114px;z-index:170}

/* ─── HEADER ─── */
.header{
  background:#fff;height:44px;
  display:flex;align-items:center;padding:0 clamp(12px,2vw,24px);gap:14px;
  box-shadow:0 2px 8px rgba(0,0,0,.10);border-bottom:1px solid var(--border);
}
.hdr-company{color:#833C0C;font-size:clamp(14px,1vw,15px);font-weight:800;white-space:nowrap}
.hdr-divider{width:1px;height:20px;background:#e5e7eb}
.hdr-badge{
  background:var(--blue);color:#fff;
  font-size:clamp(12px,0.85vw,12.5px);font-weight:700;padding:4px 14px;border-radius:20px;
  letter-spacing:.2px;white-space:nowrap;
}
.hdr-right{margin-left:auto;color:var(--muted);font-size:clamp(10px,0.75vw,11px);text-align:right;white-space:nowrap}
.hdr-right strong{color:var(--text);font-size:clamp(11px,0.8vw,12px);display:inline}

/* ─── FILTER BAR ─── */
.filterbar{
  background:var(--blue);padding:6px clamp(12px,2vw,24px) 7px;
  display:flex;flex-direction:column;gap:0;
  border-top:1px solid rgba(255,255,255,.12);
  box-shadow:0 3px 8px rgba(0,0,0,.15);
}
.filter-row{display:flex;align-items:center;gap:24px;flex-wrap:nowrap;padding:3px 0;min-width:0;overflow:hidden}
.fg{display:flex;align-items:center;gap:4px;flex-wrap:nowrap;flex-shrink:0}
.fg-label{
  display:inline-flex;align-items:center;justify-content:center;
  background:#DDEBF7;color:#1a56a0;
  font-size:clamp(10px,0.8vw,11px);font-weight:800;white-space:nowrap;
  padding:2px 8px;border-radius:5px;letter-spacing:.3px;
  margin-right:4px;box-shadow:0 1px 3px rgba(0,0,0,.18);
  border:1px solid #9cc2e8;
}
.fg-sep{color:rgba(255,255,255,.30);padding:0 4px;font-size:13px;font-weight:300}
.bu-sep{padding:0 1px;margin:0 -3px}
.team-fg{flex-shrink:0}
.filter-badge{
  margin-left:auto;flex-shrink:1;flex-basis:auto;min-width:60px;
  background:#fef08a;border-radius:6px;
  border:1.5px solid #fbbf24;
  padding:3px 14px;font-size:11.5px;font-weight:800;
  color:#dc2626;letter-spacing:.3px;
  box-shadow:0 2px 6px rgba(0,0,0,.22);
  max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;
}

/* pill button — font-weight:700 고정으로 active/inactive 전환 시 너비 불변 */
.pb{
  padding:2px 10px;border-radius:20px;cursor:pointer;font-size:clamp(10px,0.8vw,11px);
  font-family:var(--ff);white-space:nowrap;transition:background .14s,color .14s,border-color .14s;line-height:1.7;
  background:rgba(255,255,255,0.12);
  border:1px solid rgba(255,255,255,0.3);
  color:rgba(255,255,255,0.75);
  font-weight:700;
}
.pb:hover{background:rgba(255,255,255,.22);color:#fff}
.pb.active{background:#fff;border-color:#fff;color:var(--blue)}
/* 단일 선택 모드 라벨(팀/월/분기 클릭 시) 주황색 — 사용자에게 모드 변경 알림 */
.fg-label.single-mode{background:#f59e0b;border-color:#d97706;color:#fff}
/* 토글 가능한 라벨에 마우스 hover 시 cursor 표시 */
.fg-label.toggleable{cursor:pointer;user-select:none;transition:background .15s,color .15s,border-color .15s}
.fg-label.toggleable:hover{filter:brightness(0.95)}

/* checkbox */
.fck{display:flex;align-items:center;gap:5px;cursor:pointer;user-select:none}
.fck input{accent-color:#f59e0b;width:13px;height:13px}
.fck span{color:rgba(255,255,255,.9);font-size:clamp(10px,0.8vw,11.5px);font-weight:600}

/* ─── KPI STRIP ─── */
.kpi-strip{
  background:var(--bg);padding:8px 24px;
  display:grid;
  grid-template-columns: repeat(4, 1.05fr) 1.9fr 1.5fr;
  gap:8px;
  border-bottom:1px solid var(--border);
}
/* KPI 카드 공통: 2줄 높이 고정 */
.kpi-card{
  background:#F2F8EE;border:1px solid var(--border);border-radius:var(--radius);
  box-shadow:var(--shadow);padding:clamp(7px,0.6vw,10px) clamp(10px,0.9vw,14px);
  display:flex;flex-direction:column;justify-content:center;gap:4px;min-width:0;
}
/* 1행: 지표명(좌) + 금액(우) */
.kpi-row{display:flex;align-items:baseline;justify-content:space-between;gap:4px;white-space:nowrap;overflow:hidden}
.kpi-tag{font-size:clamp(11px,0.85vw,13px);font-weight:700;color:var(--muted);flex-shrink:0}
.kpi-amount{font-size:clamp(16px,1.4vw,22px);font-weight:800;color:var(--text);line-height:1}
.kpi-unit{font-size:clamp(10px,0.7vw,12px);font-weight:500;color:var(--muted);margin-left:2px}
/* 2행: 증감액·증감율 */
.kpi-row2{display:flex;align-items:center;justify-content:space-between;gap:4px;white-space:nowrap;overflow:hidden}
.kpi-vs-lbl{font-size:clamp(10px,0.75vw,11.5px);font-weight:700;color:var(--muted);flex-shrink:0}
.kpi-vs-val{font-size:clamp(10px,0.8vw,12.5px);font-weight:800;text-align:right}
.kpi-vs-val.pos{color:var(--pos-dark)}
.kpi-vs-val.neg{color:var(--neg-dark)}
.kpi-vs-val.neu{color:var(--muted)}
/* KPI 5번째 카드 증감값 강조 (폰트만) */
.kpi-vs-emph{font-weight:900;letter-spacing:.2px}
.kpi-vs-emph.pos{color:#166534}
.kpi-vs-emph.neg{color:#b91c1c}
.kpi-vs-emph.neu{color:var(--muted)}

/* ─── 달성율 바 카드 (가로 그라데이션 바) ─── */
.kpi-card-bar{
  background:#F2F8EE;border:1px solid var(--border);border-radius:var(--radius);
  box-shadow:var(--shadow);padding:clamp(7px,0.6vw,10px) clamp(10px,0.9vw,14px);
  display:flex;flex-direction:column;justify-content:center;gap:5px;min-width:0;
}
.kpi-bar-head{display:flex;align-items:baseline;justify-content:space-between;gap:6px}
.kpi-bar-lbl{font-size:clamp(10px,0.75vw,11px);font-weight:700;color:var(--muted)}
.kpi-bar-pct{font-size:clamp(15px,1.2vw,18px);font-weight:800;line-height:1}
.kpi-bar-track{
  width:100%;height:clamp(7px,0.6vw,10px);background:#e5e7eb;border-radius:6px;overflow:hidden;
  position:relative;
}
.kpi-bar-fill{
  height:100%;border-radius:6px;
  transition:width .6s ease;
  position:relative;
}
.kpi-bar-sub{font-size:clamp(10px,0.7vw,11px);color:var(--muted);text-align:right;white-space:nowrap}

/* ─── 채널 드롭다운 ─── */
.ch-select{
  height:24px;padding:0 24px 0 10px;border-radius:20px;
  border:1px solid rgba(255,255,255,0.3);
  background:rgba(255,255,255,0.12);
  color:rgba(255,255,255,0.85);font-size:11px;
  font-family:var(--ff);font-weight:600;
  cursor:pointer;outline:none;
  transition:all .14s;min-width:130px;max-width:220px;
  appearance:none;-webkit-appearance:none;
  background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='rgba(255,255,255,0.7)'/%3E%3C/svg%3E");
  background-repeat:no-repeat;background-position:right 8px center;
}
.ch-select:focus,.ch-select:hover{background-color:rgba(255,255,255,.22);color:#fff}
.ch-select option{background:#1a56a0;color:#fff;font-weight:600}

/* ─── MAIN CONTENT ─── */
.main{padding:clamp(8px,1vw,12px) clamp(12px,2vw,24px);display:flex;flex-direction:column;gap:10px;padding-bottom:40px}

/* ─── CHART CARD ─── */
.chart-card{
  background:var(--card);border:1px solid var(--border);
  border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;
}
.chart-card-hdr{padding:8px clamp(10px,1.5vw,16px);border-bottom:1px solid var(--border);display:flex;align-items:center;gap:10px;overflow:hidden}
.chart-card-title{font-size:clamp(11px,1vw,13px);font-weight:700;color:var(--text);flex:1;min-width:0}
.chart-card-desc{font-size:clamp(9px,0.8vw,11px);color:#9ca3af;margin-top:2px}
.chart-card-body{padding:clamp(8px,1.2vw,14px) clamp(10px,1.5vw,16px) 10px;overflow:hidden}
.charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
.charts-full{display:flex;flex-direction:column;gap:14px}

/* ─── TABLE CARD ─── */
.tbl-card{
  background:var(--card);border:1px solid var(--border);
  border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;
}
.tbl-toolbar{display:flex;align-items:center;padding:10px 14px;border-bottom:1px solid var(--border);gap:10px}
.tbl-title{font-size:clamp(11px,1vw,13px);font-weight:700;color:var(--text);flex:1}
.tbl-count{font-size:11px;color:var(--muted);white-space:nowrap}
.tbl-search{
  border:1px solid var(--border);border-radius:6px;padding:4px 10px;
  font-size:11px;font-family:var(--ff);outline:none;width:180px;
  transition:border-color .14s;
}
.tbl-search:focus{border-color:var(--blue)}
.tbl-outer{overflow-x:auto;max-height:540px}
.tbl-outer::-webkit-scrollbar{width:5px;height:5px}
.tbl-outer::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px}

/* ─── 반응형 미디어 쿼리 ─── */
@media (max-width:1280px) {
  .kpi-strip{grid-template-columns:repeat(4,1.05fr) 1.9fr 1.5fr;padding:6px 16px}
  .sticky-filterbar{top:44px}
  .sticky-kpi{top:114px}
  .main{padding:10px 16px}
  .filterbar{padding:6px 16px 7px}
  .header{padding:0 16px}
}
@media (max-width:1024px) {
  .kpi-strip{grid-template-columns:repeat(3,1fr) repeat(3,1fr);padding:6px 12px}
  .charts-grid{grid-template-columns:1fr}
  .main{padding:10px 12px}
}

/* TABLE ITSELF */
table{border-collapse:collapse;white-space:nowrap;font-size:14px;width:100%}
thead{position:sticky;top:0;z-index:10}
/* group header row */
thead tr.gh th{
  padding:8px 10px;text-align:center;border:1px solid #3d5a80;
  font-size:13px;font-weight:700;letter-spacing:.2px;
}
/* sub header row */
thead tr.sh th{
  padding:6px 8px;text-align:center;border:1px solid #3d5a80;
  font-size:12px;font-weight:600;
}
/* 숫자 컬럼 동일 너비 80px */
thead tr.gh th:not(.fix):not(.fix2) { min-width:80px; }
thead tr.sh th { min-width:80px; }

/* color bands for header groups */
.hg-fix   {background:#374151;color:#f9fafb}
.hg-act24 {background:#5a5a5a;color:#fff}
.hg-act25 {background:#6b6b6b;color:#fff}
.hg-plan  {background:#375623;color:#fff}
.hg-act26 {background:#1F4E78;color:#fff}
.hg-cmp25 {background:#6b3a2a;color:#fff}
.hg-cmpP  {background:#5C2508;color:#fff}
.hg-fc    {background:#0369a1;color:#fff}

tbody tr{transition:background .08s}
tbody tr:nth-child(even){background:#fafafa}
tbody tr:hover{background:#f0f4ff}
tbody td{
  padding:6px 10px;border-bottom:1px solid #f3f4f6;
  border-right:1px solid #f3f4f6;text-align:right;font-weight:400;
  min-width:80px;
}
td.fix{position:sticky;left:0;background:#fff;z-index:5;text-align:left;border-right:2px solid #e5e7eb;min-width:80px}
td.fix2{position:sticky;left:80px;background:#fff;z-index:5;text-align:left;border-right:2px solid #e5e7eb;min-width:110px}
tr:nth-child(even) td.fix,tr:nth-child(even) td.fix2{background:#fafafa}
tr:hover td.fix,tr:hover td.fix2{background:#f0f4ff}
td.team{font-weight:600;color:#1e3a5f;font-size:13px}
td.ch{color:#374151;font-size:13px}
/* value colors */
td.v-actual{color:var(--col-actual)}
td.v-plan  {color:var(--col-plan)}
td.v-act26 {color:var(--col-latest)}
td.v-cmp   {color:var(--col-cmp1)}
td.pos     {color:#1a1a2e}
td.neg     {color:#dc2626}
td.neu     {color:var(--muted)}

/* ─── 인쇄 탭 화면 레이아웃 ─── */
#pane-print {
  display:none;
  position:fixed;
  left:0; right:0; bottom:0; top:114px;
  background:#e5e7eb;
  z-index:15;
  flex-direction:column;
}
#pane-print.active { display:flex; }

.print-toolbar {
  display:flex; align-items:center; gap:14px;
  padding:8px 20px;
  background:#fff; border-bottom:1px solid var(--border);
  flex-shrink:0;
}
.print-toolbar-title { font-size:13px; font-weight:700; color:var(--text); }
.print-toolbar-sub   { font-size:12px; color:var(--muted); }
.print-btn {
  margin-left:auto;
  padding:6px 20px; border-radius:6px;
  background:var(--blue); color:#fff;
  font-size:12px; font-weight:700;
  border:none; cursor:pointer; font-family:var(--ff);
}
.print-btn:hover { background:var(--blue-dark); }

.print-preview-wrap {
  flex:1; overflow:auto; padding:20px;
  display:flex; justify-content:center;
}

/* A4 가로: 297mm × 210mm — 미리보기도 실제 인쇄와 동일 크기 */
.print-paper {
  width:277mm; height:190mm;
  background:#fff;
  box-shadow:0 4px 20px rgba(0,0,0,.2);
  border-radius:4px;
  padding:8mm 8mm 6mm;
  box-sizing:border-box;
  display:flex; flex-direction:column; gap:4px;
  font-family:'Malgun Gothic', sans-serif;
  overflow:hidden;   /* 미리보기에서 넘치면 숨김 */
}

/* 인쇄 시 두 번째 용지: 미리보기에서 두 번째 paper */
.print-paper-2 {
  width:277mm; height:190mm;
  background:#fff;
  box-shadow:0 4px 20px rgba(0,0,0,.2);
  border-radius:4px;
  padding:8mm 8mm 6mm;
  box-sizing:border-box;
  display:flex; flex-direction:column; gap:4px;
  font-family:'Malgun Gothic', sans-serif;
  overflow:hidden;
  margin-top:16px;
}
.print-paper-header {
  display:flex; justify-content:space-between; align-items:baseline;
  border-bottom:2px solid #1a56a0; padding-bottom:4px; margin-bottom:4px;
}
.print-paper-title { font-size:15px; font-weight:800; color:#1a1a2e; }
.print-paper-sub   { font-size:10px; color:#6b7280; }
.print-paper-footer {
  display:flex; justify-content:space-between;
  font-size:9px; color:#9ca3af;
  border-top:1px solid #e5e7eb; padding-top:4px; margin-top:auto;
}

.print-table-wrap { flex:1; overflow:hidden; }

/* 인쇄 표 공통 스타일 (1페이지, 2페이지 동일) */
#printTable, #printTable2 {
  width:100%; border-collapse:collapse;
  font-size:8px; font-family:'Malgun Gothic', sans-serif;
  table-layout:fixed;
}
#printTable thead th, #printTable2 thead th {
  background:#1F4E78; color:#fff;
  padding:3px 3px; text-align:center;
  border:1px solid #ccc; font-weight:700;
  font-size:7.5px; white-space:nowrap;
  overflow:hidden;
}
#printTable thead th.pt-hg-act24, #printTable2 thead th.pt-hg-act24 { background:#5a5a5a; }
#printTable thead th.pt-hg-act25, #printTable2 thead th.pt-hg-act25 { background:#6b6b6b; }
#printTable thead th.pt-hg-plan,  #printTable2 thead th.pt-hg-plan  { background:#375623; }
#printTable thead th.pt-hg-act26, #printTable2 thead th.pt-hg-act26 { background:#1F4E78; }
#printTable thead th.pt-hg-cmp25, #printTable2 thead th.pt-hg-cmp25 { background:#6b3a2a; }
#printTable thead th.pt-hg-cmpP,  #printTable2 thead th.pt-hg-cmpP  { background:#5C2508; }
#printTable thead th.pt-hg-fc,    #printTable2 thead th.pt-hg-fc    { background:#0369a1; }
#printTable thead th.pt-hg-fix,   #printTable2 thead th.pt-hg-fix   { background:#374151; }

/* 컬럼 너비 고정 — 두 표 동일하게 */
#printTable colgroup col.pc-fix,  #printTable2 colgroup col.pc-fix  { width:52px; }
#printTable colgroup col.pc-ch,   #printTable2 colgroup col.pc-ch   { width:72px; }
#printTable colgroup col.pc-num,  #printTable2 colgroup col.pc-num  { width:46px; }
#printTable colgroup col.pc-fw,   #printTable2 colgroup col.pc-fw   { width:44px; }

#printTable tbody td, #printTable2 tbody td {
  padding:2px 3px; border:1px solid #e5e7eb;
  text-align:right; font-size:8px; white-space:nowrap;
  overflow:hidden;
}
#printTable tbody td.pt-team, #printTable2 tbody td.pt-team {
  text-align:center; font-weight:700; color:#1e3a5f;
  background:#f8fafc;
}
#printTable tbody td.pt-ch, #printTable2 tbody td.pt-ch {
  text-align:left; color:#374151; padding-left:6px;
}
#printTable tbody tr.pt-sub td, #printTable2 tbody tr.pt-sub td {
  background:#dbeafe; font-weight:700; color:#1e3a5f;
}
#printTable tbody tr.pt-sub td.pt-neg, #printTable2 tbody tr.pt-sub td.pt-neg { color:#dc2626 !important; }
#printTable tbody tr.pt-sub td.pt-fc,
#printTable2 tbody tr.pt-sub td.pt-fc { background:#dbeafe!important; color:#1e3a5f!important; }

#printTable tbody tr.pt-grand td, #printTable2 tbody tr.pt-grand td {
  background:#1e3a5f; color:#e2e8f0; font-weight:700;
}
#printTable tbody tr.pt-grand td.pt-neg, #printTable2 tbody tr.pt-grand td.pt-neg { color:#fca5a5 !important; }
#printTable tbody tr.pt-grand td.pt-fc,
#printTable2 tbody tr.pt-grand td.pt-fc { background:#1e3a5f!important; color:#fff!important; }
/* 인쇄용 BU/grand fc 셀 배경 통일 */
#printTable tbody tr td.pt-fc,
#printTable2 tbody tr td.pt-fc { background:rgba(224,242,254,.4); color:#0c4a6e; }
#printTable tbody tr.pt-sub td.pt-fc,
#printTable2 tbody tr.pt-sub td.pt-fc { background:#dbeafe!important; color:#1e3a5f!important; }
#printTable tbody tr.pt-grand td.pt-fc,
#printTable2 tbody tr.pt-grand td.pt-fc { background:#1e3a5f!important; color:#fff!important; }

/* ─── @media print ─── */
@media print {
  @page { size:A4 landscape; margin:8mm; }

  /* 화면 UI 모두 숨김 */
  .sticky-header, .sticky-filterbar, .sticky-kpi,
  #pane-dashboard, #pane-table, #pane-print .print-toolbar,
  .print-preview-wrap { display:none !important; }

  /* 인쇄 탭이 활성일 때만 paper 출력 */
  #pane-print.active .print-preview-wrap {
    display:flex !important; flex-direction:column; gap:0; padding:0;
  }
  #pane-print.active { display:block !important; position:static !important;
    background:#fff; top:0; }

  .print-paper, .print-paper-2 {
    width:100%; height:auto; min-height:0;
    box-shadow:none; border-radius:0;
    padding:0; margin:0;
    overflow:visible;
    page-break-after:always;
    break-after:page;
  }
  .print-paper:last-child, .print-paper-2:last-child {
    page-break-after:avoid; break-after:avoid;
  }
  #printTable, #printTable2 { font-size:7.5px; }
  #printTable thead th, #printTable2 thead th { font-size:7px; padding:2px 3px; }
  #printTable tbody td, #printTable2 tbody td { font-size:7.5px; padding:2px 3px; }

  /* 배경색 강제 출력 */
  * { -webkit-print-color-adjust:exact !important;
      print-color-adjust:exact !important;
      color-adjust:exact !important; }
}
/* ─── 열 그룹 구분선 ─── */
.hg-sep-r, th.hg-sep-r, td.hg-sep-r { border-right:1px solid #94a3b8 !important; }
.pt-sep-r, th.pt-sep-r, td.pt-sep-r { border-right:1px solid #94a3b8 !important; }
/* 채널 열 오른쪽 구분선 */
th.hg-fix2-sep, td.hg-fix2-sep { border-right:1px solid #94a3b8 !important; }
/* 당월예상 첫 번째 열 왼쪽 구분선 */
.fc-cell-first, th.fc-cell-first { border-left:1px solid #94a3b8 !important; box-shadow:inset 1px 0 0 #94a3b8; }
.pt-fc-first, th.pt-fc-first { border-left:1px solid #94a3b8 !important; box-shadow:inset 1px 0 0 #94a3b8; }

tr.sub-row td { font-weight:700 !important; background:#dbeafe!important; color:#1e3a5f!important }
tr.sub-row td.fix, tr.sub-row td.fix2 { background:#dbeafe!important; }
tr.sub-row td.pos { color:#1a1a2e!important; font-weight:700!important; }
tr.sub-row td.neg { color:#dc2626!important; font-weight:700!important; }

/* BU 합계: 중간 파란 */
tr.bu-row td { background:#2e6da4!important; color:#fff!important; font-weight:700; }
tr.bu-row td.pos { color:#fff!important; }
tr.bu-row td.neg { color:#fca5a5!important; }
tr.bu-row td.neu { color:#bfdbfe!important; }

/* grand-row: 가장 진한 네이비 */
tr.grand-row td { background:#1e3a5f!important; color:#e2e8f0!important; font-weight:700; }
tr.grand-row td.pos { color:#e2e8f0!important; }
tr.grand-row td.neg { color:#fca5a5!important; }
tr.grand-row td.fc-cell,
tr.grand-row td.fc-cell.fc-cell-first { background:#1e3a5f!important; color:#fff!important; }

/* bu-row fc-cell */
tr.bu-row td.fc-cell,
tr.bu-row td.fc-cell.fc-cell-first { background:#2e6da4!important; color:#fff!important; }

/* 당월예상 셀 */
th.hg-fc { background:#0369a1!important; color:#fff!important; }
.fc-cell  { background:rgba(224,242,254,.35); color:#0c4a6e; }
tr.sub-row td.fc-cell,
tr.sub-row td.fc-cell.fc-cell-first { background:#dbeafe!important; color:#1e3a5f!important; }

/* 토글 행 숨김 */
tr.ch-row.collapsed { display:none; }

.empty-row{padding:40px;text-align:center;color:var(--muted)}

/* ─── 차트 주석 스트립 ─── */
.annot-strip{border-bottom:1px solid var(--border);overflow-x:auto;background:#fafafa}
.annot-strip-tbl{width:100%;border-collapse:collapse;font-size:9.5px;font-family:var(--ff)}
.annot-strip-tbl th{text-align:right;color:#9ca3af;font-weight:600;padding:2px 6px 2px 10px;white-space:nowrap;width:52px;font-size:9px}
.annot-strip-tbl td{text-align:center;font-weight:700;padding:2px 1px;white-space:nowrap;width:calc((100% - 52px)/12)}

/* GAUGE inline small */
.gauge-inline{display:flex;flex-direction:column;align-items:center;flex-shrink:0}
.gauge-inline svg{width:64px;height:38px}
.gauge-inline .g-pct{font-size:11px;font-weight:800;text-align:center;margin-top:1px;transition:color .3s}

/* Scrollbar for outer */
::-webkit-scrollbar{width:6px;height:6px}
::-webkit-scrollbar-track{background:#f1f5f9}
::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px}

/* ─── 차트 헤더 범례 ─── */
.chart-legend {
  display:flex; flex-wrap:nowrap; gap:4px 14px;
  margin-left:auto; flex-shrink:0; align-self:center;
}
.chart-legend-item {
  display:flex; align-items:center; gap:4px;
  font-size:10.5px; color:var(--muted); white-space:nowrap;
}
.chart-legend-dot {
  width:10px; height:10px; border-radius:2px; flex-shrink:0;
}
/* ─── 헤더 탭 ─── */
.hdr-tabs { display:flex; align-items:stretch; margin-left:12px; gap:0; }
.hdr-tab {
  padding:0 16px; height:44px; display:flex; align-items:center;
  font-size:clamp(11px,0.9vw,12.5px); font-weight:600;
  color:var(--muted); cursor:pointer; border:none; background:none;
  border-bottom:3px solid transparent; transition:all .15s;
  font-family:var(--ff); white-space:nowrap;
}
.hdr-tab:hover { color:var(--blue); }
.hdr-tab.active { color:var(--blue); border-bottom:3px solid var(--blue); }

/* ─── 탭 콘텐츠 ─── */
.tab-pane { display:none; }
.tab-pane.active { display:block; }

/* ─── 표 탭 — position fixed로 필터바 아래 전체 채움 ─── */
#pane-table {
  display:none;
  position:fixed;
  left:24px; right:24px; bottom:0;
  top:114px;  /* JS에서 정확히 덮어씀 */
  padding-bottom:8px;
  box-sizing:border-box;
  background:var(--bg);
  z-index:15;
  flex-direction:column;
  gap:4px;
}
#pane-table.active {
  display:flex;
}
#pane-table .tbl-toggle-bar {
  background:#f8fafc;
  border:1px solid var(--border);
  border-radius:var(--radius) var(--radius) 0 0;
  padding:5px 12px;
  display:flex; align-items:center; gap:10px;
  flex-shrink:0;
}
#pane-table .tbl-card {
  flex:1; display:flex; flex-direction:column;
  border-radius:0 0 var(--radius) var(--radius);
  border:1px solid var(--border); border-top:none;
  overflow:hidden; min-height:0;
  background:#fff;  /* 테이블 아래 빈 영역도 흰색 */
}
#pane-table .tbl-outer {
  flex:1; overflow:auto; min-height:0;
  background:#fff;
  max-height:none !important;  /* 일반 .tbl-outer max-height:540px 덮어씀 */
}
#pane-table .tbl-outer > table {
  width:100%;
  border-bottom:1px solid #f3f4f6;
}
.chart-expand-btn {
  display:flex; align-items:center; justify-content:center;
  width:24px; height:24px; border-radius:5px; border:1px solid var(--border);
  background:#fff; cursor:pointer; color:var(--muted); font-size:13px;
  margin-left:8px; flex-shrink:0; transition:all .14s;
  align-self:center;
}
.chart-expand-btn:hover { background:#f0f4ff; color:var(--blue); border-color:var(--blue); }

/* ─── 확대 오버레이 ─── */
.chart-expanded {
  position:fixed !important;
  z-index:160;
  background:#fff;
  border-radius:var(--radius);
  box-shadow:0 8px 40px rgba(0,0,0,.22);
  display:flex; flex-direction:column;
  transition:top .25s ease, left .25s ease, width .25s ease, height .25s ease;
}
.chart-expanded .chart-card-body { flex:1; height:auto !important; }
.chart-expanded .chart-card-body > div { height:100% !important; }
.chart-overlay-bg {
  position:fixed; inset:0; background:rgba(0,0,0,.25);
  z-index:150; cursor:pointer;
  animation:fadeIn .2s ease;
}
@keyframes fadeIn { from{opacity:0} to{opacity:1} }

.stbl td, .stbl th {
  padding:5px 10px; text-align:right; font-size:10.5px;
  border-right:1px solid var(--border);
}
.stbl th { text-align:left; font-weight:700; color:var(--muted); background:#fafafa; border-bottom:1px solid var(--border); }
.stbl td:first-child, .stbl th:first-child { text-align:left; position:sticky; left:0; background:#fafafa; z-index:1; min-width:72px; }
.stbl tr:last-child td { border-bottom:none; }
.stbl tr td { border-bottom:1px solid #f3f4f6; }
.stbl .c24  { color:#6b7280; }
.stbl .c25  { color:#2e6da4; }
.stbl .cp   { color:#375623; }
.stbl .c26  { color:#5C2508; font-weight:700; }

/* ─── 요약표 토글 ─── */
.stbl-toggle {
  display:flex; align-items:center; justify-content:center;
  gap:6px; width:100%; padding:5px 0;
  font-size:11px; font-weight:600; color:var(--muted);
  background:#fafafa; border:none; border-top:1px solid var(--border);
  cursor:pointer; transition:background .14s, color .14s;
  font-family:var(--ff);
}
.stbl-toggle:hover { background:#f0f4ff; color:var(--blue); }
.stbl-toggle .arrow { font-size:10px; transition:transform .3s ease; display:inline-block; }
.stbl-toggle.open .arrow { transform:rotate(180deg); }
.stbl-wrap {
  max-height:0; overflow:hidden;
  transition:max-height .35s ease, opacity .25s ease;
  opacity:0;
}
.stbl-wrap.open { max-height:400px; opacity:1; }
"""

JS_HELPERS = """
'use strict';

/* ─── DATA ─── */
const RAW = __DATA_JSON__;

/* ─── 주차 메타데이터 (Python extract_data → detect_weekly_cols 가 자동 주입) ─── */
const WEEKLY_META = __WEEKLY_META_JSON__;
const TORDER = ['RBD1팀','RBD2팀','일본사업팀','중국사업팀','동북아MC팀','Global사업팀','GEC팀','메디컬팀'];
const MLABEL = ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월'];

// 엑셀 파일 원본 순서 (행 순서 기준)
const CH_ORDER = [
  // RBD1팀
  '올리브영','군납','쿠팡','네이버','버티컬몰','제휴몰','임직원몰','자사몰-국내','기타',
  // RBD2팀: 면세, 다이소, 계열사, 코스트코(한국)
  '면세','다이소','계열사','코스트코(한국)',
  // 일본사업팀
  '큐텐/라쿠텐','아마존_일본','해외_오프라인(일본)',
  // 중국사업팀: 파트너사, 상해신이-CFC, 상해신이-CEX, 티몰(중국), 해외_오프라인_대만
  '파트너사','상해신이-CFC','상해신이-CEX','티몰(중국)','해외_오프라인_대만',
  // 동북아MC팀 (자사몰-국내 공유)
  // Global사업팀
  '해외_오프라인_미국','해외_오프라인_CIS','해외_오프라인_유럽',
  '해외_오프라인_동남아','해외_오프라인_중동','해외_오프라인_기타',
  // GEC팀
  '자사몰-해외','아마존_미국','아마존_기타','틱톡','역직구몰','해외_온라인 기타',
  // 메디컬팀
  '종합병원','클리닉','엑스퍼트몰','대리점',
];

function chSort(arr) {
  return [...arr].sort((a,b)=>{
    const ia = CH_ORDER.indexOf(a), ib = CH_ORDER.indexOf(b);
    if(ia===-1&&ib===-1) return a.localeCompare(b,'ko');
    if(ia===-1) return 1; if(ib===-1) return -1;
    return ia - ib;
  });
}

/* ─── STATE ─── */
const S = {
  teams:    new Set(['ALL']),
  years:    new Set(['2025','plan','2026']),
  months:   new Set(__DEFAULT_MONTHS__),
  channels: new Set(['ALL']),  // 채널 필터
  brands:   new Set(['ALL']),  // 브랜드 필터
  forecast: false,
  search:   '',
  // 단일 선택 모드 (라벨 클릭으로 토글). 기본은 복수 선택(false).
  singleTeam:    false,
  singleMonth:   false,
  singleQuarter: false,
};

/* ─── 단일 선택 모드 토글 (팀/월/분기 라벨 클릭) ─── */
function toggleSingleMode(kind) {
  const key = { team:'singleTeam', month:'singleMonth', quarter:'singleQuarter' }[kind];
  if (!key) return;
  S[key] = !S[key];
  // 라벨 색상 갱신
  const labelMap = { team:'teamLabel', month:'monthLabel', quarter:'quarterGroup' };
  const lbl = document.getElementById(labelMap[kind]);
  if (lbl) lbl.classList.toggle('single-mode', S[key]);
}

// 팀별 채널 목록 사전 구축
const TEAM_CHANNELS = (() => {
  const map = {};
  for (const r of RAW) {
    if (!map[r.team]) map[r.team] = new Set();
    map[r.team].add(r.channel);
  }
  // 전체 채널
  const all = new Set();
  Object.values(map).forEach(s => s.forEach(c => all.add(c)));
  map['ALL'] = all;
  return map;
})();

/* ─── 채널 드롭다운 업데이트 ─── */
function updateChannelDropdown() {
  const sel = document.getElementById('chSel');
  if (!sel) return;

  // 현재 팀에 해당하는 채널 목록
  const allT = S.teams.has('ALL');
  const available = new Set();
  if (allT) {
    TEAM_CHANNELS['ALL'].forEach(c => available.add(c));
  } else {
    S.teams.forEach(t => {
      if (TEAM_CHANNELS[t]) TEAM_CHANNELS[t].forEach(c => available.add(c));
    });
  }

  // 엑셀 원본 순서로 정렬
  const sorted = chSort([...available]);

  // 드롭다운 재구성
  const prev = sel.value;
  sel.innerHTML = '<option value="ALL">전체 채널</option>' +
    sorted.map(c => `<option value="${c}"${c===prev?' selected':''}>${c}</option>`).join('');

  // 이전 선택값이 없으면 ALL로
  if (!available.has(prev) && prev !== 'ALL') {
    sel.value = 'ALL';
    S.channels = new Set(['ALL']);
  }
}

/* ─── FILTER TOGGLE ─── */
function togTeam(btn) {
  const v = btn.dataset.v;
  if (v === 'ALL') {
    S.teams = new Set(['ALL']);
  } else if (S.singleTeam) {
    // 단일 선택 모드: 다른 팀 클릭 시 기존 자동 해제, 같은 팀 다시 클릭 시 전사로
    if (S.teams.size === 1 && S.teams.has(v)) {
      S.teams = new Set(['ALL']);
    } else {
      S.teams = new Set([v]);
    }
  } else {
    // 복수 선택(기본)
    S.teams.delete('ALL');
    S.teams.has(v) ? S.teams.delete(v) : S.teams.add(v);
    if (!S.teams.size) S.teams = new Set(['ALL']);
  }
  document.querySelectorAll('[data-f="team"]').forEach(b =>
    b.classList.toggle('active', b.dataset.v === 'ALL' ? S.teams.has('ALL') : S.teams.has(b.dataset.v))
  );
  // 팀 수동 변경 시 채널 전체채널로 리셋
  S.channels = new Set(['ALL']);
  const chSelEl = document.getElementById('chSel');
  if (chSelEl) chSelEl.value = 'ALL';
  // 팀 바뀌면 채널 드롭다운 갱신 + 채널 선택 초기화
  S.channels = new Set(['ALL']);
  updateChannelDropdown();
  run();
}

function togChannel() {
  const sel = document.getElementById('chSel');
  const v = sel ? sel.value : 'ALL';
  S.channels = new Set([v]);

  if (v !== 'ALL') {
    // 해당 채널이 속한 팀 자동 선택
    const rec = RAW.find(r => r.channel === v);
    if (rec) {
      S.teams = new Set([rec.team]);
      document.querySelectorAll('[data-f="team"]').forEach(b =>
        b.classList.toggle('active', b.dataset.v === rec.team)
      );
    }
  } else {
    // 전체 채널 복귀 시 전사로 초기화
    S.teams = new Set(['ALL']);
    document.querySelectorAll('[data-f="team"]').forEach(b =>
      b.classList.toggle('active', b.dataset.v === 'ALL')
    );
  }

  run();
}

function togBrand() {
  const sel = document.getElementById('brSel');
  const v = sel ? sel.value : 'ALL';
  S.brands = new Set([v]);
  run();
}

function togYear(btn) {
  const v = btn.dataset.v;
  S.years.has(v) ? S.years.delete(v) : S.years.add(v);
  if (!S.years.size) { S.years.add(v); }
  document.querySelectorAll('[data-f="year"]').forEach(b =>
    b.classList.toggle('active', S.years.has(b.dataset.v))
  );
  run();
}

// 분기 → 월 매핑
const Q_MONTHS = { Q1:[1,2,3], Q2:[4,5,6], Q3:[7,8,9], Q4:[10,11,12] };

// 월 버튼 + 분기 버튼 active 상태 동기화
function syncMonthButtons() {
  // 월 버튼
  document.querySelectorAll('[data-f="month"]').forEach(b => {
    const v = b.dataset.v;
    b.classList.toggle('active',
      v === 'YTD' ? S.months.has('YTD') : S.months.has(Number(v))
    );
  });
  // 분기 버튼: 해당 분기의 모든 월이 선택돼 있으면 active
  document.querySelectorAll('[data-f="quarter"]').forEach(b => {
    const qm = Q_MONTHS[b.dataset.v];
    b.classList.toggle('active', qm.every(m => S.months.has(m)));
  });
}

function togMonth(btn) {
  const v = btn.dataset.v;
  if (v === 'YTD') {
    S.months = new Set(['YTD']);
  } else {
    const n = Number(v);
    if (S.singleMonth) {
      // 단일 선택 모드
      if (S.months.size === 1 && S.months.has(n)) {
        S.months = new Set(['YTD']);
      } else {
        S.months = new Set([n]);
      }
    } else {
      // 복수 선택(기본)
      S.months.delete('YTD');
      S.months.has(n) ? S.months.delete(n) : S.months.add(n);
      if (!S.months.size) S.months = new Set(['YTD']);
    }
  }
  syncMonthButtons();
  run();
}

function togQuarter(btn) {
  const qm = Q_MONTHS[btn.dataset.v];
  if (S.singleQuarter) {
    // 단일 선택 모드: 한 번에 한 분기만
    const exactMatch = S.months.size === qm.length && qm.every(m => S.months.has(m));
    if (exactMatch) {
      S.months = new Set(['YTD']);
    } else {
      S.months = new Set(qm);
    }
  } else {
    // 복수 선택(기본): 누적/해제
    const allSelected = qm.every(m => S.months.has(m));
    S.months.delete('YTD');
    if (allSelected) {
      qm.forEach(m => S.months.delete(m));
    } else {
      qm.forEach(m => S.months.add(m));
    }
    if (!S.months.size) S.months = new Set(['YTD']);
  }
  syncMonthButtons();
  run();
}

/* ─── ACTIVE MONTHS LIST ─── */
function activeMonths() {
  if (S.months.has('YTD')) return [1,2,3,4,5,6,7,8,9,10,11,12];
  return [...S.months].map(Number).sort((a,b)=>a-b);
}

/* ─── CHART MONTHS: 차트에 표시할 월 목록 ─── */
function chartMonths() {
  if (S.months.has('YTD')) return [1,2,3,4,5,6,7,8,9,10,11,12];
  const mons = [...S.months].map(Number).sort((a,b) => a-b);
  if (mons.length === 1) {
    // 1개만 선택 시 1월~선택월
    const end = mons[0];
    return Array.from({length: end}, (_, i) => i + 1);
  }
  return mons;
}

/* ─── AGGREGATE ─── */
function aggregate(teamFilter) {
  const allT = teamFilter ? false : S.teams.has('ALL');
  const teams = teamFilter ? new Set([teamFilter]) : S.teams;
  const mons  = new Set(activeMonths());
  const allCh = S.channels.has('ALL');
  const allB  = S.brands.has('ALL');
  const map   = new Map();

  for (const r of RAW) {
    if (!allT && !teams.has(r.team)) continue;
    if (!mons.has(r.month)) continue;
    if (!allCh && !S.channels.has(r.channel)) continue;
    if (!allB  && !S.brands.has(r.brand))     continue;
    const k = r.team + '\x00' + r.channel;
    if (!map.has(k)) map.set(k, {
      team:r.team, channel:r.channel,
      y2024:0, y2025b:0, y2025:0, plan:0, actual:0,
    });
    const a = map.get(k);
    a.y2024  += r.y2024  || 0;
    a.y2025b += r.y2025b || 0;
    a.y2025  += r.y2025  || 0;
    a.plan   += r.plan   || 0;
    a.actual += r.actual || 0;
    // 주차별 누적 (fw1, fw2, ...)
    for (let wi = 1; wi <= 5; wi++) {
      const fk = `fw${wi}`;
      if (r[fk] != null) {
        a[fk] = (a[fk] || 0) + r[fk];
      }
    }
  }

  let rows = [...map.values()];
  // 24년~26년 실적 모두 0인 채널 제외
  rows = rows.filter(r =>
    (r.y2024 || 0) + (r.y2025 || 0) + (r.plan || 0) + (r.actual || 0) > 0
  );
  rows.sort((a,b) => {
    const d = TORDER.indexOf(a.team) - TORDER.indexOf(b.team);
    if (d) return d;
    const ia = CH_ORDER.indexOf(a.channel), ib = CH_ORDER.indexOf(b.channel);
    if (ia === -1 && ib === -1) return a.channel.localeCompare(b.channel, 'ko');
    if (ia === -1) return 1;
    if (ib === -1) return -1;
    return ia - ib;
  });
  return rows;
}

/* ─── FORMAT ─── */
const Kn = n => Math.round(n).toLocaleString('ko-KR');
const fv = v => (!v && v !== 0) || v === 0 ? '-' : Kn(v);
const fd = v => {
  if (v === null || v === undefined || !isFinite(v) || isNaN(v)) return '-';
  if (Math.abs(v) < 0.5) return '-';
  return v > 0 ? Kn(Math.round(v)) : `(${Kn(Math.round(Math.abs(v)))})`;
};
const fp = v => !isFinite(v) || isNaN(v) ? '-' : v < 0 ? `(${Math.abs(v).toFixed(1)}%)` : v.toFixed(1) + '%';
const fa = v => !isFinite(v) || isNaN(v) ? '-' : v.toFixed(1) + '%';
const dc = v => v > 0.5 ? 'pos' : v < -0.5 ? 'neg' : 'neu';
const pc = v => v > 0 ? 'pos' : v < 0 ? 'neg' : 'neu';

/* ─── DERIVED ─── */
function deriv(r) {
  const d25   = r.y2025  - r.y2024;
  const p25   = r.y2024  ? d25  / r.y2024  * 100 : NaN;
  const dPlan = r.actual - r.plan;
  const achPct= r.plan   ? r.actual / r.plan * 100 : NaN;
  const vs25  = r.actual - r.y2025;
  const vs25p = r.y2025  ? vs25 / r.y2025 * 100 : NaN;
  return { d25, p25, dPlan, achPct, vs25, vs25p };
}

/* ─── GRAND / TEAM TOTALS ─── */
function totals(rows) {
  const baseKeys = ['y2024','y2025b','y2025','plan','actual'];
  const fwKeys   = ['fw1','fw2','fw3','fw4','fw5'];
  const allKeys  = [...baseKeys, ...fwKeys];

  const mkEmpty = () => Object.fromEntries(allKeys.map(k => [k, 0]));
  const TT = {}, G = mkEmpty();

  for (const r of rows) {
    if (!TT[r.team]) TT[r.team] = mkEmpty();
    for (const k of allKeys) {
      const v = r[k] || 0;
      TT[r.team][k] += v;
      G[k]          += v;
    }
  }
  return { TT, G };
}

/* ═══════════════════════════════════════════════════════
   KPI STRIP
═══════════════════════════════════════════════════════ */
function buildKPI(rows) {
  const G = totals(rows).G;
  const d = deriv(G);

  // 증감값 포맷: ▲/▼ 금액 (증감율%)
  function dFmt(delta, pct) {
    const pos = delta > 0.5, neg = delta < -0.5;
    if (!pos && !neg) return { cls:'neu', txt:'–' };
    const arrow = pos ? '▲' : '▼';
    const amt   = Kn(Math.round(Math.abs(delta)));  // ▲/▼로 부호 표시 → 절대값만
    const pctStr = isFinite(pct)&&!isNaN(pct)
      ? pct.toLocaleString('ko-KR', {minimumFractionDigits:1, maximumFractionDigits:1})
      : null;
    const pStr  = pctStr !== null ? ` (${pos?'+':''}${pctStr}%)` : '';
    return { cls: pos?'pos':'neg', txt:`${arrow} ${amt}백만원${pStr}` };
  }

  // 기본 카드: 1행=지표+금액, 2행=vs라벨+증감값
  function card(tag, amount, vsLbl, delta, deltaPct) {
    const df = dFmt(delta, deltaPct);
    return `<div class="kpi-card">
      <div class="kpi-row">
        <span class="kpi-tag">${tag}</span>
        <span class="kpi-amount">${amount}<span class="kpi-unit">백만원</span></span>
      </div>
      <div class="kpi-row2">
        <span class="kpi-vs-lbl">${vsLbl}</span>
        <span class="kpi-vs-val ${df.cls}">${df.txt}</span>
      </div>
    </div>`;
  }

  // 금액만 카드 (증감 없음)
  function cardPlain(tag, amount) {
    return `<div class="kpi-card">
      <div class="kpi-row">
        <span class="kpi-tag">${tag}</span>
        <span class="kpi-amount">${amount}<span class="kpi-unit">백만원</span></span>
      </div>
    </div>`;
  }

  // 가로 달성율 바 카드 (그라데이션)
  function barCard(pctRaw, actual, plan) {
    const pct    = Math.min(Math.max(pctRaw, 0), 150);
    const w      = Math.min((pct / 150) * 100, 100).toFixed(1);
    const color  = pctRaw >= 100 ? '#375623' : pctRaw >= 80 ? '#C55A11' : '#dc2626';
    const grad   = pctRaw >= 100
      ? 'linear-gradient(90deg,#84cc16,#375623)'
      : pctRaw >= 80
      ? 'linear-gradient(90deg,#fde68a,#C55A11)'
      : 'linear-gradient(90deg,#fca5a5,#dc2626)';
    const df = dFmt(G.actual - G.plan, d.achPct - 100);
    return `<div class="kpi-card-bar">
      <div class="kpi-bar-head">
        <span class="kpi-bar-lbl">사업계획 달성율</span>
        <span class="kpi-bar-pct" style="color:${color}">${pctRaw.toFixed(1)}%</span>
      </div>
      <div class="kpi-bar-track">
        <div class="kpi-bar-fill" style="width:${w}%;background:${grad}"></div>
      </div>
      <div class="kpi-bar-sub">${df.cls!=='neu'?`<span style="color:${color};font-weight:700">${df.txt}</span>`:''}</div>
    </div>`;
  }

  const achPct = isFinite(d.achPct) ? d.achPct : 0;
  const vs24   = G.actual - G.y2024;
  const vs24p  = G.y2024 ? vs24 / G.y2024 * 100 : NaN;
  const vsPlan = G.actual - G.plan;
  const vsPlanP= G.plan   ? vsPlan / G.plan * 100 : NaN;

  const html = [
    // ① Y2024 실적 (증감 없음)
    cardPlain('Y2024 실적', fv(G.y2024)),
    // ② Y2025 실적 (증감 없음)
    cardPlain('Y2025 실적', fv(G.y2025)),
    // ③ Y2026 실적 (증감 없음)
    cardPlain('Y2026 실적', fv(G.actual)),
    // ④ Y2026 계획 (증감 없음)
    cardPlain('Y2026 계획', fv(G.plan)),
    // ⑤ Y2024 대비 증감 + Y2025 대비 증감 (2줄) — 증감 값 강조
    `<div class="kpi-card">
      <div class="kpi-row2">
        <span class="kpi-vs-lbl">Y2024 실적 대비</span>
        <span class="kpi-vs-val kpi-vs-emph ${dFmt(vs24,vs24p).cls}">${dFmt(vs24,vs24p).txt}</span>
      </div>
      <div class="kpi-row2">
        <span class="kpi-vs-lbl">Y2025 실적 대비</span>
        <span class="kpi-vs-val kpi-vs-emph ${dFmt(d.vs25,d.vs25p).cls}">${dFmt(d.vs25,d.vs25p).txt}</span>
      </div>
    </div>`,
    // ⑥ Y2026 계획 대비 증감 + 게이지바 + 달성율
    (() => {
      const color = achPct>=100?'#375623':achPct>=80?'#C55A11':'#dc2626';
      const grad  = achPct>=100
        ? 'linear-gradient(90deg,#84cc16,#375623)'
        : achPct>=80
        ? 'linear-gradient(90deg,#fde68a,#C55A11)'
        : 'linear-gradient(90deg,#fca5a5,#dc2626)';
      const w = Math.min((Math.min(Math.max(achPct,0),150)/150)*100, 100).toFixed(1);
      const df = dFmt(vsPlan, vsPlanP);
      return `<div class="kpi-card" style="gap:3px">
        <div class="kpi-row2">
          <span class="kpi-vs-lbl">Y2026 계획 대비</span>
          <span class="kpi-vs-val ${df.cls}">${df.txt}</span>
        </div>
        <div style="display:flex;align-items:center;gap:6px;margin:1px 0">
          <span class="kpi-bar-lbl" style="flex-shrink:0">달성율</span>
          <div class="kpi-bar-track" style="flex:1;margin:0">
            <div class="kpi-bar-fill" style="width:${w}%;background:${grad}"></div>
          </div>
          <span style="font-size:clamp(12px,1vw,14px);font-weight:800;color:${color};flex-shrink:0">${fa(achPct)}</span>
        </div>
      </div>`;
    })(),
  ].join('');

  document.getElementById('kpiStrip').innerHTML = html;
  const kpiTable = document.getElementById('kpiStripTable');
  if (kpiTable) kpiTable.innerHTML = html;
}

/* ═══════════════════════════════════════════════════════
   TABLE
═══════════════════════════════════════════════════════ */
function getYearSel() {
  // pill 버튼 방식: S.years Set 기반
  return S.years;
}
/* ─── BU 그룹 정의 ─── */
const BU_MAP = {
  'RBD1팀':   '동북아 BU',
  'RBD2팀':   '동북아 BU',
  '일본사업팀': '동북아 BU',
  '중국사업팀': '동북아 BU',
  '동북아MC팀': '동북아 BU',
  'Global사업팀': 'Global BU',
  'GEC팀':    'Global BU',
  '메디컬팀': 'Expert BU',
};
const BU_ORDER = ['동북아 BU','Global BU','Expert BU'];

function getFcChk() {
  const el = document.getElementById('chkFc');
  const oneMonth = S.months.size === 1 && !S.months.has('YTD');
  return el && el.checked && oneMonth;
}

function getWeeklyMeta() {
  if (!getFcChk()) return null;
  const mn = parseInt([...S.months][0]);
  const labels = WEEKLY_META[mn];
  if (!labels || !labels.length) return null;
  return { mn, labels };
}

function buildTable(rows) {
  const yr = getYearSel();
  const show24   = yr.has('2024');
  const show25   = yr.has('2025');
  const showPlan = yr.has('plan');
  const show26   = yr.has('2026');
  const showFc   = getFcChk();
  const wkMeta   = getWeeklyMeta();
  const showFcNote = document.getElementById('chkFc')?.checked && !(S.months.size===1&&!S.months.has('YTD'));
  document.getElementById('fcNote').textContent = showFcNote ? '※ 예상표기: 월 1개 선택 시 활성' : '';

  const teamW = 80, chW = 110;

  /* ── HEADER ── */
  let gh = '', sh = '';
  gh += `<th class="hg-fix fix"  rowspan="2" style="min-width:${teamW}px;width:${teamW}px;left:0">팀</th>`;
  gh += `<th class="hg-fix fix2 hg-fix2-sep" rowspan="2" style="min-width:${chW}px;width:${chW}px;left:${teamW}px">채널</th>`;

  if (show24)   gh += `<th class="hg-act24 hg-sep-r" rowspan="2" style="min-width:72px">Y2024<br>실적</th>`;
  if (show25)   gh += `<th class="hg-act25 hg-sep-r" rowspan="2" style="min-width:72px">Y2025<br>실적</th>`;
  if (show26)   gh += `<th class="hg-act26 hg-sep-r" rowspan="2" style="min-width:72px">Y2026<br>실적</th>`;
  if (showPlan) gh += `<th class="hg-plan  hg-sep-r" rowspan="2" style="min-width:72px">Y2026<br>계획</th>`;

  if (show25 && show26) {
    gh += `<th class="hg-cmp25 hg-sep-r" colspan="2">25년 대비</th>`;
    sh += `<th class="hg-cmp25" style="min-width:66px">증감액</th><th class="hg-cmp25 hg-sep-r" style="min-width:54px">증감율</th>`;
  }
  if (showPlan && show26) {
    gh += `<th class="hg-cmpP ${!showFc?'hg-sep-r':''}" colspan="2">계획 대비</th>`;
    sh += `<th class="hg-cmpP" style="min-width:66px">증감액</th><th class="hg-cmpP ${!showFc?'hg-sep-r':''}" style="min-width:54px">달성율</th>`;
  }
  if (showFc && wkMeta) {
    const { mn, labels } = wkMeta;
    const totalFcCols = labels.length + (labels.length > 1 ? 1 : 0);
    gh += `<th class="hg-fc fc-cell-first" colspan="${totalFcCols}">${mn}월 당월예상</th>`;
    labels.forEach((lbl,i) => { sh += `<th class="hg-fc ${i===0?'fc-cell-first':''}" style="min-width:72px">${lbl}</th>`; });
    if (labels.length > 1) sh += `<th class="hg-fc" style="min-width:72px">전주대비</th>`;
  }

  document.getElementById('tblHead').innerHTML =
    `<tr class="gh">${gh}</tr>` + (sh ? `<tr class="sh">${sh}</tr>` : '');

  /* ── BODY ── */
  if (!rows.length) {
    document.getElementById('tblBody').innerHTML =
      `<tr><td colspan="20" class="empty-row">조건에 맞는 데이터가 없습니다.</td></tr>`;
    document.getElementById('tblCount').textContent = '0건';
    return;
  }

  const { TT, G } = totals(rows);

  // 주차 셀 생성
  function fwCells(obj, isAgg) {
    if (!showFc || !wkMeta) return '';
    const { labels } = wkMeta;
    let s = '';
    const vals = labels.map((_, i) => {
      const v = obj[`fw${i+1}`];
      return (v != null && v !== 0) ? v : null;
    });
    const cellStyle = 'text-align:right';
    vals.forEach((v, i) => {
      const firstCls = i === 0 ? 'fc-cell-first' : '';
      const negStyle = (v != null && v < -0.5) ? 'color:#dc2626;' : '';
      s += `<td class="fc-cell ${firstCls}" style="${cellStyle};${negStyle}">${v != null ? fv(v) : '-'}</td>`;
    });
    if (labels.length > 1) {
      const last = vals[vals.length-1], prev = vals[vals.length-2];
      const diff = (last != null && prev != null) ? last - prev : null;
      const negStyle = (diff != null && diff < -0.5) ? 'color:#dc2626;' : '';
      s += `<td class="fc-cell" style="${negStyle}">${diff != null ? fd(diff) : '-'}</td>`;
    }
    return s;
  }

  // 그룹 구분선 헬퍼
  const sep = 'hg-sep-r';
  // 음수 인라인 색상 헬퍼
  const negC  = v => (v !== null && isFinite(v) && v < -0.5) ? ' style="color:#dc2626"' : '';
  const negCP = v => (v !== null && isFinite(v) && v < 0)    ? ' style="color:#dc2626"' : '';

  function dataCells(r) {
    const dv = deriv(r);
    let s = '';
    if (show24)   s += `<td class="v-actual ${sep}">${fv(r.y2024)}</td>`;
    if (show25)   s += `<td class="v-actual ${sep}">${fv(r.y2025)}</td>`;
    if (show26)   s += `<td class="v-act26 ${sep}">${fv(r.actual)}</td>`;
    if (showPlan) s += `<td class="v-plan ${sep}">${fv(r.plan)}</td>`;
    if (show25 && show26) {
      s += `<td class="v-cmp"${negC(dv.vs25)}>${fd(dv.vs25)}</td>`;
      s += `<td class="v-cmp ${sep}"${negCP(dv.vs25p)}>${fp(dv.vs25p)}</td>`;
    }
    if (showPlan && show26) {
      s += `<td class="v-cmp"${negC(dv.dPlan)}>${fd(dv.dPlan)}</td>`;
      s += `<td class="v-cmp ${!showFc?sep:''}"${negCP(dv.achPct-100)}>${fa(dv.achPct)}</td>`;
    }
    s += fwCells(r, false);
    return s;
  }

  // 소계/합계 행
  function aggCells(t, cls) {
    const dv = deriv(t);
    // CSS 클래스 방식 (인라인 style은 !important에 밀림)
    const nc  = v => (v != null && isFinite(v) && v < -0.5) ? ' neg' : '';
    const ncp = v => (v != null && isFinite(v) && v < 0)    ? ' neg' : '';
    let s = '';
    if (show24)   s += `<td class="${sep}${nc(t.y2024)}">${fd(t.y2024)}</td>`;
    if (show25)   s += `<td class="${sep}${nc(t.y2025)}">${fd(t.y2025)}</td>`;
    if (show26)   s += `<td class="${sep}${nc(t.actual)}">${fd(t.actual)}</td>`;
    if (showPlan) s += `<td class="${sep}${nc(t.plan)}">${fd(t.plan)}</td>`;
    if (show25 && show26) { s += `<td class="${nc(dv.vs25)}">${fd(dv.vs25)}</td><td class="${sep}${ncp(dv.vs25p)}">${fp(dv.vs25p)}</td>`; }
    if (showPlan && show26) { s += `<td class="${nc(dv.dPlan)}">${fd(dv.dPlan)}</td><td class="${!showFc?sep:''}${ncp(dv.achPct-100)}">${fa(dv.achPct)}</td>`; }
    s += fwCells(t, true);
    return s;
  }

  // 팀별로 그룹핑
  const teamGroups = {};
  for (const r of rows) {
    if (!teamGroups[r.team]) teamGroups[r.team] = [];
    teamGroups[r.team].push(r);
  }

  // 전사/팀 선택 여부
  const isAllTeams = S.teams.has('ALL');

  let html = '';
  const teamsInData = SALES_TEAMS_ORDER.filter(t => teamGroups[t]);

  // BU별로 묶어서 렌더 + BU 합계
  const buTotals = {};

  BU_ORDER.forEach(bu => {
    const buTeams = teamsInData.filter(t => BU_MAP[t] === bu);
    if (!buTeams.length) return;

    buTeams.forEach(team => {
      const chRows = teamGroups[team];
      const tt = TT[team];

      // 현재 선택된 연도/예상 기준으로 모든 표시값이 0인 채널 행 제외
      const isAllZero = r => {
        let tot = 0;
        if (show24)   tot += Math.abs(r.y2024  || 0);
        if (show25)   tot += Math.abs(r.y2025  || 0);
        if (show26)   tot += Math.abs(r.actual || 0);
        if (showPlan) tot += Math.abs(r.plan   || 0);
        if (showFc && wkMeta) for (let wi=1;wi<=5;wi++) tot += Math.abs(r[`fw${wi}`]||0);
        return tot === 0;
      };
      const visibleRows = chRows.filter(r => !isAllZero(r));
      if (!visibleRows.length) return; // 팀 전체가 0이면 소계도 생략

      // 채널 행들 (접기 가능)
      visibleRows.forEach((r, ri) => {
        const isFirst = ri === 0;
        html += `<tr class="ch-row${_tableCollapsed ? ' collapsed' : ''}">`;
        html += isFirst
          ? `<td class="fix team" style="left:0;vertical-align:middle" rowspan="${visibleRows.length}">${team}</td>`
          : '';
        html += `<td class="fix2 ch hg-fix2-sep" style="left:${teamW}px">${r.channel}</td>`;
        html += dataCells(r);
        html += `</tr>`;
      });

      // 소계 행
      html += `<tr class="sub-row">
        <td class="fix team" style="left:0;text-align:center">${team}</td>
        <td class="fix2 hg-fix2-sep" style="left:${teamW}px;font-weight:700;text-align:center">소 계</td>
        ${aggCells(tt, 'sub-row')}
      </tr>`;

      // BU 합산
      if (!buTotals[bu]) buTotals[bu] = {y2024:0,y2025b:0,y2025:0,plan:0,actual:0,fw1:0,fw2:0,fw3:0,fw4:0,fw5:0};
      for (const k of Object.keys(buTotals[bu])) buTotals[bu][k] += tt[k]||0;
    });

    // BU 합계 행 (전사 선택 시에만)
    if (isAllTeams && buTotals[bu]) {
      const bt = buTotals[bu];
      html += `<tr class="bu-row">
        <td class="fix" style="left:0;font-weight:700;color:#fff;background:#2e6da4;text-align:center">${bu}</td>
        <td class="fix2 hg-fix2-sep" style="left:${teamW}px;background:#2e6da4;color:#fff;font-weight:700;text-align:center">합 계</td>
        ${aggCells(bt, 'bu-row')}
      </tr>`;
    }
  });

  // 합계 행
  html += `<tr class="grand-row">
    <td class="fix" style="left:0;color:#e2e8f0!important;font-weight:700;text-align:center">🏁 전 사</td>
    <td class="fix2 hg-fix2-sep" style="left:${teamW}px;text-align:center;font-weight:700">합 계</td>
    ${aggCells(G, 'grand-row')}
  </tr>`;

  document.getElementById('tblBody').innerHTML = html;
  document.getElementById('tblCount').textContent = `${rows.length}개 채널`;
}

/* ═══════════════════════════════════════════════════════
   CHARTS
═══════════════════════════════════════════════════════ */
const CHARTS = {};

const CHART_COLORS = {
  y2024:  { border:'#6b7280', bg:'rgba(107,114,128,.08)' }, // 회색 (2024 실적)
  y2025:  { border:'#5a96c8', bg:'rgba(90,150,200,.10)' },            // 스틸 블루
  plan:   { border:'#6e9650', bg:'rgba(110,150,80,.08)'  }, // 다크 그린 (2026 계획 - 테이블 헤더 동일)
  actual: { border:'#5C2508', bg:'rgba(92,37,8,.10)'   }, // 계획대비 헤더색 (2026 실적)
};

function chartData12months(teamFilter) {
  const allCh = S.channels.has('ALL');
  const allB  = S.brands.has('ALL');
  // 특정 채널 선택 시: 채널만으로 필터링 (팀 무시)
  if (!allCh) {
    const agg = {};
    for (let m = 1; m <= 12; m++) agg[m] = {y2024:0, y2025:0, plan:0, actual:0};
    for (const r of RAW) {
      if (!S.channels.has(r.channel)) continue;
      if (!allB && !S.brands.has(r.brand)) continue;
      agg[r.month].y2024  += r.y2024  || 0;
      agg[r.month].y2025  += r.y2025  || 0;
      agg[r.month].plan   += r.plan   || 0;
      agg[r.month].actual += r.actual || 0;
    }
    return agg;
  }
  // 전체 채널: 기존 팀 필터 적용
  const allT = teamFilter ? false : S.teams.has('ALL');
  const teams = teamFilter ? new Set([teamFilter]) : S.teams;
  const agg = {};
  for (let m = 1; m <= 12; m++) agg[m] = {y2024:0, y2025:0, plan:0, actual:0};
  for (const r of RAW) {
    if (!allT && !teams.has(r.team)) continue;
    if (!allB && !S.brands.has(r.brand)) continue;
    agg[r.month].y2024  += r.y2024  || 0;
    agg[r.month].y2025  += r.y2025  || 0;
    agg[r.month].plan   += r.plan   || 0;
    agg[r.month].actual += r.actual || 0;
  }
  return agg;
}

function makeLineDatasets(agg) {
  const maxMon = Math.max(...activeMonths());
  const mk = (key, lbl) => ({
    label: lbl,
    data: [1,2,3,4,5,6,7,8,9,10,11,12].map(m => m <= maxMon ? (agg[m][key] || null) : null),
    borderColor: CHART_COLORS[key].border,
    backgroundColor: CHART_COLORS[key].bg,
    borderWidth: key==='actual' ? 2.5 : 2,
    borderDash: CHART_COLORS[key].dash,
    pointRadius: 3,
    pointHoverRadius: 5,
    tension: 0.3,
    fill: false,
  });
  const sets = [];
  if (S.years.has('2024')) sets.push(mk('y2024', '2024 실적'));
  if (S.years.has('2025')) sets.push(mk('y2025', '2025 실적'));
  if (S.years.has('2026')) sets.push(mk('actual','2026 실적'));
  if (S.years.has('plan')) sets.push(mk('plan',  '2026 계획'));
  return sets;
}

function makeCumDatasets(agg) {
  const maxMon = Math.max(...activeMonths());
  function cumData(key) {
    let cum = 0;
    return [1,2,3,4,5,6,7,8,9,10,11,12].map(m => {
      if (m > maxMon) return null;
      const v = agg[m][key];
      if (!v && v !== 0) return null;
      cum += v;
      return cum;
    });
  }
  const sets = [];
  if (S.years.has('2024')) sets.push({ label:'2024 실적', data:cumData('y2024'), borderColor:'#6b7280', backgroundColor:'rgba(107,114,128,.06)', borderWidth:1.8, pointRadius:2, tension:0.3, fill:false });
  if (S.years.has('2025')) sets.push({ label:'2025 실적', data:cumData('y2025'), borderColor:'#5a96c8', backgroundColor:'rgba(90,150,200,.08)', borderWidth:2,   pointRadius:2, tension:0.3, fill:false });
  if (S.years.has('2026')) sets.push({ label:'2026 실적', data:cumData('actual'),borderColor:'#5C2508', backgroundColor:'rgba(92,37,8,.08)',  borderWidth:2.5, pointRadius:3, tension:0.3, fill:false });
  if (S.years.has('plan')) sets.push({ label:'2026 계획', data:cumData('plan'),  borderColor:'#6e9650', backgroundColor:'rgba(110,150,80,.06)',  borderWidth:2,   pointRadius:2, tension:0.3, fill:false });
  return sets;
}

function mkLineOpts(title, getExpanded) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    layout: { padding: { left: 0, right: 16, top: 6 } },
    plugins: {
      legend: { display: false },
      tooltip: {
        mode:'index', intersect:false,
        filter: getExpanded ? () => !getExpanded() : () => true,
        callbacks: {
          label: ctx => {
            if (ctx.parsed.y == null) return ` ${ctx.dataset.label}: -`;
            const eok = ctx.parsed.y / 100;
            const fmt = Math.abs(eok) >= 10 ? Math.round(eok).toLocaleString('ko-KR') : eok.toFixed(1);
            return ` ${ctx.dataset.label}: ${fmt}억원`;
          }
        }
      }
    },
    scales: {
      x: { grid:{display:false}, ticks:{font:{size:12,family:'Malgun Gothic',weight:'bold'}}, offset: true },
      y: {
        grid:{display:false},
        grace: '40%',
        ticks:{
          font:{size:10,family:'Malgun Gothic'},
          callback: v => {
            const eok = v / 100;
            if (Math.abs(eok) >= 10) return Math.round(eok).toLocaleString('ko-KR') + '억';
            return eok.toFixed(1) + '억';
          }
        },
        afterFit(axis) { if (axis.width < 55) axis.width = 55; }
      }
    },
    interaction: { mode:'index', intersect:false }
  };
}

/* ─── 월별 매출추이 상단 주석 (계획대비 / 전기대비) ─── */
let _mainAgg = {};

const mainAnnotationPlugin = {
  id: 'mainAnnotation',
  afterDatasetsDraw(chart) {
    const { ctx, scales: { x }, chartArea } = chart;
    const maxMon = Math.max(...activeMonths());
    const isExp = _mainExpanded;
    ctx.save();

    const rowFontSize = isExp ? 13 : 9;
    const valFontSize = isExp ? 13 : 9.5;
    const lineH = isExp ? 22 : 13;
    const yRow1 = chartArea.top + 2;
    const yRow2 = chartArea.top + 2 + lineH;

    // 계획달성율 (%) / 전기대비 증감율 (▲▼%)
    const fmtPlan = (v, ref) => (!ref || !v) ? '-' : (v / ref * 100).toFixed(1) + '%';
    const fmtYoy  = (v, ref) => {
      if (!ref || !v || Math.abs(v - ref) < 0.5) return '-';
      const p = (v - ref) / ref * 100;
      return (p > 0 ? '▲' : '▼') + Math.abs(p).toFixed(1) + '%';
    };
    const colPlan = (v, ref) => (!ref || !v) ? '#9ca3af' : v >= ref - 0.5 ? '#375623' : '#C55A11';
    const colYoy  = (v, ref) => (!ref || !v || Math.abs(v - ref) < 0.5) ? '#9ca3af'
      : v > ref ? '#375623' : '#C55A11';

    // 행 라벨 — 차트 내부 왼쪽 상단
    ctx.font = `600 ${rowFontSize}px "Malgun Gothic", sans-serif`;
    ctx.fillStyle = '#9ca3af';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText('계획달성', chartArea.left + 4, yRow1);
    ctx.fillText('전기대비', chartArea.left + 4, yRow2);
    const rowLblWidth = Math.max(
      ctx.measureText('계획달성').width,
      ctx.measureText('전기대비').width
    );

    // 월별 값
    ctx.font = `700 ${valFontSize}px "Malgun Gothic", sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    const slotW = activeMonths().length > 1
      ? Math.abs(x.getPixelForValue(1) - x.getPixelForValue(0)) : chartArea.width;
    const jan1Offset = isExp ? slotW * 0.35 : 0;

    // 1) 라벨 후보 수집 + 폭 측정
    const planItems = [], yoyItems = [];
    for (let m = 1; m <= 12; m++) {
      if (m > maxMon) break;
      const d = _mainAgg[m];
      if (!d) continue;
      const xPos = x.getPixelForValue(m - 1) + (m === 1 ? jan1Offset : 0);
      const tP = fmtPlan(d.actual, d.plan);
      const tY = fmtYoy(d.actual, d.y2025);
      planItems.push({ x: xPos, text: tP, color: colPlan(d.actual, d.plan), w: ctx.measureText(tP).width });
      yoyItems.push({ x: xPos, text: tY, color: colYoy(d.actual, d.y2025), w: ctx.measureText(tY).width });
    }

    // 2) 행별 가로 충돌 해결 - 인접 라벨을 좌/우로 절반씩 밀어냄
    const hLeft = chartArea.left + 4 + rowLblWidth + 6;
    const hRight = chartArea.right - 2;
    const minGap = isExp ? 6 : 3;
    const resolveH = items => {
      for (let pass = 0; pass < 4; pass++) {
        let moved = false;
        for (let i = 1; i < items.length; i++) {
          const p = items[i - 1], c = items[i];
          const overlap = (p.x + p.w/2) + minGap - (c.x - c.w/2);
          if (overlap > 0) {
            p.x -= overlap / 2;
            c.x += overlap / 2;
            moved = true;
          }
        }
        if (!moved) break;
      }
      items.forEach(it => {
        if (it.x - it.w/2 < hLeft)  it.x = hLeft + it.w/2;
        if (it.x + it.w/2 > hRight) it.x = hRight - it.w/2;
      });
    };
    resolveH(planItems);
    resolveH(yoyItems);

    // 3) 실제 그리기
    planItems.forEach(it => { ctx.fillStyle = it.color; ctx.fillText(it.text, it.x, yRow1); });
    yoyItems.forEach(it => { ctx.fillStyle = it.color; ctx.fillText(it.text, it.x, yRow2); });

    // 확대 시: 데이터셋별 월별 금액 레이블 (선 색상, 겹침 방지)
    if (isExp) {
      ctx.font = `600 11px "Malgun Gothic", sans-serif`;
      ctx.textAlign = 'center';
      const minGap = 14;
      const offset = 6;

      for (let col = 0; col < maxMon; col++) {
        const items = [];
        chart.data.datasets.forEach((ds, di) => {
          const meta = chart.getDatasetMeta(di);
          if (meta.hidden) return;
          const val = ds.data[col];
          if (val == null) return;
          const pt = meta.data[col];
          const eok = val / 100;
          const fmt = (Math.abs(eok) >= 10
            ? Math.round(eok).toLocaleString('ko-KR')
            : eok.toFixed(1)) + '억';
          items.push({ y: pt.y, x: pt.x, text: fmt, color: ds.borderColor });
        });
        if (!items.length) continue;
        // 점들을 y 기준 정렬(위→아래). 절반은 점 위, 절반은 점 아래로 배치하여 선과 겹침 최소화.
        items.sort((a, b) => a.y - b.y);
        const half = Math.ceil(items.length / 2);
        // 위쪽 그룹: 점 위에 표기, 아래→위로 stacking
        for (let i = half - 1; i >= 0; i--) {
          let y = items[i].y - offset;
          if (i < half - 1) {
            const prevY = items[i + 1]._ry;
            if (prevY - y < minGap) y = prevY - minGap;
          }
          y = Math.max(y, chartArea.top + 14);
          items[i]._ry = y;
          items[i]._bl = 'bottom';
        }
        // 아래쪽 그룹: 점 아래에 표기, 위→아래로 stacking
        for (let i = half; i < items.length; i++) {
          let y = items[i].y + offset;
          if (i > half) {
            const prevY = items[i - 1]._ry;
            if (y - prevY < minGap) y = prevY + minGap;
          }
          y = Math.min(y, chartArea.bottom - 4);
          items[i]._ry = y;
          items[i]._bl = 'top';
        }
        items.forEach(item => {
          ctx.fillStyle = item.color;
          ctx.textBaseline = item._bl;
          ctx.fillText(item.text, item.x, item._ry);
        });
      }
    }

    ctx.restore();
  }
};

/* ─── 누적 달성율 주석 플러그인 ─── */
const cumLineAnnotationPlugin = {
  id: 'cumLineAnnotation',
  afterDatasetsDraw(chart) {
    const { ctx, scales: { x }, chartArea } = chart;
    const maxMon = Math.max(...activeMonths());
    const isExp = _cumLineExpanded;
    ctx.save();

    const rowFontSize = isExp ? 13 : 9;
    const valFontSize = isExp ? 13 : 9.5;
    const lineH = isExp ? 22 : 13;
    const yLbl = chartArea.top + 2;
    const yRow = chartArea.top + 2 + lineH;

    ctx.font = `600 ${rowFontSize}px "Malgun Gothic", sans-serif`;
    ctx.fillStyle = '#9ca3af';
    ctx.textAlign = 'left';
    ctx.textBaseline = 'middle';
    ctx.fillText('누적달성율', chartArea.left + 4, yLbl);
    const rowLblWidth = ctx.measureText('누적달성율').width;

    let cumActual = 0, cumPlan = 0;
    ctx.font = `700 ${valFontSize}px "Malgun Gothic", sans-serif`;
    ctx.textAlign = 'center';

    // 1) 라벨 후보 수집 + 폭 측정
    const items = [];
    for (let m = 1; m <= 12; m++) {
      if (m > maxMon) break;
      const d = _mainAgg[m];
      if (!d) continue;
      cumActual += d.actual || 0;
      cumPlan   += d.plan   || 0;
      if (!cumPlan) continue;
      const pct = cumActual / cumPlan * 100;
      const t = pct.toFixed(1) + '%';
      items.push({
        x: x.getPixelForValue(m - 1),
        text: t,
        color: pct >= 100 ? '#375623' : pct >= 80 ? '#C55A11' : '#dc2626',
        w: ctx.measureText(t).width,
      });
    }

    // 2) 가로 충돌 해결 + 경계 클램핑 (양방향 스윕, 안정 시까지 반복)
    // 라벨('누적달성율')은 위 행(yLbl)에, 퍼센트는 아래 행(yRow)에 그려지므로 가로 충돌 無 → 왼쪽 끝까지 사용
    const hLeft = chartArea.left + 4;
    const hRight = chartArea.right - 2;
    const minGap = isExp ? 6 : 3;
    for (let pass = 0; pass < 20; pass++) {
      let moved = false;
      // 정방향: 이전과 겹치면 현재를 오른쪽으로 밀기
      for (let i = 1; i < items.length; i++) {
        const p = items[i - 1], c = items[i];
        const minX = (p.x + p.w/2) + minGap + c.w/2;
        if (c.x < minX - 0.5) { c.x = minX; moved = true; }
      }
      // 역방향: 다음과 겹치면 현재를 왼쪽으로 밀기
      for (let i = items.length - 2; i >= 0; i--) {
        const c = items[i], n = items[i + 1];
        const maxX = (n.x - n.w/2) - minGap - c.w/2;
        if (c.x > maxX + 0.5) { c.x = maxX; moved = true; }
      }
      // 경계 클램핑
      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        if (it.x - it.w/2 < hLeft)  { it.x = hLeft  + it.w/2; moved = true; }
        if (it.x + it.w/2 > hRight) { it.x = hRight - it.w/2; moved = true; }
      }
      if (!moved) break;
    }

    // 3) 실제 그리기
    items.forEach(it => { ctx.fillStyle = it.color; ctx.fillText(it.text, it.x, yRow); });

    // 확대 시: 데이터셋별 누적 금액 레이블 (선 색상, 위/아래 분산 배치로 겹침 최소화)
    if (isExp) {
      ctx.font = `600 11px "Malgun Gothic", sans-serif`;
      ctx.textAlign = 'center';
      const minGap = 14;
      const offset = 6;

      for (let col = 0; col < maxMon; col++) {
        const items = [];
        chart.data.datasets.forEach((ds, di) => {
          const meta = chart.getDatasetMeta(di);
          if (meta.hidden) return;
          const val = ds.data[col];
          if (val == null) return;
          const pt = meta.data[col];
          const eok = val / 100;
          const fmt = (Math.abs(eok) >= 10
            ? Math.round(eok).toLocaleString('ko-KR')
            : eok.toFixed(1)) + '억';
          items.push({ y: pt.y, x: pt.x, text: fmt, color: ds.borderColor });
        });
        if (!items.length) continue;
        items.sort((a, b) => a.y - b.y);
        const half = Math.ceil(items.length / 2);
        for (let i = half - 1; i >= 0; i--) {
          let y = items[i].y - offset;
          if (i < half - 1) {
            const prevY = items[i + 1]._ry;
            if (prevY - y < minGap) y = prevY - minGap;
          }
          y = Math.max(y, chartArea.top + 14);
          items[i]._ry = y;
          items[i]._bl = 'bottom';
        }
        for (let i = half; i < items.length; i++) {
          let y = items[i].y + offset;
          if (i > half) {
            const prevY = items[i - 1]._ry;
            if (y - prevY < minGap) y = prevY + minGap;
          }
          y = Math.min(y, chartArea.bottom - 4);
          items[i]._ry = y;
          items[i]._bl = 'top';
        }
        items.forEach(item => {
          ctx.fillStyle = item.color;
          ctx.textBaseline = item._bl;
          ctx.fillText(item.text, item.x, item._ry);
        });
      }
    }

    ctx.restore();
  }
};

function mkCumLineOpts() {
  return mkLineOpts(null, () => _cumLineExpanded);
}

/* 공통 확대 위치 계산 */
function getExpandRect() {
  const stickyH = (document.querySelector('.sticky-header')?.offsetHeight   || 44)
                + (document.querySelector('.sticky-filterbar')?.offsetHeight || 70)
                + (document.querySelector('.sticky-kpi')?.offsetHeight       || 56);
  const pad = 10;
  return { top: stickyH + pad, left: pad,
           w: window.innerWidth - pad * 2, h: window.innerHeight - stickyH - pad * 2 };
}
function applyExpandCard(card) {
  const r = getExpandRect();
  card.style.top = r.top + 'px'; card.style.left = r.left + 'px';
  card.style.width = r.w + 'px'; card.style.height = r.h + 'px';
}

function initCharts() {
  /* Main line chart: 월별 매출 추이 */
  const ctx1 = document.getElementById('chartMain').getContext('2d');
  _mainAgg = chartData12months(null);
  CHARTS.main = new Chart(ctx1, {
    type: 'line',
    data: { labels: MLABEL, datasets: makeLineDatasets(_mainAgg) },
    options: mkLineOpts('월별 매출 추이', () => _mainExpanded),
    plugins: [mainAnnotationPlugin],
  });

  /* 누적 매출 라인 차트 */
  const ctx2 = document.getElementById('chartCumLine').getContext('2d');
  CHARTS.cumLine = new Chart(ctx2, {
    type: 'line',
    data: { labels: MLABEL, datasets: makeCumDatasets(_mainAgg) },
    options: mkCumLineOpts(),
    plugins: [cumLineAnnotationPlugin],
  });

  /* 누적 계획 대비 실적 막대 차트 — 전사/팀에 따라 동적 전환 */
  renderCumBar();
}

function updateCharts() {
  if (!CHARTS.main) return;
  _mainAgg = chartData12months(null);

  // 월별 추이 (labels 고정 1~12, 선택 범위 밖은 null)
  CHARTS.main.data.labels   = MLABEL;
  CHARTS.main.data.datasets = makeLineDatasets(_mainAgg);
  CHARTS.main.update();

  // 누적 라인 (datasets 전체 재빌드 — 연도 필터 반영)
  CHARTS.cumLine.data.labels   = MLABEL;
  CHARTS.cumLine.data.datasets = makeCumDatasets(_mainAgg);
  CHARTS.cumLine.update();

  // 범례 동적 갱신 (legend1: 월별추이, legend2: 누적추이)
  const _s24 = S.years.has('2024'), _s25 = S.years.has('2025'),
        _s26 = S.years.has('2026'), _spl = S.years.has('plan');
  const legendHtml = () => {
    let h = '';
    if (_s24) h += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6b7280"></span>2024 실적</span>`;
    if (_s25) h += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5a96c8"></span>2025 실적</span>`;
    if (_s26) h += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5C2508"></span>2026 실적</span>`;
    if (_spl) h += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6e9650"></span>2026 계획</span>`;
    return h;
  };
  const lg1 = document.getElementById('legend1');
  const lg2 = document.getElementById('legend2');
  if (lg1) lg1.innerHTML = legendHtml();
  if (lg2) lg2.innerHTML = legendHtml();

  // 누적 달성 차트 — 전사/팀에 따라 동적 전환
  renderCumBar();
}

/* ═══════════════════════════════════════════════════════
   renderCumBar — 동적 전환 차트
   · 전사(ALL)  → 팀별 Y2024·Y2025·Y2026계획·Y2026실적 그룹 막대
   · 팀 선택   → 선택 월 채널별 Y2024·Y2025·Y2026계획·Y2026실적 그룹 막대
═══════════════════════════════════════════════════════ */

/* 막대 그룹 상단에 2줄 텍스트 그리는 플러그인 생성 함수 */
function makeTopLabelPlugin(labels, getLines, row1Lbl = '계획달성', row2Lbl = '전기대비') {
  return {
    id: 'topLabel',
    afterDatasetsDraw(chart) {
      const { ctx, scales: { x, y }, chartArea } = chart;
      const isExp = _cumBarExpanded;
      ctx.save();

      // ── 막대 상단 금액 표시 (확대 시에만) ──
      if (isExp) {
        chart.data.datasets.forEach((ds, di) => {
          const meta = chart.getDatasetMeta(di);
          meta.data.forEach((bar, i) => {
            const val = ds.data[i];
            if (!val) return;
            const eStr = (Math.abs(val)/100).toLocaleString('ko-KR', {minimumFractionDigits:1, maximumFractionDigits:1}) + '억';
            ctx.font = `600 12px "Malgun Gothic", sans-serif`;
            ctx.fillStyle = '#374151';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            ctx.fillText(eStr, bar.x, bar.y - 2);
          });
        });
      }

      // ── 행 레이블 (계획대비 / 전기대비) ──
      const rowFontSize = isExp ? 13 : 9;
      ctx.font = `600 ${rowFontSize}px "Malgun Gothic", sans-serif`;
      ctx.textAlign = 'left';
      ctx.textBaseline = 'middle';
      const yTop = y.getPixelForValue(y.max) + 2;
      const lineH = isExp ? 22 : 14;
      ctx.fillStyle = '#6b7280';
      if (row1Lbl) ctx.fillText(row1Lbl, chartArea.left + 4, yTop + lineH * 0.6);
      if (row2Lbl) ctx.fillText(row2Lbl, chartArea.left + 4, yTop + lineH * 1.7);

      // ── 막대 그룹 상단 증감 텍스트 ──
      const valFontSize = isExp ? 13 : 9.5;
      ctx.font = `700 ${valFontSize}px "Malgun Gothic", sans-serif`;
      ctx.textBaseline = 'bottom';
      const slotW = labels.length > 1
        ? Math.abs(x.getPixelForValue(1) - x.getPixelForValue(0))
        : chartArea.width;
      const offsetX = slotW * 0.48;

      labels.forEach((lbl, i) => {
        if (!lbl) return;
        const lines = getLines(lbl, i);
        if (!lines || !lines[0]) return;
        const xPos = x.getPixelForValue(i) + offsetX;
        ctx.textAlign = 'right';
        ctx.fillStyle = lines[0].color || '#374151';
        ctx.fillText(lines[0].text, xPos, yTop + lineH);
        if (lines[1]) {
          ctx.fillStyle = lines[1].color || '#6b7280';
          ctx.fillText(lines[1].text, xPos, yTop + lineH * 2.1);
        }
      });
      ctx.restore();
    }
  };
}
const SALES_TEAMS_ORDER = ['RBD1팀','RBD2팀','일본사업팀','중국사업팀','동북아MC팀','Global사업팀','GEC팀','메디컬팀'];

function renderCumBar() {
  const ctx = document.getElementById('chartCumBar').getContext('2d');
  if (CHARTS.cumBar) { CHARTS.cumBar.destroy(); CHARTS.cumBar = null; }

  const isAll  = S.teams.has('ALL');
  const mons   = activeMonths();
  const maxMon = Math.max(...mons);
  const titleEl  = document.getElementById('chartMeta3');
  const title3El = document.getElementById('chartTitle3');

  const show24   = S.years.has('2024');
  const show25   = S.years.has('2025');
  const show26   = S.years.has('2026');
  const showPlan = S.years.has('plan');
  const nDS      = [show24, show25, show26, showPlan].filter(Boolean).length || 1;

  const eok = v => {
    return (v / 100).toLocaleString('ko-KR', {minimumFractionDigits:1, maximumFractionDigits:1}) + '억';
  };
  const fP = v => (v * 100).toFixed(1) + '%';

  /* ── 전사: 팀별 Y2024·Y2025·Y2026계획·Y2026실적 그룹 막대 ── */
  if (isAll) {
    const monLabel = mons.length === 1
      ? MLABEL[maxMon - 1]
      : `${MLABEL[mons[0]-1]}~${MLABEL[maxMon-1]}`;
    if (title3El) title3El.textContent = '📊 팀별 실적';
    titleEl.textContent = `${monLabel} 기준`;

    // 범례 (선택된 연도만)
    const lg3all = document.getElementById('legend3');
    if (lg3all) {
      let lgHtml = '';
      if (show24)   lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6b7280"></span>2024 실적</span>`;
      if (show25)   lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5a96c8"></span>2025 실적</span>`;
      if (show26)   lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5C2508"></span>2026 실적</span>`;
      if (showPlan) lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6e9650"></span>2026 계획</span>`;
      lg3all.innerHTML = lgHtml;
    }

    // 팀별 집계
    const teamMap = {};
    for (const team of SALES_TEAMS_ORDER) {
      teamMap[team] = { y2024:0, y2025:0, plan:0, actual:0 };
    }
    for (const r of RAW) {
      if (!mons.includes(r.month) || !teamMap[r.team]) continue;
      if (!S.brands.has('ALL') && !S.brands.has(r.brand)) continue;
      teamMap[r.team].y2024  += r.y2024  || 0;
      teamMap[r.team].y2025  += r.y2025  || 0;
      teamMap[r.team].plan   += r.plan   || 0;
      teamMap[r.team].actual += r.actual || 0;
    }
    const teams = SALES_TEAMS_ORDER.filter(t => {
      const d = teamMap[t];
      return (show24 && d.y2024 > 0) || (show25 && d.y2025 > 0) ||
             (show26 && d.actual > 0) || (showPlan && d.plan > 0);
    });

    // 팀별 차트: 8팀 이하이므로 스크롤 없이 꽉 채움
    const scrollDiv2 = document.getElementById('cumBarScroll');
    const innerDiv2  = document.getElementById('cumBarInner');
    if (scrollDiv2) { scrollDiv2.style.overflowX = 'hidden'; }
    if (innerDiv2)  { innerDiv2.style.minWidth = '100%'; }

    // 막대 두께: 선택 연도 수 기준
    const teamCtxW = ctx.canvas.closest('.chart-card-body')?.clientWidth || ctx.canvas.parentElement.offsetWidth || 600;
    const teamBarPx = Math.max(8, Math.floor(teamCtxW / 8 * 0.8 / nDS));

    // Y축 max = 선택 연도 최대값 × 1.3
    const teamAllVals = teams.flatMap(t => {
      const d = teamMap[t], v = [];
      if (show24)   v.push(d.y2024);
      if (show25)   v.push(d.y2025);
      if (show26)   v.push(d.actual);
      if (showPlan) v.push(d.plan);
      return v;
    }).filter(Boolean);
    const teamYMax = Math.ceil(Math.max(...teamAllVals) * 1.3 / 100) * 100;

    // 상단 텍스트: 선택 연도에 따라 조건부
    const fDiffTeam = (v, p) => {
      if (!v && v !== 0) return null;
      const sign = v >= 0 ? '▲' : '▼';
      const pStr = p != null ? `(${p >= 0 ? '+' : ''}${p.toFixed(1)}%)` : '';
      const eStr = (Math.abs(v)/100).toLocaleString('ko-KR', {minimumFractionDigits:1, maximumFractionDigits:1}) + '억';
      return { text: `${sign}${eStr} ${pStr}`, color: v >= 0 ? '#375623' : '#C55A11' };
    };
    const teamTopPlugin = makeTopLabelPlugin(teams, (team) => {
      const d = teamMap[team];
      const row1 = (show26 && showPlan) ? (() => {
        const vsPlan = d.actual - d.plan, vsPlanP = d.plan ? vsPlan/d.plan*100 : null;
        return fDiffTeam(vsPlan, vsPlanP) || { text: '-', color: '#9ca3af' };
      })() : null;
      const row2 = (show26 && show25) ? (() => {
        const vs25 = d.actual - d.y2025, vs25P = d.y2025 ? vs25/d.y2025*100 : null;
        return fDiffTeam(vs25, vs25P) || { text: '-', color: '#9ca3af' };
      })() : null;
      return [row1, row2];
    }, show26 && showPlan ? '계획달성' : null, show26 && show25 ? '전기대비' : null);

    const teamDatasets = [];
    if (show24)   teamDatasets.push({ label:'Y2024 실적', data: teams.map(t => teamMap[t].y2024  || null), backgroundColor:'rgba(107,114,128,.55)', borderColor:'#6b7280', borderWidth:0, borderRadius:3, barPercentage:0.9, categoryPercentage:0.9 });
    if (show25)   teamDatasets.push({ label:'Y2025 실적', data: teams.map(t => teamMap[t].y2025  || null), backgroundColor:'rgba(90,150,200,.65)',  borderColor:'#5a96c8', borderWidth:0, borderRadius:3, barPercentage:0.9, categoryPercentage:0.9 });
    if (show26)   teamDatasets.push({ label:'Y2026 실적', data: teams.map(t => teamMap[t].actual || null), backgroundColor:'rgba(92,37,8,.80)',     borderColor:'#5C2508', borderWidth:0, borderRadius:3, barPercentage:0.9, categoryPercentage:0.9 });
    if (showPlan) teamDatasets.push({ label:'Y2026 계획', data: teams.map(t => teamMap[t].plan   || null), backgroundColor:'rgba(110,150,80,.55)',    borderColor:'#6e9650', borderWidth:0, borderRadius:3, barPercentage:0.9, categoryPercentage:0.9 });

    CHARTS.cumBar = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: teams,
        datasets: teamDatasets,
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins:{
          legend:{ display:false },
          tooltip:{
            mode:'index', intersect:false,
            callbacks:{ label: ctx => {
              if (ctx.parsed.y == null) return null;
              return ` ${ctx.dataset.label}: ${eok(ctx.parsed.y)}`;
            }}
          }
        },
        scales:{
          x:{ grid:{display:false}, offset:true,
              ticks:{font:{size:12,family:'Malgun Gothic',weight:'bold'}} },
          y:{ grid:{display:false}, max: teamYMax,
              ticks:{ font:{size:10,family:'Malgun Gothic'}, callback: eok,
                      maxTicksLimit: 6 } }
        },
        interaction:{mode:'index',intersect:false}
      },
      plugins: [teamTopPlugin]
    });

  /* ── 팀 선택: 선택 월 채널별 Y2024·Y2025·Y2026계획·Y2026실적 그룹 막대 ── */
  } else {
    const selTeams = [...S.teams];
    const monLabel = mons.length === 1
      ? MLABEL[maxMon - 1]
      : `${MLABEL[mons[0]-1]}~${MLABEL[maxMon-1]}`;
    if (title3El) title3El.textContent = '📊 채널별 실적';
    titleEl.textContent = `${selTeams.join(', ')} · ${monLabel} 기준`;

    // 범례 (선택된 연도만)
    const lg3 = document.getElementById('legend3');
    if (lg3) {
      let lgHtml = '';
      if (show24)   lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6b7280"></span>2024 실적</span>`;
      if (show25)   lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5a96c8"></span>2025 실적</span>`;
      if (show26)   lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5C2508"></span>2026 실적</span>`;
      if (showPlan) lgHtml += `<span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6e9650"></span>2026 계획</span>`;
      lg3.innerHTML = lgHtml;
    }

    // 선택 팀 채널만 집계
    const chMap = {};
    for (const r of RAW) {
      if (!selTeams.includes(r.team)) continue;
      if (!mons.includes(r.month)) continue;
      if (!S.brands.has('ALL') && !S.brands.has(r.brand)) continue;
      if (!chMap[r.channel]) chMap[r.channel] = { y2024:0, y2025:0, plan:0, actual:0 };
      chMap[r.channel].y2024  += r.y2024  || 0;
      chMap[r.channel].y2025  += r.y2025  || 0;
      chMap[r.channel].plan   += r.plan   || 0;
      chMap[r.channel].actual += r.actual || 0;
    }

    // 선택 팀 채널만 동적으로 (CH_ORDER 기준 정렬) — 선택 연도 기준 모두 0인 채널 제외
    const activeCh = chSort(Object.keys(chMap).filter(c => {
      const d = chMap[c];
      return (show24 && d.y2024 > 0) || (show25 && d.y2025 > 0) ||
             (show26 && d.actual > 0) || (showPlan && d.plan > 0);
    }));

    // 스크롤 컨테이너 초기화
    const scrollDiv = document.getElementById('cumBarScroll');
    const innerDiv  = document.getElementById('cumBarInner');
    const canvasEl  = document.getElementById('chartCumBar');

    // 컨테이너 가시 너비 기준 슬롯 계산
    const visibleW = scrollDiv.clientWidth || 600;
    const slotPx   = Math.floor(visibleW / 8);
    const SLOTS    = 8;

    // 8채널 초과: innerDiv를 넓혀서 스크롤 / 이하: 스크롤 없이 꽉 채움
    if (activeCh.length > SLOTS) {
      innerDiv.style.minWidth = (slotPx * activeCh.length) + 'px';
      scrollDiv.style.overflowX = 'auto';
    } else {
      innerDiv.style.minWidth = '100%';
      scrollDiv.style.overflowX = 'hidden';
    }

    // 막대 두께: 선택 연도 수 기준
    const barPx = Math.max(8, Math.floor(slotPx * 0.8 / nDS));

    // 채널이 8개 미만이면 뒤에 빈 더미 슬롯 추가 → 왼쪽 정렬 + 빈 공간 확보
    const dummyCount = Math.max(0, SLOTS - activeCh.length);
    const labels  = [...activeCh, ...Array(dummyCount).fill('')];
    const mkData  = key => [
      ...activeCh.map(c => chMap[c][key] || null),
      ...Array(dummyCount).fill(null)
    ];

    // Y축 max = 선택 연도 최대값 × 1.3
    const chAllVals = activeCh.flatMap(c => {
      const d = chMap[c], v = [];
      if (show24)   v.push(d.y2024);
      if (show25)   v.push(d.y2025);
      if (show26)   v.push(d.actual);
      if (showPlan) v.push(d.plan);
      return v;
    }).filter(Boolean);
    const chYMax = chAllVals.length ? Math.ceil(Math.max(...chAllVals) * 1.3 / 100) * 100 : undefined;

    // 상단 텍스트: 선택 연도에 따라 조건부
    const fDiffCh = (v, p) => {
      if (!v && v !== 0) return null;
      const sign = v >= 0 ? '▲' : '▼';
      const pStr = p != null ? `(${p >= 0 ? '+' : ''}${p.toFixed(1)}%)` : '';
      const eStr = (Math.abs(v)/100).toLocaleString('ko-KR', {minimumFractionDigits:1, maximumFractionDigits:1}) + '억';
      return { text: `${sign}${eStr} ${pStr}`, color: v >= 0 ? '#375623' : '#C55A11' };
    };
    const chTopPlugin = makeTopLabelPlugin(labels, (lbl, i) => {
      if (!lbl || !chMap[lbl]) return null;
      const d = chMap[lbl];
      const row1 = (show26 && showPlan) ? (() => {
        const vsPlan = d.actual - d.plan, vsPlanP = d.plan ? vsPlan/d.plan*100 : null;
        return fDiffCh(vsPlan, vsPlanP) || { text: '-', color: '#9ca3af' };
      })() : null;
      const row2 = (show26 && show25) ? (() => {
        const vs25 = d.actual - d.y2025, vs25P = d.y2025 ? vs25/d.y2025*100 : null;
        return fDiffCh(vs25, vs25P) || { text: '-', color: '#9ca3af' };
      })() : null;
      return [row1, row2];
    }, show26 && showPlan ? '계획달성' : null, show26 && show25 ? '전기대비' : null);

    const chDatasets = [];
    if (show24)   chDatasets.push({ label:'Y2024 실적', data: mkData('y2024'),  backgroundColor:'rgba(107,114,128,.55)', borderColor:'#6b7280', borderWidth:0, borderRadius:3, maxBarThickness: barPx });
    if (show25)   chDatasets.push({ label:'Y2025 실적', data: mkData('y2025'),  backgroundColor:'rgba(90,150,200,.65)',  borderColor:'#5a96c8', borderWidth:0, borderRadius:3, maxBarThickness: barPx });
    if (show26)   chDatasets.push({ label:'Y2026 실적', data: mkData('actual'), backgroundColor:'rgba(92,37,8,.80)',     borderColor:'#5C2508', borderWidth:0, borderRadius:3, maxBarThickness: barPx });
    if (showPlan) chDatasets.push({ label:'Y2026 계획', data: mkData('plan'),   backgroundColor:'rgba(110,150,80,.55)',    borderColor:'#6e9650', borderWidth:0, borderRadius:3, maxBarThickness: barPx });

    CHARTS.cumBar = new Chart(ctx, {
      type: 'bar',
      data: {
        labels,
        datasets: chDatasets,
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins:{
          legend:{ display:false },
          tooltip:{
            mode:'index', intersect:false,
            filter: item => item.parsed.y != null,
            callbacks:{ label: ctx => {
              if (ctx.parsed.y == null) return null;
              return ` ${ctx.dataset.label}: ${eok(ctx.parsed.y)}`;
            }}
          }
        },
        scales:{
          x:{ grid:{display:false}, offset:true,
              ticks:{ font:{size:12,family:'Malgun Gothic',weight:'bold'}, maxRotation:30 }
          },
          y:{ grid:{display:false}, max: chYMax,
              ticks:{ font:{size:10,family:'Malgun Gothic'}, callback: eok,
                      maxTicksLimit: 6 } }
        },
        interaction:{mode:'index',intersect:false}
      },
      plugins: [chTopPlugin]
    });
  }
}

/* ─── YTD를 전사에, 분기를 연도에 세로 맞춤 ─── */
function alignQuarterToYear() {
  requestAnimationFrame(() => requestAnimationFrame(() => {
    const fb = document.querySelector('.filterbar');
    if (!fb) return;
    const fbLeft = fb.getBoundingClientRect().left;

    // ① 팀행 ㅣ와 월행 ㅣ를 정확히 정렬 (separator 기준)
    const allBtn = document.querySelector('[data-f="team"][data-v="ALL"]');
    const ytdBtn = document.querySelector('[data-f="month"][data-v="YTD"]');
    if (allBtn && ytdBtn) {
      const sep1 = allBtn.nextElementSibling; // ㅣ after 전사
      const sep2 = ytdBtn.nextElementSibling; // ㅣ after YTD
      if (sep1 && sep2) {
        ytdBtn.style.marginLeft = '0px'; // reset → force reflow on next getBCR
        const sep1Left = sep1.getBoundingClientRect().left - fbLeft;
        const sep2Left = sep2.getBoundingClientRect().left - fbLeft;
        ytdBtn.style.marginLeft = Math.max(sep1Left - sep2Left, 0) + 'px';
      }
    }

    // ② 연도 fg는 자연 위치(팀 다음) 유지 — 강제 정렬 제거
    const yrFg = document.getElementById('yearGroup')?.closest('.fg');
    if (yrFg) yrFg.style.marginLeft = '';

    // ③ 분기 ← 연도 정렬 (분기 라벨이 연도 라벨과 같은 X 위치에 오도록)
    const yrLabel = document.getElementById('yearGroup');
    const qLabel  = document.getElementById('quarterGroup');
    const lastMon = document.querySelector('[data-f="month"][data-v="12"]');
    if (yrLabel && qLabel && lastMon) {
      qLabel.style.marginLeft = '0px'; // reset, force reflow
      const yrLeft = yrLabel.getBoundingClientRect().left - fbLeft;
      const qLeft0 = qLabel.getBoundingClientRect().left - fbLeft; // 분기 라벨의 자연 위치
      qLabel.style.marginLeft = Math.max(yrLeft - qLeft0, 12) + 'px';
    }
  }));
}
let _tableCollapsed = false;
function toggleTableDetail() {
  _tableCollapsed = !_tableCollapsed;
  document.querySelectorAll('tr.ch-row').forEach(tr => {
    tr.classList.toggle('collapsed', _tableCollapsed);
  });
  const btn = document.getElementById('toggleDetailBtn');
  if (btn) btn.textContent = _tableCollapsed ? '📋 채널 펼치기' : '📋 채널 접기';
}
/* ═══════════════════════════════════════════════════════
   인쇄용 표 빌드 (2페이지: 1=동북아BU, 2=Global+Expert+합계)
═══════════════════════════════════════════════════════ */
function buildPrintTable() {
  const showFc  = getFcChk();
  const wkMeta  = getWeeklyMeta();
  const mons    = activeMonths();
  const monLbl  = S.months.has('YTD') ? 'YTD'
    : mons.length === 1 ? `${mons[0]}월` : `${mons[0]}~${mons[mons.length-1]}월`;

  // 전사 기준 집계 (필터 무관하게 항상 전사)
  const savedTeams = new Set(S.teams);
  S.teams = new Set(['ALL']);
  const allRows = aggregate(null);
  S.teams = savedTeams;
  const { TT, G } = totals(allRows);

  // 메타 텍스트
  const subTxt = `전사 기준 · ${monLbl} · __BASE_DATE__ 기준`;
  const titleTxt = `씨엠에스랩 ${monLbl} 매출 현황`;
  const metaEl = document.getElementById('printMeta');
  const sub1El = document.getElementById('printPaperSub');
  const sub2El = document.getElementById('printPaperSub2');
  const ttl1El = document.getElementById('printPaperTitle');
  const ttl2El = document.getElementById('printPaperTitle2');
  if (metaEl) metaEl.textContent = `전사 기준 · ${monLbl}`;
  if (sub1El) sub1El.textContent = subTxt;
  if (sub2El) sub2El.textContent = subTxt;
  if (ttl1El) ttl1El.textContent = titleTxt;
  if (ttl2El) ttl2El.textContent = titleTxt;

  // 포맷
  const fv2 = v => (!v&&v!==0)||v===0 ? '-' : Math.round(v).toLocaleString('ko-KR');
  const fd2 = v => { if(!v||Math.abs(v)<0.5) return '-'; return v>0?Math.round(v).toLocaleString('ko-KR'):`(${Math.round(Math.abs(v)).toLocaleString('ko-KR')})`; };
  const fp2 = v => !isFinite(v)||isNaN(v)?'-':v<0?`(${Math.abs(v).toFixed(1)}%)`:v.toFixed(1)+'%';
  const fa2 = v => !isFinite(v)||isNaN(v)?'-':v.toFixed(1)+'%';
  const dc2 = v => v>0.5?'pt-pos':v<-0.5?'pt-neg':'pt-neu';
  const pc2 = v => v>0?'pt-pos':v<0?'pt-neg':'pt-neu';

  // 주차
  const wkLabels = (showFc && wkMeta) ? wkMeta.labels : [];
  const wkCols   = wkLabels.length + (wkLabels.length > 1 ? 1 : 0);

  // 헤더 생성 (구분/채널, 중앙정렬) — colgroup으로 너비 고정
  function makeHeader() {
    // colgroup: 구분, 채널, 연도4, 증감4, 주차N
    let cg = `<colgroup>
      <col class="pc-fix"><col class="pc-ch">
      <col class="pc-num"><col class="pc-num"><col class="pc-num"><col class="pc-num">
      <col class="pc-num"><col class="pc-num"><col class="pc-num"><col class="pc-num">`;
    wkLabels.forEach(() => { cg += `<col class="pc-fw">`; });
    if (wkLabels.length > 1) cg += `<col class="pc-fw">`;
    cg += `</colgroup>`;

    let gh = `<th class="pt-hg-fix" rowspan="2" style="text-align:center">구분</th>
              <th class="pt-hg-fix" rowspan="2" style="text-align:center">채널</th>
              <th class="pt-hg-act24 pt-sep-r" rowspan="2">Y2024<br>실적</th>
              <th class="pt-hg-act25 pt-sep-r" rowspan="2">Y2025<br>실적</th>
              <th class="pt-hg-act26 pt-sep-r" rowspan="2">Y2026<br>실적</th>
              <th class="pt-hg-plan  pt-sep-r" rowspan="2">Y2026<br>계획</th>
              <th class="pt-hg-cmp25 pt-sep-r" colspan="2">25년 대비</th>
              <th class="pt-hg-cmpP ${!showFc?'pt-sep-r':''}" colspan="2">계획 대비</th>`;
    let sh = `<th class="pt-hg-cmp25">증감액</th><th class="pt-hg-cmp25 pt-sep-r">증감율</th>
              <th class="pt-hg-cmpP">증감액</th><th class="pt-hg-cmpP ${!showFc?'pt-sep-r':''}">달성율</th>`;
    if (showFc && wkMeta) {
      gh += `<th class="pt-hg-fc pt-fc-first" colspan="${wkCols}">${wkMeta.mn}월 당월예상</th>`;
      wkLabels.forEach((l,i) => { sh += `<th class="pt-hg-fc ${i===0?'pt-fc-first':''}">${l}</th>`; });
      if (wkLabels.length > 1) sh += `<th class="pt-hg-fc">전주대비</th>`;
    }
    return `${cg}<thead><tr>${gh}</tr><tr>${sh}</tr></thead>`;
  }

  // 주차 셀
  function ptFwCells(obj, isDark=false, bgColor='') {
    if (!showFc || !wkMeta) return '';
    const vals = wkLabels.map((_,i) => { const v=obj[`fw${i+1}`]; return (v!=null&&v!==0)?v:null; });
    const posC   = isDark ? '#fff'    : null;
    const negC   = isDark ? '#fca5a5' : null;
    const neuC   = isDark ? '#bfdbfe' : null;
    const dc3    = v => isDark ? (v>0.5?posC:v<-0.5?negC:neuC) : null;
    const fcCls  = isDark ? '' : 'pt-fc';
    const bgSt   = isDark && bgColor ? `background:${bgColor};` : '';

    let s = vals.map((v, i) => {
      const col = posC ? `color:${posC};` : (v!=null&&v<-0.5?'color:#dc2626;':'');
      return `<td class="${fcCls} ${i===0?'pt-fc-first':''}" style="${bgSt}${col}">${v!=null?fv2(v):'-'}</td>`;
    }).join('');
    if (wkLabels.length > 1) {
      const last=vals[vals.length-1], prev=vals[vals.length-2];
      const diff=(last!=null&&prev!=null)?last-prev:null;
      const col = diff!=null && dc3(diff) ? `color:${dc3(diff)};` : (posC?`color:${posC};`:(diff!=null&&diff<-0.5?'color:#dc2626;':''));
      s += `<td class="${fcCls}" style="${bgSt}${col}">${diff!=null?fd2(diff):'-'}</td>`;
    }
    return s;
  }

  function ptDataCells(r) {
    const dv=deriv(r);
    return `<td class="pt-sep-r">${fv2(r.y2024)}</td><td class="pt-sep-r">${fv2(r.y2025)}</td><td class="pt-sep-r">${fv2(r.actual)}</td><td class="pt-sep-r">${fv2(r.plan)}</td>
            <td class="${dc2(dv.vs25)}">${fd2(dv.vs25)}</td><td class="${pc2(dv.vs25p)} pt-sep-r">${fp2(dv.vs25p)}</td>
            <td class="${dc2(dv.dPlan)}">${fd2(dv.dPlan)}</td><td class="${!showFc?'pt-sep-r':''}">${fa2(dv.achPct)}</td>${ptFwCells(r)}`;
  }
  function ptAggCells(t, isDark=false, bgColor='') {
    const dv=deriv(t);
    const posC = isDark ? '#fff' : '#1a1a2e';
    const negC = isDark ? '#fca5a5' : '#dc2626';
    const neuC = isDark ? '#bfdbfe' : '#9ca3af';
    const dc3 = v => v > 0.5 ? posC : v < -0.5 ? negC : neuC;
    return `<td class="pt-sep-r" style="color:${posC}">${fv2(t.y2024)}</td>
            <td class="pt-sep-r" style="color:${posC}">${fv2(t.y2025)}</td>
            <td class="pt-sep-r" style="color:${posC}">${fv2(t.actual)}</td>
            <td class="pt-sep-r" style="color:${posC}">${fv2(t.plan)}</td>
            <td style="color:${dc3(dv.vs25)}">${fd2(dv.vs25)}</td>
            <td class="pt-sep-r" style="color:${dc3(dv.vs25p)}">${fp2(dv.vs25p)}</td>
            <td style="color:${dc3(dv.dPlan)}">${fd2(dv.dPlan)}</td>
            <td class="${!showFc?'pt-sep-r':''}" style="color:${posC}">${fa2(dv.achPct)}</td>${ptFwCells(t, isDark, bgColor)}`;
  }

  // 팀별 그룹핑
  const teamGroups = {};
  for (const r of allRows) {
    if (!teamGroups[r.team]) teamGroups[r.team]=[];
    teamGroups[r.team].push(r);
  }
  const teamsInData = SALES_TEAMS_ORDER.filter(t=>teamGroups[t]);

  // BU 바디 생성
  function makeBuBody(buList) {
    let body = '';
    const bk=['y2024','y2025b','y2025','plan','actual','fw1','fw2','fw3','fw4','fw5'];
    const mkE=()=>Object.fromEntries(bk.map(k=>[k,0]));
    buList.forEach(bu => {
      const buTeams = teamsInData.filter(t=>BU_MAP[t]===bu);
      if (!buTeams.length) return;
      const bt = mkE();
      const totalCols = 10 + wkCols;
      body += `<tr><td colspan="${totalCols}" style="background:#2e6da4;color:#fff;font-weight:700;text-align:left;padding:3px 6px;font-size:8.5px">◆ ${bu}</td></tr>`;
      buTeams.forEach(team => {
        const chRows = teamGroups[team];
        const tt = TT[team];
        chRows.forEach((r,ri) => {
          body += `<tr>
            ${ri===0?`<td class="pt-team" rowspan="${chRows.length}" style="vertical-align:middle;text-align:center">${team}</td>`:''}
            <td class="pt-ch" style="text-align:left">${r.channel}</td>${ptDataCells(r)}
          </tr>`;
        });
        body += `<tr class="pt-sub">
          <td class="pt-team" style="text-align:center">${team}</td>
          <td style="text-align:center;font-weight:700">소 계</td>${ptAggCells(tt)}
        </tr>`;
        bk.forEach(k=>{bt[k]+=(tt[k]||0);});
      });
      body += `<tr style="background:#2e6da4;color:#fff;font-weight:700">
        <td style="text-align:center;padding:2px 4px">${bu}</td><td style="text-align:center;font-weight:700;padding:2px 4px">합 계</td>${ptAggCells(bt, true, '#2e6da4')}
      </tr>`;
    });
    return body;
  }

  const hdr = makeHeader();

  // 1페이지: 동북아 BU
  const body1 = makeBuBody(['동북아 BU']);
  const tbl1 = document.getElementById('printTable');
  if (tbl1) tbl1.innerHTML = `${hdr}<tbody>${body1}</tbody>`;

  // 2페이지: Global BU + Expert BU + 합계
  const body2 = makeBuBody(['Global BU','Expert BU'])
    + `<tr class="pt-grand"><td style="text-align:center;font-weight:700">🏁 전 사</td><td style="text-align:center;font-weight:700">합 계</td>${ptAggCells(G, true, '#1e3a5f')}</tr>`;
  const tbl2 = document.getElementById('printTable2');
  if (tbl2) tbl2.innerHTML = `${hdr}<tbody>${body2}</tbody>`;
}

/* ─── 씨엠에스랩 로고 클릭: 초기 상태로 리셋 ─── */
let _INITIAL_SNAPSHOT = null;
function captureInitialSnapshot() {
  _INITIAL_SNAPSHOT = {
    teams:    new Set(S.teams),
    years:    new Set(S.years),
    months:   new Set(S.months),
    channels: new Set(S.channels),
    brands:   new Set(S.brands),
    forecast: !!document.getElementById('chkFc')?.checked,
  };
}
function resetToInitial() {
  if (!_INITIAL_SNAPSHOT) return;
  // 상태 복원
  S.teams    = new Set(_INITIAL_SNAPSHOT.teams);
  S.years    = new Set(_INITIAL_SNAPSHOT.years);
  S.months   = new Set(_INITIAL_SNAPSHOT.months);
  S.channels = new Set(_INITIAL_SNAPSHOT.channels);
  S.brands   = new Set(_INITIAL_SNAPSHOT.brands);

  // 버튼 active 클래스 재설정
  document.querySelectorAll('[data-f="team"]').forEach(b =>
    b.classList.toggle('active', S.teams.has(b.dataset.v)));
  document.querySelectorAll('[data-f="year"]').forEach(b =>
    b.classList.toggle('active', S.years.has(b.dataset.v)));
  syncMonthButtons(); // 월/분기 버튼 동기화

  // 채널 드롭다운 초기화
  const chSel = document.getElementById('chSel');
  if (chSel) chSel.value = 'ALL';
  updateChannelDropdown();

  // 브랜드 드롭다운 초기화
  const brSel = document.getElementById('brSel');
  if (brSel) brSel.value = 'ALL';

  // 예상표기 체크박스 초기값 복원
  const chkFc = document.getElementById('chkFc');
  if (chkFc) chkFc.checked = _INITIAL_SNAPSHOT.forecast;

  // 매출 대시보드 탭으로 이동
  switchTab('dashboard');

  // 재렌더 + 정렬
  run();
  alignQuarterToYear();

  // 스크롤 최상단
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function switchTab(name) {
  const dashPane  = document.getElementById('pane-dashboard');
  const tablePane = document.getElementById('pane-table');
  const printPane = document.getElementById('pane-print');
  const kpi       = document.querySelector('.sticky-kpi');

  document.querySelectorAll('.hdr-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + name)?.classList.add('active');

  // 모두 숨기기
  dashPane.style.display = 'none';
  tablePane.classList.remove('active');
  printPane.classList.remove('active');
  if (kpi) kpi.style.display = 'none';

  if (name === 'dashboard') {
    dashPane.style.display = '';
    if (kpi) kpi.style.display = '';

  } else if (name === 'table') {
    tablePane.classList.add('active');
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const hdrH    = document.querySelector('.sticky-header')?.offsetHeight    || 44;
      const filterH = document.querySelector('.sticky-filterbar')?.offsetHeight || 70;
      tablePane.style.top = (hdrH + filterH + 8) + 'px';
    }));
    run();

  } else if (name === 'print') {
    printPane.classList.add('active');
    // 주간회의(인쇄) 탭 진입 시 예상표기 자동 체크
    const chkFc = document.getElementById('chkFc');
    if (chkFc && !chkFc.checked) {
      chkFc.checked = true;
      S.forecast = true;
    }
    requestAnimationFrame(() => requestAnimationFrame(() => {
      const hdrH    = document.querySelector('.sticky-header')?.offsetHeight    || 44;
      const filterH = document.querySelector('.sticky-filterbar')?.offsetHeight || 70;
      printPane.style.top = (hdrH + filterH) + 'px';
    }));
    buildPrintTable();
  }
}
let _cumBarExpanded = false;
let _overlayEl = null;

function getCumBarExpandRect() {
  const stickyH = (document.querySelector('.sticky-header')?.offsetHeight   || 44)
                + (document.querySelector('.sticky-filterbar')?.offsetHeight || 70)
                + (document.querySelector('.sticky-kpi')?.offsetHeight       || 56);
  const pad = 10;
  return {
    top:  stickyH + pad,
    left: pad,
    w:    window.innerWidth  - pad * 2,
    h:    window.innerHeight - stickyH - pad * 2
  };
}

function applyCumBarExpand() {
  const card = document.getElementById('cumBarCard');
  if (!card) return;
  const r = getCumBarExpandRect();
  card.style.top    = r.top  + 'px';
  card.style.left   = r.left + 'px';
  card.style.width  = r.w    + 'px';
  card.style.height = r.h    + 'px';
}

function toggleCumBarExpand() {
  const card = document.getElementById('cumBarCard');
  const btn  = document.getElementById('cumBarExpandBtn');
  if (!card) return;

  if (!_cumBarExpanded) {
    card.classList.add('chart-expanded');
    applyCumBarExpand();

    _overlayEl = document.createElement('div');
    _overlayEl.className = 'chart-overlay-bg';
    _overlayEl.onclick = toggleCumBarExpand;
    document.body.appendChild(_overlayEl);

    btn.textContent = '✕';
    btn.title = '축소';
    _cumBarExpanded = true;
  } else {
    card.classList.remove('chart-expanded');
    card.style.top = card.style.left = card.style.width = card.style.height = '';
    if (_overlayEl) { _overlayEl.remove(); _overlayEl = null; }
    btn.textContent = '⛶';
    btn.title = '확대';
    _cumBarExpanded = false;
  }

  setTimeout(() => { if (CHARTS.cumBar) { CHARTS.cumBar.resize(); renderCumBar(); } }, 280);
}

/* ─── 월별 매출추이 확대/축소 ─── */
let _mainExpanded   = false;
let _mainOverlay    = null;

function toggleMainExpand() {
  const card = document.getElementById('mainCard');
  const btn  = document.getElementById('mainExpandBtn');
  if (!card) return;
  if (!_mainExpanded) {
    card.classList.add('chart-expanded');
    applyExpandCard(card);
    _mainOverlay = document.createElement('div');
    _mainOverlay.className = 'chart-overlay-bg';
    _mainOverlay.onclick = toggleMainExpand;
    document.body.appendChild(_mainOverlay);
    btn.textContent = '✕'; btn.title = '축소';
    _mainExpanded = true;
  } else {
    card.classList.remove('chart-expanded');
    card.style.top = card.style.left = card.style.width = card.style.height = '';
    if (_mainOverlay) { _mainOverlay.remove(); _mainOverlay = null; }
    btn.textContent = '⛶'; btn.title = '확대';
    _mainExpanded = false;
  }
  setTimeout(() => { if (CHARTS.main) CHARTS.main.resize(); }, 280);
}

/* ─── 누적 매출액 추이 확대/축소 ─── */
let _cumLineExpanded = false;
let _cumLineOverlay  = null;

function toggleCumLineExpand() {
  const card = document.getElementById('cumLineCard');
  const btn  = document.getElementById('cumLineExpandBtn');
  if (!card) return;
  if (!_cumLineExpanded) {
    card.classList.add('chart-expanded');
    applyExpandCard(card);
    _cumLineOverlay = document.createElement('div');
    _cumLineOverlay.className = 'chart-overlay-bg';
    _cumLineOverlay.onclick = toggleCumLineExpand;
    document.body.appendChild(_cumLineOverlay);
    btn.textContent = '✕'; btn.title = '축소';
    _cumLineExpanded = true;
  } else {
    card.classList.remove('chart-expanded');
    card.style.top = card.style.left = card.style.width = card.style.height = '';
    if (_cumLineOverlay) { _cumLineOverlay.remove(); _cumLineOverlay = null; }
    btn.textContent = '⛶'; btn.title = '확대';
    _cumLineExpanded = false;
  }
  setTimeout(() => { if (CHARTS.cumLine) CHARTS.cumLine.resize(); }, 280);
}

/* ─── 월별 매출추이 주석 스트립 빌더 ─── */
function buildMainAnnotStrip() {
  const strip = document.getElementById('mainAnnotStrip');
  if (!strip) return;
  const maxMon = Math.max(...activeMonths());

  const fmtPlan = (v, r) => (!r || !v) ? '-' : (v / r * 100).toFixed(1) + '%';
  const fmtYoy  = (v, r) => {
    if (!r || !v || Math.abs(v - r) < 0.5) return '-';
    const p = (v - r) / r * 100;
    return (p > 0 ? '▲' : '▼') + Math.abs(p).toFixed(1) + '%';
  };
  const colPlan = (v, r) => (!r || !v) ? '#9ca3af' : v >= r - 0.5 ? '#375623' : '#C55A11';
  const colYoy  = (v, r) => (!r || !v || Math.abs(v - r) < 0.5) ? '#9ca3af'
    : v > r ? '#375623' : '#C55A11';

  let planCells = '', yoyCells = '';
  for (let m = 1; m <= 12; m++) {
    const d = _mainAgg[m];
    const ok = d && m <= maxMon;
    const pc = ok ? colPlan(d.actual, d.plan)  : '#9ca3af';
    const yc = ok ? colYoy(d.actual,  d.y2025) : '#9ca3af';
    const pt = ok ? fmtPlan(d.actual, d.plan)  : '-';
    const yt = ok ? fmtYoy(d.actual,  d.y2025) : '-';
    planCells += `<td style="color:${pc}">${pt}</td>`;
    yoyCells  += `<td style="color:${yc}">${yt}</td>`;
  }
  strip.innerHTML = `<table class="annot-strip-tbl">
    <tr><th>계획달성</th>${planCells}</tr>
    <tr><th>전기대비</th>${yoyCells}</tr>
  </table>`;
}

/* ─── 누적 달성율 주석 스트립 빌더 ─── */
function buildCumAnnotStrip() {
  const strip = document.getElementById('cumAnnotStrip');
  if (!strip) return;
  const maxMon = Math.max(...activeMonths());
  let cumA = 0, cumP = 0, cells = '';
  for (let m = 1; m <= 12; m++) {
    const d = _mainAgg[m];
    if (d && m <= maxMon) { cumA += d.actual || 0; cumP += d.plan || 0; }
    const pct = (d && m <= maxMon && cumP) ? cumA / cumP * 100 : null;
    const txt = pct !== null ? pct.toFixed(1) + '%' : '-';
    const col = pct === null ? '#9ca3af' : pct >= 100 ? '#375623' : pct >= 80 ? '#C55A11' : '#dc2626';
    cells += `<td style="color:${col}">${txt}</td>`;
  }
  strip.innerHTML = `<table class="annot-strip-tbl">
    <tr><th>누적달성율</th>${cells}</tr>
  </table>`;
}

/* ─── 요약표 토글 ─── */
function toggleStbl(btn, wrapId) {
  const wrap = document.getElementById(wrapId);
  const isOpen = wrap.classList.toggle('open');
  btn.classList.toggle('open', isOpen);
  btn.querySelector('span:first-child').textContent = isOpen ? '요약표 닫기' : '요약표';
}

/* ═══════════════════════════════════════════════════════
   차트 하단 요약 테이블 빌드
   · tblSummary1 : 열=12개월, 행=4개 연도  (월별 금액)
   · tblSummary2 : 열=12개월, 행=4개 연도  (누적 금액)
   · tblSummary3 : 전사→열=팀, 팀→열=채널 (선택 월 합산)
═══════════════════════════════════════════════════════ */
function buildSummaryTables() {
  const agg  = chartData12months(null);
  const mons = activeMonths();
  const isAll = S.teams.has('ALL');

  const fE = v => {
    if (!v) return '-';
    const e = v / 100;
    return (Math.abs(e) >= 10 ? Math.round(e).toLocaleString('ko-KR') : e.toFixed(1)) + '억';
  };

  const KEYS = ['y2024','y2025','actual','plan'];
  const LBLS = ['24년 실적','25년 실적','26년 실적','26년 계획'];
  const CLS  = ['c24','c25','c26','cp'];
  const yearActive = key => ({ y2024: S.years.has('2024'), y2025: S.years.has('2025'), actual: S.years.has('2026'), plan: S.years.has('plan') })[key] ?? true;

  /* 공통: th 셀 스타일 */
  const thStyle    = 'padding:5px 10px;text-align:center;font-size:10.5px;font-weight:700;color:var(--muted);background:#fafafa;border-bottom:1px solid var(--border);border-right:1px solid var(--border);white-space:nowrap;';
  const thLblStyle = 'padding:5px 8px;text-align:center;font-size:10.5px;font-weight:700;color:var(--muted);background:#fafafa;border-bottom:1px solid var(--border);border-right:2px solid var(--border);white-space:nowrap;min-width:54px;width:54px;';
  const tdStyle    = (cls, hi) => `padding:5px 10px;text-align:right;font-size:10.5px;border-bottom:1px solid #f3f4f6;border-right:1px solid var(--border);white-space:nowrap;${hi?'background:#fde68a;font-weight:700;':''}`;
  const tdLblStyle = 'padding:5px 8px;text-align:center;font-size:10.5px;font-weight:700;background:#fafafa;border-bottom:1px solid #f3f4f6;border-right:2px solid var(--border);white-space:nowrap;min-width:54px;width:54px;';

  /* ── tblSummary1: 월별 금액 — 열=월, 행=연도 ── */
  (() => {
    const el = document.getElementById('tblSummary1');
    if (!el) return;
    const colW1 = 72;
    const colgroup1 = `<colgroup><col style="width:54px"><col span="12" style="width:${colW1}px"></colgroup>`;
    const hdr = `<tr><th style="${thLblStyle}">구분</th>` +
      MLABEL.map((m,i) => `<th style="${thStyle}${mons.includes(i+1)?'color:#1a56a0;':''}">${m}</th>`).join('') + '</tr>';
    const body = KEYS.map((key,ki) =>
      `<tr><td style="${tdLblStyle}" class="${CLS[ki]}">${LBLS[ki]}</td>` +
      [1,2,3,4,5,6,7,8,9,10,11,12].map((m,i) => {
        const hi = mons.includes(i+1) && yearActive(key);
        return `<td class="${CLS[ki]}" style="${tdStyle(CLS[ki],hi)}">${fE(agg[m][key]||0)}</td>`;
      }).join('') + '</tr>'
    ).join('');
    el.innerHTML = `<table style="border-collapse:collapse;font-size:11px;table-layout:fixed;width:100%">${colgroup1}<thead>${hdr}</thead><tbody>${body}</tbody></table>`;
    el.style.overflowX = 'auto';
  })();

  /* ── tblSummary2: 누적 금액 — 열=월, 행=연도 ── */
  (() => {
    const el = document.getElementById('tblSummary2');
    if (!el) return;
    const colW2 = 72;
    const colgroup2 = `<colgroup><col style="width:54px"><col span="12" style="width:${colW2}px"></colgroup>`;
    const hdr = `<tr><th style="${thLblStyle}">구분</th>` +
      MLABEL.map((m,i) => `<th style="${thStyle}${mons.includes(i+1)?'color:#1a56a0;':''}">${m}</th>`).join('') + '</tr>';
    const body = KEYS.map((key,ki) => {
      let cum = 0;
      const cells = [1,2,3,4,5,6,7,8,9,10,11,12].map((m,i) => {
        cum += agg[m][key] || 0;
        const hi = mons.includes(i+1) && yearActive(key);
        return `<td class="${CLS[ki]}" style="${tdStyle(CLS[ki],hi)}">${fE(cum)}</td>`;
      }).join('');
      return `<tr><td style="${tdLblStyle}" class="${CLS[ki]}">${LBLS[ki]}</td>${cells}</tr>`;
    }).join('');
    el.innerHTML = `<table style="border-collapse:collapse;font-size:11px;table-layout:fixed;width:100%">${colgroup2}<thead>${hdr}</thead><tbody>${body}</tbody></table>`;
    el.style.overflowX = 'auto';
  })();

  /* ── tblSummary3: 전사→팀별, 팀→채널별 ── */
  (() => {
    const el = document.getElementById('tblSummary3');
    if (!el) return;

    // 팀별/채널별 요약은 10억 초과 시에도 소수점 첫째자리까지 표기
    const fE = v => {
      if (!v) return '-';
      const e = v / 100;
      return e.toLocaleString('ko-KR', {minimumFractionDigits:1, maximumFractionDigits:1}) + '억';
    };

    const selTeams = isAll ? null : [...S.teams];

    // 긴 레이블 줄바꿈
    const LABEL_BREAKS = {
      '동북아MC팀':'동북아<br>MC팀','Global사업팀':'Global<br>사업팀',
      '해외_오프라인_미국':'해외_오프라인<br>_미국','해외_오프라인_CIS':'해외_오프라인<br>_CIS',
      '해외_오프라인_유럽':'해외_오프라인<br>_유럽','해외_오프라인_동남아':'해외_오프라인<br>_동남아',
      '해외_오프라인_중동':'해외_오프라인<br>_중동','해외_오프라인_기타':'해외_오프라인<br>_기타',
      '해외_오프라인_대만':'해외_오프라인<br>_대만','해외_오프라인(일본)':'해외_오프라인<br>(일본)',
      '해외_온라인 기타':'해외_온라인<br>기타','큐텐/라쿠텐':'큐텐/<br>라쿠텐',
      '중국법인매출':'중국법인<br>매출','중국법인수출':'중국법인<br>수출',
    };
    const breakLabel = s => {
      if (LABEL_BREAKS[s]) return LABEL_BREAKS[s];
      if (s.length <= 5) return s;
      const uIdx = s.lastIndexOf('_');
      if (uIdx > 0 && uIdx < s.length - 1) return s.slice(0, uIdx) + '<br>' + s.slice(uIdx);
      const spIdx = s.indexOf(' ');
      if (spIdx > 0) return s.slice(0, spIdx) + '<br>' + s.slice(spIdx + 1);
      const mid = Math.ceil(s.length / 2);
      return s.slice(0, mid) + '<br>' + s.slice(mid);
    };

    const SLOTS = 8;
    const wrapW = el.parentElement.clientWidth || 600;
    const lblColW = 54;
    const colW = Math.floor((wrapW - lblColW) / SLOTS);
    const thDataStyle = thStyle + `width:${colW}px;min-width:${colW}px;`;
    const tdDataStyle = (cls, hi) => tdStyle(cls, hi) + `width:${colW}px;min-width:${colW}px;`;

    const mkTable = (labels, hdr, body) => {
      const needScroll = labels.length > SLOTS;
      const tableW = lblColW + colW * Math.max(labels.length, SLOTS);
      const colgroup = `<colgroup><col style="width:${lblColW}px;min-width:${lblColW}px">` +
        labels.map(() => `<col style="width:${colW}px;min-width:${colW}px">`).join('') +
        (labels.length < SLOTS ? `<col span="${SLOTS-labels.length}" style="width:${colW}px">` : '') +
        `</colgroup>`;
      el.innerHTML = `<table style="border-collapse:collapse;font-size:11px;table-layout:fixed;width:${needScroll?tableW+'px':'100%'}">${colgroup}<thead>${hdr}</thead><tbody>${body}</tbody></table>`;
      el.style.overflowX = needScroll ? 'auto' : 'hidden';
    };

    const wkMeta = getWeeklyMeta();

    if (wkMeta) {
      /* ── 주차별 예상실적 뷰 ── */
      const { mn, labels: wkLabels } = wkMeta;

      // 팀/채널별 주차 합산
      const wkMap = {};
      for (const r of RAW) {
        if (r.month !== mn) continue;
        if (!isAll && !selTeams.includes(r.team)) continue;
        if (!S.brands.has('ALL') && !S.brands.has(r.brand)) continue;
        const key = isAll ? r.team : r.channel;
        if (!wkMap[key]) wkMap[key] = {};
        for (let w = 1; w <= wkLabels.length; w++) {
          const v = r[`fw${w}`];
          if (v != null) wkMap[key][`fw${w}`] = (wkMap[key][`fw${w}`] || 0) + v;
        }
      }

      const colLabels = isAll
        ? SALES_TEAMS_ORDER.filter(t => wkMap[t])
        : chSort(Object.keys(wkMap));

      // 데이터 있는 주차만 추출
      const activeWks = [];
      for (let w = 1; w <= wkLabels.length; w++) {
        if (colLabels.some(l => wkMap[l]?.[`fw${w}`] != null)) activeWks.push(w);
      }

      if (!activeWks.length || !colLabels.length) {
        el.innerHTML = '<div style="padding:12px;text-align:center;color:#9ca3af;font-size:11px">예상 데이터 없음</div>';
        return;
      }

      // 포맷 헬퍼 (10억 초과여도 소수점 첫째자리 유지)
      const fmtDiff = v => {
        if (v == null) return '-';
        const e = v / 100;
        const s = Math.abs(e).toLocaleString('ko-KR', {minimumFractionDigits:1, maximumFractionDigits:1}) + '억';
        return (v >= 0 ? '▲' : '▼') + s;
      };
      const diffCol = v => v == null ? '#9ca3af' : v >= 0.5 ? '#375623' : v <= -0.5 ? '#dc2626' : '#9ca3af';
      const emptyTd = colLabels.length < SLOTS
        ? `<td colspan="${SLOTS-colLabels.length}" style="border-bottom:1px solid #f3f4f6"></td>` : '';

      const hdr = `<tr><th style="${thLblStyle}">${mn}월 예상</th>` +
        colLabels.map(l => `<th style="${thDataStyle}">${breakLabel(l)}</th>`).join('') +
        (colLabels.length < SLOTS ? `<th colspan="${SLOTS-colLabels.length}" style="${thStyle}border-right:none"></th>` : '') +
        '</tr>';

      // 행: 각 주차 + 전주대비
      let body = '';
      activeWks.forEach(w => {
        const lbl = `${mn}월 ${wkLabels[w-1]}`;
        const cells = colLabels.map(l => {
          const v = wkMap[l]?.[`fw${w}`] ?? null;
          return `<td style="${tdDataStyle('',false)}">${v != null ? fE(v) : '-'}</td>`;
        }).join('');
        body += `<tr><td style="${tdLblStyle}">${lbl}</td>${cells}${emptyTd}</tr>`;
      });
      if (activeWks.length >= 2) {
        const w2 = activeWks[activeWks.length-1];
        const w1 = activeWks[activeWks.length-2];
        const cells = colLabels.map(l => {
          const v2 = wkMap[l]?.[`fw${w2}`] ?? null;
          const v1 = wkMap[l]?.[`fw${w1}`] ?? null;
          const diff = (v2 != null && v1 != null) ? v2 - v1 : null;
          return `<td style="${tdDataStyle('',false)}color:${diffCol(diff)}">${fmtDiff(diff)}</td>`;
        }).join('');
        body += `<tr><td style="${tdLblStyle}color:#6b7280">전주대비</td>${cells}${emptyTd}</tr>`;
      }

      mkTable(colLabels, hdr, body);

    } else {
      /* ── 기존 연도별 뷰 ── */
      const map = {};
      for (const r of RAW) {
        if (!mons.includes(r.month)) continue;
        if (!isAll && !selTeams.includes(r.team)) continue;
        if (!S.brands.has('ALL') && !S.brands.has(r.brand)) continue;
        const key = isAll ? r.team : r.channel;
        if (!map[key]) map[key] = {y2024:0,y2025:0,plan:0,actual:0};
        map[key].y2024  += r.y2024  || 0;
        map[key].y2025  += r.y2025  || 0;
        map[key].plan   += r.plan   || 0;
        map[key].actual += r.actual || 0;
      }
      const labels = isAll
        ? SALES_TEAMS_ORDER.filter(t => map[t] && (map[t].y2024+map[t].y2025+map[t].plan+map[t].actual) > 0)
        : chSort(Object.keys(map).filter(k => (map[k].y2024+map[k].y2025+map[k].plan+map[k].actual) > 0));

      const hdr = `<tr><th style="${thLblStyle}">구분</th>` +
        labels.map(l => `<th style="${thDataStyle}">${breakLabel(l)}</th>`).join('') +
        (labels.length < SLOTS ? `<th colspan="${SLOTS-labels.length}" style="${thStyle}border-right:none"></th>` : '') +
        '</tr>';
      const body = KEYS.map((key,ki) =>
        `<tr><td style="${tdLblStyle}" class="${CLS[ki]}">${LBLS[ki]}</td>` +
        labels.map(l => `<td class="${CLS[ki]}" style="${tdDataStyle(CLS[ki],false)}">${fE(map[l]?.[key]||0)}</td>`).join('') +
        (labels.length < SLOTS ? `<td colspan="${SLOTS-labels.length}" style="border-bottom:1px solid #f3f4f6"></td>` : '') +
        '</tr>'
      ).join('');

      mkTable(labels, hdr, body);
    }
  })();
}

/* ═══════════════════════════════════════════════════════
   MAIN RUN
═══════════════════════════════════════════════════════ */
function run() {
  S.forecast = document.getElementById('chkFc')?.checked || false;
  S.search   = document.getElementById('tblSearch')?.value.trim().toLowerCase() || '';

  const rows = aggregate(null);
  buildKPI(rows);
  buildTable(rows);
  updateCharts();
  buildSummaryTables();
  if (_mainExpanded)    applyExpandCard(document.getElementById('mainCard'));
  if (_cumLineExpanded) applyExpandCard(document.getElementById('cumLineCard'));
  if (_cumBarExpanded)  applyCumBarExpand();

  // 표 탭 활성화 중이면 top 재계산
  const paneTable = document.getElementById('pane-table');
  if (paneTable?.classList.contains('active')) {
    requestAnimationFrame(() => {
      const hdrH    = document.querySelector('.sticky-header')?.offsetHeight    || 44;
      const filterH = document.querySelector('.sticky-filterbar')?.offsetHeight || 70;
      paneTable.style.top = (hdrH + filterH + 8) + 'px';
    });
  }
  // 인쇄 탭 활성화 중이면 재빌드
  const panePrint = document.getElementById('pane-print');
  if (panePrint?.classList.contains('active')) buildPrintTable();

  const tl = S.teams.has('ALL') ? '전사' : [...S.teams].join(', ');
  const allCh = S.channels.has('ALL');
  const chLbl = allCh ? null : [...S.channels][0];
  document.getElementById('chartMeta1').textContent = chLbl ? `${chLbl} 기준` : `${tl} 기준`;
  document.getElementById('chartMeta2').textContent = chLbl ? `${chLbl} 기준` : `${tl} 기준`;
  document.getElementById('chartMeta3').textContent = `${tl} 기준`;

  // 필터 상태 배지 업데이트
  const chBadgeLbl = allCh ? '전체채널' : chLbl;
  const mons = activeMonths();
  const isContiguous = mons.every((m, i) => i === 0 || m === mons[i-1] + 1);
  const monBadgeLbl = S.months.has('YTD') ? 'YTD'
    : mons.length === 1 ? `${mons[0]}월`
    : isContiguous ? `${mons[0]}월 ~ ${mons[mons.length-1]}월`
    : mons.map(m => `${m}월`).join(', ');
  const b1 = document.getElementById('filterBadge1');
  const b2 = document.getElementById('filterBadge2');
  if (b1) b1.textContent = `${tl}  &  ${chBadgeLbl}`;
  if (b2) b2.textContent = monBadgeLbl;
}

let _raf = null;
window.addEventListener('resize', () => {
  cancelAnimationFrame(_raf);
  _raf = requestAnimationFrame(() => {
    if (CHARTS.main) Object.values(CHARTS).forEach(c => c?.resize?.());
    if (_mainExpanded)    applyExpandCard(document.getElementById('mainCard'));
    if (_cumLineExpanded) applyExpandCard(document.getElementById('cumLineCard'));
    if (_cumBarExpanded)  applyCumBarExpand();
  });
});
function freezeChannelWidth() {
  const sel = document.getElementById('chSel');
  if (!sel) return;
  sel.style.width = 'auto';
  const w = sel.getBoundingClientRect().width;
  if (w > 0) sel.style.width = w + 'px';
}
document.addEventListener('DOMContentLoaded', () => {
  updateChannelDropdown();
  initCharts();
  run();
  alignQuarterToYear();
  setTimeout(freezeChannelWidth, 80);
  captureInitialSnapshot();
});
"""

HTML_SHELL = """<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=1280, initial-scale=1.0, shrink-to-fit=yes">
<title>매출 대시보드 | CMS Lab</title>
<style>
__CSS__
</style>
</head>
<body>

<!-- ═══════ HEADER ═══════ -->
<div class="sticky-header">
<div class="header" style="background:#fff;padding:0 clamp(12px,2vw,24px)">
  <div class="hdr-company" style="color:#833C0C;cursor:pointer" onclick="resetToInitial()" title="초기 화면으로">씨엠에스랩</div>
  <div class="hdr-divider" style="background:#e5e7eb"></div>
  <div class="hdr-tabs">
    <button class="hdr-tab active" id="tab-dashboard" onclick="switchTab('dashboard')">📊 매출 대시보드</button>
    <button class="hdr-tab" id="tab-table" onclick="switchTab('table')">📋 매출현황(표)</button>
    <button class="hdr-tab" id="tab-print" onclick="switchTab('print')">🖨️ 주간회의(인쇄)</button>
  </div>
  <div class="hdr-right">
    <strong>__BASE_DATE__ 기준</strong>
  </div>
</div>
</div>

<!-- ═══════ FILTER BAR ═══════ -->
<div class="sticky-filterbar">
<div class="filterbar">

  <!-- 1행: 팀 | 연도 | 브랜드 | 채널 -->
  <div class="filter-row">
    <div class="fg team-fg">
      <span class="fg-label toggleable" id="teamLabel" onclick="toggleSingleMode('team')" title="클릭 시 단일 선택 모드 전환">팀</span>
      <button class="pb active" data-f="team" data-v="ALL"    onclick="togTeam(this)">전사</button>
      <span class="fg-sep">ㅣ</span>
      <button class="pb" data-f="team" data-v="RBD1팀"        onclick="togTeam(this)">RBD1팀</button>
      <button class="pb" data-f="team" data-v="RBD2팀"        onclick="togTeam(this)">RBD2팀</button>
      <button class="pb" data-f="team" data-v="일본사업팀"     onclick="togTeam(this)">일본사업팀</button>
      <button class="pb" data-f="team" data-v="중국사업팀"     onclick="togTeam(this)">중국사업팀</button>
      <button class="pb" data-f="team" data-v="동북아MC팀"    onclick="togTeam(this)">동북아MC팀</button>
      <span class="fg-sep bu-sep">ㅣ</span>
      <button class="pb" data-f="team" data-v="Global사업팀"   onclick="togTeam(this)">글로벌사업팀</button>
      <button class="pb" data-f="team" data-v="GEC팀"         onclick="togTeam(this)">GEC팀</button>
      <span class="fg-sep bu-sep">ㅣ</span>
      <button class="pb" data-f="team" data-v="메디컬팀"       onclick="togTeam(this)">메디컬팀</button>
    </div>
    <div class="fg">
      <span class="fg-label" id="yearGroup">연도</span>
      <button class="pb active" data-f="year" data-v="2024" onclick="togYear(this)">24년 실적</button>
      <button class="pb active" data-f="year" data-v="2025" onclick="togYear(this)">25년 실적</button>
      <button class="pb active" data-f="year" data-v="2026" onclick="togYear(this)">26년 실적</button>
      <button class="pb active" data-f="year" data-v="plan" onclick="togYear(this)">26년 계획</button>
    </div>
    <div class="fg">
      <span class="fg-label">브랜드</span>
      <select id="brSel" class="ch-select" style="width:74px;padding-right:20px" onchange="togBrand()">
        <option value="ALL">전체</option>
        <option value="CFC">CFC</option>
        <option value="CEX">CEX</option>
        <option value="DMB">DMB</option>
        <option value="SUS">SUS</option>
        <option value="기타">기타</option>
      </select>
    </div>
    <div class="fg">
      <span class="fg-label">채널</span>
      <select id="chSel" class="ch-select" onchange="togChannel()">
        <option value="ALL">전체 채널</option>
      </select>
    </div>
  </div>

  <!-- 2행: 월 | 분기 | 예상표기 -->
  <div class="filter-row" style="border-top:1px solid rgba(255,255,255,.15);padding-top:5px;margin-top:2px">
    <div class="fg">
      <span class="fg-label toggleable" id="monthLabel" onclick="toggleSingleMode('month')" title="클릭 시 단일 선택 모드 전환">월</span>
      <button class="pb active" data-f="month" data-v="YTD" onclick="togMonth(this)">YTD</button>
      <span class="fg-sep">ㅣ</span>
      <button class="pb" data-f="month" data-v="1"  onclick="togMonth(this)">1월</button>
      <button class="pb" data-f="month" data-v="2"  onclick="togMonth(this)">2월</button>
      <button class="pb" data-f="month" data-v="3"  onclick="togMonth(this)">3월</button>
      <button class="pb" data-f="month" data-v="4"  onclick="togMonth(this)">4월</button>
      <button class="pb" data-f="month" data-v="5"  onclick="togMonth(this)">5월</button>
      <button class="pb" data-f="month" data-v="6"  onclick="togMonth(this)">6월</button>
      <button class="pb" data-f="month" data-v="7"  onclick="togMonth(this)">7월</button>
      <button class="pb" data-f="month" data-v="8"  onclick="togMonth(this)">8월</button>
      <button class="pb" data-f="month" data-v="9"  onclick="togMonth(this)">9월</button>
      <button class="pb" data-f="month" data-v="10" onclick="togMonth(this)">10월</button>
      <button class="pb" data-f="month" data-v="11" onclick="togMonth(this)">11월</button>
      <button class="pb" data-f="month" data-v="12" onclick="togMonth(this)">12월</button>
    </div>
    <div class="fg">
      <span class="fg-label toggleable" id="quarterGroup" onclick="toggleSingleMode('quarter')" title="클릭 시 단일 선택 모드 전환">분기</span>
      <button class="pb" data-f="quarter" data-v="Q1" onclick="togQuarter(this)">1Q</button>
      <button class="pb" data-f="quarter" data-v="Q2" onclick="togQuarter(this)">2Q</button>
      <button class="pb" data-f="quarter" data-v="Q3" onclick="togQuarter(this)">3Q</button>
      <button class="pb" data-f="quarter" data-v="Q4" onclick="togQuarter(this)">4Q</button>
    </div>
    <div class="fg">
      <label class="fck" id="fcLabel">
        <input type="checkbox" id="chkFc" onchange="run()" checked>
        <span>예상표기</span>
      </label>
      <span id="fcNote" style="font-size:10px;color:rgba(255,255,255,.55)"></span>
    </div>
  </div>

</div>
</div>

<!-- ═══════ TAB: 매출 대시보드 ═══════ -->
<div class="tab-pane active" id="pane-dashboard">

<!-- ═══════ KPI STRIP ═══════ -->
<div class="sticky-kpi">
<div class="kpi-strip" id="kpiStrip"></div>
</div>

<!-- ═══════ MAIN CONTENT (단일 페이지) ═══════ -->
<div class="main">

  <!-- ① 월별 매출 추이 (전체 너비) -->
  <div class="chart-card" id="mainCard">
    <div class="chart-card-hdr">
      <div style="display:flex;align-items:baseline;gap:8px;min-width:0">
        <div class="chart-card-title">📈 월별 매출 추이</div>
        <div class="chart-card-desc" id="chartMeta1">전사 기준 · 4개 연도 비교</div>
      </div>
      <div class="chart-legend" id="legend1">
        <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6b7280"></span>2024 실적</span>
        <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5a96c8"></span>2025 실적</span>
        <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5C2508"></span>2026 실적</span>
        <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6e9650"></span>2026 계획</span>
      </div>
      <button class="chart-expand-btn" id="mainExpandBtn" onclick="toggleMainExpand()" title="확대">⛶</button>
    </div>
    <div class="chart-card-body" style="height:clamp(205px,27vw,255px)">
      <canvas id="chartMain"></canvas>
    </div>
    <button class="stbl-toggle" onclick="toggleStbl(this,'wrap1')">
      <span>요약표</span><span class="arrow">▼</span>
    </button>
    <div class="stbl-wrap" id="wrap1">
      <div style="overflow-x:auto">
        <table id="tblSummary1" style="width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap"></table>
      </div>
    </div>
  </div>

  <!-- ② 누적 매출액 (2열: 월별 누적 + 연도별 누적 막대) -->
  <div class="charts-grid" style="grid-template-columns:2fr 3fr">
    <div class="chart-card" id="cumLineCard">
      <div class="chart-card-hdr">
        <div style="display:flex;align-items:baseline;gap:8px;min-width:0">
          <div class="chart-card-title">📊 누적 매출액 추이</div>
          <div class="chart-card-desc" id="chartMeta2">1월~12월 누적 기준 · 연도별 비교</div>
        </div>
        <div class="chart-legend" id="legend2">
          <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6b7280"></span>2024 실적</span>
          <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5a96c8"></span>2025 실적</span>
          <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#5C2508"></span>2026 실적</span>
          <span class="chart-legend-item"><span class="chart-legend-dot" style="background:#6e9650"></span>2026 계획</span>
        </div>
        <button class="chart-expand-btn" id="cumLineExpandBtn" onclick="toggleCumLineExpand()" title="확대">⛶</button>
      </div>
      <div class="chart-card-body" style="height:clamp(200px,28vw,320px)">
        <canvas id="chartCumLine"></canvas>
      </div>
      <button class="stbl-toggle" onclick="toggleStbl(this,'wrap2')">
        <span>요약표</span><span class="arrow">▼</span>
      </button>
      <div class="stbl-wrap" id="wrap2">
        <div style="overflow-x:auto">
          <table id="tblSummary2" style="width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap"></table>
        </div>
      </div>
    </div>
    <div class="chart-card" id="cumBarCard">
      <div class="chart-card-hdr">
        <div style="display:flex;align-items:baseline;gap:8px;min-width:0">
          <div class="chart-card-title" id="chartTitle3">📊 팀별 사업계획 대비 실적</div>
          <div class="chart-card-desc" id="chartMeta3">2026년 계획 대비 실적 비교</div>
        </div>
        <div class="chart-legend" id="legend3">
          <span class="chart-legend-item"><span class="chart-legend-dot" style="background:rgba(31,78,120,.75)"></span>실적(달성)</span>
          <span class="chart-legend-item"><span class="chart-legend-dot" style="background:rgba(197,90,17,.75)"></span>실적(미달)</span>
          <span class="chart-legend-item"><span class="chart-legend-dot" style="background:rgba(110,150,80,.22);border:1.5px solid #6e9650"></span>계획</span>
        </div>
        <button class="chart-expand-btn" id="cumBarExpandBtn" onclick="toggleCumBarExpand()" title="확대">⛶</button>
      </div>
      <div class="chart-card-body" style="height:clamp(200px,28vw,320px);padding:0;position:relative">
        <div id="cumBarScroll" style="position:absolute;inset:0;overflow-x:auto;overflow-y:hidden">
          <div id="cumBarInner" style="height:100%;min-width:100%">
            <canvas id="chartCumBar"></canvas>
          </div>
        </div>
      </div>
      <button class="stbl-toggle" onclick="toggleStbl(this,'wrap3')">
        <span>요약표</span><span class="arrow">▼</span>
      </button>
      <div class="stbl-wrap" id="wrap3">
        <div style="overflow-x:auto">
          <table id="tblSummary3" style="width:100%;border-collapse:collapse;font-size:11px;white-space:nowrap"></table>
        </div>
      </div>
    </div>
  </div>

  <!-- ③ 월별 실적 & YoY 삭제됨 -->

</div>
<!-- /tab-pane dashboard -->
</div>

<!-- ═══════ TAB: 매출 현황(표) ═══════ -->
<div id="pane-table">
  <!-- KPI 스트립 (표 탭용) -->
  <div class="kpi-strip" id="kpiStripTable" style="border-bottom:1px solid var(--border);flex-shrink:0"></div>
  <div class="tbl-toggle-bar">
    <button id="toggleDetailBtn" onclick="toggleTableDetail()"
      style="padding:3px 12px;border-radius:20px;border:1px solid var(--border);background:#fff;
             font-size:12px;font-weight:600;cursor:pointer;color:var(--blue);font-family:var(--ff)">
      📋 채널 접기
    </button>
    <span style="margin-left:auto;font-size:12px;color:var(--muted);font-weight:500">(단위: 백만원)</span>
  </div>
  <div class="tbl-card">
    <div class="tbl-outer tbl-outer-full">
      <table>
        <thead id="tblHead"></thead>
        <tbody id="tblBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- 숨김 요소 (JS 참조용) -->
<span id="tblCount" style="display:none"></span>
<input id="tblSearch" type="hidden" value="">

<!-- ═══════ TAB: 인쇄용 ═══════ -->
<div id="pane-print">
  <div class="print-toolbar">
    <span class="print-toolbar-title">🖨️ 인쇄 미리보기</span>
    <span class="print-toolbar-sub" id="printMeta">전사 기준 · YTD</span>
    <button class="print-btn" onclick="window.print()">🖨️ 인쇄</button>
  </div>
  <div class="print-preview-wrap">
    <!-- 1페이지: 동북아 BU -->
    <div class="print-paper" id="printPaper">
      <div class="print-paper-header">
        <div class="print-paper-title" id="printPaperTitle">씨엠에스랩 매출 현황</div>
        <div class="print-paper-sub" id="printPaperSub">전사 기준 · YTD · __BASE_DATE__ 기준</div>
      </div>
      <div class="print-table-wrap">
        <table id="printTable"></table>
      </div>
      <div class="print-paper-footer">
        <span>CMS Lab 경영지원실</span>
        <span>__BASE_DATE__ 기준 (1/2)</span>
      </div>
    </div>
    <!-- 2페이지: Global BU + Expert BU + 합계 -->
    <div class="print-paper-2" id="printPaper2">
      <div class="print-paper-header">
        <div class="print-paper-title" id="printPaperTitle2">씨엠에스랩 매출 현황</div>
        <div class="print-paper-sub" id="printPaperSub2">전사 기준 · YTD · __BASE_DATE__ 기준</div>
      </div>
      <div class="print-table-wrap">
        <table id="printTable2"></table>
      </div>
      <div class="print-paper-footer">
        <span>CMS Lab 경영지원실</span>
        <span>__BASE_DATE__ 기준 (2/2)</span>
      </div>
    </div>
  </div>
</div>

<!-- ═══════ CHART.JS ═══════ -->
__CHARTJS__

<!-- ═══════ APP SCRIPT ═══════ -->
<script>
__JSAPP__
</script>
</body>
</html>
"""

def make_html(data_json: str, chartjs_src: str, base_date: str = '') -> str:
    # Embed Chart.js
    if chartjs_src:
        chartjs_tag = f'<script>{chartjs_src}</script>'
    else:
        chartjs_tag = '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>'

    # ── 기본 필터: 1월 ~ (현재월 + 2)월, Y2024 제외 ──
    today = datetime.date.today()
    default_max_month = min(today.month + 2, 12)
    default_months_js = ','.join(str(m) for m in range(1, default_max_month + 1))

    js_app = JS_HELPERS.replace('__DATA_JSON__', data_json)
    js_app = js_app.replace('__DEFAULT_MONTHS__', f'[{default_months_js}]')

    # WEEKLY_META 자동 주입 (extract_data가 채운 자동감지 결과 → JSON, 키는 문자열로)
    _wkcols = _DETECTED_WEEKLY_COLS or WEEKLY_COLS
    weekly_meta = {str(mn): [lbl for _, lbl in weeks] for mn, weeks in sorted(_wkcols.items())}
    js_app = js_app.replace('__WEEKLY_META_JSON__', json.dumps(weekly_meta, ensure_ascii=False))

    html = HTML_SHELL
    # Y2024 버튼 active 제거
    html = html.replace(
        '<button class="pb active" data-f="year" data-v="2024" onclick="togYear(this)">24년 실적</button>',
        '<button class="pb" data-f="year" data-v="2024" onclick="togYear(this)">24년 실적</button>'
    )
    # YTD 버튼 active 제거
    html = html.replace(
        '<button class="pb active" data-f="month" data-v="YTD" onclick="togMonth(this)">YTD</button>',
        '<button class="pb" data-f="month" data-v="YTD" onclick="togMonth(this)">YTD</button>'
    )
    # 1월 ~ default_max_month 버튼 active 추가
    for m in range(1, default_max_month + 1):
        html = html.replace(
            f'<button class="pb" data-f="month" data-v="{m}"',
            f'<button class="pb active" data-f="month" data-v="{m}"'
        )
    html = html.replace('__CSS__',     CSS)
    html = html.replace('__CHARTJS__', chartjs_tag)
    html = html.replace('__JSAPP__',   js_app)
    html = html.replace('__BASE_DATE__', base_date)
    return html


# ═══════════════════════════════════════════════════════════════════════════════
# 4. GENERATE
# ═══════════════════════════════════════════════════════════════════════════════
def generate(records: list, out_path: str, chartjs_src: str, base_date: str = ''):
    data_json = json.dumps(records, ensure_ascii=False, separators=(',', ':'))
    html = make_html(data_json, chartjs_src, base_date)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    kb = os.path.getsize(out_path) / 1024
    chs = len(set((r['team'], r['channel']) for r in records))
    print(f"✅ 생성 완료: {out_path}  ({kb:.0f} KB)")
    print(f"   레코드: {len(records):,}건 | 채널: {chs}개")
    if chartjs_src:
        print("   Chart.js: 오프라인 내장 (CDN 불필요)")
    else:
        print("   Chart.js: CDN 참조 (인터넷 필요)")


def open_browser(html_path: str):
    """생성된 HTML 파일을 기본 브라우저로 자동 오픈."""
    abs_path = os.path.abspath(html_path)
    url = 'file:///' + abs_path.replace('\\', '/')
    print(f"\n🌐 브라우저 오픈 중: {url}")
    try:
        # Windows: start 명령으로 기본 브라우저 실행
        if sys.platform == 'win32':
            os.startfile(abs_path)
        # macOS
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', abs_path])
        # Linux
        else:
            subprocess.Popen(['xdg-open', abs_path])
        print("✅ 브라우저가 열렸습니다!")
    except Exception as e:
        # 최후 수단: webbrowser 모듈
        try:
            webbrowser.open(url)
            print("✅ 브라우저가 열렸습니다!")
        except Exception as e2:
            print(f"⚠️  브라우저 자동 오픈 실패. 직접 파일을 여세요: {abs_path}")


# ═══════════════════════════════════════════════════════════════════════════════
# 5. MAIN
# ═══════════════════════════════════════════════════════════════════════════════

# ── 엑셀 파일 자동 탐색 경로 목록 (본인 환경에 맞게 추가 가능) ──────────────
EXCEL_FILENAME = '실행계획_취합자료_6월 2주차.xlsx'

AUTO_SEARCH_DIRS = [
    # 스크립트와 같은 폴더
    os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else '.',
    # 매출 Dashboard 폴더 (고정 경로)
    r'C:\Users\cmslab_admin\Desktop\1. CMSLAB\2. 관리회계\파이썬\매출 Dashboard',
    # 바탕화면
    os.path.join(os.path.expanduser('~'), 'Desktop'),
    # Downloads
    os.path.join(os.path.expanduser('~'), 'Downloads'),
    # 현재 작업 디렉터리
    os.getcwd(),
]

def read_base_date(xlsx_path: str) -> str:
    """주차 열 5행에서 기준일자를 읽어 반환. 주차 열이 비어있으면
    바로 직전 칼럼(스냅샷 날짜가 별도 칼럼에 있는 경우) → 이전 주차 열 순으로 폴백.
    extract_data()가 채워둔 자동감지 결과를 우선 사용."""
    src = _DETECTED_WEEKLY_COLS or WEEKLY_COLS
    weekly_cols = sorted(
        {col for weeks in src.values() for col, _ in weeks},
        reverse=True,
    )
    # 주차 열과 그 직전 칼럼 모두 후보로 (최신 우선)
    candidates = []
    for col in weekly_cols:
        candidates.append(col)
        if col - 1 not in candidates:
            candidates.append(col - 1)
    try:
        usecols = sorted(set(candidates))
        df = pd.read_excel(xlsx_path, sheet_name='영업그룹별 매출액',
                           header=None, usecols=usecols, nrows=5)
        for col in candidates:
            if col not in df.columns:
                continue
            val = df.loc[4, col]
            if pd.isna(val):
                continue
            if hasattr(val, 'strftime'):
                return val.strftime('%Y년 %m월 %d일')
            s = str(val).strip()
            if not s:
                continue
            for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%m/%d/%Y', '%Y.%m.%d'):
                try:
                    return datetime.datetime.strptime(s, fmt).strftime('%Y년 %m월 %d일')
                except ValueError:
                    pass
            # 원본 문자열에 "기준"이 포함돼 있으면 제거 (템플릿에서 "기준"이 추가됨)
            return s.replace('기준', '').strip()
        return ''
    except Exception:
        return ''


def find_excel(data_arg: str) -> str:
    """Excel 파일을 인수 → 자동탐색 순서로 찾아 반환."""
    # 1) 직접 지정된 경로
    if data_arg and os.path.exists(data_arg):
        return data_arg

    # 2) 인수가 파일명만인 경우 → 자동탐색 디렉터리에서 검색
    filename = os.path.basename(data_arg) if data_arg else EXCEL_FILENAME
    for d in AUTO_SEARCH_DIRS:
        candidate = os.path.join(d, filename)
        if os.path.exists(candidate):
            return candidate

    # 3) 기본 파일명으로 자동탐색
    if filename != EXCEL_FILENAME:
        for d in AUTO_SEARCH_DIRS:
            candidate = os.path.join(d, EXCEL_FILENAME)
            if os.path.exists(candidate):
                return candidate

    return None


def main():
    ap = argparse.ArgumentParser(description='매출 대시보드 생성기 v3')
    ap.add_argument('--data', default=None,
                    help=f'Excel 파일 경로 (기본: 자동탐색 {EXCEL_FILENAME})')
    ap.add_argument('--out',  default=None,
                    help='출력 HTML 파일 경로 (기본: 스크립트와 같은 폴더)')
    ap.add_argument('--no-browser', action='store_true',
                    help='브라우저 자동 오픈 비활성화')
    args = ap.parse_args()

    # ── Excel 파일 탐색 ──────────────────────────────────────────────────────
    xlsx = find_excel(args.data)
    if not xlsx:
        searched = '\n   '.join(AUTO_SEARCH_DIRS)
        print(f"❌ Excel 파일을 찾을 수 없습니다.")
        print(f"   찾은 경로:\n   {searched}")
        print(f"\n   해결방법: --data 옵션으로 직접 경로를 지정하세요.")
        print(f"   예) python generate_dashboard_v3.py --data \"C:\\경로\\파일.xlsx\"")
        sys.exit(1)
    print(f"📂 Excel 읽는 중: {xlsx}")

    # ── 출력 경로 결정 ───────────────────────────────────────────────────────
    if args.out:
        out_path = args.out
    else:
        # 스크립트와 같은 폴더에 저장
        script_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in dir() else os.getcwd()
        out_path = os.path.join(script_dir, '매출_대시보드.html')

    # ── 데이터 추출 ──────────────────────────────────────────────────────────
    records = extract_data(xlsx)
    print(f"✅ 추출 완료: {len(records):,}건")
    if _DETECTED_WEEKLY_COLS:
        summary = ', '.join(
            f'{mn}월 {len(w)}주차(col {",".join(str(c) for c,_ in w)})'
            for mn, w in sorted(_DETECTED_WEEKLY_COLS.items())
        )
        print(f"🔎 주차 자동감지: {summary}")
    else:
        print("⚠️  주차 자동감지 실패 → 하드코딩 WEEKLY_COLS 사용")

    base_date = read_base_date(xlsx)
    if base_date:
        print(f"📅 기준일자: {base_date}")
    else:
        print("⚠️  기준일자를 읽지 못했습니다 (마지막 주차 열 5행 확인 필요)")

    # ── Chart.js 로드 ────────────────────────────────────────────────────────
    chartjs = load_chartjs()
    if chartjs:
        print(f"📦 Chart.js 내장 로드 완료 ({len(chartjs)//1024}KB)")
    else:
        print("⚠️  Chart.js 로컬 파일 없음 → CDN(인터넷) 사용")

    # ── HTML 생성 ────────────────────────────────────────────────────────────
    generate(records, out_path, chartjs, base_date)

    # ── 브라우저 자동 오픈 ───────────────────────────────────────────────────
    if not args.no_browser:
        open_browser(out_path)
    else:
        print(f"\n▶  직접 파일을 열어주세요: {os.path.abspath(out_path)}")


if __name__ == '__main__':
    main()
