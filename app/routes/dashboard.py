from typing import Optional
from fastapi import APIRouter, Depends
from fastapi.responses import HTMLResponse, RedirectResponse
from sqlalchemy.orm import Session

from ..database import get_db
from ..auth import get_current_user
from ..models import User
from ..data.parser import (
    get_active_records,
    get_active_snapshot_info,
    get_cached_html,
    get_record_count,
    _AGGREGATE_THRESHOLD,
)

router = APIRouter()

_NO_DATA_HTML = """
<html><body style="font-family:sans-serif;padding:40px">
<h2>데이터가 없습니다</h2>
<p>관리자 패널에서 Excel 파일을 먼저 업로드해주세요.</p>
<p><a href="/admin">관리자 패널로 이동</a></p>
</body></html>
"""

_NAV_HEIGHT = 36   # 네비바 높이(px) — 대시보드 sticky 오프셋 계산에 사용

_NAV_STYLE = f"""
<style>
#__cms-topnav{{position:fixed;top:0;left:0;right:0;z-index:99999;
  background:#1a56a0;padding:6px 20px;display:flex;align-items:center;
  justify-content:space-between;font-family:'Malgun Gothic',sans-serif;
  font-size:13px;box-shadow:0 2px 6px rgba(0,0,0,.2)}}
#__cms-topnav .nav-title{{color:#fff;font-weight:700;letter-spacing:-.3px}}
#__cms-topnav .nav-right{{display:flex;gap:14px;align-items:center}}
#__cms-topnav .nav-user{{color:rgba(255,255,255,.75)}}
#__cms-topnav a{{color:rgba(255,255,255,.85);text-decoration:none}}
#__cms-topnav a:hover{{color:#fff}}
#__cms-topnav .nav-logout{{background:rgba(255,255,255,.15);padding:3px 12px;
  border-radius:5px;color:#fff !important}}
#__cms-topnav .nav-logout:hover{{background:rgba(255,255,255,.28)}}
body{{padding-top:{_NAV_HEIGHT}px !important}}
/* 대시보드 sticky/fixed 요소: 네비바 높이만큼 top 오프셋 */
.sticky-header   {{top:{_NAV_HEIGHT}px !important}}
.sticky-filterbar{{top:{44  + _NAV_HEIGHT}px !important}}
.sticky-kpi      {{top:{114 + _NAV_HEIGHT}px !important}}
#pane-print      {{top:{114 + _NAV_HEIGHT}px !important}}
</style>
"""

# #pane-table top은 JS가 (hdrH + filterH + 8)로 계산하는데 nav 높이를 빠뜨림.
# MutationObserver로 style 속성이 바뀔 때마다 nav 높이를 더해서 교정.
_NAV_SCRIPT = f"""<script>
(function(){{
  var H = {_NAV_HEIGHT}, _lock = false;
  var p = document.getElementById('pane-table');
  if (!p) return;
  new MutationObserver(function() {{
    if (_lock) return;
    var hdrH    = document.querySelector('.sticky-header')?.offsetHeight    ?? 44;
    var filterH = document.querySelector('.sticky-filterbar')?.offsetHeight ?? 70;
    var want = H + hdrH + filterH + 8;
    if (parseInt(p.style.top, 10) !== want) {{
      _lock = true; p.style.top = want + 'px'; _lock = false;
    }}
  }}).observe(p, {{attributes:true, attributeFilter:['style']}});
}})();
</script>"""


def _inject_nav(html: str, user: User, db=None) -> str:
    from ..models import AppConfig

    app_title = "CMS Lab 매출 대시보드"
    notice_html = ""
    try:
        if db:
            title_row = db.query(AppConfig).filter(AppConfig.key == "app_title").first()
            if title_row and title_row.value:
                app_title = title_row.value
            notice_on = db.query(AppConfig).filter(AppConfig.key == "notice_enabled").first()
            notice_text = db.query(AppConfig).filter(AppConfig.key == "notice_text").first()
            if notice_on and notice_on.value == "true" and notice_text and notice_text.value:
                notice_html = f"""<div style="background:#fef3c7;color:#92400e;padding:6px 20px;
                    font-family:sans-serif;font-size:13px;text-align:center;border-bottom:1px solid #fde68a">
                    📢 {notice_text.value}</div>"""
    except Exception:
        pass

    admin_link = '<a href="/admin">관리자</a>' if user.role == "admin" else ""
    nav = f"""{_NAV_STYLE}
<div id="__cms-topnav">
  <span class="nav-title">{app_title}</span>
  <div class="nav-right">
    <span class="nav-user">{user.name or user.email}</span>
    {admin_link}
    <a href="/logout" class="nav-logout">로그아웃</a>
  </div>
</div>{notice_html}"""
    if "<body" in html:
        html = html.replace("<body", nav + "<body", 1)
    else:
        html = nav + html
    # </body> 직전에 #pane-table top 교정 스크립트 삽입
    if "</body>" in html:
        html = html.replace("</body>", _NAV_SCRIPT + "</body>", 1)
    else:
        html = html + _NAV_SCRIPT
    return html


@router.get("/dashboard", response_class=HTMLResponse)
async def dashboard(
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        return RedirectResponse("/login", status_code=302)

    info = get_active_snapshot_info(db)
    if not info:
        return HTMLResponse(_NO_DATA_HTML)

    use_agg = get_record_count(db) > _AGGREGATE_THRESHOLD
    records = get_active_records(db, current_user.allowed_teams, aggregated=use_agg)
    if not records:
        return HTMLResponse(_NO_DATA_HTML)

    html = get_cached_html("dashboard", info["id"], current_user.allowed_teams, records, info["base_date"])
    return HTMLResponse(_inject_nav(html, current_user, db))


@router.get("/compare", response_class=HTMLResponse)
async def compare(
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        return RedirectResponse("/login", status_code=302)

    info = get_active_snapshot_info(db)
    if not info:
        return HTMLResponse(_NO_DATA_HTML)

    use_agg = get_record_count(db) > _AGGREGATE_THRESHOLD
    records = get_active_records(db, current_user.allowed_teams, aggregated=use_agg)
    if not records:
        return HTMLResponse(_NO_DATA_HTML)

    html = get_cached_html("compare", info["id"], current_user.allowed_teams, records, info["base_date"])
    return HTMLResponse(_inject_nav(html, current_user, db))
