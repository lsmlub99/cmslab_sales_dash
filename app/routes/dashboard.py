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
    make_dashboard_html,
    make_compare_html,
)

router = APIRouter()

_NO_DATA_HTML = """
<html><body style="font-family:sans-serif;padding:40px">
<h2>데이터가 없습니다</h2>
<p>관리자 패널에서 Excel 파일을 먼저 업로드해주세요.</p>
<p><a href="/admin">관리자 패널로 이동</a></p>
</body></html>
"""

_NAV_STYLE = """
<style>
#__cms-topnav{position:fixed;top:0;left:0;right:0;z-index:99999;
  background:#1a56a0;padding:6px 20px;display:flex;align-items:center;
  justify-content:space-between;font-family:'Malgun Gothic',sans-serif;
  font-size:13px;box-shadow:0 2px 6px rgba(0,0,0,.2)}
#__cms-topnav .nav-title{color:#fff;font-weight:700;letter-spacing:-.3px}
#__cms-topnav .nav-right{display:flex;gap:14px;align-items:center}
#__cms-topnav .nav-user{color:rgba(255,255,255,.75)}
#__cms-topnav a{color:rgba(255,255,255,.85);text-decoration:none}
#__cms-topnav a:hover{color:#fff}
#__cms-topnav .nav-logout{background:rgba(255,255,255,.15);padding:3px 12px;
  border-radius:5px;color:#fff !important}
#__cms-topnav .nav-logout:hover{background:rgba(255,255,255,.28)}
body{padding-top:36px !important}
</style>
"""


def _inject_nav(html: str, user: User) -> str:
    admin_link = '<a href="/admin">관리자</a>' if user.role == "admin" else ""
    nav = f"""{_NAV_STYLE}
<div id="__cms-topnav">
  <span class="nav-title">CMS Lab 매출 대시보드</span>
  <div class="nav-right">
    <span class="nav-user">{user.name or user.email}</span>
    {admin_link}
    <a href="/logout" class="nav-logout">로그아웃</a>
  </div>
</div>"""
    if "<body" in html:
        return html.replace("<body", nav + "<body", 1)
    return nav + html


@router.get("/dashboard", response_class=HTMLResponse)
async def dashboard(
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        return RedirectResponse("/login", status_code=302)

    records = get_active_records(db, current_user.allowed_teams)
    if not records:
        return HTMLResponse(_NO_DATA_HTML)

    info = get_active_snapshot_info(db)
    base_date = info["base_date"] if info else ""
    return HTMLResponse(_inject_nav(make_dashboard_html(records, base_date), current_user))


@router.get("/compare", response_class=HTMLResponse)
async def compare(
    db: Session = Depends(get_db),
    current_user: Optional[User] = Depends(get_current_user),
):
    if not current_user:
        return RedirectResponse("/login", status_code=302)

    records = get_active_records(db, current_user.allowed_teams)
    if not records:
        return HTMLResponse(_NO_DATA_HTML)

    info = get_active_snapshot_info(db)
    base_date = info["base_date"] if info else ""
    return HTMLResponse(_inject_nav(make_compare_html(records, base_date), current_user))
