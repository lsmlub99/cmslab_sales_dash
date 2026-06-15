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
    return HTMLResponse(make_dashboard_html(records, base_date))


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
    return HTMLResponse(make_compare_html(records, base_date))
