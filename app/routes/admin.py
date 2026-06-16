import os, tempfile, threading
from typing import Optional
from fastapi import APIRouter, Depends, Form, UploadFile, File, HTTPException, Request, BackgroundTasks
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from ..database import get_db, SessionLocal
from ..auth import require_admin, hash_password
from ..models import User, Snapshot, SalesRecord, Team, AppConfig
from ..data.parser import (
    extract_records_from_excel,
    save_snapshot,
    get_active_snapshot_info,
    create_upload_task,
    get_upload_task,
    _set_task,
    prewarm_html_cache,
)

router = APIRouter(prefix="/admin")
_tpl_dir = os.path.join(os.path.dirname(__file__), "..", "templates")
templates = Jinja2Templates(directory=_tpl_dir)


def get_teams(db: Session):
    rows = db.query(Team).filter(Team.is_active == True).order_by(Team.display_order).all()
    return [r.name for r in rows]


def get_config(db: Session, key: str, default: str = "") -> str:
    row = db.query(AppConfig).filter(AppConfig.key == key).first()
    return row.value if row else default


def set_config(db: Session, key: str, value: str):
    row = db.query(AppConfig).filter(AppConfig.key == key).first()
    if row:
        row.value = value
    else:
        db.add(AppConfig(key=key, value=value))
    db.commit()


# ─── 메인 페이지 ──────────────────────────────────────────────────────────────

@router.get("", response_class=HTMLResponse)
async def admin_page(
    request: Request,
    msg: str = "",
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    snapshots = (
        db.query(Snapshot)
        .order_by(Snapshot.uploaded_at.desc())
        .limit(10)
        .all()
    )
    users = db.query(User).order_by(User.created_at.desc()).all()
    info = get_active_snapshot_info(db)
    all_teams = db.query(Team).order_by(Team.display_order).all()
    config = {
        "app_title": get_config(db, "app_title", "CMS Lab 매출 대시보드"),
        "notice_enabled": get_config(db, "notice_enabled", "false"),
        "notice_text": get_config(db, "notice_text", ""),
        "allowed_domain": get_config(db, "allowed_domain", "wonik.com"),
        "admin_contact": get_config(db, "admin_contact", ""),
    }
    return templates.TemplateResponse("admin.html", {
        "request": request,
        "user": current_user,
        "snapshots": snapshots,
        "users": users,
        "active_info": info,
        "teams": get_teams(db),
        "all_teams": all_teams,
        "config": config,
        "msg": msg,
    })


# ─── Excel 업로드 ─────────────────────────────────────────────────────────────

def _run_upload_task(task_id: str, tmp_path: str, label: str, user_id: int):
    """백그라운드 스레드: Excel 파싱 → DB insert → HTML 캐시 사전 생성."""
    db = SessionLocal()
    try:
        _set_task(task_id, status="parsing", message="Excel 파일 파싱 중...")
        records, base_date = extract_records_from_excel(tmp_path)
        _set_task(task_id, message=f"파싱 완료 ({len(records):,}건). DB 저장 시작...")

        save_snapshot(db, records, label, base_date, user_id, task_id=task_id)

        _set_task(task_id, status="warming", message="대시보드 캐시 사전 생성 중...")
        prewarm_html_cache(db)

        _set_task(task_id, status="done", progress=len(records), total=len(records),
                  message=f"✅ 완료: {len(records):,}건 저장 ({label})")
    except Exception as e:
        _set_task(task_id, status="error", message=f"❌ 오류: {e}")
    finally:
        db.close()
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


@router.post("/upload")
async def upload_excel(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
    week_label: str = Form(""),
    current_user: User = Depends(require_admin),
):
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, "Excel 파일(.xlsx)만 업로드 가능합니다.")

    content = await file.read()
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(content)
        tmp_path = tmp.name

    label = week_label.strip() or os.path.splitext(file.filename)[0]
    task_id = create_upload_task()

    # 별도 스레드로 실행 (BackgroundTasks는 응답 후 실행되지만 DB 세션이 닫히므로 스레드 사용)
    t = threading.Thread(target=_run_upload_task, args=(task_id, tmp_path, label, current_user.id), daemon=True)
    t.start()

    return JSONResponse({"task_id": task_id, "label": label})


@router.get("/upload-status/{task_id}")
async def upload_status(
    task_id: str,
    current_user: User = Depends(require_admin),
):
    task = get_upload_task(task_id)
    if not task:
        raise HTTPException(404, "태스크를 찾을 수 없습니다.")
    return task


# ─── 수동 입력 (Import Form) ─────────────────────────────────────────────────

@router.get("/record")
async def get_record(
    team: str,
    channel: str,
    brand: str,
    month: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    """입력 칸 자동 채우기용 기존 레코드 조회."""
    snapshot = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snapshot:
        return {}
    r = db.query(SalesRecord).filter(
        SalesRecord.snapshot_id == snapshot.id,
        SalesRecord.team == team,
        SalesRecord.channel == channel,
        SalesRecord.brand == brand,
        SalesRecord.month == month,
    ).first()
    if not r:
        return {}
    result = {
        "y2024": float(r.y2024 or 0),
        "y2025b": float(r.y2025b or 0),
        "y2025": float(r.y2025 or 0),
        "plan": float(r.plan or 0),
        "actual": float(r.actual or 0),
    }
    for fw in ["fw1", "fw2", "fw3", "fw4", "fw5"]:
        v = getattr(r, fw)
        result[fw] = float(v) if v is not None else None
    return result


@router.post("/import-record")
async def import_record(
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    """단건 레코드 upsert (active 스냅샷 기준)."""
    body = await request.json()
    snapshot = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snapshot:
        raise HTTPException(400, "활성 스냅샷이 없습니다. 먼저 Excel을 업로드하세요.")

    r = db.query(SalesRecord).filter(
        SalesRecord.snapshot_id == snapshot.id,
        SalesRecord.team == body["team"],
        SalesRecord.channel == body["channel"],
        SalesRecord.brand == body.get("brand", "기타"),
        SalesRecord.month == int(body["month"]),
    ).first()

    _fields = ["y2024", "y2025b", "y2025", "plan", "actual", "fw1", "fw2", "fw3", "fw4", "fw5"]
    def _v(val):
        return float(val) if val not in (None, "", "null") else None

    if r:
        for f in _fields:
            if f in body:
                setattr(r, f, _v(body[f]))
    else:
        r = SalesRecord(
            snapshot_id=snapshot.id,
            team=body["team"],
            channel=body["channel"],
            brand=body.get("brand", "기타"),
            code=body.get("code", ""),
            month=int(body["month"]),
            **{f: _v(body.get(f)) for f in _fields},
        )
        db.add(r)

    db.commit()
    return {"ok": True}


@router.post("/import-bulk")
async def import_bulk(
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    """JSON 배열 다건 일괄 import."""
    items = await request.json()
    snapshot = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snapshot:
        raise HTTPException(400, "활성 스냅샷이 없습니다.")

    _fields = ["y2024", "y2025b", "y2025", "plan", "actual", "fw1", "fw2", "fw3", "fw4", "fw5"]
    def _v(val):
        return float(val) if val not in (None, "", "null") else None

    updated = inserted = 0
    for item in items:
        r = db.query(SalesRecord).filter(
            SalesRecord.snapshot_id == snapshot.id,
            SalesRecord.team == item["team"],
            SalesRecord.channel == item["channel"],
            SalesRecord.brand == item.get("brand", "기타"),
            SalesRecord.month == int(item["month"]),
        ).first()
        if r:
            for f in _fields:
                if f in item:
                    setattr(r, f, _v(item[f]))
            updated += 1
        else:
            r = SalesRecord(
                snapshot_id=snapshot.id,
                team=item["team"],
                channel=item["channel"],
                brand=item.get("brand", "기타"),
                code=item.get("code", ""),
                month=int(item["month"]),
                **{f: _v(item.get(f)) for f in _fields},
            )
            db.add(r)
            inserted += 1

    db.commit()
    return {"ok": True, "updated": updated, "inserted": inserted}


# ─── 사용자 관리 ──────────────────────────────────────────────────────────────

@router.post("/users")
async def manage_users(
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    body = await request.json()
    action = body.get("action", "create")

    if action == "create":
        if db.query(User).filter(User.email == body["email"]).first():
            raise HTTPException(400, "이미 존재하는 이메일입니다.")
        teams = body.get("allowed_teams") or None
        user = User(
            email=body["email"],
            hashed_password=hash_password(body["password"]),
            name=body.get("name", ""),
            role=body.get("role", "viewer"),
            allowed_teams=teams,
            is_active=True,
        )
        db.add(user)
        db.commit()
        return {"ok": True, "id": user.id}

    elif action == "update":
        user = db.query(User).filter(User.id == body["id"]).first()
        if not user:
            raise HTTPException(404, "사용자를 찾을 수 없습니다.")
        for field in ("name", "role", "is_active"):
            if field in body:
                setattr(user, field, body[field])
        if "allowed_teams" in body:
            user.allowed_teams = body["allowed_teams"] or None
        if body.get("password"):
            user.hashed_password = hash_password(body["password"])
        db.commit()
        return {"ok": True}

    elif action == "delete":
        user = db.query(User).filter(User.id == body["id"]).first()
        if not user:
            raise HTTPException(404, "사용자를 찾을 수 없습니다.")
        if user.id == current_user.id:
            raise HTTPException(400, "자기 자신은 삭제할 수 없습니다.")
        db.delete(user)
        db.commit()
        return {"ok": True}

    raise HTTPException(400, f"알 수 없는 action: {action}")


# ─── 팀 관리 ─────────────────────────────────────────────────────────────────

@router.post("/teams")
async def manage_teams(
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    body = await request.json()
    action = body.get("action", "")

    if action == "create":
        name = body.get("name", "").strip()
        if not name:
            raise HTTPException(400, "팀 이름을 입력하세요.")
        if db.query(Team).filter(Team.name == name).first():
            raise HTTPException(400, "이미 존재하는 팀입니다.")
        max_order = db.query(Team).count()
        db.add(Team(name=name, display_order=max_order + 1, is_active=True))
        db.commit()
        return {"ok": True}

    elif action == "update":
        team = db.query(Team).filter(Team.id == body["id"]).first()
        if not team:
            raise HTTPException(404, "팀을 찾을 수 없습니다.")
        if "name" in body:
            team.name = body["name"]
        if "display_order" in body:
            team.display_order = body["display_order"]
        if "is_active" in body:
            team.is_active = body["is_active"]
        db.commit()
        return {"ok": True}

    elif action == "delete":
        team = db.query(Team).filter(Team.id == body["id"]).first()
        if not team:
            raise HTTPException(404, "팀을 찾을 수 없습니다.")
        db.delete(team)
        db.commit()
        return {"ok": True}

    raise HTTPException(400, f"알 수 없는 action: {action}")


# ─── 시스템 설정 ──────────────────────────────────────────────────────────────

@router.post("/config")
async def update_config(
    request: Request,
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    body = await request.json()
    allowed_keys = {"app_title", "notice_enabled", "notice_text", "allowed_domain", "admin_contact"}
    for k, v in body.items():
        if k in allowed_keys:
            set_config(db, k, str(v))
    return {"ok": True}


# ─── 스케줄러 수동 트리거 ────────────────────────────────────────────────────

@router.post("/trigger-update")
async def trigger_update(
    db: Session = Depends(get_db),
    current_user: User = Depends(require_admin),
):
    from ..scheduler import run_scheduled_update
    try:
        result = run_scheduled_update(db)
        return {"ok": True, "message": result}
    except Exception as e:
        raise HTTPException(500, str(e))