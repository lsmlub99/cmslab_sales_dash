"""
외부 연동용 REST API — API key 인증
X-API-Key 헤더 또는 ?api_key= 쿼리 파라미터로 인증
"""
import re, secrets
from fastapi import APIRouter, Depends, HTTPException, Header, Query
from fastapi.responses import JSONResponse
from sqlalchemy.orm import Session
from sqlalchemy import func

from ..database import get_db
from ..models import Snapshot, SalesRecord, AppConfig
from ..data.parser import get_all_snapshots, get_active_snapshot_info

router = APIRouter(prefix="/api/v1")


# ── API Key 관리 ──────────────────────────────────────────────────────────────

def get_or_create_api_key(db: Session) -> str:
    row = db.query(AppConfig).filter(AppConfig.key == "api_key").first()
    if not row or not row.value:
        key = secrets.token_urlsafe(32)
        if row:
            row.value = key
        else:
            db.add(AppConfig(key="api_key", value=key))
        db.commit()
        return key
    return row.value


def _require_api_key(
    x_api_key: str = Header(None, alias="X-API-Key"),
    api_key_q: str = Query(None, alias="api_key"),
    db: Session = Depends(get_db),
):
    key = x_api_key or api_key_q
    if not key:
        raise HTTPException(401, detail="API key required. Use X-API-Key header or ?api_key= param.")
    stored = get_or_create_api_key(db)
    if key != stored:
        raise HTTPException(403, detail="Invalid API key.")
    return key


def _parse_base_month(base_date_str: str) -> int:
    m = re.search(r'(\d{1,2})월', base_date_str)
    return int(m.group(1)) if m else 12


# ── 엔드포인트 ────────────────────────────────────────────────────────────────

@router.get("/info", summary="현재 스냅샷 정보")
def api_info(db: Session = Depends(get_db), _=Depends(_require_api_key)):
    info = get_active_snapshot_info(db)
    return {"snapshot": info}


@router.get("/summary", summary="팀별 누적 실적 요약")
def api_summary(db: Session = Depends(get_db), _=Depends(_require_api_key)):
    snap = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snap:
        raise HTTPException(404, "No active snapshot.")

    base_month = _parse_base_month(snap.base_date)
    rows = (
        db.query(
            SalesRecord.team,
            func.sum(SalesRecord.actual).label("actual"),
            func.sum(SalesRecord.plan).label("plan"),
            func.sum(SalesRecord.y2025).label("y2025"),
        )
        .filter(SalesRecord.snapshot_id == snap.id, SalesRecord.month <= base_month)
        .group_by(SalesRecord.team)
        .order_by(SalesRecord.team)
        .all()
    )
    teams = [
        {
            "team": r.team,
            "actual": float(r.actual or 0),
            "plan": float(r.plan or 0),
            "y2025": float(r.y2025 or 0),
            "achievement_rate": round(float(r.actual or 0) / float(r.plan) * 100, 1) if r.plan else None,
        }
        for r in rows
    ]
    total_actual = sum(t["actual"] for t in teams)
    total_plan = sum(t["plan"] for t in teams)
    total_y2025 = sum(t["y2025"] for t in teams)
    return {
        "snapshot": {"week_label": snap.week_label, "base_date": snap.base_date, "base_month": base_month},
        "teams": teams,
        "total": {
            "actual": total_actual,
            "plan": total_plan,
            "y2025": total_y2025,
            "achievement_rate": round(total_actual / total_plan * 100, 1) if total_plan else None,
        },
        "unit": "백만원",
    }


@router.get("/teams/{team_name}", summary="특정 팀 월별 실적 상세")
def api_team_detail(
    team_name: str,
    db: Session = Depends(get_db),
    _=Depends(_require_api_key),
):
    snap = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snap:
        raise HTTPException(404, "No active snapshot.")

    rows = (
        db.query(
            SalesRecord.month,
            func.sum(SalesRecord.actual).label("actual"),
            func.sum(SalesRecord.plan).label("plan"),
            func.sum(SalesRecord.y2025).label("y2025"),
            func.sum(SalesRecord.y2024).label("y2024"),
        )
        .filter(SalesRecord.snapshot_id == snap.id, SalesRecord.team == team_name)
        .group_by(SalesRecord.month)
        .order_by(SalesRecord.month)
        .all()
    )
    if not rows:
        raise HTTPException(404, f"Team '{team_name}' not found.")

    monthly = [
        {
            "month": r.month,
            "actual": float(r.actual or 0),
            "plan": float(r.plan or 0),
            "y2025": float(r.y2025 or 0),
            "y2024": float(r.y2024 or 0),
        }
        for r in rows
    ]
    return {
        "team": team_name,
        "snapshot": {"week_label": snap.week_label, "base_date": snap.base_date},
        "monthly": monthly,
        "total": {
            "actual": sum(m["actual"] for m in monthly),
            "plan": sum(m["plan"] for m in monthly),
            "y2025": sum(m["y2025"] for m in monthly),
            "y2024": sum(m["y2024"] for m in monthly),
        },
        "unit": "백만원",
    }


@router.get("/snapshots", summary="스냅샷 이력 목록")
def api_snapshots(db: Session = Depends(get_db), _=Depends(_require_api_key)):
    return {"snapshots": get_all_snapshots(db)}
