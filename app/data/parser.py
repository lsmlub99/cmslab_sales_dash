"""
기존 매출 Dashboard_vf.py 의 extract_data / make_html 을 importlib로 재활용.
DB 저장 / 조회 함수도 여기서 관리.
"""
import os, sys, json, uuid, importlib.util
from typing import List, Dict, Optional
from sqlalchemy.orm import Session
from sqlalchemy import func

# 원본 스크립트 경로 (환경변수 또는 기본 상대 경로)
DASHBOARD_SRC_PATH = os.getenv(
    "DASHBOARD_SRC_PATH",
    os.path.normpath(
        os.path.join(os.path.dirname(__file__), "scripts")
    ),
)

_dashboard_mod = None
_compare_mod = None
_chartjs_cache: Optional[str] = None
_html_cache: Dict = {}      # key: (snapshot_id, teams_key, page) -> html
_upload_tasks: Dict = {}    # key: task_id -> { status, progress, total, message }

_CHUNK_SIZE = 5_000         # DB insert 청크 크기


def _load_dashboard(force: bool = False):
    """매출 Dashboard_vf.py 를 importlib로 로드. force=True 일 때만 재컴파일."""
    global _dashboard_mod
    if _dashboard_mod is not None and not force:
        return _dashboard_mod
    path = os.path.join(os.path.abspath(DASHBOARD_SRC_PATH), "매출 Dashboard_vf.py")
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"Dashboard source not found: {path}\n"
            f"DASHBOARD_SRC_PATH={DASHBOARD_SRC_PATH}"
        )
    spec = importlib.util.spec_from_file_location("dashboard_main", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["dashboard_main"] = mod
    spec.loader.exec_module(mod)
    _dashboard_mod = mod
    return mod


def _load_compare(force: bool = False):
    """매출_선택비교_vf.py 를 importlib로 로드. force=True 일 때만 재컴파일."""
    global _compare_mod
    if _compare_mod is not None and not force:
        return _compare_mod
    path = os.path.join(os.path.abspath(DASHBOARD_SRC_PATH), "매출_선택비교_vf.py")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Compare source not found: {path}")
    spec = importlib.util.spec_from_file_location("compare_main", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["compare_main"] = mod
    spec.loader.exec_module(mod)
    _compare_mod = mod
    return mod


# ─── 업로드 진행률 관리 ──────────────────────────────────────────────────────

def create_upload_task() -> str:
    task_id = str(uuid.uuid4())
    _upload_tasks[task_id] = {"status": "pending", "progress": 0, "total": 0, "message": "대기 중"}
    return task_id


def get_upload_task(task_id: str) -> Optional[Dict]:
    return _upload_tasks.get(task_id)


def _set_task(task_id: str, **kwargs):
    if task_id in _upload_tasks:
        _upload_tasks[task_id].update(kwargs)


# ─── Excel 파싱 ──────────────────────────────────────────────────────────────

def extract_records_from_excel(xlsx_path: str):
    """Excel → (records list, base_date str). Excel 업로드 시에는 강제 재로드."""
    mod = _load_dashboard(force=True)   # WEEKLY_COLS 전역 상태 때문에 업로드 시만 재로드
    records = mod.extract_data(xlsx_path)
    base_date = mod.read_base_date(xlsx_path)
    return records, base_date


# ─── DB 저장 ─────────────────────────────────────────────────────────────────

def save_snapshot(
    db: Session,
    records: List[Dict],
    week_label: str,
    base_date: str,
    uploaded_by_id: Optional[int] = None,
    task_id: Optional[str] = None,
):
    """기존 active 스냅샷 비활성화 + 이전 records 삭제 → 새 스냅샷 + 청크 단위 bulk insert.
    같은 파일을 여러 번 올려도 records가 중복 누적되지 않도록 이전 데이터를 먼저 정리한다.
    Snapshot 행(이력 메타)은 유지하고 SalesRecord만 삭제해 스토리지를 절약한다.
    """
    from ..models import Snapshot, SalesRecord

    # 이전 활성 스냅샷의 records 삭제 (Snapshot 행은 이력용으로 유지)
    old_ids = [s.id for s in db.query(Snapshot.id).filter(Snapshot.is_active == True).all()]
    if old_ids:
        db.query(SalesRecord).filter(SalesRecord.snapshot_id.in_(old_ids)).delete(synchronize_session=False)
        db.query(Snapshot).filter(Snapshot.id.in_(old_ids)).update({"is_active": False}, synchronize_session=False)
        db.commit()

    snapshot = Snapshot(
        week_label=week_label,
        base_date=base_date,
        uploaded_by=uploaded_by_id,
        is_active=True,
    )
    db.add(snapshot)
    db.flush()

    _FW = ["fw1", "fw2", "fw3", "fw4", "fw5"]
    total = len(records)
    if task_id:
        _set_task(task_id, status="inserting", total=total, message=f"DB 저장 중 (0 / {total:,}건)")

    for chunk_start in range(0, total, _CHUNK_SIZE):
        chunk = records[chunk_start: chunk_start + _CHUNK_SIZE]
        db_records = []
        for r in chunk:
            fw_vals = {fw: r.get(fw) for fw in _FW}
            db_records.append(SalesRecord(
                snapshot_id=snapshot.id,
                team=r.get("team", ""),
                channel=r.get("channel", ""),
                brand=r.get("brand", "기타"),
                code=r.get("code", ""),
                month=r.get("month", 0),
                y2024=r.get("y2024", 0),
                y2025b=r.get("y2025b", 0),
                y2025=r.get("y2025", 0),
                plan=r.get("plan", 0),
                actual=r.get("actual", 0),
                **fw_vals,
            ))
        db.bulk_save_objects(db_records)
        db.commit()

        done = min(chunk_start + _CHUNK_SIZE, total)
        if task_id:
            _set_task(task_id, progress=done, message=f"DB 저장 중 ({done:,} / {total:,}건)")

    db.refresh(snapshot)
    clear_html_cache()
    return snapshot


# ─── DB 조회 ─────────────────────────────────────────────────────────────────

def get_active_records(
    db: Session,
    allowed_teams: Optional[List[str]] = None,
    aggregated: bool = False,
) -> List[Dict]:
    """active 스냅샷 레코드 반환.
    aggregated=True 이면 DB에서 (team, brand, month) 기준으로 GROUP BY SUM → 대용량 최적화.
    """
    from ..models import Snapshot, SalesRecord

    snapshot = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snapshot:
        return []

    if aggregated:
        return _get_aggregated_records(db, snapshot.id, allowed_teams)

    q = db.query(SalesRecord).filter(SalesRecord.snapshot_id == snapshot.id)
    if allowed_teams:
        q = q.filter(SalesRecord.team.in_(allowed_teams))

    result = []
    for r in q.all():
        rec = {
            "team": r.team,
            "channel": r.channel,
            "brand": r.brand,
            "code": r.code,
            "month": r.month,
            "y2024": float(r.y2024 or 0),
            "y2025b": float(r.y2025b or 0),
            "y2025": float(r.y2025 or 0),
            "plan": float(r.plan or 0),
            "actual": float(r.actual or 0),
        }
        for fw in ["fw1", "fw2", "fw3", "fw4", "fw5"]:
            v = getattr(r, fw)
            rec[fw] = float(v) if v is not None else None
        result.append(rec)
    return result


def _get_aggregated_records(
    db: Session,
    snapshot_id: int,
    allowed_teams: Optional[List[str]] = None,
) -> List[Dict]:
    """DB에서 (team, brand, month) 기준 SUM 집계 → 대용량 데이터 대시보드 렌더링용."""
    from ..models import SalesRecord

    cols = [
        SalesRecord.team,
        SalesRecord.brand,
        SalesRecord.month,
        func.sum(SalesRecord.y2024).label("y2024"),
        func.sum(SalesRecord.y2025b).label("y2025b"),
        func.sum(SalesRecord.y2025).label("y2025"),
        func.sum(SalesRecord.plan).label("plan"),
        func.sum(SalesRecord.actual).label("actual"),
        func.sum(SalesRecord.fw1).label("fw1"),
        func.sum(SalesRecord.fw2).label("fw2"),
        func.sum(SalesRecord.fw3).label("fw3"),
        func.sum(SalesRecord.fw4).label("fw4"),
        func.sum(SalesRecord.fw5).label("fw5"),
    ]
    q = db.query(*cols).filter(SalesRecord.snapshot_id == snapshot_id)
    if allowed_teams:
        q = q.filter(SalesRecord.team.in_(allowed_teams))
    q = q.group_by(SalesRecord.team, SalesRecord.brand, SalesRecord.month)

    result = []
    for r in q.all():
        result.append({
            "team": r.team,
            "channel": "",   # 집계 시 channel/code 없음
            "brand": r.brand,
            "code": "",
            "month": r.month,
            "y2024": float(r.y2024 or 0),
            "y2025b": float(r.y2025b or 0),
            "y2025": float(r.y2025 or 0),
            "plan": float(r.plan or 0),
            "actual": float(r.actual or 0),
            "fw1": float(r.fw1) if r.fw1 is not None else None,
            "fw2": float(r.fw2) if r.fw2 is not None else None,
            "fw3": float(r.fw3) if r.fw3 is not None else None,
            "fw4": float(r.fw4) if r.fw4 is not None else None,
            "fw5": float(r.fw5) if r.fw5 is not None else None,
        })
    return result


def get_active_snapshot_info(db: Session) -> Optional[Dict]:
    """active 스냅샷 메타 정보 반환."""
    from ..models import Snapshot

    s = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not s:
        return None
    return {
        "id": s.id,
        "week_label": s.week_label,
        "base_date": s.base_date,
        "uploaded_at": s.uploaded_at.strftime("%Y-%m-%d %H:%M") if s.uploaded_at else "",
    }


def get_record_count(db: Session) -> int:
    """active 스냅샷의 레코드 수 반환."""
    from ..models import Snapshot, SalesRecord
    snapshot = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snapshot:
        return 0
    return db.query(func.count(SalesRecord.id)).filter(SalesRecord.snapshot_id == snapshot.id).scalar() or 0


# ─── HTML 생성 ───────────────────────────────────────────────────────────────

_AGGREGATE_THRESHOLD = 50_000   # 이 건수 초과 시 자동 집계 모드 사용


def _teams_key(allowed_teams: Optional[List[str]]) -> str:
    if not allowed_teams:
        return "__all__"
    return ",".join(sorted(allowed_teams))


def make_dashboard_html(records: List[Dict], base_date: str) -> str:
    global _chartjs_cache
    mod = _load_dashboard()
    data_json = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    if _chartjs_cache is None:
        _chartjs_cache = mod.load_chartjs()
    return mod.make_html(data_json, _chartjs_cache, base_date)


def make_compare_html(records: List[Dict], base_date: str) -> str:
    mod = _load_compare()
    return mod.make_html(records, base_date)


def get_cached_html(
    page: str,
    snapshot_id: int,
    allowed_teams: Optional[List[str]],
    records: List[Dict],
    base_date: str,
) -> str:
    """HTML 결과를 (snapshot_id, teams, page) 키로 캐시. 새 스냅샷 업로드 시 자동 무효화."""
    key = (snapshot_id, _teams_key(allowed_teams), page)
    if key in _html_cache:
        return _html_cache[key]
    html = make_dashboard_html(records, base_date) if page == "dashboard" else make_compare_html(records, base_date)
    _html_cache[key] = html
    return html


def clear_html_cache():
    """새 스냅샷 업로드 시 HTML 캐시 전체 무효화."""
    global _html_cache
    _html_cache = {}


def prewarm_html_cache(db: Session):
    """업로드 완료 후 모든 팀 조합의 HTML을 미리 생성해 캐시에 올린다."""
    from ..models import Team
    info = get_active_snapshot_info(db)
    if not info:
        return

    record_count = get_record_count(db)
    use_agg = record_count > _AGGREGATE_THRESHOLD

    # 전체 팀 캐시
    records = get_active_records(db, allowed_teams=None, aggregated=use_agg)
    if records:
        get_cached_html("dashboard", info["id"], None, records, info["base_date"])
        get_cached_html("compare", info["id"], None, records, info["base_date"])

    # 팀별 캐시
    teams = [t.name for t in db.query(Team).filter(Team.is_active == True).all()]
    for team in teams:
        records = get_active_records(db, allowed_teams=[team], aggregated=use_agg)
        if records:
            get_cached_html("dashboard", info["id"], [team], records, info["base_date"])
            get_cached_html("compare", info["id"], [team], records, info["base_date"])
