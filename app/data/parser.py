"""
기존 매출 Dashboard_vf.py 의 extract_data / make_html 을 importlib로 재활용.
DB 저장 / 조회 함수도 여기서 관리.
"""
import os, sys, json, importlib.util
from typing import List, Dict, Optional
from sqlalchemy.orm import Session

# 원본 스크립트 경로 (환경변수 또는 기본 상대 경로)
DASHBOARD_SRC_PATH = os.getenv(
    "DASHBOARD_SRC_PATH",
    os.path.normpath(
        os.path.join(os.path.dirname(__file__), "..", "..", "..", "sales_dash_cmslab")
    ),
)

_dashboard_mod = None
_compare_mod = None


def _load_dashboard():
    """매출 Dashboard_vf.py 를 importlib로 로드 (fresh load 보장)."""
    global _dashboard_mod
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


def _load_compare():
    """매출_선택비교_vf.py 를 importlib로 로드."""
    global _compare_mod
    path = os.path.join(os.path.abspath(DASHBOARD_SRC_PATH), "매출_선택비교_vf.py")
    if not os.path.exists(path):
        raise FileNotFoundError(f"Compare source not found: {path}")
    spec = importlib.util.spec_from_file_location("compare_main", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["compare_main"] = mod
    spec.loader.exec_module(mod)
    _compare_mod = mod
    return mod


# ─── Excel 파싱 ──────────────────────────────────────────────────────────────

def extract_records_from_excel(xlsx_path: str):
    """Excel → (records list, base_date str). 매번 fresh load 해서 전역 상태 초기화."""
    global _dashboard_mod
    _dashboard_mod = None          # WEEKLY_COLS 전역 상태 때문에 매번 재로드
    mod = _load_dashboard()
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
):
    """기존 active 스냅샷 비활성화 → 새 스냅샷 + 레코드 bulk insert."""
    from ..models import Snapshot, SalesRecord

    db.query(Snapshot).filter(Snapshot.is_active == True).update({"is_active": False})

    snapshot = Snapshot(
        week_label=week_label,
        base_date=base_date,
        uploaded_by=uploaded_by_id,
        is_active=True,
    )
    db.add(snapshot)
    db.flush()

    _FW = ["fw1", "fw2", "fw3", "fw4", "fw5"]
    db_records = []
    for r in records:
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
    db.refresh(snapshot)
    return snapshot


# ─── DB 조회 ─────────────────────────────────────────────────────────────────

def get_active_records(
    db: Session, allowed_teams: Optional[List[str]] = None
) -> List[Dict]:
    """active 스냅샷 레코드를 dict 리스트로 반환. allowed_teams=None 이면 전체."""
    from ..models import Snapshot, SalesRecord

    snapshot = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not snapshot:
        return []

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


def get_active_snapshot_info(db: Session) -> Optional[Dict]:
    """active 스냅샷 메타 정보 반환."""
    from ..models import Snapshot

    s = db.query(Snapshot).filter(Snapshot.is_active == True).first()
    if not s:
        return None
    return {
        "week_label": s.week_label,
        "base_date": s.base_date,
        "uploaded_at": s.uploaded_at.strftime("%Y-%m-%d %H:%M") if s.uploaded_at else "",
    }


# ─── HTML 생성 ───────────────────────────────────────────────────────────────

def make_dashboard_html(records: List[Dict], base_date: str) -> str:
    """기존 make_html() 을 재활용해 대시보드 HTML 생성."""
    mod = _load_dashboard()
    data_json = json.dumps(records, ensure_ascii=False, separators=(",", ":"))
    chartjs_src = mod.load_chartjs()
    return mod.make_html(data_json, chartjs_src, base_date)


def make_compare_html(records: List[Dict], base_date: str) -> str:
    """기존 비교 make_html() 재활용."""
    mod = _load_compare()
    return mod.make_html(records, base_date)
