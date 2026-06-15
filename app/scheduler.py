import os, glob, datetime
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

DATA_WATCH_DIR = os.getenv("DATA_WATCH_DIR", "")


def run_scheduled_update(db=None) -> str:
    """지정 폴더에서 가장 최신 xlsx 를 감지해 DB를 업데이트.
    DATA_WATCH_DIR 이 없으면 바로 반환 (수동 업로드 전용 운영 시).
    """
    from .database import SessionLocal
    from .models import Snapshot
    from .data.parser import extract_records_from_excel, save_snapshot

    if not DATA_WATCH_DIR or not os.path.isdir(DATA_WATCH_DIR):
        return "DATA_WATCH_DIR 가 설정되지 않았거나 존재하지 않습니다."

    files = sorted(
        glob.glob(os.path.join(DATA_WATCH_DIR, "*.xlsx")),
        key=os.path.getmtime,
        reverse=True,
    )
    if not files:
        return "xlsx 파일이 없습니다."

    latest = files[0]
    close_db = db is None
    if db is None:
        db = SessionLocal()

    try:
        # 파일 수정 시간 vs 현재 스냅샷 업로드 시간 비교
        snapshot = db.query(Snapshot).filter(Snapshot.is_active == True).first()
        if snapshot and snapshot.uploaded_at:
            file_mtime = datetime.datetime.fromtimestamp(os.path.getmtime(latest))
            if file_mtime <= snapshot.uploaded_at:
                return f"새 파일 없음 (현재 스냅샷: {snapshot.week_label})"

        records, base_date = extract_records_from_excel(latest)
        label = os.path.splitext(os.path.basename(latest))[0]
        snap = save_snapshot(db, records, label, base_date)
        return f"업데이트 완료: {len(records)}건 저장 (label={label})"
    finally:
        if close_db:
            db.close()


def start_scheduler() -> BackgroundScheduler:
    scheduler = BackgroundScheduler(timezone="Asia/Seoul")
    scheduler.add_job(
        func=run_scheduled_update,
        trigger=CronTrigger(day_of_week="fri", hour=9, minute=0),
        id="weekly_update",
        name="매주 금요일 오전 9시 자동 업데이트",
        replace_existing=True,
        misfire_grace_time=3600,
    )
    scheduler.start()
    return scheduler