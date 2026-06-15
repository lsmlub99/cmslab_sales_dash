#!/usr/bin/env python3
"""
기존 Excel 데이터를 DB에 초기 적재하는 1회성 스크립트.

Usage:
  # 기본 (파일 → DB, active=True)
  python -m app.data.seed --file "../sales_dash_cmslab/실행계획_취합자료_6월 2주차.xlsx"

  # 레이블 지정
  python -m app.data.seed --file "..." --label "6월 2주차"

  # Backup 폴더 이력 적재 (이전 파일들, is_active=False)
  python -m app.data.seed --file "Backup/파일.xlsx" --no-activate
  # 마지막 파일만 active=True 로 재실행
  python -m app.data.seed --file "최신파일.xlsx"
"""
import argparse, os, sys

# 패키지 루트를 path에 추가
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", ".."))

from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), "..", "..", ".env"))

from app.database import SessionLocal, Base, engine
from app.models import Snapshot, SalesRecord
from app.data.parser import extract_records_from_excel, save_snapshot


def main():
    ap = argparse.ArgumentParser(description="Excel → DB 초기 적재")
    ap.add_argument("--file", required=True, help="Excel 파일 경로")
    ap.add_argument("--label", default="", help="스냅샷 레이블 (예: '6월 2주차')")
    ap.add_argument(
        "--no-activate",
        action="store_true",
        help="기존 active 스냅샷 유지 (이력 보존용)",
    )
    args = ap.parse_args()

    if not os.path.exists(args.file):
        print(f"[ERROR] File not found: {args.file}")
        sys.exit(1)

    Base.metadata.create_all(bind=engine)
    db = SessionLocal()

    try:
        print(f"[SEED] Parsing: {args.file}")
        records, base_date = extract_records_from_excel(args.file)
        label = args.label or os.path.splitext(os.path.basename(args.file))[0]

        if args.no_activate:
            snapshot = Snapshot(
                week_label=label,
                base_date=base_date,
                is_active=False,
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
        else:
            snapshot = save_snapshot(db, records, label, base_date)

        active_str = "[ACTIVE]" if snapshot.is_active else "[HISTORY]"
        print(f"{active_str} Done: {len(records)} records saved")
        print(f"   label={label}, base_date={base_date}, snapshot_id={snapshot.id}")
    finally:
        db.close()


if __name__ == "__main__":
    main()
