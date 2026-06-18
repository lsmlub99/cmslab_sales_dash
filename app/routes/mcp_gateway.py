"""
CMS Lab 통합 MCP 게이트웨이 — FastAPI 마운트용
/mcp 경로로 Streamable HTTP MCP 서버 노출.

툴 목록:
  - sales_*       : 내부 DB 직접 조회
  - oliveyoung_*  : 채널 인사이트 MCP 프록시 (8개)
  - coupang_*     : 채널 인사이트 MCP 프록시 (4개)
  - naver_*       : 채널 인사이트 MCP 프록시 (4개)

환경변수:
  CHANNEL_MCP_URL      채널 MCP 서버 URL (기본: Vercel 엔드포인트)
  CHANNEL_MCP_API_KEY  채널 MCP Bearer 토큰
"""

import os
import re
from collections import defaultdict
from typing import Optional

import httpx
from mcp.server.fastmcp import FastMCP
from sqlalchemy.orm import Session
from sqlalchemy import func

from ..database import SessionLocal
from ..models import Snapshot, SalesRecord, AppConfig

CHANNEL_MCP_URL     = os.getenv("CHANNEL_MCP_URL",    "https://oliveyoung-review.vercel.app/api/mcp")
CHANNEL_MCP_API_KEY = os.getenv("CHANNEL_MCP_API_KEY","")

mcp = FastMCP("CMS Lab 통합 분석")


# ── 채널 MCP 프록시 ───────────────────────────────────────────────────────────

async def _channel(tool_name: str, arguments: dict | None = None) -> str:
    if arguments is None:
        arguments = {}
    try:
        from mcp.client.streamable_http import streamablehttp_client
        from mcp import ClientSession

        headers: dict[str, str] = {}
        if CHANNEL_MCP_API_KEY:
            headers["Authorization"] = f"Bearer {CHANNEL_MCP_API_KEY}"

        async with streamablehttp_client(CHANNEL_MCP_URL, headers=headers) as (read, write, _):
            async with ClientSession(read, write) as session:
                await session.initialize()
                result = await session.call_tool(tool_name, arguments)

        if result.content:
            return result.content[0].text
        return "(결과 없음)"
    except Exception as e:
        return f"[오류] 채널 MCP 호출 실패 ({tool_name}): {e}"


# ── Sales DB 헬퍼 ─────────────────────────────────────────────────────────────

def _get_db() -> Session:
    return SessionLocal()


def _sales_summary() -> str:
    db = _get_db()
    try:
        snap = db.query(Snapshot).filter(Snapshot.is_active == True).first()
        if not snap:
            return "현재 활성 스냅샷이 없습니다."

        m = re.search(r'(\d{1,2})월', snap.base_date)
        base_month = int(m.group(1)) if m else 12

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
        if not rows:
            return "데이터가 없습니다."

        lines = [
            f"기준: {snap.week_label} ({snap.base_date})",
            f"누적: 1~{base_month}월",
            "",
            "팀명 | 실적(백만원) | 계획(백만원) | 달성률 | 전년동기",
            "─" * 55,
        ]
        ta = tp = ty = 0.0
        for r in rows:
            a, p, y = float(r.actual or 0), float(r.plan or 0), float(r.y2025 or 0)
            ach = f"{a/p*100:.1f}%" if p else "N/A"
            lines.append(f"{r.team:<10} {a:>10,.0f}  {p:>10,.0f}  {ach:>7}  {y:>10,.0f}")
            ta += a; tp += p; ty += y

        ach_total = f"{ta/tp*100:.1f}%" if tp else "N/A"
        lines += ["─" * 55, f"{'전사합계':<10} {ta:>10,.0f}  {tp:>10,.0f}  {ach_total:>7}  {ty:>10,.0f}"]
        return "\n".join(lines)
    finally:
        db.close()


def _sales_team_detail(team_name: str) -> str:
    db = _get_db()
    try:
        snap = db.query(Snapshot).filter(Snapshot.is_active == True).first()
        if not snap:
            return "현재 활성 스냅샷이 없습니다."

        rows = (
            db.query(SalesRecord)
            .filter(SalesRecord.snapshot_id == snap.id, SalesRecord.team == team_name)
            .order_by(SalesRecord.month)
            .all()
        )
        if not rows:
            return f"'{team_name}' 팀 데이터가 없습니다. 팀 이름을 확인하세요."

        lines = [
            f"[{team_name}] 월별 실적 — {snap.week_label} ({snap.base_date})",
            "",
            "월  | 실적(백만원) | 계획(백만원) | 달성률 | 전년동기",
            "─" * 52,
        ]
        for r in rows:
            a, p, y = float(r.actual or 0), float(r.plan or 0), float(r.y2025 or 0)
            ach = f"{a/p*100:.1f}%" if p else "N/A"
            lines.append(f"{r.month:>2}월  {a:>10,.0f}  {p:>10,.0f}  {ach:>7}  {y:>10,.0f}")
        return "\n".join(lines)
    finally:
        db.close()


def _sales_snapshots() -> str:
    db = _get_db()
    try:
        snaps = db.query(Snapshot).order_by(Snapshot.uploaded_at.desc()).all()
        if not snaps:
            return "스냅샷이 없습니다."
        lines = ["id  | 주차      | 기준일     | 활성  | 업로드시각(KST)", "─" * 60]
        for s in snaps:
            active = "✓" if s.is_active else " "
            lines.append(f"{s.id:<4}  {s.week_label:<10}  {s.base_date:<10}  {active:<5}  {s.uploaded_at}")
        return "\n".join(lines)
    finally:
        db.close()


# ── Sales 툴 ──────────────────────────────────────────────────────────────────

@mcp.tool()
def sales_get_summary() -> str:
    """팀별 Y2026 누적 실적 요약 (계획 대비 달성률, 전년 동기 대비 포함)."""
    return _sales_summary()


@mcp.tool()
def sales_get_team_detail(team_name: str) -> str:
    """특정 팀의 월별 실적 상세.
    team_name: 팀 이름 (예: RBD1, RBD2 — 정확한 이름 필요)
    """
    return _sales_team_detail(team_name)


@mcp.tool()
def sales_get_snapshots() -> str:
    """업로드된 주차별 스냅샷 이력 목록."""
    return _sales_snapshots()


# ── 올리브영 툴 ───────────────────────────────────────────────────────────────

@mcp.tool()
async def oliveyoung_get_stats() -> str:
    """올리브영 전체 현황 — 리뷰 수, 평균 별점, 재구매율."""
    return await _channel("get_stats")


@mcp.tool()
async def oliveyoung_get_market_rankings(category: str = "") -> str:
    """올리브영 카테고리별 베스트 순위 Top 20.
    category: 카테고리명 (비워두면 전체)
    """
    return await _channel("get_market_rankings", {"category": category} if category else {})


@mcp.tool()
async def oliveyoung_get_promo_status() -> str:
    """올영픽·오늘의특가 프로모션 입점 현황."""
    return await _channel("get_promo_status")


@mcp.tool()
async def oliveyoung_get_negative_alerts() -> str:
    """최근 7일 부정 리뷰 급증 상품 알림."""
    return await _channel("get_negative_alerts")


@mcp.tool()
async def oliveyoung_get_product_stats() -> str:
    """올리브영 상품별 리뷰 수, 평균 별점, 재구매율."""
    return await _channel("get_product_stats")


@mcp.tool()
async def oliveyoung_get_insights(goods_no: str = "") -> str:
    """긍·부정 키워드 Top 8, 피부 타입 분포 분석.
    goods_no: 상품번호 (비워두면 전체 통합)
    """
    return await _channel("get_insights", {"goods_no": goods_no} if goods_no else {})


@mcp.tool()
async def oliveyoung_get_new_products() -> str:
    """최근 30일 신규 등록 상품 및 리뷰 증가 속도."""
    return await _channel("get_new_products")


@mcp.tool()
async def oliveyoung_get_today_ranking() -> str:
    """오늘 시간별 자사 상품 순위 타임라인."""
    return await _channel("get_today_ranking")


# ── 쿠팡 툴 ──────────────────────────────────────────────────────────────────

@mcp.tool()
async def coupang_get_stats() -> str:
    """쿠팡 전체 현황 — 리뷰 수, 평균 별점, 로켓배송 비율 등."""
    return await _channel("get_coupang_stats")


@mcp.tool()
async def coupang_get_product_stats() -> str:
    """쿠팡 상품별 리뷰 수·평점 목록."""
    return await _channel("get_coupang_product_stats")


@mcp.tool()
async def coupang_get_rankings() -> str:
    """쿠팡 검색·카테고리 순위."""
    return await _channel("get_coupang_rankings")


@mcp.tool()
async def coupang_get_reviews(product_id: str = "") -> str:
    """쿠팡 실구매 리뷰 내용.
    product_id: 상품 ID (비워두면 최신 전체)
    """
    return await _channel("get_coupang_reviews", {"product_id": product_id} if product_id else {})


# ── 네이버 툴 ─────────────────────────────────────────────────────────────────

@mcp.tool()
async def naver_get_trends() -> str:
    """네이버 DataLab 검색 트렌드 (최근 8주)."""
    return await _channel("get_naver_trends")


@mcp.tool()
async def naver_get_search_ranks() -> str:
    """네이버 쇼핑 검색 결과 순위."""
    return await _channel("get_naver_search_ranks")


@mcp.tool()
async def naver_get_market() -> str:
    """선케어 카테고리 경쟁사 현황."""
    return await _channel("get_naver_market")


@mcp.tool()
async def naver_get_insight() -> str:
    """AI 생성 시장 분석 인사이트."""
    return await _channel("get_naver_insight")


# ── FastAPI 마운트용 ASGI 앱 ──────────────────────────────────────────────────

def get_mcp_app():
    """main.py에서 app.mount('/mcp', get_mcp_app()) 로 사용."""
    return mcp.get_asgi_app()
