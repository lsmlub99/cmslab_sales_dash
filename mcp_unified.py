"""
CMS Lab 통합 MCP 게이트웨이
────────────────────────────────────────────────
판매 대시보드(내부 REST API) + 올리브영/쿠팡/네이버 채널 인사이트(외부 MCP)를
하나의 MCP 서버로 묶어 Claude Desktop / Claude Code 에서 단일 서버로 사용.

실행 방법:
  - Claude Desktop (stdio):  python mcp_unified.py
  - HTTP 서버:               python mcp_unified.py --http
  - Render 배포:             start 커맨드 → python mcp_unified.py --http

필요 환경변수:
  SALES_API_URL        sales dashboard 배포 URL  (예: https://xxx.onrender.com)
  SALES_API_KEY        /api/v1/* 인증 키
  CHANNEL_MCP_URL      채널 MCP 서버 URL  (기본값: Vercel 엔드포인트)
  CHANNEL_MCP_API_KEY  채널 MCP Bearer 토큰
  MCP_PORT             HTTP 모드 포트 (기본 8001)
"""

import os
import sys
import asyncio
import httpx
from mcp.server.fastmcp import FastMCP

# ── 환경 설정 ──────────────────────────────────────────────────────────────────
SALES_API_URL       = os.getenv("SALES_API_URL",      "http://localhost:8000").rstrip("/")
SALES_API_KEY       = os.getenv("SALES_API_KEY",      "")
CHANNEL_MCP_URL     = os.getenv("CHANNEL_MCP_URL",    "https://oliveyoung-review.vercel.app/api/mcp")
CHANNEL_MCP_API_KEY = os.getenv("CHANNEL_MCP_API_KEY","")
MCP_PORT            = int(os.getenv("MCP_PORT", "8001"))

mcp = FastMCP("CMS Lab 통합 분석")


# ── 내부 헬퍼 ─────────────────────────────────────────────────────────────────

async def _sales(endpoint: str) -> str:
    """Sales Dashboard REST API 호출."""
    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            resp = await client.get(
                f"{SALES_API_URL}/api/v1/{endpoint}",
                headers={"X-API-Key": SALES_API_KEY},
            )
            resp.raise_for_status()
            return resp.text
    except httpx.HTTPStatusError as e:
        return f"[오류] HTTP {e.response.status_code}: {e.response.text[:200]}"
    except Exception as e:
        return f"[오류] Sales API 접속 실패: {e}"


async def _channel(tool_name: str, arguments: dict | None = None) -> str:
    """채널 인사이트 MCP 서버 툴 호출 (Streamable HTTP 프로토콜)."""
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


# ── Sales Dashboard 툴 3개 ────────────────────────────────────────────────────

@mcp.tool()
async def sales_get_summary() -> str:
    """팀별 Y2026 누적 실적 요약. 계획 대비 달성률, 전년 동기 대비 포함."""
    return await _sales("summary")


@mcp.tool()
async def sales_get_team_detail(team_name: str) -> str:
    """특정 팀의 월별 실적 상세 (실적/계획/전년동기).
    team_name: 팀 이름 (예: RBD1, RBD2, ...)
    """
    return await _sales(f"teams/{team_name}")


@mcp.tool()
async def sales_get_snapshots() -> str:
    """업로드된 주차별 스냅샷 이력 목록 (week_label, base_date, is_active)."""
    return await _sales("snapshots")


# ── 올리브영 툴 8개 ───────────────────────────────────────────────────────────

@mcp.tool()
async def oliveyoung_get_stats() -> str:
    """올리브영 전체 현황 — 리뷰 수, 평균 별점, 재구매율."""
    return await _channel("get_stats")


@mcp.tool()
async def oliveyoung_get_market_rankings(category: str = "") -> str:
    """올리브영 카테고리별 베스트 순위 Top 20.
    category: 조회할 카테고리명 (비워두면 전체)
    """
    args = {"category": category} if category else {}
    return await _channel("get_market_rankings", args)


@mcp.tool()
async def oliveyoung_get_promo_status() -> str:
    """올영픽·오늘의특가 등 프로모션 입점 현황."""
    return await _channel("get_promo_status")


@mcp.tool()
async def oliveyoung_get_negative_alerts() -> str:
    """최근 7일 부정 리뷰 급증 상품 알림."""
    return await _channel("get_negative_alerts")


@mcp.tool()
async def oliveyoung_get_product_stats() -> str:
    """상품별 리뷰 수, 평균 별점, 재구매율 전체 목록."""
    return await _channel("get_product_stats")


@mcp.tool()
async def oliveyoung_get_insights(goods_no: str = "") -> str:
    """긍·부정 키워드 Top 8, 피부 타입 분포 분석.
    goods_no: 상품번호 (비워두면 전체 통합)
    """
    args = {"goods_no": goods_no} if goods_no else {}
    return await _channel("get_insights", args)


@mcp.tool()
async def oliveyoung_get_new_products() -> str:
    """최근 30일 신규 등록 상품 및 리뷰 증가 속도."""
    return await _channel("get_new_products")


@mcp.tool()
async def oliveyoung_get_today_ranking() -> str:
    """오늘 시간별 자사 상품 순위 타임라인."""
    return await _channel("get_today_ranking")


# ── 쿠팡 툴 4개 ───────────────────────────────────────────────────────────────

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
    """쿠팡 실구매 리뷰 내용 조회.
    product_id: 상품 ID (비워두면 전체 최신)
    """
    args = {"product_id": product_id} if product_id else {}
    return await _channel("get_coupang_reviews", args)


# ── 네이버 툴 4개 ─────────────────────────────────────────────────────────────

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
    """선케어 카테고리 경쟁사 현황 분석."""
    return await _channel("get_naver_market")


@mcp.tool()
async def naver_get_insight() -> str:
    """AI 생성 시장 분석 인사이트."""
    return await _channel("get_naver_insight")


# ── 진입점 ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if "--http" in sys.argv:
        # HTTP 서버 모드 (Render 배포 또는 원격 MCP)
        print(f"[MCP Gateway] HTTP 서버 시작: 0.0.0.0:{MCP_PORT}")
        mcp.run(transport="streamable-http", host="0.0.0.0", port=MCP_PORT)
    else:
        # stdio 모드 (Claude Desktop / Claude Code 로컬 연결)
        mcp.run()
