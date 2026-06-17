"""
탭 레지스트리 — 앱에 존재하는 모든 탭을 여기서 관리.
새 탭 추가 시 이 목록에만 추가하면 어드민 권한 설정 UI에 자동 반영됨.
"""

TABS = [
    {"id": "dashboard", "label": "매출 대시보드", "route": "/dashboard"},
    {"id": "compare",   "label": "매출현황(표)",  "route": "/compare"},
]


def can_access_tab(user, tab_id: str) -> bool:
    """사용자가 해당 탭에 접근 가능한지 확인. allowed_tabs=NULL이면 전체 허용."""
    if not user.allowed_tabs:
        return True
    return tab_id in user.allowed_tabs
