# CMS Lab 매출 대시보드 — 인수인계 문서

> 내부 매출 데이터 시각화 및 관리 시스템. Excel 업로드 → 자동 파싱 → 팀/채널별 대시보드 제공.

---

## 시스템 구성

| 구분 | 내용 |
|---|---|
| **서버** | FastAPI + Uvicorn (Render 배포) |
| **DB** | Supabase PostgreSQL (`sales_dashboard` 스키마) |
| **프론트** | Jinja2 서버사이드 렌더링 (별도 빌드 없음) |
| **스케줄러** | APScheduler (자정 캐시 초기화) |
| **AI** | OpenAI GPT-4o-mini (어드민 AI 탭) |
| **MCP** | FastMCP Streamable HTTP (`/mcp` 엔드포인트) |

---

## 디렉토리 구조

```
sales_dashboard_web/
├── app/
│   ├── main.py              # FastAPI 앱 진입점, 마이그레이션
│   ├── models.py            # SQLAlchemy ORM 모델
│   ├── database.py          # DB 연결, 세션
│   ├── auth.py              # JWT 인증, 비밀번호 해시
│   ├── tab_registry.py      # 탭 목록 및 권한 헬퍼
│   ├── scheduler.py         # 자정 캐시 초기화 스케줄
│   ├── data/
│   │   └── parser.py        # Excel 파싱, HTML 생성, 캐시
│   ├── routes/
│   │   ├── auth_routes.py   # 로그인/로그아웃/회원가입
│   │   ├── dashboard.py     # 대시보드/비교 페이지
│   │   ├── admin.py         # 어드민 패널 (업로드, 사용자, 설정)
│   │   ├── api.py           # REST API (/api/v1/*)
│   │   ├── chat.py          # AI 챗봇 엔드포인트
│   │   └── mcp_gateway.py   # 통합 MCP 게이트웨이
│   └── templates/
│       └── admin.html       # 어드민 패널 UI
└── mcp_unified.py           # MCP 로컬 stdio 실행용 (Claude Desktop)
```

---

## Render 배포 설정

### 환경변수 (Render → Environment)

| 변수 | 설명 | 필수 |
|---|---|---|
| `DATABASE_URL` | Supabase Session Pooler URL | ✅ |
| `SECRET_KEY` | JWT 서명 키 (랜덤 문자열) | ✅ |
| `FIRST_ADMIN_EMAIL` | 최초 관리자 이메일 | ✅ |
| `FIRST_ADMIN_PASSWORD` | 최초 관리자 비밀번호 | ✅ |
| `OPENAI_API_KEY` | OpenAI API 키 (어드민 AI 탭용) | 선택 |
| `CHANNEL_MCP_API_KEY` | 채널 인사이트 MCP Bearer 토큰 | 선택 |
| `MCP_SECRET` | `/mcp` 엔드포인트 인증 토큰 | 선택 |

### Start Command
```
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

---

## 주요 기능

### 데이터 업로드
- 어드민 패널 → 데이터 관리 탭에서 Excel(.xlsx) 업로드
- 스냅샷 단위로 저장 (주차별 이력 보존)
- 활성 스냅샷이 대시보드에 표시됨

### 사용자/권한 관리
- **소속팀**: 특정 팀 데이터만 열람 가능 (NULL = 전체)
- **접근 탭**: 접근 허용 탭 목록 (NULL = 전체)
- **소속 그룹**: 그룹의 탭 권한을 상속
- 권한 우선순위: 개인 설정 → 그룹 기본값 → 전체 허용

### REST API (`/api/v1/*`)
- `GET /api/v1/summary` — 팀별 누적 실적 요약
- `GET /api/v1/teams/{team_name}` — 팀 월별 상세
- `GET /api/v1/snapshots` — 스냅샷 이력
- 인증: `X-API-Key` 헤더 (어드민 패널 → 시스템 설정에서 발급)

### 통합 MCP 게이트웨이 (`/mcp`)
- Sales 툴 3개 (내부 DB 직접 조회)
- 올리브영 툴 8개 / 쿠팡 4개 / 네이버 4개 (채널 MCP 프록시)
- 인증: `Authorization: Bearer {MCP_SECRET}`

---

## 로컬 개발

```bash
# 의존성 설치
pip install -r requirements.txt

# .env 파일 생성
DATABASE_URL=postgresql+psycopg2://...
SECRET_KEY=dev-secret-key
FIRST_ADMIN_EMAIL=admin@example.com
FIRST_ADMIN_PASSWORD=admin1234

# 실행
uvicorn app.main:app --reload
```

---

## DB 스키마 (`sales_dashboard` 스키마)

| 테이블 | 설명 |
|---|---|
| `users` | 사용자 계정, 권한, 탭/팀 접근 설정 |
| `teams` | 그룹 (탭 권한 기본값 포함) |
| `snapshots` | 업로드 이력 (주차별) |
| `sales_records` | 실적 데이터 (팀/월/실적/계획/전년) |
| `upload_history` | 파일 업로드 로그 |
| `app_config` | 앱 설정 (공지, API 키, 앱 이름 등) |

마이그레이션은 `app/main.py` `_run_migrations()`에서 자동 실행 (멱등).

---

## 문의
- 개발: 임승민 (lsmlub99@cms-lab.co.kr)
