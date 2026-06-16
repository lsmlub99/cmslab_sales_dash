import os
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.responses import RedirectResponse
from dotenv import load_dotenv

load_dotenv()

from .database import Base, engine
from .routes import auth_routes, dashboard, admin
from .models import User
from .auth import hash_password
from .scheduler import start_scheduler


def _create_first_admin():
    """관리자 계정이 없으면 환경변수로 자동 생성."""
    from .database import SessionLocal
    db = SessionLocal()
    try:
        if db.query(User).filter(User.role == "admin").count() == 0:
            email = os.getenv("FIRST_ADMIN_EMAIL", "admin@cms-lab.co.kr")
            password = os.getenv("FIRST_ADMIN_PASSWORD", "changeme123!")
            db.add(User(
                email=email,
                hashed_password=hash_password(password),
                name="관리자",
                role="admin",
                allowed_teams=None,
                is_active=True,
            ))
            db.commit()
            print(f"[Init] 관리자 계정 생성: {email}")
    finally:
        db.close()


@asynccontextmanager
async def lifespan(app: FastAPI):
    Base.metadata.create_all(bind=engine)
    _create_first_admin()
    scheduler = start_scheduler()
    yield
    scheduler.shutdown(wait=False)


app = FastAPI(title="CMS Lab 매출 대시보드", lifespan=lifespan)

app.include_router(auth_routes.router)
app.include_router(dashboard.router)
app.include_router(admin.router)


@app.get("/")
async def root():
    return RedirectResponse("/dashboard")


@app.get("/health")
async def health():
    return {"status": "ok"}