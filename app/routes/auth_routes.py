import os
from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from ..database import get_db
from ..models import User
from ..auth import verify_password, hash_password, create_access_token, get_current_user

ALLOWED_DOMAIN = os.getenv("ALLOWED_EMAIL_DOMAIN", "wonik.com")

_FALLBACK_TEAMS = [
    "RBD1팀", "RBD2팀", "동북아MC팀", "Global사업팀",
    "GEC팀", "일본사업팀", "중국사업팀", "메디컬팀",
]


def _get_teams(db) -> list:
    try:
        from ..models import Team
        rows = db.query(Team).filter(Team.is_active == True).order_by(Team.display_order).all()
        return [r.name for r in rows] or _FALLBACK_TEAMS
    except Exception:
        return _FALLBACK_TEAMS


def _get_allowed_domain(db) -> str:
    try:
        from ..models import AppConfig
        row = db.query(AppConfig).filter(AppConfig.key == "allowed_domain").first()
        return row.value if row and row.value else ALLOWED_DOMAIN
    except Exception:
        return ALLOWED_DOMAIN


def _get_admin_contact(db) -> str:
    try:
        from ..models import AppConfig
        row = db.query(AppConfig).filter(AppConfig.key == "admin_contact").first()
        return row.value if row else ""
    except Exception:
        return ""


router = APIRouter()
_tpl_dir = os.path.join(os.path.dirname(__file__), "..", "templates")
templates = Jinja2Templates(directory=_tpl_dir)


@router.get("/login", response_class=HTMLResponse)
async def login_page(request: Request, current_user=Depends(get_current_user)):
    if current_user:
        return RedirectResponse("/dashboard", status_code=302)
    return templates.TemplateResponse("login.html", {"request": request, "error": None})


@router.post("/login")
async def login(
    request: Request,
    email: str = Form(...),
    password: str = Form(...),
    db: Session = Depends(get_db),
):
    user: User = db.query(User).filter(User.email == email).first()

    if user and verify_password(password, user.hashed_password):
        if not user.is_active:
            return templates.TemplateResponse(
                "login.html",
                {"request": request, "error": "승인 대기 중입니다. 관리자에게 문의하세요.", "pending": True},
            )

    if not user or not verify_password(password, user.hashed_password) or not user.is_active:
        return templates.TemplateResponse(
            "login.html",
            {"request": request, "error": "이메일 또는 비밀번호가 올바르지 않습니다."},
        )

    token = create_access_token({"sub": str(user.id)})
    response = RedirectResponse("/dashboard", status_code=302)
    response.set_cookie("access_token", token, httponly=True, max_age=60 * 60 * 8, samesite="lax")
    return response


@router.get("/register", response_class=HTMLResponse)
async def register_page(request: Request, current_user=Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user:
        return RedirectResponse("/dashboard", status_code=302)
    domain = _get_allowed_domain(db)
    teams = _get_teams(db)
    return templates.TemplateResponse("register.html", {
        "request": request, "error": None, "success": False,
        "domain": domain, "teams": teams, "admin_contact": "",
    })


@router.post("/register")
async def register(
    request: Request,
    name: str = Form(...),
    email: str = Form(...),
    team: str = Form(""),
    password: str = Form(...),
    password2: str = Form(...),
    db: Session = Depends(get_db),
):
    domain = _get_allowed_domain(db)
    teams = _get_teams(db)
    contact = _get_admin_contact(db)
    ctx = {"request": request, "success": False, "domain": domain, "teams": teams, "admin_contact": contact}

    if not email.lower().endswith(f"@{domain}"):
        return templates.TemplateResponse("register.html", {**ctx, "error": f"회사 이메일(@{domain})만 가입 가능합니다."})
    if not team or team not in teams:
        return templates.TemplateResponse("register.html", {**ctx, "error": "부서를 선택해주세요."})
    if password != password2:
        return templates.TemplateResponse("register.html", {**ctx, "error": "비밀번호가 일치하지 않습니다."})
    if len(password) < 8:
        return templates.TemplateResponse("register.html", {**ctx, "error": "비밀번호는 8자 이상이어야 합니다."})

    if db.query(User).filter(User.email == email).first():
        return templates.TemplateResponse("register.html", {**ctx, "error": "이미 등록된 이메일입니다."})

    db.add(User(
        email=email,
        hashed_password=hash_password(password),
        name=name.strip(),
        role="viewer",
        allowed_teams=[team],
        is_active=False,
        email_verified=True,
    ))
    db.commit()
    return templates.TemplateResponse("register.html", {
        **ctx, "error": None, "success": True,
    })



@router.get("/check-email")
async def check_email(email: str, db: Session = Depends(get_db)):
    """이메일 사용 가능 여부 AJAX 체크."""
    domain = _get_allowed_domain(db)
    if not email.lower().endswith(f"@{domain}"):
        return {"available": False, "reason": "domain", "message": f"@{domain} 도메인만 가입 가능합니다"}
    if db.query(User).filter(User.email == email).first():
        return {"available": False, "reason": "taken", "message": "이미 등록된 이메일입니다"}
    return {"available": True, "message": "사용 가능한 이메일입니다"}


@router.get("/logout")
async def logout():
    response = RedirectResponse("/login", status_code=302)
    response.delete_cookie("access_token")
    return response
