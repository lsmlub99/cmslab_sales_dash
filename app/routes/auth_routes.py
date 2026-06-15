import os
from fastapi import APIRouter, Depends, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from ..database import get_db
from ..models import User
from ..auth import verify_password, hash_password, create_access_token, get_current_user

ALLOWED_DOMAIN = os.getenv("ALLOWED_EMAIL_DOMAIN", "cms-lab.co.kr")

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

    if user and verify_password(password, user.hashed_password) and not user.is_active:
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
async def register_page(request: Request, current_user=Depends(get_current_user)):
    if current_user:
        return RedirectResponse("/dashboard", status_code=302)
    return templates.TemplateResponse("register.html", {"request": request, "error": None, "success": False, "domain": ALLOWED_DOMAIN})


@router.post("/register")
async def register(
    request: Request,
    name: str = Form(...),
    email: str = Form(...),
    password: str = Form(...),
    password2: str = Form(...),
    db: Session = Depends(get_db),
):
    if not email.lower().endswith(f"@{ALLOWED_DOMAIN}"):
        return templates.TemplateResponse("register.html", {
            "request": request,
            "error": f"회사 이메일(@{ALLOWED_DOMAIN})만 가입 가능합니다.",
            "success": False, "domain": ALLOWED_DOMAIN,
        })
    if password != password2:
        return templates.TemplateResponse("register.html", {
            "request": request, "error": "비밀번호가 일치하지 않습니다.",
            "success": False, "domain": ALLOWED_DOMAIN,
        })
    if len(password) < 8:
        return templates.TemplateResponse("register.html", {
            "request": request, "error": "비밀번호는 8자 이상이어야 합니다.",
            "success": False, "domain": ALLOWED_DOMAIN,
        })
    if db.query(User).filter(User.email == email).first():
        return templates.TemplateResponse("register.html", {
            "request": request, "error": "이미 등록된 이메일입니다.",
            "success": False, "domain": ALLOWED_DOMAIN,
        })

    db.add(User(
        email=email,
        hashed_password=hash_password(password),
        name=name.strip(),
        role="viewer",
        is_active=False,
    ))
    db.commit()
    return templates.TemplateResponse("register.html", {
        "request": request, "error": None, "success": True, "domain": ALLOWED_DOMAIN,
    })


@router.get("/logout")
async def logout():
    response = RedirectResponse("/login", status_code=302)
    response.delete_cookie("access_token")
    return response
