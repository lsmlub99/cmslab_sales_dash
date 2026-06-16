import os
from sendgrid import SendGridAPIClient
from sendgrid.helpers.mail import Mail

SENDGRID_API_KEY = os.getenv("SENDGRID_API_KEY", "")
FROM_EMAIL = os.getenv("FROM_EMAIL", "")
APP_URL = os.getenv("APP_URL", "https://cmslab-sales-dash.onrender.com")


def send_verification_email(to_email: str, name: str, token: str) -> bool:
    verify_url = f"{APP_URL}/verify-email/{token}"

    html = f"""
<div style="font-family:'Malgun Gothic',sans-serif;max-width:480px;margin:0 auto;
     padding:32px;border:1px solid #e5e7eb;border-radius:10px">
  <h2 style="color:#1a56a0;margin-bottom:8px">이메일 인증</h2>
  <p style="color:#374151;margin-bottom:24px">안녕하세요, {name}님.<br>
  아래 버튼을 클릭해 이메일을 인증해주세요.</p>
  <a href="{verify_url}"
     style="display:inline-block;background:#1a56a0;color:#fff;text-decoration:none;
            padding:12px 28px;border-radius:7px;font-weight:600;font-size:14px">
    이메일 인증하기
  </a>
  <p style="color:#9ca3af;font-size:12px;margin-top:24px">
    이 메일은 CMS Lab 매출 대시보드 가입 요청에 의해 발송됐습니다.<br>
    본인이 요청하지 않은 경우 이 메일을 무시하세요.
  </p>
</div>"""

    if not SENDGRID_API_KEY:
        print(f"[Email] SENDGRID_API_KEY 미설정 — 인증 링크: {verify_url}")
        return False

    try:
        sg = SendGridAPIClient(SENDGRID_API_KEY)
        sg.send(Mail(
            from_email=FROM_EMAIL,
            to_emails=to_email,
            subject="[CMS Lab 매출 대시보드] 이메일 인증",
            html_content=html,
        ))
        return True
    except Exception as e:
        print(f"[Email] SendGrid 오류: {e}")
        return False
