FROM python:3.11-slim

WORKDIR /srv

# 시스템 의존성 (psycopg2-binary 빌드용)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libpq-dev gcc && \
    rm -rf /var/lib/apt/lists/*

# Python 의존성
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 원본 스크립트 복사 (extract_data / make_html 재활용)
COPY ../sales_dash_cmslab/매출\ Dashboard_vf.py   ./sales_dash_cmslab/매출\ Dashboard_vf.py
COPY ../sales_dash_cmslab/매출_선택비교_vf.py      ./sales_dash_cmslab/매출_선택비교_vf.py
COPY ../sales_dash_cmslab/chart.umd.js            ./sales_dash_cmslab/chart.umd.js

# 웹 앱 복사
COPY app/ ./app/

ENV DASHBOARD_SRC_PATH=/srv/sales_dash_cmslab
ENV PORT=8000

EXPOSE 8000

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000"]