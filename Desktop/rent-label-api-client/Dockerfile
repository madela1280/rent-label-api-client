FROM python:3.11-slim

# Tesseract + 한글/영어 데이터
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        tesseract-ocr \
        tesseract-ocr-kor \
        tesseract-ocr-eng && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python 의존성
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 앱 소스
COPY . .

# Render 기본 포트 사용
ENV PORT=10000
CMD ["uvicorn","app:app","--host","0.0.0.0","--port","10000"]
