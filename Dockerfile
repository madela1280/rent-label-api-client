FROM python:3.10-slim

# OS 의존 패키지 (tesseract 포함, 경량 설치)
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr tesseract-ocr-kor tesseract-ocr-eng libgl1 libglib2.0-0 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 런타임 환경
ENV TESSERACT_CMD=/usr/bin/tesseract
ENV PYTHONUNBUFFERED=1

# FastAPI 실행 (Render의 PORT 환경변수 대응)
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]

