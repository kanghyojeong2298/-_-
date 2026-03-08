# Python 3.11 slim 이미지 사용
FROM python:3.11-slim

# 작업 디렉토리 설정
WORKDIR /app

# 시스템 패키지 설치 (pdfplumber, PDF 처리 등에 필요)
RUN apt-get update && apt-get install -y \
    gcc \
    g++ \
    poppler-utils \
    libpoppler-dev \
    libpoppler-cpp-dev \
    pkg-config \
    && rm -rf /var/lib/apt/lists/*

# Python 패키지 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 앱 파일 복사
COPY . .

# output 디렉토리 생성
RUN mkdir -p output

# Streamlit 설정: 포트 8080 (Cloud Run 기본값)
ENV PORT=8080
EXPOSE 8080

# 실행 명령
CMD ["python", "-m", "streamlit", "run", "app.py", \
     "--server.port=8080", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--server.enableCORS=false", \
     "--server.enableXsrfProtection=false"]
