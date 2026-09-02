FROM python:3.9-slim

# 安装编译 lxml 所需的系统依赖（这是关键修复）
RUN apt-get update && apt-get install -y \
    gcc \
    libxml2-dev \
    libxslt1-dev \
    zlib1g-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/

COPY . .
RUN mkdir -p /app/backend/tmp/uploads /app/backend/tmp/output

EXPOSE 5000

CMD ["gunicorn", "-b", "0.0.0.0:5000", "-w", "2", "--timeout", "120", "backend.app:app"]