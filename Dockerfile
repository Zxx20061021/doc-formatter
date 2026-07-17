FROM python:3.11-slim

WORKDIR /app

# 最小系统依赖
RUN apt-get update && apt-get install -y --no-install-recommends \
    libjpeg-dev zlib1g-dev libpng-dev \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 创建运行时目录
RUN mkdir -p uploads temp meta

ENV PORT=8000
EXPOSE $PORT

# 单 worker 确保文件存储一致性
CMD gunicorn app:app --bind 0.0.0.0:${PORT} --workers 1 --timeout 180 --access-logfile -
