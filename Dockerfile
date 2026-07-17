FROM python:3.11-slim

WORKDIR /app

# 安装 LibreOffice + 中文字体 + 系统依赖
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    libreoffice-core-nogui \
    fonts-noto-cjk \
    fonts-wqy-zenhei \
    libjpeg-dev zlib1g-dev libpng-dev \
    && rm -rf /var/lib/apt/lists/* \
    && apt-get clean

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# 创建运行时目录
RUN mkdir -p uploads temp meta

# Render/Railway 使用 PORT 环境变量
ENV PORT=8000
EXPOSE $PORT

# 单 worker 确保文件存储一致性，超时 180 秒适配大文件转换
CMD gunicorn app:app --bind 0.0.0.0:${PORT} --workers 1 --timeout 180 --access-logfile -
