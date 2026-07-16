FROM python:3.11-slim

WORKDIR /app

# Install system dependencies for Pillow
RUN apt-get update && apt-get install -y --no-install-recommends \
    libjpeg-dev zlib1g-dev libpng-dev \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create temp directories
RUN mkdir -p uploads temp

EXPOSE $PORT

CMD gunicorn app:app --bind 0.0.0.0:${PORT:-8000} --workers 2 --timeout 120
