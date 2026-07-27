FROM python:3.11-slim
WORKDIR /app

# Install system utilities required to fetch and manage web dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    wget \
    gnupg \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Install Playwright browser engines and system shared libraries inside the container
RUN playwright install chromium
RUN playwright install-deps chromium

COPY . .
CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 3600 app:app
