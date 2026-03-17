# Official Playwright Python image — includes Chromium + all system deps
FROM mcr.microsoft.com/playwright/python:v1.49.1-jammy

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .

EXPOSE 5000

# 1 worker, 600s timeout — Playwright is blocking and not safe with multiple workers
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "600", "--workers", "1", "--threads", "1", "app:app"]
