# Official Playwright Python image — includes Chromium + all system deps
FROM mcr.microsoft.com/playwright/python:v1.49.1-jammy

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Ensure Playwright browsers are installed for the pip-installed version
RUN playwright install chromium

COPY app.py .

EXPOSE 5000

# gthread worker allows concurrent threads: one for render job, one for SSE stream, one for download
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--timeout", "600", "--workers", "1", "--worker-class", "gthread", "--threads", "4", "app:app"]
