FROM python:3.12-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY geocodare_od.py .
COPY web ./web

# Baza de date, cache-ul și fișierele încărcate — montați aici un volum persistent
ENV DATA_DIR=/data \
    PORT=8000 \
    COOKIE_SECURE=1 \
    TRUST_PROXY=1 \
    PYTHONUNBUFFERED=1
VOLUME /data
EXPOSE 8000

CMD ["python", "-m", "web"]
