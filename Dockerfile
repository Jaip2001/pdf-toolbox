FROM python:3.11-slim

# pymupdf + pillow ke liye system libs
RUN apt-get update && apt-get install -y --no-install-recommends \
        libgl1 libglib2.0-0 libjpeg62-turbo zlib1g \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY . .

ENV PORT=8080
EXPOSE 8080
CMD ["sh", "-c", "gunicorn -c gunicorn_conf.py wsgi:app"]
