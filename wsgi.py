"""WSGI entry point (gunicorn / waitress / Render / Railway sab isi ko use karte hain)"""
import os
from app import app

if __name__ == '__main__':
    # Windows / simple local run
    from waitress import serve
    port = int(os.environ.get('PORT', 8000))
    print(f'* serving on http://0.0.0.0:{port}')
    serve(app, host='0.0.0.0', port=port, threads=8, channel_timeout=900)
