"""Production server config (gunicorn) — timeouts bade rakhe hain kyunki
bade PDF convert karne me 1-2 minute lag sakte hain."""
import multiprocessing, os

bind = f"0.0.0.0:{os.environ.get('PORT', '8000')}"
workers = int(os.environ.get('WEB_CONCURRENCY', min(4, max(1, multiprocessing.cpu_count() // 2))))
threads = int(os.environ.get('WEB_THREADS', 4))
timeout = 600            # 10 min — bade PDF ke liye
graceful_timeout = 60
keepalive = 5
max_requests = 200
max_requests_jitter = 40
accesslog = '-'
errorlog = '-'
loglevel = 'info'
