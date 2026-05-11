bind = "0.0.0.0:8000"
workers = 3
worker_class = "sync"
worker_connections = 1000
timeout = 3600
keepalive = 2
errorlog = "logs/gunicorn_error.log"
accesslog = "logs/gunicorn_access.log"
capture_output = True
loglevel = "info"

# Fix for "Unhandled signal: cld" errors when using subprocess in workers
import signal

def post_worker_init(worker):
    # Reset SIGCHLD signal handler to default in the worker process
    # This prevents Gunicorn from complaining about unhandled signals 
    # when subprocesses (like the Perl script) exit.
    signal.signal(signal.SIGCHLD, signal.SIG_DFL)
