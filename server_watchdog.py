import socket
import subprocess
import sys
import time
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_PATH = os.path.join(SCRIPT_DIR, 'app.py')

# Ensure we use python.exe for app.py
base_dir = os.path.dirname(sys.executable)
candidate_python = os.path.join(base_dir, 'python.exe')
PYTHON_EXE = candidate_python if os.path.exists(candidate_python) else sys.executable

WATCHDOG_LOCK_PORT = 5025  # Ensures single instance

def acquire_lock():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(('127.0.0.1', WATCHDOG_LOCK_PORT))
        s.listen(1)
        return s
    except socket.error:
        # Already running
        sys.exit(0)

def main():
    lock_sock = acquire_lock()
    log_file = os.path.join(SCRIPT_DIR, 'server_watchdog.log')

    def log(msg):
        entry = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n"
        try:
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(entry)
        except Exception:
            pass

    log("=== Geam Digital Server Watchdog Started ===")

    while True:
        try:
            log(f"Launching app.py using {PYTHON_EXE}")
            # Run without shell, inheriting environment
            proc = subprocess.Popen(
                [PYTHON_EXE, APP_PATH],
                cwd=SCRIPT_DIR,
                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
            )
            proc.wait()
            ret = proc.returncode
            log(f"app.py exited with code {ret}. Restarting in 2 seconds...")
        except Exception as e:
            log(f"Exception while running app: {e}")
        time.sleep(2)

if __name__ == '__main__':
    main()
