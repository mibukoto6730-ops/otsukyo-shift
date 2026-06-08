"""大津京こと薬局 シフト作成アプリ ランチャー
ダブルクリックで起動（コンソールなし）
"""
import subprocess
import sys
import webbrowser
import time
import os
import socket

APP_DIR = os.path.dirname(os.path.abspath(__file__))

def find_free_port(start=8502):
    for p in range(start, start + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("localhost", p)) != 0:
                return p
    return start

port = find_free_port()

proc = subprocess.Popen(
    [sys.executable, "-m", "streamlit", "run", "app.py",
     "--server.port", str(port),
     "--server.headless", "true",
     "--browser.gatherUsageStats", "false"],
    cwd=APP_DIR,
    creationflags=subprocess.CREATE_NO_WINDOW,
)

time.sleep(3)
webbrowser.open(f"http://localhost:{port}")

proc.wait()
