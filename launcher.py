#!/usr/bin/env python3
"""
运城培优成绩追踪看板 — 桌面应用入口
启动本地服务器 + 自动打开浏览器
支持 macOS / Windows / Linux
"""

import sys, os, time, threading, webbrowser, socket

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

HOST = "127.0.0.1"
PORT = 8899

def find_free_port(start=8899):
    for p in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex((HOST, p)) != 0:
                return p
    return start

def main():
    global PORT
    PORT = find_free_port(PORT)

    from app import app
    import uvicorn

    log_level = "warning" if getattr(sys, 'frozen', False) else "info"
    config = uvicorn.Config(app, host=HOST, port=PORT, log_level=log_level)
    server = uvicorn.Server(config)

    t = threading.Thread(target=server.run, daemon=True)
    t.start()

    for _ in range(200):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex((HOST, PORT)) == 0:
                break
        time.sleep(0.1)

    url = f"http://{HOST}:{PORT}"
    print(f"\n  📊 运城培优成绩追踪看板")
    print(f"  📍 {url}")
    print(f"  🛑 关闭此窗口即可停止服务\n")
    webbrowser.open(url)

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n  👋 服务已停止")
        server.should_exit = True

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        import multiprocessing
        multiprocessing.freeze_support()
    main()
