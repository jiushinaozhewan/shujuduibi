"""Streamlit 启动器（PyInstaller 入口）
- 打包后：sys._MEIPASS/app.py 是 streamlit 脚本
- 自动寻找可用端口；用默认浏览器打开
"""
from __future__ import annotations
import sys
import os
import socket
import webbrowser
import threading
import time
from pathlib import Path


def find_free_port(preferred: int = 8501) -> int:
    for p in (preferred, 8502, 8503, 8504, 0):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            s.bind(("127.0.0.1", p))
            port = s.getsockname()[1]
            s.close()
            return port
        except OSError:
            continue
    return preferred


def open_browser_later(url: str, delay: float = 2.0):
    def _open():
        time.sleep(delay)
        webbrowser.open_new(url)
    threading.Thread(target=_open, daemon=True).start()


def main():
    # 定位 app.py：frozen 时从 _MEIPASS，否则当前目录
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    else:
        base = Path(__file__).resolve().parent
    app_path = base / "app.py"
    if not app_path.exists():
        print(f"[ERROR] 找不到 app.py: {app_path}")
        input("按回车退出...")
        sys.exit(1)

    port = find_free_port(8501)
    url = f"http://localhost:{port}"
    print("=" * 50)
    print("  跨表核对 · Web 版")
    print(f"  地址：{url}")
    print("  浏览器将在 2 秒后自动打开")
    print("  关闭此窗口即退出")
    print("=" * 50)
    open_browser_later(url)

    # 构造 streamlit 命令行参数并调用其 cli 入口
    sys.argv = [
        "streamlit",
        "run",
        str(app_path),
        f"--server.port={port}",
        "--server.headless=true",
        "--global.developmentMode=false",
        "--browser.gatherUsageStats=false",
        "--server.fileWatcherType=none",
    ]
    import streamlit.web.cli as stcli
    sys.exit(stcli.main())


if __name__ == "__main__":
    main()
