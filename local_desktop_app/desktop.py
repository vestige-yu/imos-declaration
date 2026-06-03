#!/usr/bin/env python3
import os
import socket
import threading
import traceback
import webbrowser

import app


def log_startup_error(message):
    try:
        app.DATA_DIR.mkdir(parents=True, exist_ok=True)
        log_path = app.DATA_DIR / "startup.log"
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(message.rstrip() + "\n\n")
    except Exception:
        pass


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def open_browser_and_wait(url):
    webbrowser.open(url)
    threading.Event().wait()


def main():
    port = int(os.environ.get("PORT") or find_free_port())
    server, url = app.run_in_thread(host="127.0.0.1", port=port)
    try:
        try:
            import webview
        except ImportError:
            log_startup_error("pywebview unavailable, opened in the default browser.")
            open_browser_and_wait(url)
            return

        try:
            window = webview.create_window("IMOS 报关单生成", url, width=1280, height=860)
            webview.start()
            return window
        except Exception:
            log_startup_error("pywebview startup failed:\n" + traceback.format_exc())
            open_browser_and_wait(url)
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()
