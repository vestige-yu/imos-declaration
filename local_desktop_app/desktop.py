#!/usr/bin/env python3
import os
import socket
import threading
import webbrowser

import app


def find_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def main():
    port = int(os.environ.get("PORT") or find_free_port())
    server, url = app.run_in_thread(host="127.0.0.1", port=port)
    try:
        try:
            import webview
        except ImportError:
            webbrowser.open(url)
            threading.Event().wait()
            return

        window = webview.create_window("IMOS 报关单生成", url, width=1280, height=860)
        webview.start()
        return window
    finally:
        server.shutdown()
        server.server_close()


if __name__ == "__main__":
    main()
