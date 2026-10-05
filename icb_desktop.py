"""
ICB Desktop — Launches the ICB application in a native window.
"""

import threading
import time
import sys
import socket
import os
import subprocess

PORT = 5050

def port_is_open(port, host="127.0.0.1", timeout=0.5):
    """Check if the Flask server is ready."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(timeout)
        try:
            s.connect((host, port))
            return True
        except (ConnectionRefusedError, OSError):
            return False

def free_port(port):
    """Free the port in a cross-platform way (Windows & macOS/Linux)."""
    if not port_is_open(port):
        return
    try:
        if os.name == 'nt':  # Windows
            cmd = f'for /f "tokens=5" %a in (\'netstat -aon ^| findstr :{port}\') do taskkill /f /pid %a'
            subprocess.run(cmd, shell=True, capture_output=True)
        else:  # macOS / Linux
            result = subprocess.run(["lsof", "-ti", f":{port}"], capture_output=True, text=True)
            pids = result.stdout.strip().split("\n")
            for pid in pids:
                if pid:
                    subprocess.run(["kill", "-9", pid], capture_output=True)
        time.sleep(0.5)
    except Exception:
        pass

def main():
    # --- 0. Free the port if already in use ---
    free_port(PORT)

    # --- 1. Start the Flask server in the background ---
    from icb_app import start_server

    server_thread = threading.Thread(target=start_server, args=(PORT,), daemon=True)
    server_thread.start()

    # Wait for the server to be ready (max 15 seconds)
    print("Starting ICB server...")
    for _ in range(30):
        if port_is_open(PORT):
            break
        time.sleep(0.5)
    else:
        print("ERROR: Flask server did not start in time.")
        sys.exit(1)

    print(f"Server ready on port {PORT}.")

    # --- 2. Open the native desktop window ---
    import webview

    webview.create_window(
        title="ICB Analysis — Capgemini",
        url=f"http://127.0.0.1:{PORT}",
        width=1400,
        height=900,
        resizable=True,
        min_size=(900, 600),
        text_select=True,
    )

    webview.start()
    print("Window closed. Stopping the application.")

if __name__ == "__main__":
    main()