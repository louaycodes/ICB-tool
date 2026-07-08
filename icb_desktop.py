"""
ICB Desktop — Lance l'application ICB dans une fenetre native.
Double-cliquez sur ce fichier ou executez : python3 icb_desktop.py
"""

import threading
import time
import sys
import socket

PORT = 5050


def port_is_open(port, host="127.0.0.1", timeout=0.5):
    """Verifie si le serveur Flask est pret."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(timeout)
        try:
            s.connect((host, port))
            return True
        except (ConnectionRefusedError, OSError):
            return False


def main():
    # --- 0. Liberer le port s'il est deja occupe ---
    if port_is_open(PORT):
        import subprocess
        try:
            result = subprocess.run(
                ["lsof", "-ti", f":{PORT}"], capture_output=True, text=True
            )
            pids = result.stdout.strip().split("\n")
            for pid in pids:
                if pid:
                    subprocess.run(["kill", "-9", pid], capture_output=True)
            print(f"Port {PORT} libere (ancien processus arrete).")
            import time as _t
            _t.sleep(0.5)
        except Exception:
            pass

    # --- 1. Demarrer le serveur Flask en arriere-plan ---
    from icb_app import start_server

    server_thread = threading.Thread(target=start_server, args=(PORT,), daemon=True)
    server_thread.start()

    # Attendre que le serveur soit pret (max 15 secondes)
    print("Demarrage du serveur ICB...")
    for _ in range(30):
        if port_is_open(PORT):
            break
        time.sleep(0.5)
    else:
        print("ERREUR : Le serveur Flask n'a pas demarre a temps.")
        sys.exit(1)

    print(f"Serveur pret sur le port {PORT}.")

    # --- 2. Ouvrir la fenetre desktop native ---
    import webview

    window = webview.create_window(
        title="ICB Analysis — Capgemini",
        url=f"http://127.0.0.1:{PORT}",
        width=1400,
        height=900,
        resizable=True,
        min_size=(900, 600),
        text_select=True,
    )

    # Bloquer ici jusqu'a fermeture de la fenetre
    webview.start()

    print("Fenetre fermee. Arret de l'application.")


if __name__ == "__main__":
    main()
