"""Control Logístico — app de escritorio (Python + HTML).

Ejecutar:  python logistica.py            -> abre la app en su propia ventana
           python logistica.py --navegador -> abre en una pestaña normal
           python logistica.py --sin-ventana --puerto 8765 (solo servidor)

La ventana se abre con pywebview si está instalado; si no, con Edge/Chrome en
modo aplicación (--app). Al cerrar la ventana, el servidor se detiene solo.
"""
import argparse
import json
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

sys.path.insert(0, str(Path(__file__).parent))
import motor  # noqa: E402

WEB = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "web"
CON = motor.conectar()
ULTIMO_PING = [time.time()]


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, data, code=200):
        body = json.dumps(data, ensure_ascii=False, default=str).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n) if n else b""

    def _q(self):
        return {k: v[0] for k, v in parse_qs(urlparse(self.path).query).items()}

    def do_GET(self):
        ruta = urlparse(self.path).path
        q = self._q()
        try:
            if ruta in ("/", "/index.html"):
                body = (WEB / "index.html").read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            elif ruta == "/api/ping":
                ULTIMO_PING[0] = time.time()
                self._json({"ok": True})
            elif ruta == "/api/meta":
                self._json(motor.meta(CON))
            elif ruta == "/api/tablero":
                self._json(motor.tablero(CON, q.get("desde"), q.get("hasta")))
            elif ruta == "/api/remisiones":
                self._json(motor.listar_remisiones(CON, **{k: q.get(k, "") for k in
                                                           ("q", "fase", "semaforo", "pais", "digitador", "estado")}))
            elif ruta == "/api/remision":
                d = motor.detalle(CON, q.get("file", ""))
                self._json(d if d else {"error": "No existe"}, 200 if d else 404)
            elif ruta == "/api/exportar":
                data = motor.exportar_excel(CON)
                nombre = f"control_logistico_{datetime.now():%Y%m%d_%H%M}.xlsx"
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f'attachment; filename="{nombre}"')
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)
            else:
                self._json({"error": "Ruta no encontrada"}, 404)
        except Exception as e:  # noqa: BLE001
            self._json({"error": str(e)}, 500)

    def do_POST(self):
        ruta = urlparse(self.path).path
        q = self._q()
        try:
            if ruta == "/api/importar":
                self._json(motor.importar_consolidado(CON, self._body(), q.get("hoja") or None))
                return
            d = json.loads(self._body() or b"{}")
            if ruta == "/api/evento":
                self._json(motor.guardar_evento(CON, d))
            elif ruta in ("/api/reembarque", "/api/remision", "/api/sae"):
                self._json(motor.guardar(CON, ruta.rsplit("/", 1)[1], d))
            elif ruta == "/api/parametros":
                motor.guardar_params(CON, d)
                self._json({"ok": True})
            elif ruta == "/api/persona":
                with motor._lock:
                    pid = motor.guardar_persona(CON, d.get("nombre"), d.get("rol") or "DIGITADOR")
                    CON.commit()
                self._json({"id": pid})
            elif ruta == "/api/demo":
                motor.cargar_demo(CON)
                self._json({"ok": True})
            elif ruta == "/api/vaciar":
                motor.vaciar(CON)
                self._json({"ok": True})
            else:
                self._json({"error": "Ruta no encontrada"}, 404)
        except ValueError as e:
            self._json({"error": str(e)}, 400)
        except Exception as e:  # noqa: BLE001
            self._json({"error": f"{type(e).__name__}: {e}"}, 500)

    def do_DELETE(self):
        if urlparse(self.path).path == "/api/evento":
            motor.borrar_evento(CON, int(self._q().get("id", 0)))
            self._json({"ok": True})
        else:
            self._json({"error": "Ruta no encontrada"}, 404)


def puerto_libre(preferido):
    for p in [preferido] + list(range(preferido + 1, preferido + 50)):
        with socket.socket() as s:
            if s.connect_ex(("127.0.0.1", p)) != 0:
                return p
    return 0


def buscar_navegador_app():
    candidatos = []
    if os.name == "nt":
        for base in (os.environ.get("PROGRAMFILES(X86)"), os.environ.get("PROGRAMFILES"), os.environ.get("LOCALAPPDATA")):
            if base:
                candidatos += [Path(base) / "Microsoft/Edge/Application/msedge.exe",
                               Path(base) / "Google/Chrome/Application/chrome.exe"]
    elif sys.platform == "darwin":
        candidatos += [Path("/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"),
                       Path("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome")]
    for c in candidatos:
        if c.exists():
            return str(c)
    for n in ("microsoft-edge", "google-chrome", "chromium", "chromium-browser"):
        if shutil.which(n):
            return shutil.which(n)
    return None


def vigilar_cierre(servidor, espera=150):
    """Apaga el servidor cuando la ventana deja de enviar pings."""
    while True:
        time.sleep(10)
        if time.time() - ULTIMO_PING[0] > espera:
            print("Ventana cerrada: deteniendo la app.")
            servidor.shutdown()
            return


def main():
    ap = argparse.ArgumentParser(description="Control Logístico")
    ap.add_argument("--puerto", type=int, default=8765)
    ap.add_argument("--navegador", action="store_true", help="abrir en una pestaña normal del navegador")
    ap.add_argument("--sin-ventana", action="store_true", help="solo iniciar el servidor")
    a = ap.parse_args()

    puerto = puerto_libre(a.puerto)
    srv = ThreadingHTTPServer(("127.0.0.1", puerto), Handler)
    url = f"http://127.0.0.1:{puerto}/"
    print(f"Control Logístico en {url}  (Ctrl+C para salir)")

    if a.sin_ventana:
        srv.serve_forever()
        return

    if not a.navegador:
        try:
            import webview  # pywebview: ventana nativa
            threading.Thread(target=srv.serve_forever, daemon=True).start()
            webview.create_window("Control Logístico", url, width=1440, height=900, min_size=(1000, 650))
            webview.start()
            srv.shutdown()
            return
        except ImportError:
            pass

    nav = None if a.navegador else buscar_navegador_app()
    if nav:
        perfil = motor.DB_PATH.parent / "perfil_ventana"
        subprocess.Popen([nav, f"--app={url}", "--window-size=1440,900", f"--user-data-dir={perfil}",
                          "--no-first-run", "--no-default-browser-check"])
        threading.Thread(target=vigilar_cierre, args=(srv,), daemon=True).start()
    else:
        webbrowser.open(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
