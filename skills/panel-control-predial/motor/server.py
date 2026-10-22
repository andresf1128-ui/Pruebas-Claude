"""Servidor local: sirve el panel y vuelve a leer la carpeta del proyecto cuando detecta cambios."""
import json
import os
import sys
import threading
import time
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import sync

AQUI = Path(__file__).resolve().parent
cfg, P = sync.leer_config(sys.argv[1] if len(sys.argv) > 1 else None)
LOCK = threading.Lock()
ESTADO = {'ultimo': None, 'error': None, 'sincronizando': False}


def firma():
    """Firma barata de lo que alimenta el panel: carpeta del proyecto, matriz y export de cuadernos."""
    n = tam = mx = 0
    for d, _, fs in os.walk(cfg['carpeta_proyecto']):
        for f in fs:
            try:
                st = os.stat(os.path.join(d, f))
            except OSError:
                continue
            n += 1
            tam += st.st_size
            mx = max(mx, st.st_mtime)
    extra = [os.path.getmtime(p) for p in (cfg['matriz_xlsx'], cfg.get('cuadernos_json'), sync.ESTADO_DIR / 'marcas.json') if p and os.path.exists(p)]
    return (n, tam, mx, tuple(extra))


def actualizar(motivo):
    with LOCK:
        ESTADO['sincronizando'] = True
        try:
            sync.construir(cfg, P)
            ESTADO.update(ultimo=time.strftime('%Y-%m-%d %H:%M:%S'), error=None, motivo=motivo)
        except Exception as e:  # el panel sigue mostrando el último dato bueno
            ESTADO['error'] = f'{type(e).__name__}: {e}'
        finally:
            ESTADO['sincronizando'] = False


def vigilar():
    ant = None
    while True:
        try:
            f = firma()
            if f != ant:
                actualizar('arranque' if ant is None else 'cambio detectado')
                ant = f
        except Exception as e:
            ESTADO['error'] = f'{type(e).__name__}: {e}'
        time.sleep(int(cfg.get('intervalo_segundos', 60)))


class H(SimpleHTTPRequestHandler):
    def __init__(self, *a, **k):
        super().__init__(*a, directory=str(AQUI / 'web'), **k)

    def _json(self, obj, code=200):
        b = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(b)

    def do_GET(self):
        if self.path.startswith('/api/estado'):
            return self._json(ESTADO)
        if self.path.startswith('/data.json'):
            self.path = '/data.json'
        super().do_GET()

    def do_POST(self):
        if self.path.startswith('/api/marcar'):  # marca manual de un formato: {mi, formato, estado}
            try:
                b = json.loads(self.rfile.read(int(self.headers.get('Content-Length', 0))))
                ids = {f['id'] for f in P.get('formatos', []) if f.get('marca_manual')}
                if b['formato'] not in ids or b['estado'] not in ('', 'Pendiente', 'En proceso', 'Completo', 'Revisado'):
                    return self._json({'error': 'marca no válida'}, 400)
                sync.ESTADO_DIR.mkdir(exist_ok=True)
                p = sync.ESTADO_DIR / 'marcas.json'
                m = sync.leer_json(p, {})
                m.setdefault(b['mi'], {})[b['formato']] = b['estado']
                p.write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
                actualizar('marca manual')
                return self._json(ESTADO)
            except Exception as e:
                return self._json({'error': str(e)}, 400)
        if self.path.startswith('/api/actualizar'):
            actualizar('manual')
            return self._json(ESTADO)
        self.send_error(404)

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        super().end_headers()

    def log_message(self, *a):
        pass


if __name__ == '__main__':
    threading.Thread(target=vigilar, daemon=True).start()
    puerto = int(cfg.get('puerto', 8765))
    srv = ThreadingHTTPServer(('127.0.0.1', puerto), H)  # solo este equipo
    print(f'Panel en http://127.0.0.1:{puerto}  ·  vigila: {cfg["carpeta_proyecto"]}  (Ctrl+C para salir)')
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
