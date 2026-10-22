"""Se corre en TU equipo (solo Python, sin instalar nada). Lee la carpeta sincronizada del proyecto y los
borradores del formato 1.1 y deja dos archivos para subir al chat:
  inventario.json  (carpetas, nombres y tamaños; no copia el contenido de tus archivos)
  borradores.json  (texto de cada borrador .docx del formato, para revisar si está completo)
Uso:  python exportar_inventario.py "C:/.../1 Proyecto Hidroarma" --borradores "C:/.../BORRADORES_3.1_SONSON..." --borradores "C:/.../AGUADAS..."
"""
import argparse
import json
import os
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path


def nfc(s):
    return unicodedata.normalize('NFC', s)


def texto_docx(ruta):
    try:
        x = zipfile.ZipFile(ruta).read('word/document.xml').decode('utf8', 'ignore')
    except Exception:
        return None
    pars = re.findall(r'<w:p[ >].*?</w:p>', x, re.S)
    return '\n'.join(t for t in (''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)) for p in pars) if t.strip())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('carpeta')
    ap.add_argument('--borradores', action='append', default=[], help='carpeta con formatos diligenciados (se puede repetir)')
    ap.add_argument('--salida', default='.')
    a = ap.parse_args()
    raiz = Path(a.carpeta)
    files, folders = [], []
    for d, _, fs in os.walk(raiz):
        rel = nfc(os.path.relpath(d, raiz)).replace('\\', '/')
        rel = '' if rel == '.' else rel
        if rel:
            folders.append(rel)
        for f in fs:
            if f.startswith('~$') or f.lower() in ('desktop.ini', 'thumbs.db'):
                continue
            try:
                st = (Path(d) / f).stat()
            except OSError:
                continue
            files.append({'dir': rel, 'name': nfc(f), 'len': st.st_size, 'mod': int(st.st_mtime)})
    out = Path(a.salida)
    out.mkdir(parents=True, exist_ok=True)
    (out / 'inventario.json').write_text(json.dumps({'generado': datetime.now().isoformat(timespec='seconds'), 'raiz': raiz.name, 'files': files, 'folders': folders}, ensure_ascii=False), encoding='utf-8')
    bor = []
    for base in a.borradores:
        for d, _, fs in os.walk(base):
            for f in fs:
                if f.startswith('~$') or not f.lower().endswith('.docx'):
                    continue
                p = Path(d) / f
                bor.append({'archivo': nfc(f), 'carpeta': nfc(os.path.relpath(d, base)).replace('\\', '/'), 'mod': int(p.stat().st_mtime), 'texto': texto_docx(p)})
    (out / 'borradores.json').write_text(json.dumps(bor, ensure_ascii=False), encoding='utf-8')
    print(f'inventario.json: {len(files)} archivos, {len(folders)} carpetas · borradores.json: {len(bor)} documentos · en {out.resolve()}')


if __name__ == '__main__':
    main()
