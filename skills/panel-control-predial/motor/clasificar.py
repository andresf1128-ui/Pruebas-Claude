"""Clasificación nominal de archivos por su nombre, según las reglas del proyecto (primera regla que coincide)."""
import re

IMG = ('.jpg', '.jpeg', '.png', '.heic', '.webp')


def compilar(reglas):
    return [(re.compile(r['nombre']) if r.get('nombre') else None, tuple(r.get('ext', ())), r['clase']) for r in reglas]


def clasificar(nombre, reglas, defecto='Otro soporte'):
    n = nombre.lower()
    ext = n.rsplit('.', 1)[-1] if '.' in n else ''
    for rx, exts, clase in reglas:
        if exts and ext not in exts:
            continue
        if rx and not rx.search(n):
            continue
        return clase
    return defecto
