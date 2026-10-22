"""Prueba de aceptación de cifras (criterio 1). Recalcula desde los archivos crudos, sin pasar por el panel,
y compara con las invariantes y con los valores de referencia del proyecto.
Uso: python verificar.py [config.json]   → código de salida 0 si todo cuadra."""
import collections
import json
import os
import sys

import sync
from clasificar import IMG


def main():
    cfg, P = sync.leer_config(sys.argv[1] if len(sys.argv) > 1 else None)
    d = sync.construir(cfg, P)
    m, F = d['meta'], d['folios']
    # conteo independiente directo del disco
    n = ni = 0
    claves = set()
    for dd, _, fs in os.walk(cfg['carpeta_proyecto']):
        for f in fs:
            if f.startswith('~$') or f.lower() in ('desktop.ini', 'thumbs.db'):
                continue
            n += 1
            if f.lower().endswith(IMG):
                ni += 1
                claves.add((f.lower(), os.path.getsize(os.path.join(dd, f))))
    res = []

    def chk(nombre, obtenido, esperado):
        res.append((obtenido == esperado, nombre, obtenido, esperado))

    chk('archivos del panel = archivos en disco', m['archivos'], n)
    chk('imágenes del panel = imágenes en disco', m['imagenes'], ni)
    chk('fotos distintas del panel = distintas en disco', m['fotos_distintas'], len(claves))
    chk('folios sin repetir', len({f['mi'] for f in F}), len(F))
    chk('suma de estados = folios', sum(collections.Counter(f['est'] for f in F).values()), len(F))
    chk('municipios = suma por gestión', sum(1 for f in F if f['mun'] in P['municipios']), len(F))
    chk('gestiones válidas', sum(1 for f in F if f['gest'] in P['gestiones']), len(F))
    chk('todo folio con carpeta propia tiene radicado', sum(1 for f in F if f['rel'] == 'Propia' and not f['rad']), 0)
    ref = P.get('referencia', {})
    if 'folios' in ref: chk('[ref] folios', len(F), ref['folios'])
    if 'archivos' in ref: chk('[ref] archivos', m['archivos'], ref['archivos'])
    if 'imagenes' in ref: chk('[ref] imágenes', m['imagenes'], ref['imagenes'])
    if 'fotos_distintas' in ref: chk('[ref] fotos distintas', m['fotos_distintas'], ref['fotos_distintas'])
    for k, campo in (('estados', 'est'), ('municipios', 'mun'), ('gestiones', 'gest')):
        for v, esp in ref.get(k, {}).items():
            chk(f'[ref] {k}: {v}', sum(1 for f in F if f[campo] == v), esp)
    bad = [r for r in res if not r[0]]
    for ok, nom, o, e in res:
        print(('OK   ' if ok else 'FALLA'), nom, '' if ok else f'(panel {o} · esperado {e})')
    print(f'\n{len(res) - len(bad)} de {len(res)} controles cuadran.')
    sys.exit(1 if bad else 0)


if __name__ == '__main__':
    main()
