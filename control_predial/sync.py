"""Lee la carpeta sincronizada del proyecto, la matriz maestra y (opcional) el export de
NotebookLM, y escribe web/data.json con los folios, alertas y controles de calidad."""
import collections
import json
import os
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

import openpyxl

from clasificar import IMG, clasificar

AQUI = Path(__file__).resolve().parent
MATRICULA = re.compile(r'(?<![\d-])(\d{3}-\d{1,6})(?![\d-])')
RADICADO = re.compile(r'^(\d{8})\b')
IGNORAR = {'.lock', '.freelist', '.atx', '.gdbtable', '.gdbtablx', '.gdbindexes', '.horizon', '.spx'}  # no se ignoran del conteo; solo informativo
DOCS = [('CTL', ['CTL']), ('Escritura', ['Escritura', 'Escritura de servidumbre']), ('F. catastral', ['Ficha catastral']),
        ('F. predial', ['Ficha predial']), ('Plano', ['Plano']), ('Avalúo/acta', ['Avalúo', 'Acta de avalúo']),
        ('Soporte judicial', ['Soporte expropiatorio', 'Sentencia'])]
COLS = {  # campo -> encabezado de la matriz maestra
    'lado': 'Lado de trabajo', 'mun': 'Municipio según ubicación del Anexo', 'gest': 'Gestión / derecho real',
    'rel_x': 'Relación con carpeta del Drive', 'ubic': 'Ubicación / procedencia literal', 'area': 'Área interpretada (m²)',
    'valor': 'Valor registrado', 'kmz': 'Geometría KMZ asociada', 'campo': 'Afectación según campo',
    'vereda': 'Dato de campo asociado', 'estreg': 'Comentarios / estado registral anotado', 'alx': 'Alertas de conciliación'}


def nfc(s):
    return unicodedata.normalize('NFC', s or '')


def cargar_matriz(ruta):
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    ws = wb['Matriz maestra']
    filas = list(ws.iter_rows(min_row=5, values_only=True))
    enc = list(filas[0])
    out = {}
    for f in filas[1:]:
        if not f[0]:
            continue
        r = dict(zip(enc, f))
        out[str(r['Matrícula']).strip()] = {k: r.get(h) for k, h in COLS.items()}
    return out


def escanear(raiz):
    """Una entrada por archivo: ruta relativa de carpeta, nombre, tamaño."""
    archivos, carpetas = [], set()
    raiz = Path(raiz)
    for d, subs, fs in os.walk(raiz):
        rel = nfc(os.path.relpath(d, raiz)).replace('\\', '/')
        rel = '' if rel == '.' else rel
        carpetas.add(rel)
        for f in fs:
            if f.startswith('~$') or f.lower() in ('desktop.ini', 'thumbs.db'):
                continue
            p = Path(d) / f
            try:
                st = p.stat()
            except OSError:
                continue
            archivos.append({'dir': rel, 'name': nfc(f), 'len': st.st_size, 'mod': int(st.st_mtime)})
    return archivos, carpetas


def cargar_cuadernos(ruta):
    if not ruta or not Path(ruta).exists():
        return None
    return json.loads(Path(ruta).read_text(encoding='utf-8'))


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def construir(cfg):
    t0 = time.time()
    matriz = cargar_matriz(cfg['matriz_xlsx'])
    archivos, carpetas = escanear(cfg['carpeta_proyecto'])
    cuad = cargar_cuadernos(cfg.get('cuadernos_json'))
    calidad = []

    # --- asignar archivos a folios por el nombre de las carpetas ---
    por_folio = collections.defaultdict(list)       # mi -> archivos (cualquier carpeta que lo nombre)
    propias = collections.defaultdict(set)          # mi -> carpetas de radicado
    rad_de = {}
    sin_folio = collections.defaultdict(int)
    carpetas_top = collections.defaultdict(int)
    for a in archivos:
        a['cls'] = clasificar(a['name'])
        segs = a['dir'].split('/') if a['dir'] else []
        top = segs[0] if segs else '(raíz)'
        carpetas_top[top] += 1
        mis = set()
        for s in segs:
            mis.update(MATRICULA.findall(s))
        for mi in mis:
            if mi in matriz:
                por_folio[mi].append(a)
                if segs and RADICADO.match(segs[0]):
                    propias[mi].add(top)
                    rad_de.setdefault(mi, RADICADO.match(top).group(1))
            else:
                sin_folio[top] += 1
    for top, n in sorted(sin_folio.items()):
        calidad.append(['Media', 'Carpeta sin folio en la matriz', f'«{top}» nombra una matrícula que no está en la matriz ({n} archivos).'])
    for top in sorted(carpetas_top):
        if RADICADO.match(top) and not any(top in v for v in propias.values()) and top not in sin_folio:
            calidad.append(['Media', 'Carpeta de radicado sin matrícula', f'«{top}» no nombra ninguna matrícula reconocible (revisar el nombre).'])
    vacias = sorted(c for c in carpetas if c and not any(a['dir'] == c or a['dir'].startswith(c + '/') for a in archivos))
    for c in vacias:
        calidad.append(['Baja', 'Carpeta vacía', f'«{c}» no tiene archivos.'])

    # radicados compartidos y rótulos
    rad_folios = collections.defaultdict(list)
    for mi, r in rad_de.items():
        rad_folios[r].append(mi)
    for r, mis in rad_folios.items():
        if len(mis) > 1:
            calidad.append(['Info', 'Carpeta compartida por varios folios', f'Radicado {r}: {", ".join(sorted(mis))}. Sus archivos y fotos se cuentan para cada folio.'])
    for top in sorted(carpetas_top):
        if not RADICADO.match(top):
            continue
        mis = MATRICULA.findall(top)
        if re.search(r'[A-Za-zñÑ]\d{3}-\d', top):
            calidad.append(['Baja', 'Error de escritura en carpeta', f'«{top}»: falta un espacio antes de la matrícula.'])
        pal = re.findall(r'[A-Za-zÁÉÍÓÚÑáéíóúñ]+', top)
        lado = next((p for p in pal if p.lower().startswith(('cañaver', 'cañaver', 'encimad'))), '')
        if lado and lado.lower() not in ('cañaveral', 'encimadas'):
            calidad.append(['Baja', 'Error de escritura en carpeta', f'«{top}»: escribir «Cañaveral» o «Encimadas».'])
        elif mis and lado:
            esperado = 'cañaveral' if mis[0].startswith('028') else 'encimadas'
            if lado.lower() != esperado:
                calidad.append(['Baja', 'Rótulo de lado cruzado', f'«{top}»: la matrícula {mis[0]} va con {esperado.capitalize()}.'])

    # --- cuadernos de NotebookLM ---
    nb = {}
    if cuad is not None:
        for c in cuad:
            t = c.get('title', '')
            m = MATRICULA.search(t)
            fotos = {s['title'].strip().lower() for s in c.get('sources', []) if s['title'].lower().endswith(IMG)}
            if not m:
                continue
            nb[m.group(1)] = {'titulo': t.strip(), 'fotos': fotos, 'fuentes': len(c.get('sources', []))}
        for mis in rad_folios.values():  # carpeta compartida: los folios sin cuaderno propio usan el del folio hermano
            con = next((nb[m] for m in mis if m in nb), None)
            for m in mis:
                if con and m not in nb:
                    nb[m] = con

    # --- fotos: ocurrencias por (nombre, tamaño) para marcar las compartidas ---
    clave_folios = collections.defaultdict(set)
    for mi, fs in por_folio.items():
        for a in fs:
            if a['name'].lower().endswith(IMG):
                clave_folios[(a['name'].lower(), a['len'])].add(mi)
    dup_misma = 0
    for top in carpetas_top:
        vistos = collections.Counter((a['name'].lower(), a['len']) for a in archivos
                                     if a['name'].lower().endswith(IMG) and (a['dir'].split('/')[0] if a['dir'] else '(raíz)') == top)
        for (n, _), k in vistos.items():
            if k > 1:
                dup_misma += k - 1
                calidad.append(['Baja', 'Foto repetida dentro de la misma carpeta', f'«{n}» está {k} veces en «{top}».'])

    folios = []
    for mi, r in matriz.items():
        fs = por_folio.get(mi, [])
        propia = bool(propias.get(mi))
        rel_x = str(r['rel_x'] or '')
        matriz_de = '' if rel_x == 'Carpeta propia' else rel_x.split()[-1] if rel_x else ''
        if propia != (rel_x == 'Carpeta propia'):
            calidad.append(['Media', 'Matriz y carpeta no coinciden',
                            f'{mi}: la matriz dice «{rel_x}» pero en el Drive {"hay" if propia else "no hay"} carpeta de radicado.'])
        en_prop = [a for a in fs if a['dir'].split('/')[0] in propias.get(mi, ())]
        imgs = {(a['name'].lower(), a['len']): a['name'].lower() for a in fs if a['name'].lower().endswith(IMG)}
        sp_nombres = set(imgs.values())
        comp = sum(1 for k in imgs if len(clave_folios[k]) > 1)
        docs = {k: any(a['cls'] in cl for a in en_prop) for k, cl in DOCS}
        docs['Plano'] = docs['Plano'] or any('plano' in a['name'].lower() for a in en_prop)  # «Ficha predial y plano.pdf»
        gest = r['gest'] or ''
        c = nb.get(mi)
        sp_only = sorted(sp_nombres - c['fotos']) if c else []
        nb_only = sorted(c['fotos'] - sp_nombres) if c else []
        if cuad is None:
            est = 'Servidumbre sin cuaderno' if False else 'Sin dato de NotebookLM'
        elif c:
            est = 'Fotos por conciliar' if (sp_only or nb_only) else 'Conciliado'
        elif gest == 'Servidumbre':
            est = 'Sin cuaderno · servidumbre'
        elif not propia and rel_x != 'Carpeta propia':
            est = 'Folio derivado · soporte en folio matriz'
        else:
            est = 'Falta cuaderno'
        area = num(r['area'])
        valor = str(r['valor'] or '')
        valor = re.sub(r'[^\d]', '', valor) if valor.startswith('$') else ''
        al = []
        if est == 'Falta cuaderno':
            al.append(['Alta', 'Sin cuaderno ni fotos en un folio con carpeta propia' if not imgs else 'Folio con fotos pero sin cuaderno'])
        if est == 'Fotos por conciliar':
            al.append(['Media', f'{len(sp_only)} foto(s) en el Drive sin cargar al cuaderno, {len(nb_only)} solo en cuaderno'])
        if propia:
            if gest == 'Expropiación':
                if not docs['Soporte judicial']:
                    al.append(['Media', 'Expropiación sin soporte judicial nombrado'])
            else:
                if not docs['CTL']:
                    al.append(['Alta', 'Falta CTL en la carpeta'])
                if not docs['Escritura']:
                    al.append(['Media', 'Falta escritura en la carpeta'])
        if area is None and gest != 'Expropiación':
            al.append(['Media', 'Área no interpretable en el Anexo'])
        if not propia:
            al.append(['Baja', 'Sin carpeta propia: soporte en folio matriz ' + matriz_de])
        if not r['kmz']:
            al.append(['Baja', 'Sin geometría KMZ asociada'])
        folios.append({
            'mi': mi, 'lado': str(r['lado']).split(' – ')[0], 'mun': r['mun'], 'gest': gest,
            'rel': 'Propia' if propia else 'Derivado', 'matriz': matriz_de, 'est': est, 'rad': rad_de.get(mi, ''),
            'archivos': len(en_prop), 'fotos': len(imgs), 'comp': comp,
            'cuad': bool(c), 'fcuad': len(c['fotos']) if c else None,
            'ubic': r['ubic'] or '', 'area': area, 'valor': valor, 'kmz': bool(r['kmz']),
            'campo': r['campo'] or '', 'vereda': r['vereda'] or '', 'estreg': (r['estreg'] or '')[:220],
            'docs': docs, 'al': al})
    por_mi = {f['mi']: f for f in folios}
    for f in folios:  # los derivados se apoyan en los documentos de su folio matriz
        if f['rel'] == 'Derivado' and f['matriz'] in por_mi:
            f['docs'] = dict(por_mi[f['matriz']]['docs'])
    if cuad is None:
        calidad.append(['Media', 'Sin export de NotebookLM', 'No se encontró el archivo de cuadernos; el estado del cotejo no se puede calcular.'])

    # --- cambios frente a la sincronización anterior ---
    snap_p = AQUI / 'estado' / 'snapshot.json'
    ahora = {f"{a['dir']}/{a['name']}": a['len'] for a in archivos}
    cambios = {'nuevos': [], 'modificados': [], 'eliminados': [], 'primera': True}
    if snap_p.exists():
        ant = json.loads(snap_p.read_text(encoding='utf-8'))
        cambios = {'primera': False, 'nuevos': sorted(set(ahora) - set(ant)), 'eliminados': sorted(set(ant) - set(ahora)),
                   'modificados': sorted(k for k in ahora if k in ant and ant[k] != ahora[k])}
    snap_p.parent.mkdir(exist_ok=True)
    snap_p.write_text(json.dumps(ahora, ensure_ascii=False), encoding='utf-8')

    orden = {'Alta': 0, 'Media': 1, 'Baja': 2, 'Info': 3}
    calidad.sort(key=lambda x: (orden[x[0]], x[1]))
    meta = {
        'generado': datetime.now().isoformat(timespec='seconds'), 'segundos': round(time.time() - t0, 1),
        'archivos': len(archivos), 'carpetas': len(carpetas) - (1 if '' in carpetas else 0),
        'imagenes': sum(1 for a in archivos if a['name'].lower().endswith(IMG)),
        'cuadernos': None if cuad is None else len(cuad),
        'fuentes_cuadernos': None if cuad is None else sum(len(c.get('sources', [])) for c in cuad),
        'fotos_distintas': len({(a['name'].lower(), a['len']) for a in archivos if a['name'].lower().endswith(IMG)}),
        'carpeta': str(cfg['carpeta_proyecto']),
        'cambios': {k: (v if k == 'primera' else len(v)) for k, v in cambios.items()}}
    salida = {'meta': meta, 'folios': folios, 'calidad': calidad,
              'cambios': {k: v[:200] if isinstance(v, list) else v for k, v in cambios.items()}}
    destino = AQUI / 'web' / 'data.json'
    tmp = destino.with_suffix('.tmp')
    tmp.write_text(json.dumps(salida, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    tmp.replace(destino)
    return salida


def leer_config(ruta=None):
    p = Path(ruta) if ruta else AQUI / 'config.json'
    if not p.exists():
        sys.exit(f'Falta {p}. Copia config.example.json a config.json y completa las rutas.')
    cfg = json.loads(p.read_text(encoding='utf-8'))
    for k in ('carpeta_proyecto', 'matriz_xlsx'):
        if not Path(cfg.get(k, '')).exists():
            sys.exit(f'No existe la ruta de «{k}»: {cfg.get(k)}')
    return cfg


if __name__ == '__main__':
    r = construir(leer_config(sys.argv[1] if len(sys.argv) > 1 else None))
    print(json.dumps(r['meta'], ensure_ascii=False, indent=1))
