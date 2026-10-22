"""Lee la carpeta sincronizada del proyecto, la matriz maestra y (opcional) el export de NotebookLM,
y escribe web/data.json con folios, alertas, control de formatos y controles de calidad.
Todo lo propio de un proyecto vive en el archivo `proyecto` (ver ../proyectos/)."""
import collections
import json
import os
import re
import sys
import time
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl

from clasificar import IMG, clasificar, compilar

AQUI = Path(__file__).resolve().parent
ESTADO_DIR = AQUI / 'estado'


def nfc(s):
    return unicodedata.normalize('NFC', s or '')


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def cargar_matriz(ruta, P):
    m = P['matriz']
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    filas = list(wb[m['hoja']].iter_rows(min_row=m['fila_encabezado'], values_only=True))
    enc = list(filas[0])
    out = {}
    for f in filas[1:]:
        if not f[0]:
            continue
        r = dict(zip(enc, f))
        out[str(r[m['clave']]).strip()] = {k: r.get(h) for k, h in m['columnas'].items()}
    return out


def escanear(raiz):
    archivos, carpetas = [], set()
    for d, _, fs in os.walk(raiz):
        rel = nfc(os.path.relpath(d, raiz)).replace('\\', '/')
        rel = '' if rel == '.' else rel
        carpetas.add(rel)
        for f in fs:
            if f.startswith('~$') or f.lower() in ('desktop.ini', 'thumbs.db'):
                continue
            try:
                st = (Path(d) / f).stat()
            except OSError:
                continue
            archivos.append({'dir': rel, 'name': nfc(f), 'len': st.st_size, 'mod': int(st.st_mtime)})
    return archivos, carpetas


def leer_json(ruta, defecto=None):
    p = Path(ruta) if ruta else None
    return json.loads(p.read_text(encoding='utf-8')) if p and p.exists() else defecto


def texto_docx(ruta):
    try:
        x = zipfile.ZipFile(ruta).read('word/document.xml').decode('utf8', 'ignore')
    except Exception:
        return None
    pars = re.findall(r'<w:p[ >].*?</w:p>', x, re.S)
    return '\n'.join(t for t in (''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', p)) for p in pars) if t.strip())


def revisar_formato(texto, mi, v):
    """Devuelve la lista de motivos por los que un borrador NO está completo (vacía = completo)."""
    motivos = []
    for sec in v.get('secciones_requeridas', []):
        if not re.search(sec, texto, re.I):
            motivos.append(f'falta la sección «{sec.split("|")[0].strip("()")}»')
    for r in v.get('restos_de_plantilla', []):
        if r.lower() in texto.lower():
            motivos.append(f'queda texto de ejemplo de la plantilla: «{r}»')
    if v.get('exigir_matricula_en_texto') and mi not in texto:
        motivos.append(f'no menciona la matrícula {mi}')
    return motivos


def escanear_formatos(cfg, P, MAT, matriz):
    """Borradores diligenciados en carpetas locales: {formato: {matrícula: [{archivo, ruta, motivos}]}}."""
    out = {}
    for fo in P.get('formatos', []):
        rx = re.compile(fo['patron_archivo'], re.I)
        exts = tuple('.' + e for e in fo.get('extensiones', ['docx']))
        hall = collections.defaultdict(list)
        for base in cfg.get('formatos_carpetas', {}).get(fo['id'], []):
            if not Path(base).exists():
                continue
            for d, _, fs in os.walk(base):
                for f in fs:
                    if f.startswith('~$') or not f.lower().endswith(exts) or not rx.search(f) and not MAT.search(f):
                        continue
                    ruta = Path(d) / f
                    mis = {m for m in MAT.findall(nfc(f)) if m in matriz}
                    texto = texto_docx(ruta) if f.lower().endswith('.docx') else None
                    if not mis and texto:  # el nombre no trae matrícula: se busca en el contenido
                        c = collections.Counter(m for m in MAT.findall(texto) if m in matriz)
                        mis = {c.most_common(1)[0][0]} if c else set()
                    for mi in mis:
                        mot = revisar_formato(texto, mi, fo.get('verificacion', {})) if texto is not None else ['no se pudo leer el archivo']
                        hall[mi].append({'archivo': nfc(f), 'carpeta': nfc(os.path.relpath(d, base)).replace('\\', '/'),
                                         'mod': int(ruta.stat().st_mtime), 'motivos': mot})
        out[fo['id']] = hall
    return out


def evaluar_alertas(P, f, ctx):
    out = []
    for r in P['alertas']:
        c, g = r['cuando'], f['gest']
        if 'gestiones' in r and g not in r['gestiones'] or g in r.get('excepto_gestiones', ()):
            continue
        ok = False
        if c == 'estado':
            ok = f['est'] == r['valor']
        elif c == 'falta_doc':
            d = next((x for x in P['documentos'] if x['nombre'] == r['doc']), None)
            aplica = d and g not in d.get('excepto_gestiones', ()) and ('solo_gestiones' not in d or g in d['solo_gestiones'])
            ok = bool(aplica and f['rel'] == 'Propia' and not f['docs'].get(r['doc']))
        elif c == 'sin_area':
            ok = f['area'] is None
        elif c == 'derivado':
            ok = f['rel'] == 'Derivado'
        elif c == 'sin_kmz':
            ok = not f['kmz']
        if ok:
            out.append([r['prioridad'], r['texto'].format(**ctx)])
    return out


def construir(cfg, P):
    t0 = time.time()
    matriz = cargar_matriz(cfg['matriz_xlsx'], P)
    archivos, carpetas = escanear(cfg['carpeta_proyecto'])
    cuad = leer_json(cfg.get('cuadernos_json'))
    marcas = leer_json(ESTADO_DIR / 'marcas.json', {})
    reglas = compilar(P['clases'])
    MAT, RAD = re.compile(P['patron_matricula']), re.compile(P['patron_radicado'])
    calidad = []

    por_folio = collections.defaultdict(list)
    propias = collections.defaultdict(set)
    rad_de = {}
    sin_folio = collections.defaultdict(int)
    carpetas_top = collections.defaultdict(int)
    for a in archivos:
        a['cls'] = clasificar(a['name'], reglas, P.get('clase_por_defecto', 'Otro soporte'))
        segs = a['dir'].split('/') if a['dir'] else []
        top = segs[0] if segs else '(raíz)'
        a['top'] = top
        carpetas_top[top] += 1
        mis = set()
        for s in segs:
            mis.update(MAT.findall(s))
        for mi in mis:
            if mi in matriz:
                por_folio[mi].append(a)
                if segs and RAD.match(segs[0]):
                    propias[mi].add(top)
                    rad_de.setdefault(mi, RAD.match(top).group(1))
            else:
                sin_folio[top] += 1
    for top, n in sorted(sin_folio.items()):
        calidad.append(['Media', 'Carpeta sin folio en la matriz', f'«{top}» nombra una matrícula que no está en la matriz ({n} archivos).'])
    vinculadas = {t for v in propias.values() for t in v}
    for top in sorted(carpetas_top):
        if RAD.match(top) and top not in vinculadas and top not in sin_folio:
            calidad.append(['Media', 'Carpeta de radicado sin matrícula', f'«{top}» no nombra ninguna matrícula reconocible (revisar el nombre).'])
    for c in sorted(c for c in carpetas if c and not any(a['dir'] == c or a['dir'].startswith(c + '/') for a in archivos)):
        calidad.append(['Baja', 'Carpeta vacía', f'«{c}» no tiene archivos.'])

    rad_folios = collections.defaultdict(list)
    for mi, r in rad_de.items():
        rad_folios[r].append(mi)
    for r, mis in rad_folios.items():
        if len(mis) > 1:
            calidad.append(['Info', 'Carpeta compartida por varios folios', f'Radicado {r}: {", ".join(sorted(mis))}. Sus archivos y fotos se cuentan para cada folio.'])
    palabras = set(P.get('rotulos_lado', {}).values())
    for top in sorted(carpetas_top):
        if not RAD.match(top):
            continue
        mis = MAT.findall(top)
        if re.search(r'[A-Za-zñÑ]\d{3}-\d', top):
            calidad.append(['Baja', 'Error de escritura en carpeta', f'«{top}»: falta un espacio antes de la matrícula.'])
        pal = [p.lower() for p in re.findall(r'[A-Za-zÁÉÍÓÚÑáéíóúñ]+', top)]
        lado = next((p for p in pal if any(p[:6] == w[:6] for w in palabras)), '')
        if lado and lado not in palabras:
            calidad.append(['Baja', 'Error de escritura en carpeta', f'«{top}»: revisar la palabra «{lado}».'])
        elif mis and lado:
            esperado = P['rotulos_lado'].get(mis[0][:3])
            if esperado and lado != esperado:
                calidad.append(['Baja', 'Rótulo de lado cruzado', f'«{top}»: la matrícula {mis[0]} va con {esperado.capitalize()}.'])

    # --- cuadernos ---
    nb = {}
    if cuad is not None:
        for c in cuad:
            m = MAT.search(c.get('title', ''))
            if m:
                nb[m.group(1)] = {'fotos': {s['title'].strip().lower() for s in c.get('sources', []) if s['title'].lower().endswith(IMG)}}
        for mis in rad_folios.values():
            con = next((nb[m] for m in mis if m in nb), None)
            for m in mis:
                if con and m not in nb:
                    nb[m] = con

    # --- fotos compartidas / repetidas ---
    clave_folios = collections.defaultdict(set)
    for mi, fs in por_folio.items():
        for a in fs:
            if a['name'].lower().endswith(IMG):
                clave_folios[(a['name'].lower(), a['len'])].add(mi)
    vistos = collections.Counter((a['top'], a['name'].lower(), a['len']) for a in archivos if a['name'].lower().endswith(IMG))
    for (top, n, _), k in vistos.items():
        if k > 1:
            calidad.append(['Baja', 'Foto repetida dentro de la misma carpeta', f'«{n}» está {k} veces en «{top}».'])

    # --- formatos de control ---
    borradores = escanear_formatos(cfg, P, MAT, matriz)

    folios = []
    for mi, r in matriz.items():
        fs = por_folio.get(mi, [])
        propia = bool(propias.get(mi))
        rel_x = str(r['rel_x'] or '')
        matriz_de = '' if rel_x == P['matriz'].get('valor_propia', 'Carpeta propia') else (rel_x.split()[-1] if rel_x else '')
        if propia != (rel_x == P['matriz'].get('valor_propia', 'Carpeta propia')):
            calidad.append(['Media', 'Matriz y carpeta no coinciden',
                            f'{mi}: la matriz dice «{rel_x}» pero en el Drive {"hay" if propia else "no hay"} carpeta de radicado.'])
        en_prop = [a for a in fs if a['top'] in propias.get(mi, ())]
        imgs = {(a['name'].lower(), a['len']): a['name'].lower() for a in fs if a['name'].lower().endswith(IMG)}
        sp = set(imgs.values())
        gest = r['gest'] or ''
        docs = {}
        for d in P['documentos']:
            ok = any(a['cls'] in d['clases'] for a in en_prop)
            if d.get('tambien_nombre'):
                ok = ok or any(d['tambien_nombre'] in a['name'].lower() for a in en_prop)
            docs[d['nombre']] = ok
        c = nb.get(mi)
        sp_only = sorted(sp - c['fotos']) if c else []
        nb_only = sorted(c['fotos'] - sp) if c else []
        if cuad is None:
            est = 'Sin dato de NotebookLM'
        elif c:
            est = 'Fotos por conciliar' if (sp_only or nb_only) else 'Conciliado'
        elif gest == 'Servidumbre':
            est = 'Sin cuaderno · servidumbre'
        elif not propia and rel_x != P['matriz'].get('valor_propia', 'Carpeta propia'):
            est = 'Folio derivado · soporte en folio matriz'
        else:
            est = 'Falta cuaderno'
        valor = str(r['valor'] or '')
        valor = re.sub(r'[^\d]', '', valor) if valor.startswith('$') else ''
        fmt = {}
        for fo in P.get('formatos', []):
            if gest not in fo.get('gestiones', P['gestiones']):
                continue
            cand = sorted(borradores[fo['id']].get(mi, []), key=lambda x: (bool(x['motivos']), -x['mod']))
            mejor = cand[0] if cand else None
            estado = 'Sin borrador' if not mejor else ('Incompleto' if mejor['motivos'] else 'Completo')
            fmt[fo['id']] = {'estado': estado, 'archivo': mejor['archivo'] if mejor else '', 'carpeta': mejor['carpeta'] if mejor else '',
                             'motivos': mejor['motivos'] if mejor else [], 'n': len(cand), 'manual': marcas.get(mi, {}).get(fo['id'], '')}
        f = {
            'mi': mi, 'lado': str(r['lado']).split(' – ')[0], 'mun': r['mun'], 'gest': gest,
            'rel': 'Propia' if propia else 'Derivado', 'matriz': matriz_de, 'est': est, 'rad': rad_de.get(mi, ''),
            'archivos': len(en_prop), 'fotos': len(imgs), 'comp': sum(1 for k in imgs if len(clave_folios[k]) > 1),
            'cuad': bool(c), 'fcuad': len(c['fotos']) if c else None,
            'ubic': r['ubic'] or '', 'area': num(r['area']), 'valor': valor, 'kmz': bool(r['kmz']),
            'campo': r['campo'] or '', 'vereda': r['vereda'] or '', 'estreg': (r['estreg'] or '')[:220],
            'docs': docs, 'fmt': fmt}
        f['al'] = evaluar_alertas(P, f, {'sp_only': len(sp_only), 'nb_only': len(nb_only), 'matriz': matriz_de})
        folios.append(f)
    por_mi = {f['mi']: f for f in folios}
    for f in folios:
        if f['rel'] == 'Derivado' and f['matriz'] in por_mi:
            f['docs'] = dict(por_mi[f['matriz']]['docs'])
    for fo in P.get('formatos', []):
        carp = cfg.get('formatos_carpetas', {}).get(fo['id'], [])
        if not carp:
            calidad.append(['Media', f'Formato {fo["id"]} sin carpeta', 'Falta «formatos_carpetas» en config.json: no se buscan borradores diligenciados.'])
        for b in carp:
            if not Path(b).exists():
                calidad.append(['Media', f'Formato {fo["id"]}: carpeta no encontrada', b])
        for mi, lst in borradores[fo['id']].items():
            if len(lst) > 1:
                calidad.append(['Baja', f'Formato {fo["id"]}: varios borradores', f'{mi}: {", ".join(x["archivo"] for x in lst)}'])
        for mi in borradores[fo['id']]:
            if mi in por_mi and fo['id'] not in por_mi[mi]['fmt']:
                calidad.append(['Baja', f'Formato {fo["id"]}: gestión sin formato', f'{mi} tiene borrador pero su gestión no lleva este formato.'])
    if cuad is None:
        calidad.append(['Media', 'Sin export de NotebookLM', 'No se encontró el archivo de cuadernos; el estado del cotejo no se puede calcular.'])

    # --- cambios frente a la lectura anterior ---
    snap_p = ESTADO_DIR / 'snapshot.json'
    ahora = {f"{a['dir']}/{a['name']}": a['len'] for a in archivos}
    cambios = {'nuevos': [], 'modificados': [], 'eliminados': [], 'primera': True}
    ant = leer_json(snap_p)
    if ant is not None:
        cambios = {'primera': False, 'nuevos': sorted(set(ahora) - set(ant)), 'eliminados': sorted(set(ant) - set(ahora)),
                   'modificados': sorted(k for k in ahora if k in ant and ant[k] != ahora[k])}
    ESTADO_DIR.mkdir(exist_ok=True)
    snap_p.write_text(json.dumps(ahora, ensure_ascii=False), encoding='utf-8')

    orden = {'Alta': 0, 'Media': 1, 'Baja': 2, 'Info': 3}
    calidad.sort(key=lambda x: (orden[x[0]], x[1]))
    meta = {
        'proyecto': P['nombre'], 'generado': datetime.now().isoformat(timespec='seconds'), 'segundos': round(time.time() - t0, 1),
        'archivos': len(archivos), 'carpetas': len(carpetas) - (1 if '' in carpetas else 0),
        'imagenes': sum(1 for a in archivos if a['name'].lower().endswith(IMG)),
        'fotos_distintas': len({(a['name'].lower(), a['len']) for a in archivos if a['name'].lower().endswith(IMG)}),
        'cuadernos': None if cuad is None else len(cuad),
        'fuentes_cuadernos': None if cuad is None else sum(len(c.get('sources', [])) for c in cuad),
        'municipios': P['municipios'], 'gestiones': P['gestiones'],
        'documentos': [{k: d.get(k) for k in ('nombre', 'solo_gestiones', 'excepto_gestiones')} for d in P['documentos']],
        'formatos': [{'id': f['id'], 'nombre': f['nombre'], 'gestiones': f.get('gestiones', P['gestiones']), 'marca_manual': f.get('marca_manual', False)} for f in P.get('formatos', [])],
        'carpeta': str(cfg['carpeta_proyecto'])}
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
    pp = Path(cfg.get('proyecto', ''))
    if not pp.is_absolute():
        pp = (p.parent / pp).resolve() if (p.parent / pp).exists() else (AQUI / pp).resolve()
    if not pp.exists():
        sys.exit(f'No existe el archivo de proyecto: {cfg.get("proyecto")}')
    return cfg, json.loads(pp.read_text(encoding='utf-8'))


if __name__ == '__main__':
    cfg, P = leer_config(sys.argv[1] if len(sys.argv) > 1 else None)
    print(json.dumps(construir(cfg, P)['meta'], ensure_ascii=False, indent=1))
