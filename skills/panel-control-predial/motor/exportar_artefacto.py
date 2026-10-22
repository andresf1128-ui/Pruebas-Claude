"""Convierte el resultado del motor (data.json) en los conjuntos de datos del artefacto del tablero."""
import json
from pathlib import Path

DOCS = [('doc_ctl', 'CTL'), ('doc_escritura', 'Escritura'), ('doc_catastral', 'F. catastral'), ('doc_predial', 'F. predial'),
        ('doc_plano', 'Plano'), ('doc_avaluo', 'Avalúo/acta'), ('doc_judicial', 'Soporte judicial')]
RANK = {'Alta': 3, 'Media': 2, 'Baja': 1}


def convertir(d, corte, salida):
    dd = {x['nombre']: x for x in d['meta']['documentos']}

    def na(g, k):
        x = dd[k]
        return bool((x.get('solo_gestiones') and g not in x['solo_gestiones']) or (x.get('excepto_gestiones') and g in x['excepto_gestiones']))

    rows, al, fm = [], [], []
    for f in d['folios']:
        mx = max(f['al'], key=lambda a: RANK[a[0]])[0] if f['al'] else 'Ninguna'
        r = {'mi': f['mi'], 'municipio': f['mun'], 'lado': f['lado'], 'gestion': f['gest'], 'carpeta': f['rel'], 'matriz': f['matriz'] or '—',
             'estado': f['est'], 'radicado': f['rad'] or '—', 'archivos': f['archivos'] if f['rel'] == 'Propia' else 0, 'fotos': f['fotos'],
             'compartidas': f['comp'], 'cuaderno': 'Sí' if f['cuad'] else 'No', 'fotos_cuaderno': f['fcuad'] if f['fcuad'] is not None else 0,
             'area_m2': f['area'] if f['area'] is not None else 0, 'area_dato': 'Sí' if f['area'] is not None else 'No',
             'valor': ('$' + format(int(f['valor']), ',').replace(',', '.')) if f['valor'] else '—', 'kmz': 'Sí' if f['kmz'] else 'No',
             'ubicacion': f['ubic'] or '—', 'alerta_max': mx, 'n_alertas': len(f['al'])}
        for col, k in DOCS:
            r[col] = 'N/A' if na(f['gest'], k) else ('Sí' if f['docs'][k] else 'No')
        rows.append(r)
        for i, a in enumerate(f['al']):
            al.append({'id': f"{f['mi']}|{i + 1}", 'mi': f['mi'], 'municipio': f['mun'], 'gestion': f['gest'], 'prioridad': a[0], 'alerta': a[1]})
        for fid, x in f['fmt'].items():
            fm.append({'mi': f['mi'], 'municipio': f['mun'], 'gestion': f['gest'], 'formato': fid, 'estado': x['estado'],
                       'archivo': x['archivo'] or '—', 'falta': '; '.join(x['motivos'][:3]) or '—'})
    cal = [{'id': str(i + 1), 'nivel': c[0], 'control': c[1], 'detalle': c[2]} for i, c in enumerate(d['calidad'])]
    m = d['meta']
    fu = [{'generado': corte, 'fuente': 'SharePoint «1 Proyecto Hidroarma», cuadernos de NotebookLM y matriz maestra (Excel)',
           'actualizacion': 'Se actualiza cuando se sube un archivo nuevo', 'archivos': m['archivos'], 'carpetas': m['carpetas'],
           'imagenes': m['imagenes'], 'fotos_distintas': m['fotos_distintas'], 'cuadernos': m['cuadernos'], 'fuentes_cuadernos': m['fuentes_cuadernos']}]
    c = d['cambios']
    cam = []
    if not c.get('primera'):
        for tipo, lst in (('Nuevo', c['nuevos']), ('Modificado', c['modificados']), ('Eliminado', c['eliminados'])):
            for i, ruta in enumerate(lst[:300]):
                cam.append({'id': f'{tipo}|{i + 1}', 'tipo': tipo, 'ruta': ruta, 'corte': corte})
    out = Path(salida)
    out.mkdir(parents=True, exist_ok=True)
    res = {'folios': rows, 'alertas': al, 'calidad': cal, 'fuente': fu, 'formatos': fm, 'cambios': cam}
    for n, v in res.items():
        (out / f'{n}.json').write_text(json.dumps(v, ensure_ascii=False), encoding='utf-8')
    return {n: len(v) for n, v in res.items()}
