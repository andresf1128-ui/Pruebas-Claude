"""Cruza los archivos subidos (inventario, matriz, cuadernos, borradores) y deja los conjuntos de datos del artefacto.
Uso:
  python actualizar.py --inventario inventario.json --matriz matriz.xlsx --cuadernos cuadernos.json \
      [--borradores borradores.json] --corte "22-oct-2026" --salida salida/ [--proyecto ../proyectos/hidroarma.json] \
      [--snapshot ../estado/hidroarma_snapshot.json] [--no-guardar-snapshot]
Luego se suben los 6 JSON de `salida/` al artefacto y se actualizan sus documentos (ver SKILL.md, «Actualizar el artefacto»)."""
import argparse
import json
import sys
from pathlib import Path

import exportar_artefacto
import sync

AQUI = Path(__file__).resolve().parent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--inventario', required=True)
    ap.add_argument('--matriz', required=True)
    ap.add_argument('--cuadernos')
    ap.add_argument('--borradores')
    ap.add_argument('--corte', required=True)
    ap.add_argument('--salida', required=True)
    ap.add_argument('--proyecto', default=str(AQUI.parent / 'proyectos' / 'hidroarma.json'))
    ap.add_argument('--snapshot', default=str(AQUI.parent / 'estado' / 'hidroarma_snapshot.json'))
    ap.add_argument('--no-guardar-snapshot', action='store_true')
    a = ap.parse_args()
    P = json.loads(Path(a.proyecto).read_text(encoding='utf-8'))
    out = Path(a.salida)
    out.mkdir(parents=True, exist_ok=True)
    cfg = {'inventario_json': a.inventario, 'matriz_xlsx': a.matriz, 'cuadernos_json': a.cuadernos, 'borradores_json': a.borradores,
           'corte': a.corte, 'snapshot': a.snapshot, 'guardar_snapshot': not a.no_guardar_snapshot, 'salida_data': str(out / 'data_motor.json')}
    d = sync.construir(cfg, P)
    n = exportar_artefacto.convertir(d, a.corte, out)
    m = d['meta']
    print(f"Corte {a.corte}: {m['archivos']} archivos, {m['carpetas']} carpetas, {m['imagenes']} imágenes ({m['fotos_distintas']} distintas), "
          f"cuadernos {m['cuadernos']}/{m['fuentes_cuadernos']} fuentes")
    c = d['cambios']
    print('Cambios:', 'primera lectura' if c.get('primera') else f"{len(c['nuevos'])} nuevos, {len(c['modificados'])} modificados, {len(c['eliminados'])} eliminados")
    print('Conjuntos:', n)


if __name__ == '__main__':
    sys.exit(main())
