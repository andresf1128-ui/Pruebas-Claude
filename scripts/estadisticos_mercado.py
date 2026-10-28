#!/usr/bin/env python3
"""Estadística del estudio de mercado rural (Resolución IGAC 941 de 2026).

Lee un CSV de ofertas, aplica el porcentaje de negociación, descuenta cultivos permanentes y
construcciones, calcula el valor del terreno por hectárea y los estadísticos (n, promedio, mediana,
desviación muestral, coeficiente de variación, límites y cumplimiento del CV máximo, art. 21).

No homogeneiza por factores: el anexo técnico 2.1 (Nota 2) lo excluye. La depuración se hace sobre el
valor de oferta ya negociado (art. 17 c). El porcentaje de negociación debe estar justificado (art. 17 b).

CSV de entrada (encabezados obligatorios):
    id, valor_oferta, area_ha            valor total ofertado y área del terreno en hectáreas
Opcionales:
    negociacion                          porcentaje de negociación de la oferta (0,10 = 10 %). Si falta, se usa el parámetro
    valor_construcciones                 valor depreciado de construcciones y anexos (se resta al valor negociado)
    valor_cultivos                       valor de cultivos permanentes e infraestructura productiva (se resta)
    incluir                              SI/NO (por defecto SI)

Uso:
    python3 scripts/estadisticos_mercado.py ofertas.csv [--salida resultado.csv]
"""
import argparse
import csv
import statistics
import sys
from pathlib import Path

import yaml

RAIZ = Path(__file__).resolve().parent.parent


def cargar_parametros(ruta=None):
    ruta = Path(ruta) if ruta else RAIZ / "config" / "parametros-941.yaml"
    with open(ruta, encoding="utf-8") as f:
        p = yaml.safe_load(f)["mercado_rural"]
    return p["descuento_ofertas"], p["cv_maximo"]


def numero(texto, campo, fila):
    try:
        return float(str(texto).replace("$", "").replace(" ", "").replace(",", "."))
    except ValueError:
        raise ValueError(f"Fila {fila}: valor no numérico en '{campo}': {texto!r}")


def leer_ofertas(ruta):
    with open(ruta, encoding="utf-8-sig", newline="") as f:
        lector = csv.DictReader(f)
        faltan = {"id", "valor_oferta", "area_ha"} - set(lector.fieldnames or [])
        if faltan:
            raise ValueError(f"Faltan columnas obligatorias: {', '.join(sorted(faltan))}")
        return list(lector)


def calcular(filas, negociacion_defecto, cv_max):
    detalle = []
    for i, f in enumerate(filas, start=2):
        valor = numero(f["valor_oferta"], "valor_oferta", i)
        area = numero(f["area_ha"], "area_ha", i)
        if area <= 0 or valor <= 0:
            raise ValueError(f"Fila {i}: valor y área deben ser mayores que cero")
        neg = numero(f["negociacion"], "negociacion", i) if (f.get("negociacion") or "").strip() else negociacion_defecto
        if not 0 <= neg < 1:
            raise ValueError(f"Fila {i}: la negociación debe estar entre 0 y 1 (por ejemplo 0,10)")
        constr = numero(f.get("valor_construcciones") or 0, "valor_construcciones", i)
        cultivos = numero(f.get("valor_cultivos") or 0, "valor_cultivos", i)
        negociado = valor * (1 - neg)
        terreno = negociado - constr - cultivos
        if terreno <= 0:
            raise ValueError(f"Fila {i}: construcciones y cultivos superan el valor negociado")
        incluir = (f.get("incluir") or "SI").strip().upper() in ("SI", "SÍ", "S", "1", "TRUE")
        detalle.append({
            "id": f["id"], "valor_oferta": valor, "negociacion": neg,
            "valor_negociado": negociado, "valor_construcciones": constr, "valor_cultivos": cultivos,
            "valor_terreno": terreno, "area_ha": area, "valor_terreno_ha": terreno / area,
            "incluir": "SI" if incluir else "NO",
        })
    datos = [d["valor_terreno_ha"] for d in detalle if d["incluir"] == "SI"]
    est = {"n": len(datos)}
    if datos:
        prom = statistics.fmean(datos)
        est["promedio"] = prom
        est["mediana"] = statistics.median(datos)
        if len(datos) > 1:
            d = statistics.stdev(datos)
            est["desviacion"] = d
            est["cv"] = d / prom
            est["limite_inferior"] = prom - d
            est["limite_superior"] = prom + d
            est["cumple_cv"] = est["cv"] <= cv_max
    est["cv_maximo"] = cv_max
    return detalle, est


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("ofertas")
    ap.add_argument("--salida", help="CSV con el detalle calculado")
    ap.add_argument("--parametros", help="YAML alternativo de parámetros")
    a = ap.parse_args()

    desc, cv = cargar_parametros(a.parametros)
    if desc.get("requiere_justificacion_escrita"):
        print("AVISO: el porcentaje de negociación debe justificarse en el informe (art. 17 b de la 941)", file=sys.stderr)
    if cv["estado"] != "confirmado":
        print(f"AVISO: cv_maximo = {cv['valor']} está por confirmar", file=sys.stderr)
    try:
        detalle, est = calcular(leer_ofertas(a.ofertas), desc["valor"], cv["valor"])
    except ValueError as e:
        sys.exit(f"ERROR: {e}")

    for k, v in est.items():
        print(f"{k}: {v:.4f}" if isinstance(v, float) else f"{k}: {v}")
    if a.salida:
        with open(a.salida, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(detalle[0].keys()))
            w.writeheader()
            w.writerows(detalle)
        print(f"Detalle escrito en {a.salida}")


if __name__ == "__main__":
    main()
