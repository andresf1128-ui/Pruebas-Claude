#!/usr/bin/env python3
"""Estadística del estudio de mercado rural.

Lee un CSV de ofertas, aplica el descuento de negociación, calcula el valor unitario
y los estadísticos (n, promedio, mediana, desviación muestral, coeficiente de variación,
límites y cumplimiento del CV máximo). Los parámetros salen de config/parametros-941.yaml.

CSV de entrada (encabezados obligatorios):
    id, valor_oferta, area_ha            valor total ofertado y área del terreno en hectáreas
Opcionales:
    valor_construcciones                 valor depreciado de las construcciones (se resta al valor negociado)
    incluir                              SI/NO (por defecto SI)
    factor_homogeneizacion               producto de los factores aplicados al $/ha (por defecto 1)

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


def calcular(filas, descuento, cv_max):
    detalle = []
    for i, f in enumerate(filas, start=2):
        valor = numero(f["valor_oferta"], "valor_oferta", i)
        area = numero(f["area_ha"], "area_ha", i)
        if area <= 0 or valor <= 0:
            raise ValueError(f"Fila {i}: valor y área deben ser mayores que cero")
        constr = numero(f.get("valor_construcciones") or 0, "valor_construcciones", i)
        factor = numero(f.get("factor_homogeneizacion") or 1, "factor_homogeneizacion", i)
        negociado = valor * (1 - descuento)
        terreno = negociado - constr
        if terreno <= 0:
            raise ValueError(f"Fila {i}: las construcciones superan el valor negociado")
        unitario = terreno / area
        incluir = (f.get("incluir") or "SI").strip().upper() in ("SI", "SÍ", "S", "1", "TRUE")
        detalle.append({
            "id": f["id"], "valor_oferta": valor, "descuento": descuento,
            "valor_negociado": negociado, "valor_construcciones": constr,
            "valor_terreno": terreno, "area_ha": area, "valor_ha": unitario,
            "factor_homogeneizacion": factor, "valor_ha_homogeneizado": unitario * factor,
            "incluir": "SI" if incluir else "NO",
        })
    datos = [d["valor_ha_homogeneizado"] for d in detalle if d["incluir"] == "SI"]
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
    avisos = []
    for nombre, p in (("descuento_ofertas", desc), ("cv_maximo", cv)):
        if p["estado"] != "confirmado":
            avisos.append(f"{nombre} = {p['valor']} está 'por_confirmar' frente al texto oficial de la 941")
    try:
        detalle, est = calcular(leer_ofertas(a.ofertas), desc["valor"], cv["valor"])
    except ValueError as e:
        sys.exit(f"ERROR: {e}")

    for av in avisos:
        print(f"AVISO: {av}", file=sys.stderr)
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
