#!/usr/bin/env python3
"""Copia, ordena y rotula las fotos de la visita de campo.

Estructura de origen aceptada:
    origen/P/*.jpg  origen/C1/*.jpg  origen/C2/*.jpg   subcarpeta = unidad (P = predio, C# = construcción)
    origen/*.jpg                                         sin subcarpeta: unidad P

Nombre de salida:  CODIGO_Predio_Unidad_NNN.jpg   (NNN consecutivo por unidad, en orden de toma)
Las categorías (fachada, cubierta, etc.) se agregan después con la skill organizar-fotos-avaluos.

Los originales no se tocan: todo se copia. Genera manifiesto_fotos.csv con fecha, hora,
coordenadas y ruta original, y lista las fotos sin fecha o sin GPS.
"""
import argparse
import csv
import re
import shutil
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

from PIL import ExifTags, Image, ImageOps

EXT = {".jpg", ".jpeg", ".png"}
TAGS = {v: k for k, v in ExifTags.TAGS.items()}


def limpio(texto):
    t = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9]+", "", t.title())


def a_grados(valor, ref):
    d, m, s = (float(x) for x in valor)
    g = d + m / 60 + s / 3600
    return -g if ref in ("S", "W") else g


def leer_exif(ruta):
    fecha, lat, lon = None, None, None
    try:
        with Image.open(ruta) as im:
            ex = im.getexif()
            sub = ex.get_ifd(TAGS["ExifOffset"])
            bruto = sub.get(TAGS["DateTimeOriginal"]) or ex.get(TAGS["DateTime"])
            if bruto:
                fecha = datetime.strptime(str(bruto), "%Y:%m:%d %H:%M:%S")
            gps = ex.get_ifd(TAGS["GPSInfo"])
            if gps and 2 in gps and 4 in gps:
                lat = a_grados(gps[2], gps.get(1, "N"))
                lon = a_grados(gps[4], gps.get(3, "E"))
    except Exception:
        pass
    return fecha, lat, lon


def unidad_de(ruta, origen):
    rel = ruta.relative_to(origen).parts
    return rel[0].upper() if len(rel) > 1 else "P"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--origen", required=True)
    ap.add_argument("--salida", required=True)
    ap.add_argument("--codigo", required=True, help="código del avalúo, p. ej. AV-045")
    ap.add_argument("--predio", required=True, help="nombre corto del predio")
    a = ap.parse_args()

    origen, salida = Path(a.origen), Path(a.salida)
    if not origen.is_dir():
        sys.exit(f"ERROR: no existe la carpeta de origen {origen}")
    if salida.exists() and any(salida.iterdir()):
        sys.exit(f"ERROR: la carpeta de salida {salida} no está vacía; use otra para no sobrescribir")
    salida.mkdir(parents=True, exist_ok=True)

    fotos = []
    for p in sorted(origen.rglob("*")):
        if p.suffix.lower() in EXT and p.is_file():
            fecha, lat, lon = leer_exif(p)
            fotos.append({"ruta": p, "unidad": unidad_de(p, origen), "fecha": fecha,
                          "lat": lat, "lon": lon, "mtime": datetime.fromtimestamp(p.stat().st_mtime)})
    if not fotos:
        sys.exit("ERROR: no se encontraron fotos")

    fotos.sort(key=lambda f: (f["unidad"], f["fecha"] or f["mtime"], f["ruta"].name))
    contador, filas, sin_fecha, sin_gps = {}, [], [], []
    base = f"{a.codigo}_{limpio(a.predio)}"
    for f in fotos:
        u = f["unidad"]
        contador[u] = contador.get(u, 0) + 1
        nombre = f"{base}_{u}_{contador[u]:03d}.jpg"
        destino = salida / u / nombre
        destino.parent.mkdir(exist_ok=True)
        with Image.open(f["ruta"]) as im:
            exif = im.getexif()
            img = ImageOps.exif_transpose(im).convert("RGB")
            img.save(destino, "JPEG", quality=95, exif=exif.tobytes() if len(exif) else b"")
        if f["fecha"] is None:
            sin_fecha.append(f["ruta"].name)
        if f["lat"] is None:
            sin_gps.append(f["ruta"].name)
        filas.append({
            "archivo": nombre, "unidad": u, "original": str(f["ruta"]),
            "fecha_hora": (f["fecha"] or "").isoformat(sep=" ") if f["fecha"] else "",
            "latitud": f"{f['lat']:.6f}" if f["lat"] is not None else "",
            "longitud": f"{f['lon']:.6f}" if f["lon"] is not None else "",
        })

    with open(salida / "manifiesto_fotos.csv", "w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(filas[0].keys()))
        w.writeheader()
        w.writerows(filas)

    print(f"{len(filas)} fotos copiadas a {salida}")
    for u, n in sorted(contador.items()):
        print(f"  {u}: {n}")
    if sin_fecha:
        print(f"AVISO: {len(sin_fecha)} sin fecha EXIF (se ordenaron por fecha de archivo): {', '.join(sin_fecha[:10])}")
    if sin_gps:
        print(f"AVISO: {len(sin_gps)} sin coordenadas GPS: {', '.join(sin_gps[:10])}")


if __name__ == "__main__":
    main()
