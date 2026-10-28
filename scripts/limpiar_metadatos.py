#!/usr/bin/env python3
"""Limpia metadatos de autoría y herramientas en los entregables y verifica que no queden rastros.

Soporta .docx, .xlsx, .pptx (OOXML) y .pdf (con pypdf si está instalado). Trabaja sobre una COPIA
(archivo_limpio.ext) y deja el original intacto.

Qué hace en OOXML:
  - docProps/core.xml: fija creator y lastModifiedBy al perito (--autor) y vacía keywords, description y category.
  - docProps/app.xml: vacía Application, AppVersion, Company si no se indica --empresa.
  - Elimina comments*.xml y sus referencias de rels/Content_Types, y lista si quedan cambios controlados (w:ins/w:del).
Qué verifica (todos los formatos): busca términos prohibidos en el texto y en los metadatos y los reporta.

Uso:
    python3 scripts/limpiar_metadatos.py informe.docx --autor "Nombre del perito" [--empresa "Firma"]
    python3 scripts/limpiar_metadatos.py informe.docx --solo-verificar
"""
import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

PROHIBIDOS = [
    r"\bclaude\b", r"\banthropic\b", r"\bchatgpt\b", r"\bopenai\b", r"\bgemini\b", r"\bcopilot\b",
    r"inteligencia artificial", r"\bIA\b", r"\bLLM\b", r"asistente virtual", r"generado por",
    r"\bTODO\b", r"por verificar", r"por confirmar", r"pendiente de", r"\bborrador\b",
]
PATRON = re.compile("|".join(PROHIBIDOS), re.IGNORECASE)
OOXML = {".docx", ".xlsx", ".pptx"}


def vaciar_elemento(xml, etiqueta, valor=""):
    patron = re.compile(rf"<{etiqueta}(\s[^>]*)?>.*?</{etiqueta}>|<{etiqueta}(\s[^>]*)?/>", re.DOTALL)
    nuevo = f"<{etiqueta}>{valor}</{etiqueta}>"
    if patron.search(xml):
        return patron.sub(nuevo, xml, count=1)
    return xml


def limpiar_ooxml(origen, destino, autor, empresa):
    with zipfile.ZipFile(origen) as zin, zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            nombre = item.filename
            datos = zin.read(nombre)
            if re.match(r"(word|xl|ppt)/(comments\w*|commentAuthors)\.xml$", nombre):
                continue  # se eliminan los comentarios
            if nombre == "docProps/core.xml":
                x = datos.decode("utf-8")
                x = vaciar_elemento(x, "dc:creator", autor)
                x = vaciar_elemento(x, "cp:lastModifiedBy", autor)
                for t in ("cp:keywords", "dc:description", "cp:category", "dc:subject"):
                    x = vaciar_elemento(x, t)
                datos = x.encode("utf-8")
            elif nombre == "docProps/app.xml":
                x = datos.decode("utf-8")
                x = vaciar_elemento(x, "Application")
                x = vaciar_elemento(x, "AppVersion")
                x = vaciar_elemento(x, "Company", empresa)
                datos = x.encode("utf-8")
            elif nombre.endswith(".rels") or nombre == "[Content_Types].xml":
                x = datos.decode("utf-8")
                x = re.sub(r"<Relationship\b[^>]*comments\w*\.xml[^>]*/>", "", x)
                x = re.sub(r"<Relationship\b[^>]*commentAuthors\.xml[^>]*/>", "", x)
                x = re.sub(r"<Override\b[^>]*(comments\w*|commentAuthors)\.xml[^>]*/>", "", x)
                datos = x.encode("utf-8")
            elif nombre.endswith(".xml") and nombre.startswith("word/"):
                x = datos.decode("utf-8")
                x = re.sub(r"<w:commentRangeStart\b[^>]*/>|<w:commentRangeEnd\b[^>]*/>", "", x)
                x = re.sub(r"<w:r>(?:(?!</w:r>).)*?<w:commentReference\b[^>]*/>.*?</w:r>", "", x, flags=re.DOTALL)
                datos = x.encode("utf-8")
            zout.writestr(item, datos)


def texto_de(ruta):
    ext = ruta.suffix.lower()
    partes = {}
    if ext in OOXML:
        with zipfile.ZipFile(ruta) as z:
            for n in z.namelist():
                if n.endswith((".xml", ".rels")):
                    xml = z.read(n).decode("utf-8", "ignore")
                    partes[n] = re.sub(r"<[^>]+>", " ", xml)
                    if n.startswith("word/") and re.search(r"<w:(ins|del)\b", xml):
                        partes[n + " [cambios controlados]"] = "cambios controlados presentes"
    elif ext == ".pdf":
        try:
            from pypdf import PdfReader
            r = PdfReader(str(ruta))
            partes["metadatos"] = " ".join(str(v) for v in (r.metadata or {}).values())
            partes["texto"] = " ".join((p.extract_text() or "") for p in r.pages)
        except ImportError:
            print("AVISO: instale pypdf para verificar PDF", file=sys.stderr)
    return partes


def verificar(ruta):
    hallazgos = []
    for parte, texto in texto_de(ruta).items():
        if "cambios controlados" in parte:
            hallazgos.append((parte, "hay cambios controlados sin aceptar"))
        for m in PATRON.finditer(texto):
            ini = max(0, m.start() - 30)
            hallazgos.append((parte, texto[ini:m.end() + 30].strip()))
    return hallazgos


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("archivo")
    ap.add_argument("--autor", default="")
    ap.add_argument("--empresa", default="")
    ap.add_argument("--solo-verificar", action="store_true")
    a = ap.parse_args()
    ruta = Path(a.archivo)
    if not ruta.is_file():
        sys.exit(f"ERROR: no existe {ruta}")
    objetivo = ruta
    if not a.solo_verificar:
        if ruta.suffix.lower() not in OOXML:
            sys.exit("Limpieza automática solo para .docx/.xlsx/.pptx; en PDF use exportar de nuevo desde el documento limpio.")
        if not a.autor:
            sys.exit("ERROR: indique --autor (nombre del perito)")
        objetivo = ruta.with_name(ruta.stem + "_limpio" + ruta.suffix)
        limpiar_ooxml(ruta, objetivo, a.autor, a.empresa)
        print(f"Copia limpia: {objetivo}")
    hall = verificar(objetivo)
    if hall:
        print(f"{len(hall)} hallazgo(s) en {objetivo.name}:")
        for parte, ctx in hall:
            print(f"  - [{parte}] ...{ctx}...")
        sys.exit(1)
    print("Sin rastros de herramientas ni textos de trabajo.")


if __name__ == "__main__":
    main()
