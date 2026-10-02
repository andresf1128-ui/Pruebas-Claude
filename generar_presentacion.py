#!/usr/bin/env python3
"""Genera la presentación (10 diapositivas) a partir del Excel del estudio de mercado.

Uso:  python generar_presentacion.py [estudio.xlsx] [salida.pptx]

Lee los VALORES CALCULADOS guardados por Excel (guarde el libro antes de ejecutar).
Los colores replican el libro: azul marino 1F3864, azul 2F5597, grises BFBFBF/D9D9D9,
verde claro E2EFDA (resultado clave) y amarillo FFF2CC (entrada).
"""
import sys
from pathlib import Path

import openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt, Emu

# ---------- Identidad ----------
AVALUADOR = "Andrés Felipe Cuartas Montoya"
RAA = "RAA 1.128.391.782"

NAVY, BLUE = RGBColor(0x1F, 0x38, 0x64), RGBColor(0x2F, 0x55, 0x97)
G1, G2 = RGBColor(0xBF, 0xBF, 0xBF), RGBColor(0xD9, 0xD9, 0xD9)
GREEN, YELLOW = RGBColor(0xE2, 0xEF, 0xDA), RGBColor(0xFF, 0xF2, 0xCC)
OKGREEN, RED = RGBColor(0x37, 0x86, 0x3C), RGBColor(0xC0, 0x39, 0x2B)
WHITE, INK, MUTED = RGBColor(255, 255, 255), RGBColor(0x26, 0x2B, 0x33), RGBColor(0x6B, 0x72, 0x80)
FONT = "Calibri"
W, H = Inches(13.333), Inches(7.5)


# ---------- Utilidades de formato ----------
def num(v, d=0):
    return f"{v:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def cop(v):
    return "$ " + num(v)


def mill(v, d=1):
    return f"$ {num(v / 1e6, d)} M"


def pct(v, d=1):
    return num(v * 100, d) + " %"


def fnum(v, default=0.0):
    return float(v) if isinstance(v, (int, float)) else default


# ---------- Lectura del Excel ----------
def leer(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    m, c, d, l, p = (wb[n] for n in ("MERCADO RURAL", "COMBINACIONES", "DEPRECIACIÓN", "LIQUIDACIÓN", "PROYECCIÓN"))
    D = {}
    D["fecha"] = "/".join(str(m[a].value) for a in ("M2", "L2", "K2") if m[a].value is not None) if False else \
        f"{int(fnum(m['M2'].value)):02d}/{int(fnum(m['L2'].value)):02d}/{int(fnum(m['K2'].value))}"
    D["solicitante"] = m["F3"].value or "—"
    D["direccion"] = m["E4"].value or "—"
    D["depto"], D["municipio"] = m["D6"].value or "—", m["H6"].value or "—"
    D["referencia"] = str(d["A2"].value or "").split("·")[-1].strip()

    ofertas, r = [], 9
    while isinstance(m.cell(r, 1).value, (int, float)):
        g = lambda col: m.cell(r, col).value
        ofertas.append(dict(no=int(g(1)), ubic=g(2), pedido=fnum(g(4)), neg=fnum(g(5)), depurado=fnum(g(6)),
                            area=fnum(g(7)), ha=fnum(g(8)), constr=fnum(g(9)), desc=g(10), fuente=g(11),
                            fecha=g(15), comp=str(g(16) or "")))
        r += 1
    D["ofertas"] = ofertas
    D["obs"] = next((str(m.cell(x, 1).value) for x in range(r, r + 40)
                     if str(m.cell(x, 1).value or "").startswith("OBSERVACIONES")), "")

    est = {}
    for x in range(r, r + 12):
        k = str(m.cell(x, 6).value or "")
        if k:
            est[k.split(" (")[0]] = m.cell(x, 8).value
    D["est"] = est

    D["cv_max"] = fnum(c["B5"].value, 0.10)
    D["filtro"] = c["B6"].value
    D["elegibles"] = c["B7"].value
    D["combos"] = [dict(nombre=c.cell(x, 1).value, n=c.cell(x, 2).value, ofertas=c.cell(x, 3).value,
                        prom=fnum(c.cell(x, 6).value), de=fnum(c.cell(x, 7).value), cv=fnum(c.cell(x, 8).value),
                        cumple=c.cell(x, 12).value, m2=fnum(c.cell(x, 14).value))
                   for x in range(13, 16) if c.cell(x, 1).value]
    D["sel"] = dict(nombre=c["B18"].value, ofertas=c["B19"].value, cv=fnum(c["B20"].value), cumple=c["B21"].value,
                    ha=fnum(c["B22"].value), m2=fnum(c["B23"].value))

    D["proy"] = dict(baja=fnum(p["D5"].value), alta=fnum(p["F5"].value), amp=fnum(p["H5"].value),
                     estado=p["L5"].value, resumen=p["L8"].value, nota=p["A21"].value,
                     nbaja=p["D8"].value, nalta=p["F8"].value,
                     ref=[[p.cell(x, k).value for k in range(1, 8)] for x in (18, 19)])

    constr = []
    for x in range(5, 15):
        if d.cell(x, 2).value:
            constr.append(dict(item=d.cell(x, 2).value, edad=d.cell(x, 3).value, vida=d.cell(x, 4).value,
                               estado=d.cell(x, 5).value, repos=fnum(d.cell(x, 6).value), fd=fnum(d.cell(x, 12).value),
                               adop=fnum(d.cell(x, 14).value), area=fnum(d.cell(x, 15).value),
                               total=fnum(d.cell(x, 16).value), oferta=d.cell(x, 1).value))
    D["constr"] = constr
    D["constr_total"] = fnum(d["P15"].value)

    liq = dict(terreno=None, principal=[], anexas=[], tot_p=0, tot_a=0, total=0, letras="")
    sec = None
    for x in range(3, 30):
        b = str(l.cell(x, 2).value or "")
        if b.startswith("Notas"):
            break
        row = (b, l.cell(x, 3).value, fnum(l.cell(x, 4).value), fnum(l.cell(x, 5).value), fnum(l.cell(x, 6).value))
        if b == "TERRENO":
            sec = "t"
        elif b == "CONSTRUCCIÓN PRINCIPAL":
            sec = "p"
        elif b == "CONSTRUCCIONES ANEXAS":
            sec = "a"
        elif b == "DESCRIPCIÓN":
            continue
        elif b.startswith("VALOR TOTAL DE LA CONSTRUCCIÓN PRINCIPAL"):
            liq["tot_p"] = row[4]
        elif b.startswith("VALOR TOTAL DE LAS CONSTRUCCIONES ANEXAS"):
            liq["tot_a"] = row[4]
        elif b == "VALOR TOTAL DEL AVALÚO":
            liq["total"] = row[4]
        elif b == "VALOR TOTAL EN LETRAS":
            liq["letras"] = str(l.cell(x, 3).value or "")
        elif b == "TERRENO":
            continue
        elif b and sec == "t" and liq["terreno"] is None:
            liq["terreno"] = row
        elif b and sec == "p":
            liq["principal"].append(row)
        elif b and sec == "a":
            liq["anexas"].append(row)
    D["liq"] = liq
    return D


# ---------- Primitivas de diseño ----------
def rect(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    sh = s.shapes.add_shape(shape, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = fill
    if line is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = line
        sh.line.width = Pt(0.75)
    sh.shadow.inherit = False
    return sh


def text(s, x, y, w, h, t, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False):
    tb = s.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = Inches(0.05)
    lines = t if isinstance(t, list) else [t]
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = ln
        r.font.size, r.font.bold, r.font.italic = Pt(size), bold, italic
        r.font.color.rgb, r.font.name = color, FONT
    return tb


def base(prs, titulo, subtitulo, n, total):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, Inches(1.05), NAVY)
    rect(s, 0, Inches(1.05), W, Inches(0.06), BLUE)
    text(s, Inches(0.5), Inches(0.12), Inches(10.5), Inches(0.55), titulo, 26, True, WHITE)
    text(s, Inches(0.5), Inches(0.62), Inches(11), Inches(0.35), subtitulo, 13, False, G2)
    text(s, Inches(11.3), Inches(0.3), Inches(1.6), Inches(0.5), f"{n:02d} / {total:02d}", 14, True, G2, PP_ALIGN.RIGHT)
    rect(s, 0, Inches(7.1), W, Inches(0.4), G2)
    text(s, Inches(0.5), Inches(7.13), Inches(9), Inches(0.3),
         f"{AVALUADOR} · {RAA}", 10, False, NAVY, anchor=MSO_ANCHOR.MIDDLE)
    text(s, Inches(8.8), Inches(7.13), Inches(4.1), Inches(0.3),
         "Estudio de mercado · Gestión valuatoria", 10, False, MUTED, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)
    return s


def kpi(s, x, y, w, h, label, value, fill=GREEN, vcolor=NAVY, size=26):
    rect(s, x, y, w, h, fill)
    rect(s, x, y, Inches(0.08), h, NAVY)
    text(s, x + Inches(0.2), y + Inches(0.08), w - Inches(0.3), Inches(0.35), label, 11, True, MUTED)
    text(s, x + Inches(0.2), y + Inches(0.42), w - Inches(0.3), h - Inches(0.5), value, size, True, vcolor,
         anchor=MSO_ANCHOR.MIDDLE)


def tabla(s, x, y, w, rows, widths, hdr_size=11, size=10.5, row_h=0.36, align=None, hl_last=False):
    """rows[0] = encabezado. align: lista de PP_ALIGN por columna."""
    gt = s.shapes.add_table(len(rows), len(rows[0]), x, y, w, Inches(row_h * len(rows)))
    t = gt.table
    tot = sum(widths)
    for i, wd in enumerate(widths):
        t.columns[i].width = int(w * wd / tot)
    for i, row in enumerate(rows):
        t.rows[i].height = Inches(row_h)
        for j, v in enumerate(row):
            cell = t.cell(i, j)
            cell.margin_left = cell.margin_right = Inches(0.06)
            cell.margin_top = cell.margin_bottom = Inches(0.02)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            last = hl_last and i == len(rows) - 1
            cell.fill.fore_color.rgb = NAVY if i == 0 else (GREEN if last else (WHITE if i % 2 else RGBColor(0xF2, 0xF2, 0xF2)))
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if i == 0 else (align[j] if align else PP_ALIGN.LEFT)
            r = p.add_run()
            r.text = str(v)
            r.font.name = FONT
            r.font.size = Pt(hdr_size if i == 0 else size)
            r.font.bold = i == 0 or last
            r.font.color.rgb = WHITE if i == 0 else INK
    return t


def estilo_chart(ch, legend=True):
    ch.font.size, ch.font.name = Pt(11), FONT
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False


def recortar(t, n):
    t = str(t or "")
    return t if len(t) <= n else t[: n - 1].rstrip() + "…"


def es_descartada(o):
    return o["comp"].upper().startswith("DESCARTADA")


# ---------- Diapositivas ----------
def construir(D, salida):
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    TOTAL = 10
    of, st, sel, py, lq = D["ofertas"], D["est"], D["sel"], D["proy"], D["liq"]

    # 1. Portada
    s = prs.slides.add_slide(prs.slide_layouts[6])
    rect(s, 0, 0, W, H, NAVY)
    rect(s, 0, Inches(5.2), W, Inches(2.3), RGBColor(0x18, 0x2C, 0x50))
    rect(s, Inches(0.7), Inches(1.55), Inches(0.12), Inches(2.35), G1)
    text(s, Inches(1.0), Inches(0.9), Inches(11), Inches(0.5), "GESTIÓN VALUATORIA", 16, True, G1)
    text(s, Inches(1.0), Inches(1.5), Inches(11.5), Inches(1.6), ["Estudio de mercado", "predios rurales"], 46, True, WHITE)
    text(s, Inches(1.0), Inches(3.2), Inches(11.5), Inches(0.9),
         recortar(D["direccion"].split(". Coord")[0], 110), 18, False, G2)
    text(s, Inches(1.0), Inches(3.95), Inches(11), Inches(0.4),
         f"{D['municipio']} · {D['depto']}", 14, False, G1)
    rect(s, Inches(1.0), Inches(5.55), Inches(0.9), Inches(0.05), GREEN)
    text(s, Inches(1.0), Inches(5.7), Inches(8), Inches(0.5), "Avaluador", 12, False, G1)
    text(s, Inches(1.0), Inches(6.0), Inches(9), Inches(0.5), AVALUADOR, 24, True, WHITE)
    text(s, Inches(1.0), Inches(6.5), Inches(9), Inches(0.4), RAA, 15, False, GREEN)
    text(s, Inches(9.3), Inches(5.7), Inches(3.5), Inches(0.5), "Fecha del estudio", 12, False, G1, PP_ALIGN.RIGHT)
    text(s, Inches(9.3), Inches(6.0), Inches(3.5), Inches(0.5), D["fecha"], 24, True, WHITE, PP_ALIGN.RIGHT)
    text(s, Inches(7.3), Inches(6.5), Inches(5.5), Inches(0.4), f"Solicitante: {D['solicitante']}", 13, False, G1, PP_ALIGN.RIGHT)

    # 2. Resumen ejecutivo
    s = base(prs, "Resumen ejecutivo", "Resultado del estudio y estado frente a la Res. 941", 2, TOTAL)
    nvig = sum(1 for o in of if not es_descartada(o))
    cumple = str(sel["cumple"]).upper().startswith("S")
    kpi(s, Inches(0.5), Inches(1.4), Inches(3.0), Inches(1.4), "VALOR TOTAL DEL AVALÚO", mill(lq["total"], 1), GREEN, NAVY, 28)
    kpi(s, Inches(3.65), Inches(1.4), Inches(3.0), Inches(1.4), "VALOR ADOPTADO TERRENO", mill(fnum(st.get("VALOR ADOPTADO")), 1) + " /ha", GREEN, NAVY, 24)
    kpi(s, Inches(6.8), Inches(1.4), Inches(3.0), Inches(1.4), "OFERTAS EN LA TABLA", f"{len(of)} · {nvig} vigente" + ("" if nvig == 1 else "s"), G2, NAVY, 22)
    kpi(s, Inches(9.95), Inches(1.4), Inches(2.9), Inches(1.4), "CV MEJOR COMBINACIÓN", pct(sel["cv"]), YELLOW if not cumple else GREEN, RED if not cumple else OKGREEN, 28)
    text(s, Inches(0.5), Inches(3.0), Inches(6), Inches(0.4), "Lectura del estudio", 16, True, NAVY)
    puntos = [
        f"Valor en letras: {lq['letras'].capitalize()}.",
        f"CV de la muestra general: {pct(fnum(st.get('COEFICIENTE DE VARIACIÓN')))} frente al máximo admisible de {pct(D['cv_max'], 0)}.",
        f"Combinación seleccionada: {sel['nombre']} — ofertas {sel['ofertas']} → {mill(sel['ha'])} /ha · {cop(sel['m2'])} /m².",
        f"Rango de orientación: {mill(py['baja'])} – {mill(py['alta'])} por ha ({py['resumen']}).",
        f"Estado Res. 941: {py['estado']}. " + ("Se requieren más ofertas del mismo segmento." if not cumple else "La combinación cumple el CV máximo."),
    ]
    for i, t_ in enumerate(puntos):
        y = Inches(3.5 + i * 0.62)
        rect(s, Inches(0.5), y + Inches(0.1), Inches(0.14), Inches(0.14), BLUE, shape=MSO_SHAPE.OVAL)
        text(s, Inches(0.8), y, Inches(12), Inches(0.6), t_, 15, False, INK)
    rect(s, Inches(0.5), Inches(6.55), Inches(12.35), Inches(0.42), YELLOW if not cumple else GREEN)
    text(s, Inches(0.6), Inches(6.55), Inches(12.2), Inches(0.42),
         "Valores orientativos: " + ("no sustentan un valor adoptado hasta cumplir CV ≤ " + pct(D["cv_max"], 0) + "." if not cumple else "cumplen el CV máximo admisible."),
         12, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)

    # 3. Predio y alcance
    s = base(prs, "Predio objeto y alcance", "Identificación del inmueble y marco metodológico", 3, TOTAL)
    ident = [("Dirección / ubicación", D["direccion"]), ("Departamento", D["depto"]), ("Municipio", D["municipio"]),
             ("Solicitante", D["solicitante"]), ("Fecha del estudio", D["fecha"]), ("Referencia de liquidación", D["referencia"]),
             ("Superficie del terreno", f"{num(lq['terreno'][2], 2)} ha" if lq["terreno"] else "—")]
    for i, (k, v) in enumerate(ident):
        y = Inches(1.4 + i * 0.7)
        rect(s, Inches(0.5), y, Inches(2.5), Inches(0.62), G2)
        text(s, Inches(0.6), y, Inches(2.4), Inches(0.62), k, 12, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        rect(s, Inches(3.0), y, Inches(4.4), Inches(0.62), RGBColor(0xF7, 0xF7, 0xF7))
        text(s, Inches(3.1), y, Inches(4.25), Inches(0.62), recortar(v, 95), 11, False, INK, anchor=MSO_ANCHOR.MIDDLE)
    rect(s, Inches(7.8), Inches(1.4), Inches(5.05), Inches(4.8), GREEN)
    text(s, Inches(8.0), Inches(1.5), Inches(4.7), Inches(0.4), "Método y criterios", 16, True, NAVY)
    mets = [("Método de mercado", "Comparación de ofertas depuradas por % de negociación."),
            ("Estadística", "Promedio, desviación, CV, límites y asimetría de la muestra."),
            ("Res. 941 · rural", f"CV máximo admisible {pct(D['cv_max'], 0)}; sin homogenización por factores."),
            ("Combinaciones", "Grupos de 3 o más ofertas por menor CV."),
            ("Construcciones", "Depreciación Ross-Heidecke."),
            ("Liquidación", "Terreno + construcciones, redondeo a la centena.")]
    for i, (k, v) in enumerate(mets):
        y = Inches(2.0 + i * 0.68)
        text(s, Inches(8.0), y, Inches(4.7), Inches(0.3), k, 12, True, BLUE)
        text(s, Inches(8.0), y + Inches(0.27), Inches(4.7), Inches(0.4), v, 11, False, INK)

    # 4. Ofertas (tabla)
    s = base(prs, "Ofertas de mercado", "Datos tomados de la hoja MERCADO RURAL (valores depurados por % de negociación)", 4, TOTAL)
    rows = [["No.", "Ubicación", "Área ha", "Valor pedido", "% neg.", "Valor depurado", "Valor $/ha", "Comparabilidad"]]
    for o in of:
        rows.append([o["no"], recortar(o["ubic"], 48), num(o["area"], 2), mill(o["pedido"], 0), pct(o["neg"], 0),
                     mill(o["depurado"], 0), mill(o["ha"], 2), recortar(o["comp"].split(" – ")[0].split(" en ")[0], 12)])
    n = len(of)
    tabla(s, Inches(0.4), Inches(1.3), Inches(12.55), rows, [0.5, 4.2, 0.9, 1.2, 0.7, 1.3, 1.3, 1.6],
          size=10 if n > 9 else 11, row_h=min(0.55, 5.2 / (n + 1)),
          align=[PP_ALIGN.CENTER, PP_ALIGN.LEFT] + [PP_ALIGN.RIGHT] * 5 + [PP_ALIGN.CENTER])
    text(s, Inches(0.4), Inches(6.6), Inches(12.5), Inches(0.45),
         "M = millones de pesos. ALTA = comparable en uso agropecuario; DESCARTADA = turístico, playa, campestre o dato no confiable.",
         10, False, MUTED, italic=True)

    # 5. Gráfico y estadísticos
    s = base(prs, "Valor por hectárea y estadísticos", "Dispersión de la muestra general ($ millones por ha, escala logarítmica)", 5, TOTAL)
    cd = CategoryChartData()
    cd.categories = [f"Of. {o['no']}" for o in of]
    cd.add_series("Valor $/ha (millones)", [round(o["ha"] / 1e6, 3) for o in of])
    adop = fnum(st.get("VALOR ADOPTADO")) / 1e6
    cd.add_series("Valor adoptado", [round(adop, 3)] * len(of))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.4), Inches(1.3), Inches(7.9), Inches(5.6), cd)
    ch = gf.chart
    estilo_chart(ch)
    ch.plots[0].gap_width = 60
    ch.plots[0].has_data_labels = False
    ser = ch.plots[0].series[0]
    ser.format.fill.solid()
    ser.format.fill.fore_color.rgb = BLUE
    for i, o in enumerate(of):
        pt = ser.points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = G1 if es_descartada(o) else NAVY
    ser.data_labels.show_value = True
    ser.data_labels.font.size = Pt(9)
    ser.data_labels.number_format = '#,##0.0'
    ser.data_labels.number_format_is_linked = False
    ser.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    s2 = ch.plots[0].series[1]
    s2.format.fill.solid()
    s2.format.fill.fore_color.rgb = RGBColor(0x92, 0xD0, 0x50)
    va = ch.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = G2
    va.tick_labels.font.size = Pt(10)
    va.minimum_scale = 1
    scaling = va._element.find('{http://schemas.openxmlformats.org/drawingml/2006/chart}scaling')
    from lxml import etree
    ns = 'http://schemas.openxmlformats.org/drawingml/2006/chart'
    lb = etree.SubElement(scaling, f'{{{ns}}}logBase')
    lb.set('val', '10')
    scaling.remove(lb)
    scaling.insert(0, lb)
    if scaling.find(f'{{{ns}}}orientation') is None:
        ori = etree.Element(f'{{{ns}}}orientation')
        ori.set('val', 'minMax')
        scaling.insert(1, ori)
    ch.category_axis.tick_labels.font.size = Pt(10)
    items = [("PROMEDIO", lambda v: mill(v)), ("DESVIACIÓN ESTÁNDAR", lambda v: mill(v)),
             ("COEFICIENTE DE VARIACIÓN", pct), ("LÍMITE INFERIOR", lambda v: mill(v)),
             ("LÍMITE SUPERIOR", lambda v: mill(v)), ("COEFICIENTE DE ASIMETRÍA", lambda v: num(v, 2)),
             ("VALOR ADOPTADO", lambda v: mill(v))]
    for i, (k, f) in enumerate(items):
        y = Inches(1.35 + i * 0.72)
        key = k == "VALOR ADOPTADO"
        rect(s, Inches(8.6), y, Inches(4.3), Inches(0.64), GREEN if key else G2)
        text(s, Inches(8.7), y, Inches(2.3), Inches(0.64), k.capitalize(), 11, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        v = st.get(k)
        text(s, Inches(10.6), y, Inches(2.25), Inches(0.64), f(fnum(v)) if v is not None else "—", 14, True,
             RED if k == "COEFICIENTE DE VARIACIÓN" and fnum(v) > D["cv_max"] else NAVY, PP_ALIGN.RIGHT, MSO_ANCHOR.MIDDLE)

    # 6. Combinaciones
    s = base(prs, "Combinaciones de ofertas", f"Grupos de 3 o más ofertas · CV máximo admisible {pct(D['cv_max'], 0)} (Res. 941)", 6, TOTAL)
    cd = CategoryChartData()
    cd.categories = [str(c_["nombre"]).replace("Combinación", "Comb.") for c_ in D["combos"]]
    cd.add_series("Coeficiente de variación", [round(c_["cv"], 4) for c_ in D["combos"]])
    cd.add_series("CV máximo admisible", [D["cv_max"]] * len(D["combos"]))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.4), Inches(1.3), Inches(5.6), Inches(3.9), cd)
    ch = gf.chart
    estilo_chart(ch)
    ch.plots[0].gap_width = 70
    for k_, col in enumerate((NAVY, RGBColor(0x92, 0xD0, 0x50))):
        ch.plots[0].series[k_].format.fill.solid()
        ch.plots[0].series[k_].format.fill.fore_color.rgb = col
    ch.plots[0].series[0].data_labels.show_value = True
    ch.plots[0].series[0].data_labels.number_format = '0.0%'
    ch.plots[0].series[0].data_labels.number_format_is_linked = False
    ch.value_axis.tick_labels.number_format = '0%'
    ch.value_axis.tick_labels.number_format_is_linked = False
    ch.value_axis.major_gridlines.format.line.color.rgb = G2
    rows = [["Combinación", "Ofertas", "Promedio $/ha", "Desv. est.", "CV", "Cumple", "$/m²"]]
    for c_ in D["combos"]:
        rows.append([c_["nombre"], c_["ofertas"], mill(c_["prom"]), mill(c_["de"]), pct(c_["cv"]), c_["cumple"], cop(c_["m2"])])
    tabla(s, Inches(6.2), Inches(1.3), Inches(6.75), rows, [1.7, 1.1, 1.4, 1.1, 0.8, 1.0, 1.1], hdr_size=10, size=10, row_h=0.55,
          align=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT])
    text(s, Inches(6.2), Inches(3.6), Inches(6.7), Inches(1.5),
         [f"Ofertas elegibles: {D['elegibles']} de {len(of)} · filtro por límites: {D['filtro']}.",
          "Se priorizan las combinaciones que cumplen el CV máximo y, si ninguna cumple, las de menor CV."], 12, False, INK)
    rect(s, Inches(0.4), Inches(5.4), Inches(12.55), Inches(1.5), GREEN)
    text(s, Inches(0.6), Inches(5.45), Inches(4), Inches(0.35), "COMBINACIÓN SELECCIONADA", 12, True, MUTED)
    text(s, Inches(0.6), Inches(5.8), Inches(4), Inches(0.5), f"{sel['nombre']} · ofertas {sel['ofertas']}", 16, True, NAVY)
    for i, (k, v) in enumerate([("CV", pct(sel["cv"])), ("¿Cumple?", sel["cumple"]), ("Valor por ha", mill(sel["ha"])), ("Valor por m²", cop(sel["m2"]))]):
        x = Inches(5.0 + i * 2.0)
        text(s, x, Inches(5.5), Inches(1.9), Inches(0.3), k, 11, True, MUTED)
        text(s, x, Inches(5.85), Inches(1.9), Inches(0.8), v, 20, True, RED if k == "¿Cumple?" and str(v).upper() == "NO" else NAVY)

    # 7. Proyección
    s = base(prs, "Proyección del valor esperado", "Rango de orientación entre dos combinaciones de referencia · sin valor puntual", 7, TOTAL)
    kpi(s, Inches(0.5), Inches(1.4), Inches(3.9), Inches(1.3), "REFERENCIA BAJA · $/ha", mill(py["baja"]), GREEN, NAVY, 28)
    text(s, Inches(0.5), Inches(2.72), Inches(3.9), Inches(0.3), str(py["nbaja"]), 10, False, MUTED)
    kpi(s, Inches(4.7), Inches(1.4), Inches(3.9), Inches(1.3), "REFERENCIA ALTA · $/ha", mill(py["alta"]), GREEN, NAVY, 28)
    text(s, Inches(4.7), Inches(2.72), Inches(3.9), Inches(0.3), str(py["nalta"]), 10, False, MUTED)
    kpi(s, Inches(8.9), Inches(1.4), Inches(3.95), Inches(1.3), "AMPLITUD DEL RANGO", mill(py["amp"]), G2, NAVY, 28)
    # barra de rango
    lo, hi = py["baja"] / 1e6, py["alta"] / 1e6
    adop = fnum(st.get("VALOR ADOPTADO")) / 1e6
    vmin, vmax = min(lo, adop) * 0.8, max(hi, adop) * 1.2
    bx, bw, by = Inches(0.9), Inches(11.5), Inches(4.35)
    rect(s, bx, by, bw, Inches(0.5), G2)
    px = lambda v: int(bx + bw * (v - vmin) / (vmax - vmin))
    rect(s, px(lo), by, max(px(hi) - px(lo), Emu(60000)), Inches(0.5), NAVY)
    for v, lab, col, dy in ((lo, f"Baja {num(lo, 1)} M", NAVY, 0.6), (hi, f"Alta {num(hi, 1)} M", NAVY, 0.6)):
        text(s, px(v) - Inches(1.0), by + Inches(dy), Inches(2.0), Inches(0.35), lab, 12, True, col, PP_ALIGN.CENTER)
    rect(s, px(adop) - Emu(20000), by - Inches(0.3), Emu(40000), Inches(1.1), RGBColor(0x92, 0xD0, 0x50))
    text(s, px(adop) - Inches(1.3), by - Inches(0.65), Inches(2.6), Inches(0.35), f"Valor adoptado {num(adop, 1)} M", 12, True, OKGREEN, PP_ALIGN.CENTER)
    text(s, Inches(0.9), Inches(3.2), Inches(8), Inches(0.35), "Posición del valor adoptado frente al rango de orientación ($ millones/ha)", 12, True, NAVY)
    rows = [["Referencia", "Combinación", "Ofertas", "Promedio $/ha", "$/m²", "CV", "¿Cumple?"]]
    for r_ in py["ref"]:
        rows.append([r_[0], r_[1], r_[2], mill(fnum(r_[3])), cop(fnum(r_[4])), pct(fnum(r_[5])), r_[6]])
    tabla(s, Inches(0.9), Inches(5.55), Inches(11.5), rows, [1.6, 1.6, 1.3, 1.6, 1.2, 0.9, 1.0], row_h=0.34,
          align=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER])
    text(s, Inches(0.9), Inches(6.55), Inches(11.5), Inches(0.5), recortar(py["nota"], 260), 10, False, MUTED, italic=True)

    # 8. Depreciación
    s = base(prs, "Depreciación de construcciones", "Método Ross-Heidecke · valor actual por estado de conservación", 8, TOTAL)
    rows = [["Ítem", "Edad", "Vida útil", "Estado", "Reposición $/m²", "FD", "Valor adoptado $/m²", "Área m²", "Valor total"]]
    for c_ in D["constr"]:
        rows.append([c_["item"], c_["edad"], c_["vida"], c_["estado"], cop(c_["repos"]), num(c_["fd"], 3), cop(c_["adop"]), num(c_["area"], 0), cop(c_["total"])])
    rows.append(["TOTAL", "", "", "", "", "", "", num(sum(c_["area"] for c_ in D["constr"]), 0), cop(D["constr_total"])])
    tabla(s, Inches(0.4), Inches(1.3), Inches(12.55), rows, [2.0, 0.7, 0.9, 0.8, 1.5, 0.8, 1.7, 0.9, 1.6], size=11,
          row_h=min(0.4, 3.2 / len(rows)), hl_last=True,
          align=[PP_ALIGN.LEFT] + [PP_ALIGN.CENTER] * 3 + [PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
    cd = CategoryChartData()
    cd.categories = [c_["item"] for c_ in D["constr"]]
    cd.add_series("Reposición $/m²", [c_["repos"] / 1e3 for c_ in D["constr"]])
    cd.add_series("Valor depreciado $/m²", [c_["adop"] / 1e3 for c_ in D["constr"]])
    top = Inches(1.3 + 0.5 * min(len(rows), 9) + 0.15) if len(rows) < 9 else Inches(5.3)
    top = Inches(1.3) + Inches(min(0.4, 3.2 / len(rows)) * len(rows)) + Inches(0.15)
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.4), top, Inches(12.55), H - top - Inches(0.5), cd)
    estilo_chart(gf.chart)
    gf.chart.plots[0].gap_width = 80
    for k_, col in enumerate((G1, NAVY)):
        gf.chart.plots[0].series[k_].format.fill.solid()
        gf.chart.plots[0].series[k_].format.fill.fore_color.rgb = col
    gf.chart.value_axis.major_gridlines.format.line.color.rgb = G2
    gf.chart.has_title = True
    gf.chart.chart_title.text_frame.text = "Reposición vs. valor depreciado ($ miles por m²)"
    gf.chart.chart_title.text_frame.paragraphs[0].runs[0].font.size = Pt(12)

    # 9. Liquidación
    s = base(prs, "Liquidación del avalúo", "Terreno + construcciones · totales redondeados a la centena", 9, TOTAL)
    rows = [["Concepto", "Unidad", "Cantidad", "Valor unitario", "Valor total"]]
    if lq["terreno"]:
        t_ = lq["terreno"]
        rows.append(["TERRENO – " + recortar(t_[0], 40), t_[1], num(t_[2], 2), cop(t_[3]), cop(t_[4])])
    for t_ in lq["principal"]:
        rows.append([t_[0].title(), t_[1], num(t_[2], 0), cop(t_[3]), cop(t_[4])])
    for t_ in lq["anexas"]:
        rows.append([t_[0].title(), t_[1], num(t_[2], 0), cop(t_[3]), cop(t_[4])])
    rows.append(["VALOR TOTAL DEL AVALÚO", "", "", "", cop(lq["total"])])
    tabla(s, Inches(0.4), Inches(1.3), Inches(8.2), rows, [3.4, 0.8, 1.0, 1.6, 1.8], size=11,
          row_h=min(0.46, 5.3 / len(rows)), hl_last=True,
          align=[PP_ALIGN.LEFT, PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT])
    terr = lq["terreno"][4] if lq["terreno"] else 0
    cd = CategoryChartData()
    cd.categories = ["Terreno", "Constr. principal", "Constr. anexas"]
    cd.add_series("Valor", [terr / 1e6, lq["tot_p"] / 1e6, lq["tot_a"] / 1e6])
    gf = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(8.8), Inches(1.3), Inches(4.2), Inches(3.4), cd)
    ch = gf.chart
    estilo_chart(ch)
    ch.plots[0].has_data_labels = True
    dl = ch.plots[0].data_labels
    dl.number_format, dl.number_format_is_linked, dl.show_percentage, dl.show_value = '0.0%', False, True, False
    dl.font.size, dl.font.color.rgb, dl.font.bold = Pt(11), WHITE, True
    for i, col in enumerate((NAVY, BLUE, G1)):
        pt = ch.plots[0].series[0].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = col
    kpi(s, Inches(8.8), Inches(4.85), Inches(4.15), Inches(1.0), "VALOR TOTAL", cop(lq["total"]), GREEN, NAVY, 20)
    rect(s, Inches(0.4), Inches(6.05), Inches(12.55), Inches(0.9), G2)
    text(s, Inches(0.55), Inches(6.08), Inches(12.3), Inches(0.85),
         ["Valor en letras", lq["letras"]], 13, True, NAVY, anchor=MSO_ANCHOR.MIDDLE)

    # 10. Observaciones y firma
    s = base(prs, "Observaciones y conclusiones", "Salvedades del estudio y responsable técnico", 10, TOTAL)
    obs = [x.strip() for x in D["obs"].replace("OBSERVACIONES:", "").split("\n") if x.strip()]
    y = Inches(1.35)
    for o_ in obs[:5]:
        o_ = o_.lstrip("0123456789) ").strip()
        h_ = Inches(0.4 + 0.22 * (len(o_) // 105))
        rect(s, Inches(0.5), y + Inches(0.08), Inches(0.12), Inches(0.12), BLUE, shape=MSO_SHAPE.OVAL)
        text(s, Inches(0.75), y, Inches(8.1), h_, recortar(o_, 520), 12, False, INK)
        y += h_ + Inches(0.1)
    rect(s, Inches(9.2), Inches(1.35), Inches(3.7), Inches(5.5), NAVY)
    rect(s, Inches(9.5), Inches(4.7), Inches(3.1), Inches(0.03), G1)
    text(s, Inches(9.4), Inches(1.6), Inches(3.3), Inches(0.4), "Avaluador", 12, False, G1, PP_ALIGN.CENTER)
    text(s, Inches(9.4), Inches(2.0), Inches(3.3), Inches(1.2), AVALUADOR, 22, True, WHITE, PP_ALIGN.CENTER)
    text(s, Inches(9.4), Inches(3.3), Inches(3.3), Inches(0.5), RAA, 14, True, GREEN, PP_ALIGN.CENTER)
    text(s, Inches(9.4), Inches(4.8), Inches(3.3), Inches(0.4), "Firma", 11, False, G1, PP_ALIGN.CENTER)
    text(s, Inches(9.4), Inches(5.9), Inches(3.3), Inches(0.8), ["Fecha del estudio", D["fecha"]], 12, True, WHITE, PP_ALIGN.CENTER)

    prs.save(salida)
    return salida


if __name__ == "__main__":
    xlsx = Path(sys.argv[1] if len(sys.argv) > 1 else "Estudio_Mercado_Rural_Puente_Bomba_Tigreras.xlsx")
    out = Path(sys.argv[2] if len(sys.argv) > 2 else "Presentacion_Estudio_Mercado_Rural.pptx")
    data = leer(xlsx)
    if not data["ofertas"] or data["liq"]["total"] == 0:
        sys.exit("El libro no tiene valores calculados: ábralo en Excel, guárdelo y vuelva a ejecutar.")
    print("Generada:", construir(data, out))
