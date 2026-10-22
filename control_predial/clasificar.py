"""Clasificación nominal de archivos por su nombre (misma lógica que la hoja «Documentos Drive»)."""
import re

IMG = ('.jpg', '.jpeg', '.png', '.heic', '.webp')


def clasificar(nombre, ruta=''):
    n = nombre.lower()
    r = (ruta or '').lower()
    ext = n.rsplit('.', 1)[-1] if '.' in n else ''
    if n.startswith('forma de llegada'):
        return 'KMZ' if ext == 'kmz' else 'Acceso / llegada'
    if 'acta' in n and 'reuni' in n:
        return 'Acta de reunión'
    if n.startswith('acta') and 'aval' in n:
        return 'Acta de avalúo'
    if 'cornare' in n:
        return 'Ambiental'
    if 'sentencia' in n:
        return 'Sentencia'
    if ext == 'kmz':
        return 'Soporte expropiatorio' if 'expropiacion' in n else 'KMZ'
    if n.startswith('aval') and ext == 'pdf':
        return 'Avalúo'
    if ('resoluci' in n or 'oficio' in n) and ('riesgo' in n or 'autoriza' in n or 'expropia' in n):
        return 'Soporte expropiatorio'
    if 'certificado' in n and ext == 'pdf':
        return 'CTL'
    if 'escritura' in n and 'servidumbre' in n:
        return 'Escritura de servidumbre'
    if n.startswith('escritura'):
        return 'Escritura'
    if ext in ('xlsx', 'xls') and (n[:1].isdigit() or 'anexo predial' in n):
        return 'Formato / control'
    if 'ficha catastral' in n:
        return 'Ficha catastral'
    if 'ficha predial' in n:
        return 'Ficha predial'
    if n.startswith('plano'):
        return 'Plano'
    if 'acuerdo no' in n:
        return 'Norma urbanística (PBOT)'
    if ext == 'docx' and ('formato' in n or 'informe' in n):
        return 'Formato / informe'
    if ext in ('jpg', 'jpeg', 'png', 'heic', 'webp'):
        return 'Fotografía / imagen'
    return 'Otro soporte'
