# Verificación de cifras · HIDROARMA Control Predial 360 v2 (corte 21-oct-2026)

Alcance: consistencia interna del libro (recalculado desde las tablas crudas con Python).
**No verificado contra SharePoint ni NotebookLM**: la fuente madre exige autenticación y el entorno no tiene salida a sharepoint.com (HTTP 403 del proxy).

## Cifras que cuadran
| Indicador | Libro | Recalculado |
|---|---|---|
| Folios en matriz | 93 | 93 (sin duplicados) |
| Sonsón / Aguadas | 50 / 43 | 50 / 43; la regla 028→Sonsón y 102→Aguadas se cumple en los 93 |
| Compraventa / Expropiación / Servidumbre | 35 / 16 / 42 | 35 / 16 / 42 |
| Carpeta propia / asociados a folio matriz | 86 / 7 | 86 / 7 |
| Archivos inventariados | 1.719 | 1.719 (1.694 ZIP + 25 SharePoint) |
| Estados del cotejo | 35 / 10 / 1 / 41 / 6 | 35 / 10 / 1 / 41 / 6 (suman 93) |
| Fuentes NotebookLM | 579 | 579 |
| Archivos por carpeta (SharePoint vs inventario) | Δ = 0 | 0 diferencias en los 93 folios |
| Banderas CTL, escritura, ficha catastral, ficha predial, avalúo/acta, soporte judicial | | 0 diferencias frente al inventario de archivos |
| Bandera "Plano nombrado" | 66 Sí | coincide (detecta "plano" en el nombre del archivo, incluidos los de «Ficha predial y plano.pdf») |
| Área literal vs interpretada | | 0 diferencias |

## Cifras que NO cuadran o inducen a error
1. **Alertas de prioridad alta: 84 de 85 (Inicio!J5).** El indicador no discrimina: 63 son "sin geometría KMZ" y 7 "sin carpeta propia", todas en prioridad Alta. Debe reclasificarse (el tablero 360 sí muestra 3 altas reales).
2. **Fotos en la matriz = 480; fotos únicas reales = 472 (471 por nombre).** La carpeta 20264371 la comparten 028-27277 y 028-30143 y sus 8 fotos se cuentan dos veces. Además 102-9927 cuenta 4 y tiene 3 únicas (foto duplicada, H-17).
3. **Archivos en carpetas de radicado: 910 sumando folios, 901 reales** (la misma carpeta compartida cuenta 9 dos veces).
4. **Fotos en cuaderno: 462 en «Cotejo» vs 454 en «Cuadernos».** Diferencia de 8 = cuaderno compartido «PREDIO 028-30143 (compartido)» contado en dos folios.
5. **«45 folios con cuaderno» vs 44 cuadernos de folio.** Los 45 folios incluyen 028-27277 que usa el cuaderno de 028-30143. El 45 coincide con "45 cuadernos" por casualidad (el cuaderno 45 es la plantilla del informe de visita).
6. **Carpetas de radicado: 86 = 85 carpetas distintas de folio + 1 carpeta adicional (20264416).** Coincide con el 86 de "folios con carpeta propia" solo por azar; en realidad 86 folios comparten 85 carpetas.
7. **Radicado 20264371 repetido** en dos folios (H alerta "matrícula repetida", sin cerrar).
8. **Hallazgo H-22 "Fotos por folio coincide con SharePoint en todos los folios" es inexacto**: hay 3 excepciones (028-30110 y 102-7693 tienen sus fotos fuera de su carpeta; 102-9927 por duplicado).
9. **H-24: título "33 cuadernos", texto "34 de los 44", tablero "35 conciliados"** — tres cifras distintas para el mismo concepto. La que se deriva de los datos es 35 folios (34 cuadernos distintos de folio, por el compartido).
10. **H-01 (prioridad Alta) figura como «Conforme»** pese a que 102-13526 y 102-13564 tienen la misma evidencia duplicada y están entre los 10 "Fotos por conciliar". Inconsistencia de estado.
11. **«Fotos en carpeta propia» > 0 en 45 folios**, pero 028-30110 no tiene carpeta propia y 102-7693 guarda sus fotos fuera de su carpeta: el rótulo de la columna no es exacto.
12. **Datos de valor/área**: 21 folios sin valor numérico (13 expropiación, 7 servidumbre, 1 compraventa) y 13 sin área interpretada; cualquier total de valor o hectáreas debe declarar esa exclusión.

## Pendiente de validar contra la fuente madre
- Que los 1.719 archivos / 143 carpetas sigan iguales hoy en SharePoint.
- Contenido del PDF «Acuerdo No. 80-17-14» (H-03) y los 25 archivos nuevos.
- Conteo en vivo de cuadernos NotebookLM (45 / 579 fuentes).
- Datos del Anexo Predial Final (valores, titulares, áreas): no se recibió el archivo origen.

---

# Adenda · verificación contra la captura en vivo de SharePoint y NotebookLM (ZIP «Cotejo_evidencia_2026-10-21»)

Fuente: `sharepoint_inventario_en_vivo_2026-10-21.json` y `notebooklm_cuadernos_en_vivo_2026-10-21.json`, recalculados de forma independiente (no se usaron los totales del Excel).
Limitación: es la captura del 21-oct, no una lectura de hoy; el ZIP lo generó el mismo proceso que construyó el libro, así que confirma que el libro refleja la captura, no que la captura siga vigente.

## Confirmado
| Dato | Captura cruda | Excel |
|---|---|---|
| Archivos | 1.719 (sin duplicados de ruta, 0 errores de lectura) | 1.719 |
| Carpetas | 143 | 143 |
| Archivos por folio (93 folios) | | 0 diferencias |
| Imágenes (.jpg/.jpeg/.png) | 472 | 472 |
| Cuadernos NotebookLM | 45 | 45 |
| Fuentes | 579, las 579 con estado procesado | 579 |
| Fotos en cuadernos / audio / fuentes externas | 454 / 71 / 9 | 454 / 71 / 9 |
| Fuentes por cuaderno | | 0 diferencias (45 de 45) |
| Fotos solo en SharePoint / solo en cuaderno | 24 / 7 | 24 / 7 |
| Carpeta vacía (H-19) | «REGISTRO FOTOGRÁFICO» (0 archivos) | consistente |

## Fotos repetidas entre carpetas: evidencia compartida (aclarado por el usuario)
Son tomas a distancia de predios casi colindantes, por eso la misma foto está en la carpeta de cada folio al que sirve. **No es un error ni se debe depurar.**
Hay 11 grupos de imágenes idénticas (nombre y tamaño), 27 copias de más:
- 1792278815580/700/787/860.jpg: 028-3148 (Sonsón) y 102-13491, 102-13570, 102-13573, 102-13587 (Aguadas).
- 1792087615367/452.jpg: 028-8567, 028-8570, 028-24662, 028-24663. 1792087612750.jpg: 028-24662 y 028-24663.
- TimePhoto_20261014_143623/625/646.jpg: 028-26765 y 028-26767.
- «FORMA DE LLEGADA.jpg» dos veces dentro de 102-9927 (esta sí es copia accidental, H-17).

Consecuencia solo para el conteo: las fotos distintas son ~445 y las asignaciones foto-folio son 472. Los tableros deben mostrar ambas cifras ("fotos distintas" y "fotos asignadas a folios") y marcar cada foto como compartida, no como duplicada. El modelo de datos de la app trata la relación foto↔folio como muchos a muchos.
