---
name: avaluo-rural-fase6-calidad
description: Fase 6 y control transversal del avalúo rural. Inspección de calidad de fondo (norma, jurídico, áreas, mercado, liquidación) y de forma (tipo de letra, ortografía, orden, estilo uniforme, cifras, imágenes, metadatos). Emite dictamen APROBADO o lista de hallazgos. Se corre al cerrar cada fase y al final, antes de entregar.
---

# Fase 6 · Control de calidad

Complementa `control-calidad-documentos` (regla de oro: ningún documento se guarda, entrega ni se cita sin dictamen APROBADO; si no se completa el control, el archivo queda como `_BORRADOR` y se informa qué falta). Esta skill agrega la lista específica del avalúo rural. Es un revisor independiente: revisa lo que otras fases produjeron, no lo que cree haber hecho.

## Procedimiento
1. Fuente maestra: documentos del cliente y fuentes oficiales mandan sobre archivos de control internos.
2. Correr las listas A a H. Cada punto tiene resultado: CUMPLE / NO CUMPLE / NO APLICA, con evidencia (archivo, celda, página).
3. Dictamen:
   - **APROBADO**: sin NO CUMPLE.
   - **APROBADO CON OBSERVACIONES**: solo observaciones de forma sin riesgo, ya corregibles.
   - **NO APROBADO**: con la lista de hallazgos ordenada por gravedad (crítico / mayor / menor), el responsable de la corrección y qué fase se repite.
4. Tras corregir, repetir el control completo del documento (no solo el punto).

## A. Norma (Resolución 941 de 2026 y análisis normativo)
- La tabla de requisitos de la 941 está completa y cada obligación tiene dónde se cumple.
- Cada cita normativa tiene artículo/numeral, enlace oficial, fecha de consulta y está vigente (no derogada ni modificada).
- Las áreas y porcentajes de cruces coinciden con el cálculo en EPSG:9377.
- El uso del suelo y las restricciones citadas corresponden al municipio y a la subzona del predio.
- Los parámetros de `config/parametros-941.yaml` están todos en `confirmado` antes de entregar el informe definitivo. Si alguno está por confirmar: NO APROBADO.
- Los 13 numerales del art. 14 están presentes y completos (ver `referencias/RESUMEN_941.md`, sección 3).
- Visita técnica hecha y firmada por el avaluador RAA, o la orden escrita del solicitante archivada (arts. 10 y 13.4).
- Certificado RAA vigente con las categorías requeridas anexo; centroide con coordenadas tomadas en sitio y sistema de referencia indicado.
- Ningún modelo de IA determina el valor. Si se usó algún modelo automatizado, existe el manual auditable con métricas de error y la validación del avaluador (art. 20, parágrafo 2).

## B. Jurídico
- Todo documento aportado está listado y revisado; fecha de impresión del CTL literal.
- Cada gravamen y medida con su estado (vigente/cancelado) confirmado en el original.
- Tabla de conciliación de áreas y linderos; diferencias explicadas.
- Titular, matrícula, cédula catastral y nombre del predio coinciden en todo el informe.

## C. Áreas y geometría
- Una sola área oficial por predio, EPSG:9377, con el mismo valor en texto, cuadros, mapas y liquidación.
- Suma de áreas de zonas, usos y cruces = área del predio (tolerancia declarada).
- Unidades correctas (ha vs. m²), conversiones de la escritura verificadas.
- Mapas con fuente, escala o referencia, leyenda legible.

## D. Mercado
- Cada oferta con soporte (enlace, captura, contacto) y fecha.
- El porcentaje de negociación **está justificado por escrito** (art. 17 b) y el CV máximo (10 %, art. 21) viene de la celda de parámetros, no escrito en fórmulas.
- La depuración se hizo sobre el valor negociado, se descontaron construcciones y cultivos permanentes de las ofertas y **no se aplicaron factores de homogeneización** (anexo 2.1, Nota 2).
- La tabla del estudio tiene las columnas del anexo 2.1 y cada oferta tiene URL, captura de pantalla, fecha y teléfono.
- Recalcular los estadísticos con `scripts/estadisticos_mercado.py`: debe coincidir con el libro.
- Exclusiones documentadas. Factores de homogeneización justificados.
- El valor adoptado está dentro del rango y se sustenta.

## E. Liquidación
- Suma de partes = total. Redondeos solo al final.
- Depreciación con el modelo continuo Ross-Heidecke (art. 30), sin modelos en escalera; edad, vida útil (Tabla 2 del anexo o prolongada sustentada) y estado coherentes con el formato de construcciones y las fotos.
- Suelo de protección o ronda hídrica: sin castigos por la sola condición (art. 49).
- Memoria de cálculo con todas las ecuaciones, variables, datos y resultados (art. 14.10 c).
- Tipología IGAC coherente con las fotos y la descripción.
- Valor por hectárea y proporciones razonables frente al mercado y al catastro.
- Cultivos/plantaciones sin doble conteo con el terreno.

## F. Campo y soportes
- Todas las construcciones tienen formato de descripción diligenciado y fotos.
- Fotos rotuladas con el patrón `CODIGO_Predio_Unidad_NNN`, manifiesto completo, coordenadas coherentes con el predio.
- Registro fotográfico del informe corresponde al predio correcto (nada de fotos de otro predio).

## G. Forma
Revisión mecánica, con apoyo de scripts cuando sea posible:
- Tipo de letra, tamaño, interlineado, márgenes y estilos de título **idénticos** en todo el documento y a `config/estilo.yaml` (en .docx revisar `styles.xml` y runs con formato directo).
- Numeración de secciones y de cuadros/figuras continua; índice actualizado; referencias cruzadas válidas.
- Ortografía y gramática en español de Colombia; nombres propios, topónimos, siglas y normas escritos igual en todo el texto; tildes en mayúsculas.
- Cifras: separador de miles y decimal es-CO; moneda y unidades uniformes; mismas cifras en texto y cuadros.
- Fechas en el mismo formato; coherencia entre fecha de visita, de informe y de vigencia.
- Imágenes: 14 cm de ancho, borde biselado, fuente bajo cada una, resolución adecuada, sin deformaciones.
- Tablas sin celdas cortadas, encabezados repetidos, alineación coherente.
- Sin marcas de revisión, comentarios, cambios controlados, resaltados de trabajo, ni textos «pendiente», «por verificar», «TODO», «borrador».
- Encabezados y pies de página correctos en todas las secciones; paginación continua.

## H. Entrega limpia
- Correr `python3 scripts/limpiar_metadatos.py documento --solo-verificar` en cada .docx, .xlsx, .pptx y .pdf del paquete. Debe terminar sin hallazgos.
- Nombres de archivo según el estándar del encargo.
- Firma, registro RAA y vigencia vigentes.
- Paquete final completo: informe, anexos, libro de mercado, KMZ, fotos y sustentación.

## Informe de calidad (salida)
Cuadro con: fase, ítem, resultado, evidencia, gravedad, responsable, estado. Al final: dictamen y fecha. Si es NO APROBADO, el paquete se queda en `_BORRADOR` y no sale al cliente.
