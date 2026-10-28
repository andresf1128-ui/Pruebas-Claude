---
name: avaluo-rural-fase1-campo
description: Fase 1 del avalúo rural. Prepara y procesa la captura de información en campo (ficha, trazas GPS, fotos rotuladas por predio y por construcción, formato de descripción de construcciones con la tipología del IGAC). Úsala antes y después de la visita técnica.
---

# Fase 1 · Captura de información en campo

Objetivo: llegar a la visita con todo preparado y salir con datos completos, georreferenciados y fotos uniformes y rotuladas, sin depender de la memoria.

## 1. Antes de salir (preparación)
1. Abrir el expediente del encargo y listar los predios con matrícula, ficha/cédula catastral y vereda (los toma de los documentos del cliente).
2. Preparar la cartografía offline: polígono catastral del predio, vías de acceso, ortofoto o imagen satelital reciente (apoyarse en `copernicus-sentinel-avaluos`) y las capas normativas que se conocen. Exportar a KMZ/GeoJSON con `kmz-sig-profesional`.
3. Cargar el formulario de campo (ver `referencias/ficha-campo.md`) en la herramienta elegida (sección 2).
4. Cargar la lista de verificación de campo: linderos, acceso, servicios, construcciones, cultivos, fuentes de agua, conflictos o afectaciones visibles, vecinos, ofertas de la zona.
5. Confirmar la **visita técnica**: la 941 la exige (art. 10) y debe hacerla el avaluador RAA que firma el informe (art. 13.4). Si no es posible ingresar al inmueble, se informa por escrito al solicitante; solo él puede ordenar, bajo su responsabilidad y por escrito, avaluar con la información suministrada, y esa orden se archiva y se declara en el informe. Sin esa orden, o con información insuficiente, no se continúa. Pedir al solicitante nombre, teléfonos y correo de quien acompañará la visita (art. 8, parágrafo 1).
6. Plazo: 30 días hábiles desde el recibo de la solicitud con toda la información (art. 9), salvo otro término pactado en el contrato.

## 2. Herramientas recomendadas
| Necesidad | Herramienta | Nota |
|---|---|---|
| Mapa georreferenciado offline, ubicación en tiempo real, marcas y trazas | **Avenza Maps** (mapas GeoPDF/GeoTIFF), **QField** (con proyecto QGIS) o **ArcGIS Field Maps** | Avenza es ágil para navegar y marcar puntos; QField/Field Maps permiten formularios con capas. Verifique en su versión qué formatos exporta cada una. |
| Formulario estructurado con validaciones, fotos y GPS | **QField**, **KoboToolbox/ODK Collect** (gratis, offline), **Survey123** o **Jotform** | Los formularios con campos obligatorios evitan olvidos. Recomendado: un solo formulario por predio y un formulario hijo por construcción. |
| Linderos y áreas de campo | GNSS de mano o móvil con corrección (si se tiene), cinta/telémetro para construcciones | Declarar el método y la precisión en el informe. |
| Fotos con fecha, hora y coordenada | Cámara del móvil con GPS activado (o la cámara del propio formulario) | Que la foto se tome desde el formulario asegura el vínculo predio/construcción. |
| Dron (si aplica y está permitido) | Ortomosaico y vistas oblicuas | Respetar la normativa aeronáutica y los permisos del predio. |
| Notas de voz | Grabadora del móvil | Se transcriben en la ficha; el audio no se entrega. |

Si el perito no ha definido herramienta, proponer: **QField o Kobo para el formulario + Avenza Maps para navegación**. Es una sugerencia que se ajusta a su práctica.

## 3. Reglas de fotos (uniformes, rápidas, rotuladas)
La 941 pide un registro fotográfico y/o fílmico **amplio y suficiente** de todo lo que entra en la liquidación: terreno, construcciones, anexos y cultivos (art. 13.5).

**Durante la visita** (consigna única para quien fotografía): tomar siempre en horizontal, con GPS activado, y seguir este guion por unidad. Una unidad es el predio o cada construcción.
- **Predio (P):** acceso/entrada, panorámicas de linderos (N, S, E, O), vías internas, topografía, cobertura y uso (pastos, cultivos, bosque), fuentes de agua, cercas, mojones, servicios, y detalles relevantes (erosión, invasión, afectaciones).
- **Construcción (C1, C2, …):** fachada principal, otras fachadas, cubierta, estructura visible, muros, pisos, cielos, carpintería, cocina, baños, instalaciones, patologías y estado general, anexos.
- Una foto con el **rótulo físico** (papel con código del avalúo, predio y unidad) al iniciar cada unidad facilita la clasificación de respaldo.

**Descarga y organización (rápida y uniforme):**
1. Copiar el contenido completo del móvil a una carpeta `…/Originales/` del expediente (cable, o carpeta de la nube del móvil). Los originales no se tocan.
2. Si las fotos se tomaron desde el formulario, ya vienen vinculadas a predio y construcción. Si no, se agrupan en subcarpetas por unidad (`P`, `C1`, `C2`) o se separan por hora.
3. Ejecutar:
   ```
   python3 scripts/organizar_fotos.py --origen "…/Originales" --salida "…/Organizado" --codigo AV-045 --predio "La Esperanza"
   ```
   El script copia (no mueve), ordena por fecha/hora EXIF, rotula así `CODIGO_Predio_Unidad_Categoria_NNN.jpg` (ej. `AV-045_LaEsperanza_C1_Fachada_003.jpg`), normaliza la orientación y genera `manifiesto_fotos.csv` con fecha, hora, coordenadas y la ruta original.
4. Revisar las fotos sin GPS o sin fecha (las lista el manifiesto) y la categoría asignada. Para la clasificación visual por categoría, usar `organizar-fotos-avaluos`.
5. Las fotos del informe se seleccionan del conjunto organizado; no se editan sus colores ni se alteran los contenidos. El recorte se limita al encuadre.

## 4. Predios con construcciones
Por cada construcción diligenciar el **formato de descripción de construcciones** del perito y clasificarla con el **catálogo de tipologías constructivas del IGAC**:
1. Colocar la plantilla vigente del formato en `plantillas/descripcion_construcciones.xlsx` (o .docx). Sin plantilla el agente no inventa el formato: se la pide al perito.
2. Colocar el catálogo de tipologías del IGAC (la versión que use el perito) en `referencias/catalogo_tipologias_igac.*`. La tipología se asigna **por comparación con las definiciones y fotos del catálogo**, citando su código; si dos tipologías compiten, se anotan las dos con el criterio de decisión del perito.
3. Datos mínimos por construcción (art. 11.2 de la 941): código (C1…), **tipología constructiva**, uso, características físicas predominantes, **condición física y jurídica**, área construida (con método de medición), número de pisos, **edad** (fuente: licencia, escritura, testimonio, dato del propietario), elementos y materiales de estructura y acabados, obras complementarias, **estado de conservación** (escala de Heidecke, de 1 a 5 con medios puntos), **vida útil** de referencia o prolongada con su sustento, reformas o intervenciones que afecten la vida útil, anexos y fotos asociadas (nombres del manifiesto). Las construcciones sin licencia se registran con la constancia correspondiente; las áreas a valorar son las que indique el solicitante (art. 12.6).
4. Estimar sobre el terreno el estado de conservación y consignar las patologías con foto.
5. Verificar que la suma de áreas construidas cuadre con lo que dicen las escrituras, el catastro y la licencia; las diferencias pasan a la fase 3 y al informe.

## 5. Cierre de la fase
- Ficha de campo completa (sin campos obligatorios vacíos).
- Trazas y puntos exportados a KMZ/GeoJSON; áreas recalculadas en EPSG:9377 si se midieron en campo.
- Fotos organizadas + `manifiesto_fotos.csv`.
- Formato de construcciones diligenciado por cada construcción.
- Lista de pendientes de campo (datos que faltaron, personas por contactar).
- Pedir la aprobación del perito para pasar a la fase 2.
