---
name: avaluo-rural-fase4-mercado
description: Fase 4 del avalúo rural. Estudio de mercado conforme a los arts. 16 a 21 y al anexo técnico 2.1 de la Resolución 941 de 2026 - ofertas comparables, porcentaje de negociación justificado (criterio del perito 10 %), depuración sobre el valor negociado, sin homogeneización por factores, CV máximo rural de 10 %. Úsala después de la fase 3.
---

# Fase 4 · Estudio de mercado rural

Base normativa: `referencias/RESUMEN_941.md`, sección 4. Parámetros en `config/parametros-941.yaml`.

## Reglas de la 941 que gobiernan esta fase
1. **Depuración sobre el valor negociado** (art. 17 c): primero se aplica el porcentaje de negociación a cada oferta y sobre ese valor se depura.
2. **El porcentaje de negociación debe estar justificado** (art. 17 b). El 10 % es criterio del perito, no una cifra de la norma: cada informe lleva el párrafo que lo sustenta (naturaleza del mercado rural de la zona, evidencia de transacciones, diferencia entre pedido y cierre). Sin esa justificación el parámetro queda por confirmar y el informe no se marca definitivo.
3. **No se homogeneiza por factores** (anexo técnico 2.1, Nota 2). El tratamiento es depuración, clasificación, comparación y análisis. Tampoco se «iguala» ofertas con el predio (art. 17, parágrafo). Lo único que admite ajuste con fórmula es la forma de un lote irregular (art. 17 i, fórmulas del anexo).
4. **CV máximo rural: 10,0 %** (art. 21), leído en conjunto con tipología, mercado, calidad y uniformidad de los datos. Si no se adopta la media, se sustenta con media, mediana, atípicos y comparabilidad (art. 21, parágrafo 1). Si el mercado es heterogéneo o hay pocos datos, se justifica el criterio y se apoya en otros métodos (parágrafo 2).
5. **Valores de construcciones de las ofertas** con los lineamientos del método del costo (art. 17 d) y **descuento de cultivos permanentes e infraestructura productiva** (art. 17 j); si faltan cantidades, áreas estimadas a partir de lo que informe el vendedor.
6. **Evidencia de cada dato** (art. 17 f, g, h; art. 14.13): URL, captura de pantalla con la información visible, teléfono, fotos, fecha de consulta. Los datos se corroboran.
7. Todo cálculo con memoria de cálculo completa: ecuaciones, variables, datos y resultados (art. 14.10 c).

## 1. Universo comparable
1. Caracterizar el predio objeto con lo de las fases 1 a 3: municipio y vereda, área, uso actual y permitido, AHT y clases agrológicas, pendiente, agua, acceso y distancia a centros poblados, infraestructura productiva, norma urbanística y ambiental.
2. Criterios de comparabilidad del art. 19 (rural): áreas, vocación del suelo, calidad y fertilidad (AHT, valor potencial, clases agrológicas, aptitud), pendiente, agua, accesibilidad, distancia a centros poblados, infraestructura productiva, norma urbanística y ambiental. Se escriben antes de buscar.
3. Radio de búsqueda y ampliación escalonada, declarando por qué.

## 2. Búsqueda de ofertas
Fuentes con registro individual (enlace o contacto, fecha de captura y captura de pantalla):
- Portales y avisos (Fincaraíz, Metrocuadrado, Ciencuadras, MercadoLibre, prensa regional).
- Lonjas, inmobiliarias, corredores, notarios y habitantes de la vereda.
- Ofertas halladas en la visita (ficha de campo, sección F). La 941 pide recolectar información económica en campo (art. 13.7).
- Transacciones documentables; pesan más que las ofertas.
- Observatorios inmobiliarios catastrales y páginas de entidades (fuentes secundarias permitidas, art. 13.7).
Herramientas de apoyo: `brightdata-plugin:search` y `scrape`, o navegador automatizado; todo dato recolectado se confirma con el anuncio original.

## 3. Ficha y tabla del estudio (anexo 2.1, rural NPH)
Columnas obligatorias, en este orden:
número de oferta · municipio y vereda · fecha de captura · tipo de oferta (venta, transacción, arriendo) · valor de oferta · % de negociación · valor negociado · área de terreno (ha) · áreas construidas y anexos (m²) con valor unitario y total · área de cultivo con valor por ha y total · **valor terreno $/ha** · infraestructura (vías internas, canales, pozos, reservorios) · norma urbanística y determinantes ambientales · clases de suelo y AHT con valor potencial · registro fotográfico · observaciones o descripción · coordenadas aproximadas · teléfono de contacto · enlace URL · captura de pantalla.
Lo que la fuente no permita conocer se consigna con las salvedades correspondientes (anexo 2.1), sin inventar.

## 4. Depuración y estadística
1. Valor negociado = valor de oferta × (1 − % de negociación).
2. Valor del terreno = valor negociado − construcciones y anexos (método del costo) − cultivos permanentes e infraestructura.
3. $/ha = valor del terreno ÷ área en hectáreas.
4. Exclusión de datos inconsistentes o atípicos, con la razón de cada una en ¿INCLUIR?. No se excluye para «hacer cumplir» el CV.
5. Estadísticos (art. 20, opcionales y complementarios): n, promedio, mediana, desviación estándar, CV, límites, asimetría; se pueden agregar IQR, MAD o diagramas de caja.
6. Libro de mercado rural del perito (`plantillas/ESTUDIO_MERCADO_RURAL.xlsx`) con fórmulas vivas, siguiendo el patrón de `mercado-nph-estadisticos-depreciacion`: unidad **$/ha**; celda de CV máximo 0,10; celda de % de negociación referenciada (no escrita en cada fórmula); sin columna de factores de homogeneización.
7. Verificación independiente: `python3 scripts/estadisticos_mercado.py ofertas.csv`, que debe coincidir con el libro.

Si el CV supera el 10 %: revisar depuración, buscar más datos, o justificar el criterio y apoyarlo en otro método (art. 21, parágrafo 2). Nunca se ajustan valores para cuadrar.

## 5. Valor adoptado y zonificación
- Valor unitario de terreno: la media aritmética salvo que se sustente otro criterio (art. 21). Redondeo al final según el criterio del perito.
- Si el predio tiene zonas con diferencias físicas o normativas, se divide en franjas o zonas con soporte técnico (arts. 13.8 y 46), con áreas en EPSG:9377.
- **Suelo de protección y ronda hídrica** (art. 49): no se aplican castigos ni porcentajes de reducción por estar en suelo de protección o por restricciones de uso; se analiza y sustenta la incidencia de condiciones favorables o desfavorables. La ronda hídrica conserva el valor del terreno rural agropecuario contiguo con características similares. Cuando no hay actividades rentables permitidas, el valor sale del mercado de inmuebles rurales agropecuarios aledaños de menor productividad.
- Métodos complementarios para validar: costo (construcciones y anexos) y, cuando la actividad lo justifique, renta o capitalización de ingresos (arts. 46 a 48). Para renta agropecuaria faltan las páginas del anexo técnico con los modelos de costos y ejemplos (tablas 9 a 17): no se aplica hasta tenerlas.

## 6. Qué suele pasarse por alto
- Fecha de las ofertas y tendencia del mercado; índices inmobiliarios solo con sustento (art. 13.7).
- Ventas con condiciones atípicas (apuro, parentesco, permuta).
- Diferencia entre el área ofertada y la registral o catastral de la oferta.
- Aguas, concesiones, vías y energía de cada oferta.
- Cultivos permanentes y plantaciones, siempre descontados.
- Servidumbres y determinantes ambientales de cada oferta.
- Proyectos de infraestructura cercanos.
- Mercados ilíquidos: documentar la búsqueda realizada y las salvedades.

## 7. Salidas
- Libro de mercado rural completo y verificado, sin errores de fórmula.
- Informe del estudio de mercado: texto, tabla del anexo 2.1, memoria de cálculo, justificación del % de negociación y soportes.
- Pedir la aprobación del perito para pasar a la fase 5.
