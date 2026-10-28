---
name: avaluo-rural-fase4-mercado
description: Fase 4 del avalúo rural. Estudio de mercado - búsqueda de ofertas comparables, depuración, descuento del 10 % a las ofertas, homogeneización y estadística con coeficiente de variación máximo de la Resolución 941 de 2026 (parámetro del perito 10 %), usando el libro de mercado rural del perito. Úsala después de la fase 3.
---

# Fase 4 · Estudio de mercado rural

Los parámetros salen de `config/parametros-941.yaml`: descuento de ofertas (10 %) y CV máximo (10 %), ambos instrucción del perito y por confirmar contra la 941. No se usa el 7,5 % de la 620 ni el de las hojas urbanas (NPH).

## 1. Definir el universo comparable
1. Caracterizar el predio objeto: municipio y vereda, área, uso actual y permitido, topografía, aptitud (UPRA), acceso, disponibilidad de agua, infraestructura, afectaciones y la hipótesis de **mejor uso** de la fase 2.
2. Criterios de comparabilidad (se escriben antes de buscar): misma zona geoeconómica o vecina con condiciones similares, rango de área (por ejemplo 0,5x a 2x), uso, accesibilidad y fecha reciente (la 941 puede fijar ventana de tiempo; revisar).
3. Radio de búsqueda y ampliación escalonada si faltan datos, declarando por qué.

## 2. Búsqueda de ofertas
Fuentes, en este orden y con registro de cada una (enlace o contacto, fecha de captura y captura de pantalla):
- Portales inmobiliarios (Fincaraíz, Metrocuadrado, Ciencuadras, MercadoLibre, avisos de prensa regional).
- Redes de inmobiliarias y lonjas de la zona; corredores, notarios y personas de la vereda (ofertas verbales, con nombre y teléfono).
- Ofertas halladas en la visita de campo (ficha de campo, sección F).
- Transacciones reales cuando se puedan documentar (escrituras públicas, registros de la SNR), que pesan más que las ofertas.
Herramientas de apoyo: `brightdata-plugin:search` y `scrape` para recolectar anuncios, o navegador automatizado; todo dato recolectado se **confirma con el anuncio original**.

## 3. Ficha por oferta
No. de oferta, fuente, fecha, ubicación (vereda y coordenadas), área de terreno (ha), área y descripción de construcciones, uso, acceso, agua, topografía, condiciones de venta, teléfono/contacto, valor ofertado, enlace y captura.
Las construcciones de la oferta se valoran con el método de depreciación del perito y su valor se resta al valor negociado para obtener el **valor del terreno** (como hace la hoja de mercado del perito).

## 4. Depuración
- Descartar ofertas inconsistentes (área o precio dudosos, duplicadas, sin ubicación verificable, atípicas por condiciones jurídicas o de uso). Cada exclusión se registra con su razón en la columna ¿INCLUIR?.
- No se excluye para «hacer cumplir» el CV: se excluye por razones técnicas documentadas.

## 5. Descuento y homogeneización
1. **Descuento de negociación**: valor negociado = valor ofertado × (1 − 0,10). El parámetro viene de la configuración; en la hoja es una celda de DATOS referenciada, no un número escrito en cada fórmula.
2. **Valor del terreno** = valor negociado − valor de construcciones depreciadas (si aplica).
3. **$/ha** = valor del terreno ÷ área en hectáreas.
4. **Homogeneización** del $/ha frente al predio objeto con factores justificados (ubicación/acceso, área o tamaño, topografía, agua, aptitud, forma, infraestructura, fecha). Cada factor tiene criterio, rango y fuente; sin factores arbitrarios. El producto de los factores se aplica a cada oferta.

## 6. Estadística
Usar el libro de mercado rural del perito (`plantillas/ESTUDIO_MERCADO_RURAL.xlsx`) con fórmulas vivas, según el patrón de la skill `mercado-nph-estadisticos-depreciacion` (columna de construcciones vinculada a DEPRECIACIÓN, ¿INCLUIR?, alertas, bloque de estadísticos). Cambios para lo rural:
- Unidad: **$/ha** (no $/m²).
- Celda de DATOS del CV máximo: **0,10** (no 0,075). Celda de descuento: **0,10**.
- Mismos estadísticos: n, promedio, mediana, desviación estándar muestral, CV, límites (promedio ± desviación), asimetría, ¿CV ≤ máximo?, promedio redondeado.

Verificación independiente: correr `python3 scripts/estadisticos_mercado.py ofertas.csv` con los mismos datos y confirmar que coincide con la hoja (si no coincide, hay un error en una de las dos).

Si el CV supera el máximo: revisar depuración y homogeneización, buscar más ofertas, o justificar (con la 941) un procedimiento alternativo. No se «ajustan» valores para cuadrar.

## 7. Resultado y sustentación del valor
- Valor adoptado de terreno por hectárea: promedio (o mediana) justificado, redondeado según el criterio del perito y la norma, y dentro de los límites.
- Si el predio tiene zonas de distinto valor (planas/ladera, con agua/sin agua, con aptitud/sin ella), se establece **zonificación de valor** con sus áreas calculadas en EPSG:9377, y se aplican factores de zona.
- Análisis de sensibilidad breve (rango de valor si se usan promedio, mediana o límites).
- Anexos: fichas de ofertas, capturas, mapa de ubicación de las ofertas (`kmz-sig-profesional`).

## 8. Consideraciones que suelen pasarse por alto (revisar siempre)
- Fecha y vigencia de las ofertas y tendencia del mercado.
- Condiciones de pago o de venta atípicas (apuro, parentesco, permuta).
- Diferencia entre área oferta y área registral/catastral de la oferta.
- Valor de aguas, concesiones, vías y energía.
- Cultivos permanentes y plantaciones (se valoran por separado; no se mezclan con terreno).
- Servidumbres y condicionantes ambientales de la oferta.
- Efecto de proyectos de infraestructura cercanos.
- Mercados ilíquidos: pocos datos; documentar la búsqueda realizada.

## Salidas
- Libro de mercado rural completo y verificado (sin errores de fórmula).
- Informe del estudio de mercado (texto + cuadros + soportes).
- Pedir la aprobación del perito para pasar a la fase 5.
