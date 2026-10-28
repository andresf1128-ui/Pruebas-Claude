---
name: avaluo-rural-fase2-norma
description: Fase 2 del avalúo rural. Análisis normativo extremadamente riguroso del predio y de la Resolución 941 de 2026, con consulta en vivo de fuentes oficiales (POT/PBOT/EOT, autoridad ambiental, UPRA/SIPRA, IGAC, DANE, ANT, ANM, entre otras) y salidas gráficas. Úsala después de la fase 1.
---

# Fase 2 · Análisis de la norma

La norma es lo más sensible del avalúo: rigor absoluto, nada de suposiciones. Apoyarse en `sustentacion-fuentes-normativas-avaluo` (sustento predio por predio) y `salidas-graficas-norma-predios` (mapas y formato), y en `areas-oficiales-magna-sirgas` para toda área.

## Reglas no negociables
1. **Solo fuentes oficiales**, consultadas en vivo el día del análisis, con enlace y fecha de consulta. Sin consulta en vivo no se escribe «verificado»; se escribe «no re-verificado hoy» (y eso no va al cliente: va a la lista de pendientes del perito).
2. Un resumen de una herramienta web puede fallar: los textos normativos clave se cotejan con el documento original (capa de texto y comprobación visual).
3. Separar siempre: **fuente abierta** (enlace público) vs. **insumo suministrado** (cartografía, GDB, CTL del cliente).
4. Si una cita se aplica fuera de su contexto (otra subzona, otro municipio, derogada o modificada), se dice y se busca la correcta. Verificar la **vigencia** de cada norma (derogatorias y modificaciones).
5. Áreas y porcentajes de cruces sobre el polígono exacto, en EPSG:9377.
6. Entregable sin notas de «por verificar» ni «pendiente». Lo que no esté publicado se explica con la norma de jerarquía aplicable (por ejemplo, determinante ambiental de superior jerarquía, art. 10 de la Ley 388 de 1997, a confirmar en el texto).

## A. Resolución IGAC 941 de 2026 (lectura previa obligatoria)
Antes de trabajar el predio, leer el texto oficial completo y su anexo técnico (normograma del IGAC o Diario Oficial) y llenar `config/parametros-941.yaml` con artículo y numeral. Extraer, como mínimo, lo que la 941 diga sobre:
- Objeto, alcance, definiciones, vigencia de la norma y régimen de transición para encargos en curso.
- Exigencias al avaluador (inscripción RAA, responsabilidad, independencia) y a la solicitud/encargo.
- **Visita técnica** (el art. 10 sería el aplicable, según fuentes secundarias: confirmar).
- Contenido mínimo del informe y de sus anexos; información que debe declarar.
- Métodos para predios rurales (comparación de mercado, capitalización de renta, costo de reposición para construcciones, residual cuando aplique) y condiciones de uso de cada uno.
- Tratamiento de las ofertas: depuración, factor de negociación, homogeneización, **coeficiente de variación máximo** y número mínimo de datos.
- Depreciación de construcciones y vida útil.
- Cálculo de valores de terreno, construcciones, cultivos, plantaciones y anexos.
- Vigencia del avalúo.
- **Uso de inteligencia artificial** y de herramientas automatizadas: qué permite, qué exige declarar y qué prohíbe (las fuentes secundarias reportan un «uso controlado»; leer el artículo y aplicarlo).
- Uso actual y cobertura (el art. 46 se cita en `copernicus-sentinel-avaluos`: confirmar el contenido).
- Casos especiales (servidumbres, expropiación, procesos con el Estado, tierras con restricciones).

Resultado: una **tabla de requisitos de la 941** (artículo → obligación → dónde se cumple en el encargo). Esa tabla alimenta la lista de control de la fase 6.

## B. Matriz de fuentes oficiales por tema
| Tema | Fuente oficial a consultar | Qué se extrae |
|---|---|---|
| Ordenamiento territorial | POT/PBOT/EOT del municipio (acuerdo y cartografía) | Clasificación del suelo (rural), categorías de protección, áreas de producción, usos principales/compatibles/prohibidos, normas de parcelación, densidades, unidades de planificación |
| Determinantes ambientales | Corporación Autónoma Regional (CAR/CORNARE, etc.), Ministerio de Ambiente, SIAC | POMCA, zonificación ambiental, rondas hídricas, áreas de conservación, reservas de la Ley 2ª de 1959, DMI, distritos de manejo |
| Áreas protegidas | RUNAP – Parques Nacionales Naturales | Traslape con áreas protegidas y regímenes |
| Aptitud y frontera agrícola | **UPRA – SIPRA** | Frontera agrícola, aptitud por cadena productiva, zonificaciones, UAF de referencia, ordenamiento productivo |
| Suelos y cartografía | **IGAC** (Geoportal, Colombia en Mapas, estudios agrológicos, cartografía básica, catastro) | Clasificación agrológica, capacidad de uso, límites, vías, hidrografía, información catastral |
| Estadísticas y contexto | **DANE** (censos, agropecuario, divipola, estratificación, proyecciones) | Contexto socioeconómico y productivo del municipio |
| Tierras | Agencia Nacional de Tierras (ANT), URT, comunidades étnicas | Baldíos, resguardos, consejos comunitarios, procesos de restitución, zonas de reserva campesina |
| Minería e hidrocarburos | ANM (Anna Minería), ANH, ANLA | Títulos, solicitudes, áreas estratégicas, licencias |
| Amenazas y riesgos | Servicio Geológico Colombiano, IDEAM, UNGRD, mapas municipales de gestión del riesgo | Movimientos en masa, inundación, sismicidad, incendios |
| Infraestructura | INVÍAS, ANI, ANLA, UPME, operadores de red | Vías proyectadas, líneas de transmisión, gasoductos y oleoductos, franjas de servidumbre |
| Registro y catastro | VUR/SNR, gestor catastral | Se atiende en la fase 3 |

Cada consulta deja: entidad, capa o documento, enlace, fecha y hora de consulta, resultado (incluido «sin traslape») y captura. Usar `copernicus-sentinel-avaluos` para uso actual y cobertura y para comprobar el estado reciente del predio.

## C. Análisis del predio
1. Clasificación del suelo y categoría rural. Usos del suelo según el POT: principal, compatible, condicionado, prohibido, con citas del artículo del acuerdo.
2. Restricciones y determinantes ambientales con **área y porcentaje** del predio afectado por cada una; consecuencia sobre el aprovechamiento.
3. Frontera agrícola y aptitud (UPRA) frente al uso actual y al uso permitido.
4. Normas de subdivisión y UAF aplicables (Ley 160 de 1994 y desarrollos; confirmar vigencia) y su efecto en la divisibilidad y en el valor.
5. Afectaciones y servidumbres de infraestructura y su cuantificación.
6. Amenazas y riesgos y su incidencia.
7. **Mejor uso (highest and best use)** consistente con la norma y el mercado, para fijar la hipótesis de valoración.
8. Estratificación y servicios públicos cuando apliquen.

## D. Salidas
- Texto del análisis normativo por predio (formato del encargo) con citas.
- Salidas gráficas normativas (mapas) con la fuente bajo cada imagen y áreas en EPSG:9377.
- Tabla de requisitos de la 941 cumplidos.
- Lista de pendientes para el perito (a un lado de lo que va al cliente).
- Pedir la aprobación del perito para pasar a la fase 3.
