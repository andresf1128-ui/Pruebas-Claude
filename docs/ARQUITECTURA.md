# Agente de avalúos rurales · arquitectura y puesta en marcha

## Qué contiene este repositorio
```
.claude/agents/avaluador-rural.md        orquestador (reglas, fases, higiene de entregables)
.claude/skills/avaluo-rural-fase1-campo/         captura en campo, fotos, construcciones
.claude/skills/avaluo-rural-fase2-norma/         Resolución 941 + fuentes oficiales
.claude/skills/avaluo-rural-fase3-juridico/      revisión documental y conciliación
.claude/skills/avaluo-rural-fase4-mercado/       ofertas, descuento 10 %, CV
.claude/skills/avaluo-rural-fase5-liquidacion-informe/
.claude/skills/avaluo-rural-fase6-calidad/       inspección de fondo y forma
config/parametros-941.yaml               parámetros con estado confirmado / por_confirmar
config/estilo.yaml                       estilo único de los entregables
scripts/organizar_fotos.py               copia, ordena y rotula fotos; manifiesto con EXIF/GPS
scripts/estadisticos_mercado.py          negociación, $/ha y estadística con CV (sin homogeneización)
referencias/RESUMEN_941.md               tabla de requisitos de la 941 por artículo
referencias/Res_941_2026_texto_extraido_parcial.txt   texto OCR de la resolución (hasta la pág. 24 del anexo)
scripts/limpiar_metadatos.py             limpia y verifica rastros en docx/xlsx/pptx/pdf
```

## Skills que ya tiene el perito y que el agente reutiliza
`areas-oficiales-magna-sirgas`, `kmz-sig-profesional`, `copernicus-sentinel-avaluos`, `salidas-graficas-norma-predios`, `sustentacion-fuentes-normativas-avaluo`, `nota-ctl-hidroarma`, `generalidades-sector-avaluo`, `organizar-fotos-avaluos`, `mercado-nph-estadisticos-depreciacion` (patrón de hojas), `control-calidad-documentos`, y las de formato (`docx`, `pdf`, `xlsx`, `pptx`).
Esas skills están en su cuenta de Claude, no en este repositorio; las skills de fase las nombran y las invocan.

## Estado frente a la Resolución 941
Ya está en `referencias/` el texto extraído (parcial) y la tabla de requisitos por artículo (`RESUMEN_941.md`). Lo confirmado, con su artículo, está en `config/parametros-941.yaml`:
- CV máximo rural 10 % (art. 21), visita técnica obligatoria (arts. 10 y 13.4), plazo de 30 días hábiles (art. 9), depreciación Ross-Heidecke continua (art. 30), vigencia de la norma desde su publicación (art. 60).
- Lo que cambió respecto de lo supuesto al inicio: **no se homogeneiza por factores** (anexo 2.1, Nota 2); el **porcentaje de negociación debe justificarse** (art. 17 b), el 10 % es criterio del perito; la IA solo se regula en el art. 20, parágrafo 2.

## Qué falta para dejarlo al 100 %
1. **Páginas faltantes del anexo técnico de la 941** (el texto extraído se corta en la página 24 del anexo): tabla de estados de conservación, coeficientes y ejemplos de Ross-Heidecke (2.3.3 y 2.3.4), y modelos de renta agropecuaria (tablas 9 a 17). Sin ellas no se calcula la depreciación ni se aplica renta agropecuaria. También quedan sin leer los arts. 22 a 26, 31 a 45 y 50 a 56.
2. **Justificación escrita del 10 % de negociación**, para sustentarla en el informe (art. 17 b).
3. **Vigencia del avalúo:** no aparece en los artículos leídos de la 941; verificar en el Decreto 1170 de 2015.
4. **Plantillas del perito** en `plantillas/`: formato de descripción de construcciones, libro de mercado rural, informe y liquidación.
5. **Catálogo de tipologías constructivas del IGAC** en la versión que use el perito.
6. **Estilo** (tipo de letra, tamaños, márgenes) en `config/estilo.yaml`.
7. **Herramienta de campo elegida** (QField, Kobo, Avenza, Field Maps) para construir el formulario con los campos de `.claude/skills/avaluo-rural-fase1-campo/referencias/ficha-campo.md`.

## Usar el agente
En Claude Code: invocar el subagente `avaluador-rural` y darle el radicado, los documentos y la carpeta de fotos. Trabaja una fase a la vez y espera aprobación del perito entre fases.

## Nota sobre la inteligencia artificial en los entregables
El agente limpia los entregables de marcas y metadatos de herramientas. La 941 solo regula la IA cuando se usa un modelo automatizado como herramienta complementaria de valoración (art. 20, parágrafo 2), con manual auditable, métricas de error y validación del avaluador. Aquí ningún modelo determina el valor; los cálculos son scripts y fórmulas visibles. El informe es del perito y él responde por él.
