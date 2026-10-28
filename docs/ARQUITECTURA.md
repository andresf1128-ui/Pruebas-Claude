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
scripts/estadisticos_mercado.py          descuento, $/ha y estadística con CV
scripts/limpiar_metadatos.py             limpia y verifica rastros en docx/xlsx/pptx/pdf
```

## Skills que ya tiene el perito y que el agente reutiliza
`areas-oficiales-magna-sirgas`, `kmz-sig-profesional`, `copernicus-sentinel-avaluos`, `salidas-graficas-norma-predios`, `sustentacion-fuentes-normativas-avaluo`, `nota-ctl-hidroarma`, `generalidades-sector-avaluo`, `organizar-fotos-avaluos`, `mercado-nph-estadisticos-depreciacion` (patrón de hojas), `control-calidad-documentos`, y las de formato (`docx`, `pdf`, `xlsx`, `pptx`).
Esas skills están en su cuenta de Claude, no en este repositorio; las skills de fase las nombran y las invocan.

## Qué falta para dejarlo al 100 % (el agente no puede inventarlo)
1. **Texto oficial de la Resolución 941 de 2026** y su anexo técnico. Desde esta sesión no se pudo abrir igac.gov.co (fuera de los dominios permitidos de la red). Cuando se habilite el dominio, o se aporte el PDF, hay que llenar `config/parametros-941.yaml` con artículo y numeral y marcar `confirmado`. Hasta entonces el CV de 10 % y el descuento de 10 % constan como instrucción del perito.
2. **Plantillas del perito**: formato de descripción de construcciones, libro de mercado rural (ESTUDIO_MERCADO_RURAL), formato de informe e hoja de liquidación, en `plantillas/`.
3. **Catálogo de tipologías constructivas del IGAC** en la versión que use el perito, en `.claude/skills/avaluo-rural-fase1-campo/referencias/`.
4. **Estilo** (tipo de letra, tamaños, márgenes) en `config/estilo.yaml`.
5. **Herramienta de campo elegida** (QField, Kobo, Avenza, Field Maps) para construir el formulario con los campos de `referencias/ficha-campo.md`.

## Usar el agente
En Claude Code: invocar el subagente `avaluador-rural` y darle el radicado, los documentos y la carpeta de fotos. Trabaja una fase a la vez y espera aprobación del perito entre fases.

## Nota sobre la inteligencia artificial en los entregables
El agente limpia los entregables de marcas y metadatos de herramientas, y exige en la fase 2 leer qué dice la 941 sobre el uso de IA para cumplirlo. Si la resolución obliga a declarar su uso, se declara; el informe es del perito y él responde por él.
