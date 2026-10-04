# Reglas del estudio de mercado – Lotes en Bello (Niquía / Bellavista)

## Objetivo
Reunir ofertas de **lotes** para valor del suelo, cerca del inmueble objeto
(lat 6.348283, lon -75.552109; Niquía / Bellavista, Bello, Antioquia).

## Criterios de la oferta (todos obligatorios)
- Lote de ~100 m² (aceptable 80–120 m²). Sin segundos pisos ni superiores.
- Régimen NPH: nada en propiedad horizontal, unidad cerrada ni con administración.
- Radio máximo ~2 km del objeto. Anotar barrio y coordenadas.
- Oferta vigente. Registrar fecha de publicación y fecha de captura.
- **Enlace directo al anuncio** (no la página de listado del portal).

## Reparto de trabajo
- **Codex:** buscar candidatos y llenar `ofertas/lotes_candidatos.csv`. Solo datos
  que vea en el anuncio. Si un dato no aparece, dejar la celda vacía; nunca inventar.
  Guardar en `evidencia_texto` el texto literal del anuncio (precio, área, ubicación).
- **Claude:** verificar cada fila (consistencia, duplicados, área, piso, distancia,
  enlace directo), calcular $/m² e ingresar las válidas en la hoja `MERCADO NPH`
  (filas 23 en adelante, solo celdas de entrada, sin alterar fórmulas).

## Reglas de Git
- Codex trabaja en la rama `codex/lotes-bello`; no toca `claude/keen-hopper-yo97py`.
- No modificar el archivo .xlsx. Solo escribir en `ofertas/`.
