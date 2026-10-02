# Estudio de mercado rural – Puente Bomba / El Ebanal (Tigreras)

Avaluador: Andrés Felipe Cuartas Montoya · RAA 1.128.391.782

## Presentación automática (10 diapositivas)

1. Guarde el Excel con los valores del mercado (en Excel, para que se guarden los resultados de las fórmulas).
2. Reemplace `Estudio_Mercado_Rural_Puente_Bomba_Tigreras.xlsx` en este repositorio y haga commit/push.
3. GitHub Actions ejecuta `generar_presentacion.py` y actualiza `Presentacion_Estudio_Mercado_Rural.pptx`.

Manual: `pip install -r requirements.txt && python generar_presentacion.py`

Diapositivas: portada · resumen · predio y método · ofertas · oferta comparable · ofertas descartadas ·
valor/ha y estadísticos · depreciación Ross-Heidecke · liquidación · observaciones y firma. Las combinaciones y la proyección son de análisis interno y no se presentan.
Colores tomados del libro (azul marino 1F3864, azul 2F5597, grises, verde claro E2EFDA, amarillo FFF2CC).

## Usar como modelo para nuevos estudios

Reemplace las ofertas y construcciones en el Excel (sin cambiar los nombres de las hojas ni su estructura) y suba el archivo:
la presentación se regenera sola con los nuevos datos. Acepta cualquier número de ofertas y de construcciones.
El avaluador (nombre y RAA) está definido al inicio de `generar_presentacion.py`.
