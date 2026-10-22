# Control Predial Hidroarma · panel local que se actualiza solo

Lee la carpeta «1 Proyecto Hidroarma» sincronizada con OneDrive, la matriz maestra y (opcional) el export de NotebookLM.
Cada 60 s revisa si algo cambió y recalcula todo: archivos y fotos por folio, documentos presentes, estado del cotejo,
alertas y controles de calidad. No envía datos a ningún servicio: corre solo en tu equipo (127.0.0.1).

## Puesta en marcha (Windows)
1. Instala Python 3.10+ (python.org) si no lo tienes.
2. En OneDrive, deja la carpeta **1 Proyecto Hidroarma** siempre disponible en este equipo
   (clic derecho → «Mantener siempre en este dispositivo»), para que los archivos estén sincronizados sin descargarse a demanda.
3. Copia `config.example.json` a `config.json` y completa las rutas:
   - `carpeta_proyecto`: la carpeta sincronizada.
   - `matriz_xlsx`: el Excel con la hoja «Matriz maestra» (encabezados en la fila 5).
   - `cuadernos_json`: export de NotebookLM (opcional; sin él no se calcula el estado del cotejo).
4. Doble clic en `iniciar.bat`. Se abre el panel en http://127.0.0.1:8765

## Qué se actualiza solo
Carpetas y archivos nuevos, movidos o eliminados; fotos (incluidas las compartidas entre predios colindantes);
documentos presentes (CTL, escritura, fichas, plano, avalúo, soporte judicial); carpetas vacías, mal escritas,
con el lado cruzado o sin folio; cambios frente a la lectura anterior (pestaña «Calidad de datos»).

## Qué no
- **NotebookLM** no tiene API pública. Se lee del archivo `cuadernos_json`; actualízalo cuando cambien los cuadernos.
- **Valores, titulares y áreas** vienen de la matriz maestra; si cambia el Anexo, actualiza ese Excel.

## Reglas que puedes editar (`sync.py`)
- Prioridad de alertas (bloque `al.append([...])`): hoy falta CTL = Alta, falta escritura = Media, sin KMZ = Baja.
- Un folio sin carpeta propia hereda los documentos de su folio matriz.
- Una carpeta compartida por varios folios (p. ej. 20264371) cuenta sus archivos y fotos para cada uno.
- Una foto repetida en varias carpetas es evidencia compartida; solo se alerta si está repetida dentro de la misma carpeta.

## Verificación hecha
Contra la captura de SharePoint del 21-oct-2026 (1.719 archivos, 45 cuadernos): coinciden los archivos y fotos por folio,
las banderas de documentos, el estado del cotejo (35 conciliados · 10 por conciliar · 1 falta cuaderno · 41 sin cuaderno · 6 derivados)
y la clasificación de los 1.719 archivos. Probado además agregando archivos a la carpeta: el panel los detecta en segundos.
