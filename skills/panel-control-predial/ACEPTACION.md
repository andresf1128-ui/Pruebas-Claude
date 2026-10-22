# Aceptación del panel

El panel no se da por bueno ni se replica a otro proyecto hasta cumplir estos cuatro criterios.
Anotar fecha y quién confirma en cada uno.

| # | Criterio | Cómo se comprueba | Estado |
|---|---|---|---|
| 1 | **Las cifras coinciden con la realidad.** | En el equipo real, con la carpeta sincronizada: `python verificar.py` sin fallas y 5 cifras del panel contrastadas a mano con SharePoint (archivos, un folio al azar, fotos, estado del cotejo, un formato). | ☐ |
| 2 | **Las reglas de alerta son las del usuario.** | Revisar `alertas` del `proyecto.json` con el usuario: qué es Alta, Media o Baja. En Hidroarma hoy son reglas de ejemplo (falta CTL = Alta, falta escritura = Media, sin KMZ = Baja). | ☐ |
| 3 | **El formato 1.1 está resuelto.** | Formato = «1.1 Formato análisis normativo rural». Carpetas de borradores configuradas en `formatos_carpetas`. «Completo» = secciones presentes, sin texto de ejemplo y menciona la matrícula (ajustable en `verificacion`). Probar con 3 borradores reales: uno completo, uno incompleto y uno sin matrícula en el nombre. | ☐ |
| 4 | **Uso real de varios días.** | Dejarlo corriendo contra el Drive real de 3 a 5 días hábiles. Anotar sorpresas y corregirlas; terminar sin discrepancias nuevas. | ☐ |

## Para replicar a otro proyecto
Repetir 1 a 4 con el `proyecto.json` nuevo. El criterio 2 y la definición de «completo» del 3 son siempre del dueño del proyecto.
