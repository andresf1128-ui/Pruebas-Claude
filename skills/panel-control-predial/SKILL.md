---
name: panel-control-predial
description: Monta un panel local de control predial (o documental por expedientes) que se actualiza solo desde una carpeta de OneDrive/SharePoint sincronizada, la matriz maestra en Excel y un export de NotebookLM. Usar cuando el usuario quiera replicar el «Control Predial 360» de Hidroarma en otro proyecto, controlar formatos por municipio y gestión, o pasar un Excel de seguimiento a una herramienta que se mantenga sola.
---

# Panel de control predial

Motor reutilizable: lee la carpeta del proyecto, la matriz y los borradores de formatos, recalcula todo con reglas
declaradas en un `proyecto.json` y sirve un panel local (sin enviar datos a ningún servicio).
Referencia viva: `proyectos/hidroarma.json`. El código está en `motor/` y no contiene nada propio de un proyecto.

## Cuándo usarlo
- Hay una carpeta por expediente/predio con nombre `radicado + matrícula` y un Excel maestro con un renglón por predio.
- Se necesita ver avance, faltantes y alertas por municipio y gestión, y controlar formatos diligenciados.

## Pasos para un proyecto nuevo
1. **Entrevista corta** (preguntar solo lo que no se pueda leer de los archivos): ruta de la carpeta sincronizada; Excel maestro y
   su hoja y encabezados; cómo se escribe la matrícula y el radicado en los nombres de carpeta; municipios y gestiones;
   documentos que debe tener cada expediente; formatos a controlar y dónde están los borradores diligenciados;
   qué es «completo» para cada formato.
2. **Copiar** `proyectos/_plantilla.json` (o `hidroarma.json`) a `proyectos/<nombre>.json` y ajustar: `matriz`, `patron_matricula`,
   `patron_radicado`, `rotulos_lado`, `gestiones`, `municipios`, `clases` (nombre de archivo → tipo de documento), `documentos`,
   `alertas`, `formatos`.
3. **Configurar** `motor/config.json` a partir de `motor/config.example.json` (rutas locales del usuario).
4. **Verificar antes de entregar** (criterio 1 de `ACEPTACION.md`): `python verificar.py` debe terminar sin fallas.
   Si falla, no se entrega el panel: se corrige la regla o se explica la diferencia.
5. **Iniciar** con `iniciar.bat` (Windows) o `python server.py`. Abre `http://127.0.0.1:8765`.
6. **Aceptar** con el usuario recorriendo `ACEPTACION.md`. Solo con los cuatro criterios firmados se da por replicable.

## Reglas que no se negocian
- No inventar datos: lo que no se pueda leer queda como «Sin dato», con el motivo.
- Una sola fuente de cálculo (`sync.py`); nada de cifras escritas a mano en la interfaz.
- Las fotos o archivos compartidos entre predios colindantes son evidencia compartida: se cuentan por folio y también
  como distintos; solo se alerta si están repetidos dentro de la misma carpeta.
- Un «completo» de formato se prueba por contenido (secciones requeridas, sin texto de ejemplo, menciona la matrícula)
  y se puede confirmar o corregir a mano; la marca manual nunca borra el motivo automático.
- Si falla una lectura, el panel conserva el último dato bueno y lo avisa.

## Límites conocidos
- NotebookLM no tiene API pública: se lee de un export (`cuadernos_json`) que hay que renovar.
- La lectura directa de SharePoint (Microsoft Graph) requiere que un administrador registre una app; el motor actual
  usa la carpeta sincronizada, que no lo requiere.
- El panel lee la carpeta de un equipo; para varios usuarios hay que publicar `web/` y `data.json` desde un equipo central.
