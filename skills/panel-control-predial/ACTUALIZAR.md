# Cómo actualizar el tablero (para el usuario)

Tablero: https://claude.ai/artifact/E5hiEtfxisEBbRVi4FZyyg

Cada vez que quieras ponerlo al día:

1. En tu PC, con la carpeta «1 Proyecto Hidroarma» sincronizada, abre una terminal en `motor/` y corre:
   `python exportar_inventario.py "C:/.../1 Proyecto Hidroarma" --borradores "C:/.../BORRADORES_3.1_SONSON..." --borradores "C:/.../AGUADAS..."`
   (Sin `--borradores` si aún no hay formatos diligenciados.) Genera `inventario.json` y `borradores.json`. No copia el contenido de tus archivos
   de la carpeta: solo nombres y tamaños. `borradores.json` sí lleva el texto de los formatos .docx.
2. Exporta de NotebookLM el listado de cuadernos y fuentes (`cuadernos.json`) si cambió algo ahí.
3. Sube al chat: `inventario.json`, `borradores.json`, el Excel madre y `cuadernos.json`, y escribe «actualiza».
4. Claude cruza todo, actualiza el tablero y te dice qué cambió.

El tablero muestra siempre la fecha del último corte. Lo que no subas no se toca.
