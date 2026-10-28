---
name: avaluo-rural-fase5-liquidacion-informe
description: Fase 5 del avalúo rural. Liquidación del valor (terreno, construcciones con depreciación, cultivos y anexos), elaboración del informe conforme a la Resolución 941 de 2026 y preparación de la sustentación al cliente con las consideraciones más importantes del predio. Úsala después de la fase 4.
---

# Fase 5 · Liquidación, informe y sustentación

## 1. Liquidación
Todo cálculo sale de la hoja de liquidación (fórmulas vivas) o de scripts, nunca de cabeza.

1. **Terreno**: por zona de valor, área (ha, EPSG:9377) × valor $/ha adoptado en la fase 4. La suma de áreas de las zonas debe igualar el área adoptada del predio (diferencia 0).
2. **Construcciones**: por cada una, área × costo de reposición a nuevo × (1 − depreciación). Depreciación con el **modelo continuo Ross-Heidecke** (art. 30), que combina edad y estado de conservación; los modelos discontinuos o en escalera quedan excluidos. Edad, vida útil y estado de conservación vienen del formato de la fase 1. Vida útil de referencia según la Tabla 2 del anexo (temporales 5, corta 30, media 50, larga 70, permanentes 100 años) y **vida útil prolongada** solo cuando se superó la vida de referencia (o desde el 90 %), con calificaciones de Heidecke entre 2,5 y 4,5 y sustento. Los coeficientes y las ecuaciones salen del anexo técnico (2.3.3 y 2.3.4): ese texto falta en `referencias/`, así que la depreciación no se calcula hasta tenerlo. El costo a nuevo es de reposición o reproducción, con costos directos e indirectos y presupuestos que consideren el lugar (arts. 27 y 28). Ver `mercado-nph-estadisticos-depreciacion` para el patrón de la hoja DEPRECIACIÓN.
3. **Cultivos y plantaciones**: método adecuado a su especie y etapa (costo de establecimiento y mantenimiento depreciado, o renta capitalizada), con rendimiento y precio de fuentes oficiales (DANE, UPRA, Agronet/SIPSA, Fedecafé, etc.); nunca se suma doble con el terreno.
4. **Anexos y mejoras**: cercas, vías internas, pozos, riego, galpones, tanques; valor de reposición depreciado.
5. **Servidumbres, afectaciones o limitaciones**: ajuste o indemnización cuando la 941 lo ordene y el caso lo requiera, con su método.
6. **Valor total** = terreno + construcciones + cultivos + anexos; redondeo solo al final y según el criterio de la 941 y del perito.
7. Controles: suma de partes = total; valor/ha total frente al promedio de mercado (razonabilidad); proporciones (terreno vs. construcciones); comparación con el avalúo catastral y con avalúos previos, explicando las diferencias.

## 2. Informe
El contenido mínimo lo fija el **art. 14 de la 941** (13 numerales; detalle en `referencias/RESUMEN_941.md`, sección 3). Orden de referencia, que cubre los 13 numerales y el formato del encargo:
1. Carátula, índice y datos del solicitante y del avaluador (RAA).
2. Información general: solicitante, marco jurídico del encargo, tipo de inmueble y de avalúo (comercial rural), departamento, municipio y vereda, nombre del predio, fecha(s) de visita técnica y fecha del informe. Número Predial Nacional (NPN) cuando exista.
3. Identificación del predio: ubicación con **coordenadas del centroide tomadas en sitio y su sistema de referencia**, linderos, matrícula, ficha o NPN, área, distancia en km desde la cabecera municipal al acceso con tipo y estado de la vía.
4. Documentos consultados y componente jurídico (fase 3), con las advertencias.
5. Análisis normativo (fase 2), con salidas gráficas y fuentes.
6. Descripción del sector y del predio (generalidades del sector, usos, vías, servicios, clima, suelos).
7. Descripción de construcciones, cultivos y anexos (con tipología del IGAC).
8. Estudio de mercado (fase 4) y valor del terreno.
9. Metodología de valoración y liquidación.
10. Resultado del avalúo (cuadro final).
11. Consideraciones, limitaciones y supuestos.
12. Declaraciones del avaluador (independencia, ausencia de conflicto de intereses, responsabilidad).
13. Firma y fecha del avaluador.
14. Anexos: solicitud, registro fotográfico, mapas, documentos con su fecha de expedición, ofertas (URL, captura de pantalla y fecha de consulta; las de campo con foto, descripción, áreas y fecha) y el **certificado RAA vigente con las categorías requeridas**.

Además: **memoria de cálculo** con todas las ecuaciones, variables, índices, tasas, datos y resultados (art. 14.10 c); en un inmueble no sometido a PH la liquidación se presenta con **valores unitarios de terreno y construcción** (art. 13.9 b); las áreas a valorar son las indicadas por el solicitante y la responsabilidad sobre ellas es suya (arts. 12.6 y 13.6); y el informe debe permitir a un tercero concluir el valor sin acudir a fuentes externas (art. 5.4).

Formato: estilo único (`config/estilo.yaml`); imágenes a 14 cm, borde biselado y fuente bajo cada una; cifras en formato es-CO; tablas con las mismas columnas y alineación en todo el documento. Apoyarse en `docx` o `pdf` para generar el archivo y en `generalidades-sector-avaluo` para las generalidades del sector.

Redacción: texto técnico, claro y sin adornos, en primera persona del perito donde corresponda. No incluir notas de trabajo ni marcas de herramientas. Antes de entregar correr `scripts/limpiar_metadatos.py`.

## 3. Sustentación al cliente
Documento corto (1 a 3 páginas, o presentación) para explicar el avalúo, con las consideraciones más importantes del predio:
- Resultado y cómo se llegó a él en tres pasos (norma, mercado, liquidación).
- Los 5 a 8 factores que más pesan en el valor (acceso, agua, restricciones ambientales, aptitud/uso permitido, jurídica, estado de construcciones).
- Qué restringe o limita el aprovechamiento y su efecto cuantificado.
- Qué se hizo para cumplir la 941 (visita, métodos, CV, descuento).
- Advertencias del componente jurídico y qué le conviene hacer al cliente (por ejemplo corregir un área, cancelar un gravamen, formalizar una servidumbre).
- Preguntas probables y respuestas preparadas, con sus soportes a la mano.
Para una presentación, usar `pptx`, `canva` o `Gamma`, con el mismo estilo del informe.

## 4. Salidas
- Hoja de liquidación verificada.
- Informe (.docx y .pdf) y anexos.
- Documento o presentación de sustentación.
- Pedir la aprobación del perito para pasar al control de calidad final (fase 6).
